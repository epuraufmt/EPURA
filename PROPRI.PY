import os
import sys
import time
import pandas as pd
import numpy as np
from multiprocessing import Pool, Manager
from functools import partial
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException,
    ElementClickInterceptedException, ElementNotInteractableException, StaleElementReferenceException
)

# --- CONFIGURAÇÕES ---

# **ATENÇÃO**: Verifique e ajuste estes caminhos!
CHROMEDRIVER_PATH = r"C:\Users\CPU 01\Desktop\chromedriver-win64\chromedriver.exe"
INPUT_XLSX_PATH = "PEDRA90.xlsx"
OUTPUT_XLSX_PATH = "PEDRA90OK.xlsx"

PORTAL_URL_INICIAL = "https://portalfazenda.cuiaba.mt.gov.br/portalfazenda/PortalContribuinte/FiltrarGuias"

ID_MODAL_FILTRAR_GUIAS = "mdlFiltrarGuias"
ID_CAMPO_INSCRICAO_IMOBILIARIA = "ContentPageBody_txtInscricao"
ID_BOTAO_AVANCAR = "ContentPageBody_lkbAvancar"

SELETOR_NOME_PROPRIETARIO = "span#ContentPageBody_lblContribuinte"
SELETOR_CPF_CNPJ = "span#ContentPageBody_lblCpfCnpj"
SELETOR_TOTAL_VALOR_SITE = "span#ContentPageBody_lblTotal"

SELETOR_BOTAO_OK_SWEETALERT = "button.swal2-confirm.swal2-styled"
SELETOR_CONTEUDO_SWEETALERT = "div#swal2-content"
TEXTO_ERRO_CADASTRO_NAO_ENCONTRADO = "Nenhum cadastro em atividade encontrado no munícipio."
TEXTO_NENHUM_DADO_APRESENTADO = "Não existem dados para serem apresentados"

WAIT_TIMEOUT = 30 # Tempo máximo de espera para elementos na página

# Nomes das colunas no seu arquivo Excel.
# **IMPORTANTE**: Certifique-se de que 'inscricao' corresponda exatamente ao cabeçalho da sua coluna de inscrição no Excel.
INSCRICAO_FIELD_NAME = "inscricao"
NOME_PROPRIETARIO_FIELD_NAME = "PROPIETARIO"
CPF_FIELD_NAME = "CPF"
PROPI_CLASSIFICATION_FIELD_NAME = "PROPI" # Nome da coluna para a classificação do proprietário

# Novos campos para os dados de IPTU
IPTU_S_OU_N_FIELD_NAME = "IPTU_S_OU_N"
IPTU_QNTS_ANOS_FIELD_NAME = "IPTU_QNTS_ANOS"
REAIS_FIELD_NAME = "REAIS"

MUNICIPIO_CUIABA = "MUNICIPIO DE CUIABA"
DE_MATO_GROSSO = "DE MATO GROSSO"

NUM_PROCESSOS = 20 # Número de navegadores Chrome (processos) a serem executados em paralelo

# --- Funções Auxiliares ---

def parse_currency_to_float(currency_str):
    """Converte uma string de moeda (R$X.XXX,XX) para float."""
    if not isinstance(currency_str, str):
        return 0.0
    # Extrai o valor numérico da string como 'Qtd.: 48 | Total: R$117.435,16'
    match = re.search(r'R\$([\d\.,]+)', currency_str) # Regex mais flexível para R$X,XX ou R$X.XXX,XX
    if match:
        cleaned_str = match.group(1).replace(".", "").replace(",", ".").strip()
        try:
            return float(cleaned_str)
        except ValueError:
            return 0.0
    return 0.0

def format_float_to_currency(float_val):
    """Formata um float para string de moeda (R$X.XXX,XX)."""
    # Garante 2 casas decimais e formatação de milhar e decimal corretas para português
    return f"R${float_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def click_sweetalert_ok_button(driver, process_id):
    """
    Tenta clicar no botão 'OK' de um aviso SweetAlert.
    Retorna True se o botão foi clicado, False caso contrário.
    """
    try:
        ok_button = WebDriverWait(driver, 0.5, poll_frequency=0.05).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, SELETOR_BOTAO_OK_SWEETALERT))
        )
        driver.execute_script("arguments[0].click();", ok_button)
        time.sleep(1)
        return True
    except TimeoutException:
        return False
    except Exception as e:
        return False

def get_sweetalert_content(driver, process_id):
    """
    Tenta obter o texto do conteúdo de um SweetAlert visível.
    Retorna o texto se encontrado, None caso contrário.
    """
    try:
        content_element = WebDriverWait(driver, 1).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, SELETOR_CONTEUDO_SWEETALERT))
        )
        return content_element.text.strip()
    except TimeoutException:
        return None
    except Exception as e:
        return None

def get_proprietario_from_site_selenium(driver, inscricao_completa, process_id):
    """Consulta o site para obter o nome do proprietário da inscrição, CPF/CNPJ, classificar PROPI e dados de IPTU."""
    proprietario_nome_real = None
    cpf_cnpj_info = ""
    propi_classification = None

    iptu_s_ou_n = "N" # Default para 'N'
    iptu_qnts_anos = 0
    reais_total_calculado = 0.0 # Este será o valor final do IPTU vencido puro.

    try:
        driver.get(f"{PORTAL_URL_INICIAL}?modulo=imobiliario")
        time.sleep(3)

        click_sweetalert_ok_button(driver, process_id)

        modal = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.visibility_of_element_located((By.ID, ID_MODAL_FILTRAR_GUIAS))
        )

        campo_inscricao = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.element_to_be_clickable((By.ID, ID_CAMPO_INSCRICAO_IMOBILIARIA))
        )

        campo_inscricao.click()
        campo_inscricao.clear()
        campo_inscricao.send_keys(inscricao_completa)

        driver.execute_script("arguments[0].dispatchEvent(new Event('input'))", campo_inscricao)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", campo_inscricao)
        time.sleep(1)

        botao_avancar = driver.find_element(By.ID, ID_BOTAO_AVANCAR)
        driver.execute_script("arguments[0].click();", botao_avancar)

        sweetalert_text = get_sweetalert_content(driver, process_id)

        if sweetalert_text and TEXTO_ERRO_CADASTRO_NAO_ENCONTRADO in sweetalert_text:
            print(f"[Processo {process_id}] ERRO: Cadastro não encontrado para {inscricao_completa}.")
            click_sweetalert_ok_button(driver, process_id)
            return inscricao_completa, "ERRO", "", None, iptu_s_ou_n, iptu_qnts_anos, reais_total_calculado
        elif sweetalert_text:
            # Apenas clica no OK sem logar, se não for o erro de cadastro não encontrado
            click_sweetalert_ok_button(driver, process_id)

        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.url_contains("PortalContribuinte/EmitirGuias")
        )
        time.sleep(2)

        # Etapa: Extrair nome real do proprietário
        nome_proprietario_element = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, SELETOR_NOME_PROPRIETARIO))
        )
        proprietario_nome_real = nome_proprietario_element.text.strip()

        # Etapa: Extrair CPF/CNPJ
        try:
            cpf_cnpj_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, SELETOR_CPF_CNPJ))
            )
            cpf_cnpj_info = cpf_cnpj_element.text.strip()
        except TimeoutException:
            cpf_cnpj_info = ""

        # Etapa: Classificar o campo PROPI (a classificação)
        proprietario_upper = proprietario_nome_real.upper()
        if ("IGREJA" in proprietario_upper or "INSTITUTO" in proprietario_upper) and \
           DE_MATO_GROSSO.upper() not in proprietario_upper:
            propi_classification = "PJPRIV"
        elif proprietario_nome_real and (MUNICIPIO_CUIABA.upper() in proprietario_nome_real.upper() or DE_MATO_GROSSO.upper() in proprietario_upper):
            propi_classification = "PJPUBLICA"
        elif cpf_cnpj_info:
            if len(cpf_cnpj_info) == 14: # CPF tem 14 caracteres (incluindo pontos e traço)
                propi_classification = "PF"
            elif len(cpf_cnpj_info) == 18: # CNPJ tem 18 caracteres (incluindo pontos, barras e traço)
                propi_classification = "PJPRIV"
            else:
                propi_classification = None
        else:
            propi_classification = None

        # --- Lógica de IPTU ---
        # Verifica "Não existem dados para serem apresentados"
        try:
            no_data_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f"//td[contains(@class, 'dataTables_empty') and contains(text(), '{TEXTO_NENHUM_DADO_APRESENTADO}')]"))
            )
            if no_data_element:
                return inscricao_completa, proprietario_nome_real, cpf_cnpj_info, propi_classification, \
                       iptu_s_ou_n, iptu_qnts_anos, reais_total_calculado
        except TimeoutException:
            pass # Continua para a lógica de débitos, pois a mensagem não foi encontrada

        # Se não há "Não existem dados", então há débitos a processar
        iptu_s_ou_n = "S"
        unique_red_years = set() # Usado para contar anos únicos com débitos vencidos
        
        # --- NÃO VAMOS PEGAR O TOTAL DO SITE DIRETAMENTE AQUI ---
        # Vamos construir o total que queremos (IPTU vencido) iterando as linhas.
        # reais_total_calculado permanece 0.0 no início e só soma o que realmente é IPTU vencido.

        # Loop de paginação para verificar débitos
        while True:
            try:
                WebDriverWait(driver, WAIT_TIMEOUT).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table#tableDebitos tbody tr"))
                )
            except TimeoutException:
                print(f"[Processo {process_id}] Aviso: Tabela de débitos não encontrada para {inscricao_completa}. Encerrando verificação de IPTU.")
                break

            rows = driver.find_elements(By.CSS_SELECTOR, "table#tableDebitos tbody tr")

            # --- Termos flexíveis para exclusão (linhas que não contam para IPTU puro) ---
            termos_auto_infracao = ["auto infração", "auto de infração", "infraçao", "multa"]
            termos_coleta_lixo = ["coleta de lixo", "taxa de coleta", "lixo"]
            termos_credito = ["crédito", "credito"]

            for row in rows:
                try:
                    row_text = row.text.lower()
                    
                    # Tenta obter o valor da linha. A penúltima TD parece ser o valor no HTML de exemplo.
                    valor_float = 0.0
                    try:
                        all_tds = row.find_elements(By.TAG_NAME, "td")
                        if len(all_tds) >= 2: 
                            valor_str = all_tds[-2].text.strip() # Penúltima TD
                            valor_float = parse_currency_to_float(valor_str)
                            if valor_float == 0.0: # Se o parse_currency_to_float falhou ou retornou 0, tenta a antepenúltima
                                valor_str = all_tds[-3].text.strip() # Ant-penúltima TD (pode ser o valor "real" vs. "corrigido")
                                valor_float = parse_currency_to_float(valor_str)
                        # print(f"DEBUG: Linha: '{row_text}', Valor encontrado: {valor_float}") # Para depuração

                    except Exception as e:
                        print(f"[Processo {process_id}] ERRO ao obter valor de linha para {inscricao_completa}: {e}")
                        valor_float = 0.0 # Garante que o valor seja 0 em caso de erro

                    is_auto_infracao = any(termo in row_text for termo in termos_auto_infracao)
                    is_coleta_lixo = any(termo in row_text for termo in termos_coleta_lixo)
                    is_credito = any(termo in row_text for termo in termos_credito)
                    
                    is_red_color = "color: red" in row.get_attribute("style")

                    # Lógica de contabilização para IPTU puro vencido:
                    # Só somamos se NÃO for lixo/auto/crédito E ESTIVER em vermelho.
                    if not is_auto_infracao and not is_coleta_lixo and not is_credito:
                        if is_red_color:
                            reais_total_calculado += valor_float # SOMA apenas IPTU vencido
                            try:
                                year_element = row.find_element(By.ID, "numAno")
                                year = year_element.text.strip()
                                if year:
                                    unique_red_years.add(year)
                            except NoSuchElementException:
                                pass
                            except Exception as e:
                                print(f"[Processo {process_id}] ERRO ao obter ano em linha vermelha para {inscricao_completa}: {e}")
                                pass
                    # Se for lixo/auto/crédito, ou se for azul, simplesmente IGNORAMOS essa linha para o *nosso* total.
                    # Não deduzimos mais do total do site, estamos construindo o nosso total.

                except StaleElementReferenceException:
                    print(f"[Processo {process_id}] Alerta: StaleElementReferenceException ao processar linha para {inscricao_completa}. Tentando próxima página.")
                    break # Quebra o loop interno para tentar avançar a página
                except Exception as row_e:
                    print(f"[Processo {process_id}] ERRO geral ao processar linha da tabela para {inscricao_completa}: {row_e}")
                    continue

            # Verifica e clica no botão "Próximo" para paginar
            try:
                next_button = driver.find_element(By.ID, "tableDebitos_next")
                if "disabled" in next_button.get_attribute("class"):
                    break # Sem mais páginas
                else:
                    driver.execute_script("arguments[0].click();", next_button)
                    time.sleep(2)
                    WebDriverWait(driver, WAIT_TIMEOUT).until(
                        EC.staleness_of(rows[0]) if rows else EC.presence_of_element_located((By.CSS_SELECTOR, "table#tableDebitos tbody tr"))
                    )
            except NoSuchElementException:
                break # Botão "Próximo" não encontrado, sem mais páginas
            except Exception as e:
                print(f"[Processo {process_id}] ERRO ao manipular botão 'Próximo' para {inscricao_completa}: {e}. Encerrando paginação.")
                break

        iptu_qnts_anos = len(unique_red_years)
        
        # Garante que o valor total não seja negativo (embora não deva ser com essa lógica)
        if reais_total_calculado < 0:
            reais_total_calculado = 0.0

        # Lógica final para IPTU_S_OU_N: se não sobrou IPTU vencido e o valor é 0, marca como 'N'
        if reais_total_calculado > 0.0 or iptu_qnts_anos > 0:
            iptu_s_ou_n = "S"
        else:
            iptu_s_ou_n = "N"


        return inscricao_completa, proprietario_nome_real, cpf_cnpj_info, propi_classification, \
               iptu_s_ou_n, iptu_qnts_anos, reais_total_calculado

    except (TimeoutException, NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException) as e:
        print(f"[Processo {process_id}] ERRO de interação na página para {inscricao_completa}: {type(e).__name__}. URL: {driver.current_url}.")
        return inscricao_completa, "NÃO ENCONTRADO", "", None, "N", 0, 0.0
    except WebDriverException as e:
        print(f"[Processo {process_id}] ERRO de WebDriver para inscrição {inscricao_completa}: {e}.")
        return inscricao_completa, "NÃO ENCONTRADO", "", None, "N", 0, 0.0
    except Exception as e:
        print(f"[Processo {process_id}] ERRO INESPERADO para inscrição {inscricao_completa}: {e}.")
        return inscricao_completa, "NÃO ENCONTRADO", "", None, "N", 0, 0.0

def worker_function(records_chunk, process_id):
    """Função worker para processamento paralelo de um chunk de registros."""
    processed_chunk_results = []

    service = webdriver.chrome.service.Service(CHROMEDRIVER_PATH)
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") # Executa Chrome em modo invisível
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3") # Define o nível de log do ChromeDriver para apenas FATAL erros

    driver = webdriver.Chrome(service=service, options=options)

    try:
        for record in records_chunk:
            inscricao = record[INSCRICAO_FIELD_NAME]

            # Normaliza a inscrição: remove sufixos '#XXX' e adiciona '001'
            inscricao_str = str(inscricao).strip()
            if '#' in inscricao_str:
                base_inscricao = inscricao_str.split('#')[0]
                inscricao_completa = f"{base_inscricao}001"
            else:
                inscricao_completa = f"{inscricao_str}001"

            # Busca os dados no site
            inscricao_ret, proprietario_nome_real, cpf_cnpj_info, propi_classification, \
            iptu_s_ou_n, iptu_qnts_anos, reais_total_float = \
                get_proprietario_from_site_selenium(driver, inscricao_completa, process_id)

            # Prepara o dicionário de resultados
            result_dict = {
                INSCRICAO_FIELD_NAME: inscricao,
                NOME_PROPRIETARIO_FIELD_NAME: proprietario_nome_real,
                CPF_FIELD_NAME: cpf_cnpj_info,
                PROPI_CLASSIFICATION_FIELD_NAME: propi_classification,
                IPTU_S_OU_N_FIELD_NAME: iptu_s_ou_n,
                IPTU_QNTS_ANOS_FIELD_NAME: iptu_qnts_anos,
                REAIS_FIELD_NAME: reais_total_float
            }

            processed_chunk_results.append(result_dict)

            time.sleep(0.5) # Pequeno atraso para evitar sobrecarga do servidor

    except Exception as e:
        print(f"[Processo {process_id}] ERRO CRÍTICO no worker: {e}")
    finally:
        if driver:
            try:
                driver.quit() # Garante que o navegador seja fechado
            except Exception as close_e:
                print(f"[Processo {process_id}] ERRO ao fechar navegador Chrome: {close_e}")

    return processed_chunk_results

# --- Função Auxiliar para Salvar XLSX Grande (em Múltiplas Abas) ---
def save_large_df_to_excel(df, output_path, chunk_size=1048500):
    """
    Salva um DataFrame grande em múltiplas abas de um arquivo Excel.
    Cada aba terá no máximo 'chunk_size' linhas (ajustado para um pouco abaixo do limite para segurança).
    """
    num_chunks = int(np.ceil(len(df) / chunk_size))

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for i in range(num_chunks):
            start_row = i * chunk_size
            end_row = min((i + 1) * chunk_size, len(df))
            chunk_df = df.iloc[start_row:end_row]

            sheet_name = f'Parte_{i+1}'
            print(f"Salvando {len(chunk_df)} linhas na aba '{sheet_name}'...")
            chunk_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"DataFrame salvo em '{output_path}' em {num_chunks} abas.")


# --- Função Principal ---

def main():
    global INPUT_XLSX_PATH

    print("--- Iniciando processamento de XLSX com paralelismo ---")

    # Se o caminho de entrada não foi definido, pede ao usuário
    if not INPUT_XLSX_PATH:
        input_path = input("Por favor, digite o caminho completo do arquivo XLSX de entrada (ex: C:\\caminho\\arquivo.xlsx): ")
        if not os.path.exists(input_path):
            print(f"ERRO: Arquivo XLSX não encontrado no caminho: '{input_path}'. Verifique o caminho digitado.")
            sys.exit(1)
        INPUT_XLSX_PATH = input_path

    try:
        df_full = pd.read_excel(INPUT_XLSX_PATH)
        print(f"Arquivo XLSX '{INPUT_XLSX_PATH}' lido com sucesso. Total de linhas: {len(df_full)}")
    except FileNotFoundError:
        print(f"ERRO: Arquivo XLSX de entrada não encontrado em '{INPUT_XLSX_PATH}'. Verifique o caminho.")
        sys.exit(1)
    except Exception as e:
        print(f"ERRO ao ler o XLSX de entrada: {e}")
        sys.exit(1)

    # Verifica se a coluna de inscrição existe
    if INSCRICAO_FIELD_NAME not in df_full.columns:
        print(f"ERRO: O campo '{INSCRICAO_FIELD_NAME}' (nome da coluna de inscrição) não foi encontrado nas colunas do XLSX. Abortando.")
        print(f"Colunas encontradas no arquivo Excel: {df_full.columns.tolist()}")
        sys.exit(1)

    # Garante que as colunas necessárias existam no DataFrame, inicializando-as se ausentes
    required_cols_with_defaults = {
        NOME_PROPRIETARIO_FIELD_NAME: "",
        CPF_FIELD_NAME: "",
        PROPI_CLASSIFICATION_FIELD_NAME: "",
        IPTU_S_OU_N_FIELD_NAME: "N",
        IPTU_QNTS_ANOS_FIELD_NAME: 0,
        REAIS_FIELD_NAME: 0.0
    }

    for col, default_val in required_cols_with_defaults.items():
        if col not in df_full.columns:
            print(f"Aviso: Campo '{col}' não encontrado no XLSX. Criando coluna e preenchendo com valores padrão.")
            df_full[col] = default_val

    # Lógica para identificar registros que precisam ser processados (inscrições únicas)
    # Processa se PROPI não estiver classificado, ou se IPTU for 'N', ou anos for 0, ou reais for 0.0
    df_unique_to_process = df_full[
        ~(df_full[PROPI_CLASSIFICATION_FIELD_NAME].astype(str).str.upper().isin(["PF", "PJPRIV", "PJPUBLICA"])) |
        (df_full[IPTU_S_OU_N_FIELD_NAME].astype(str).str.upper() == "N") |
        (df_full[IPTU_QNTS_ANOS_FIELD_NAME] == 0) |
        (df_full[REAIS_FIELD_NAME] == 0.0)
    ].copy()

    # Pega apenas as inscrições únicas desse subset para processamento
    df_unique_to_process = df_unique_to_process.drop_duplicates(subset=[INSCRICAO_FIELD_NAME])

    total_records_full = len(df_full)
    total_records_to_process_unique = len(df_unique_to_process)

    if total_records_to_process_unique == 0:
        print(f"\n--- Não há registros únicos para processar (todos os dados já foram classificados ou preenchidos). ---")
        print(f"Salvando o arquivo original (sem alterações) em: '{OUTPUT_XLSX_PATH}'")
        df_full.to_excel(OUTPUT_XLSX_PATH, index=False)
        sys.exit(0)

    print(f"Total de registros no arquivo de entrada: {total_records_full}")
    print(f"Total de **inscrições únicas** para processar: {total_records_to_process_unique}")

    records_list_to_process_unique = df_unique_to_process.to_dict(orient='records')

    # Divide os registros em "chunks" para os processos paralelos
    chunks = []
    if total_records_to_process_unique > 0:
        chunk_size = (total_records_to_process_unique + NUM_PROCESSOS - 1) // NUM_PROCESSOS
        for i in range(0, total_records_to_process_unique, chunk_size):
            chunks.append(records_list_to_process_unique[i:i + chunk_size])
    else:
        chunks = []

    print(f"Dividindo {total_records_to_process_unique} registros únicos em {len(chunks)} chunks para {NUM_PROCESSOS} processos.")

    processed_results_from_workers = []

    # Executa os workers em paralelo
    with Pool(NUM_PROCESSOS) as pool:
        results = pool.starmap(worker_function, [(chunk, i + 1) for i, chunk in enumerate(chunks)])

        for processed_chunk_data in results:
            processed_results_from_workers.extend(processed_chunk_data)

    df_processed_updates = pd.DataFrame(processed_results_from_workers)

    # Remover duplicatas dos resultados processados antes de mesclar (camada de segurança)
    duplicadas_processadas = df_processed_updates[df_processed_updates.duplicated(subset=[INSCRICAO_FIELD_NAME])]
    if not duplicadas_processadas.empty:
        print(f"ALERTA: Foram encontradas {len(duplicadas_processadas)} inscrições duplicadas nos resultados processados. Removendo...")
        df_processed_updates.drop_duplicates(subset=[INSCRICAO_FIELD_NAME], keep='last', inplace=True)
        print(f"Após remover duplicatas dos resultados processados: {len(df_processed_updates)} linhas únicas para merge.")

    # Lógica de atualização: mescla os dados processados de volta ao DataFrame completo
    if not df_processed_updates.empty:
        cols_to_update = [
            NOME_PROPRIETARIO_FIELD_NAME,
            CPF_FIELD_NAME,
            PROPI_CLASSIFICATION_FIELD_NAME,
            IPTU_S_OU_N_FIELD_NAME,
            IPTU_QNTS_ANOS_FIELD_NAME,
            REAIS_FIELD_NAME
        ]

        df_results_for_merge = df_processed_updates[[INSCRICAO_FIELD_NAME] + cols_to_update]

        # Realiza o merge para adicionar as colunas atualizadas com sufixo '_new'
        df_final = pd.merge(df_full,
                            df_results_for_merge,
                            on=INSCRICAO_FIELD_NAME,
                            how='left',
                            suffixes=('', '_new'))

        # Itera sobre as colunas e substitui os valores originais pelos novos (se existirem)
        for col_name in cols_to_update:
            new_col_name = f"{col_name}_new"
            if new_col_name in df_final.columns:
                df_final[col_name] = np.where(df_final[new_col_name].notna(),
                                                df_final[new_col_name],
                                                df_final[col_name])
                df_final = df_final.drop(columns=[new_col_name]) # Remove a coluna temporária

        # Limpeza final para garantir que nenhuma coluna '_new' residual permaneça
        df_final = df_final.loc[:, ~df_final.columns.str.contains('_new$', case=False)]
    else:
        df_final = df_full.copy() # Se não houve atualizações, usa o DataFrame original sem alterações

    # Contagem final de registros processados/classificados
    classified_count_total = sum(1 for _, row in df_final.iterrows()
                              if pd.notna(row[PROPI_CLASSIFICATION_FIELD_NAME]) and
                              str(row[PROPI_CLASSIFICATION_FIELD_NAME]).strip().upper() in ["PF", "PJPRIV", "PJPUBLICA"])

    iptu_processed_count = sum(1 for _, row in df_final.iterrows()
                               if str(row[IPTU_S_OU_N_FIELD_NAME]).strip().upper() == "S" or \
                                  (str(row[IPTU_S_OU_N_FIELD_NAME]).strip().upper() == "N" and \
                                   row[IPTU_QNTS_ANOS_FIELD_NAME] == 0 and \
                                   row[REAIS_FIELD_NAME] == 0.0)) # Considera como verificado mesmo se 'N' e 0

    try:
        # Tenta salvar o arquivo XLSX
        df_final.to_excel(OUTPUT_XLSX_PATH, index=False)
        print(f"\n--- XLSX ATUALIZADO SALVO EM: '{OUTPUT_XLSX_PATH}' ---")
    except Exception as e:
        # Em caso de erro (ex: This sheet is too large!), tenta salvar em múltiplas abas
        print(f"ERRO ao salvar o XLSX de saída: {e}. Tentando salvar em múltiplas abas.")
        try:
            save_large_df_to_excel(df_final, OUTPUT_XLSX_PATH)
            print(f"AVISO: O arquivo foi salvo em múltiplas abas em '{OUTPUT_XLSX_PATH}' devido ao seu tamanho.")
        except Exception as e_multi_sheet:
            print(f"ERRO CRÍTICO: Falha ao salvar mesmo em múltiplas abas: {e_multi_sheet}")
            print(f"Considere salvar como CSV: df_final.to_csv('{OUTPUT_XLSX_PATH.replace('.xlsx', '.csv')}', index=False)")


    print("\n--- PROCESSAMENTO DE DADOS CONCLUÍDO ---")
    print(f"{classified_count_total} registros com classificação ('PF', 'PJPRIV', 'PJPUBLICA') preenchida com sucesso.")
    print(f"{iptu_processed_count} registros de IPTU verificados (S ou N).")

if __name__ == "__main__":
    main()
