from config import(get_env_from_file, update_env_file)
from dotenv import load_dotenv
import wx
import json
import webview
import atexit
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_autoinstaller
import os
import time
import sys
import traceback
import threading
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
caminho_arquivo_selecionado = None
load_dotenv()
def disable_asserts():
    wx.DisableAsserts()
    
atexit.register(disable_asserts)

import logging
import traceback
processo_handler = logging.FileHandler('processo.log')
processo_handler.setLevel(logging.INFO)
processo_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
processo_handler.setFormatter(processo_formatter)

# Configuração do logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
logger.addHandler(processo_handler)

def Log(message, level='info', exc_info=False):
    if level == 'info':
        logger.info(message, exc_info=exc_info)
    elif level == 'warning':
        logger.warning(message, exc_info=exc_info)
    elif level == 'error':
        logger.error(message, exc_info=exc_info)
    else:
        logger.info(message, exc_info=exc_info)


logging.basicConfig(
    filename='dev.log',
    level=logging.ERROR,  # Captura apenas mensagens de erro e acima
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = log_exception

def logar(login, senha, tipo, head=True):
    driver = None
    try:
        env_vars_to_clear = [
            'RelatorioPrestador',
            'RelatorioTomador',
            'DeclaracaoPrestador',
            'DeclaracaoTomador',
            'PortalContribuinte',
            'GuiaIssPrestador',
            'GuiaIssTomador',
            'NfsTomador',
            'NfsPrestador',
            'Ver'
        ]

        for var in env_vars_to_clear:
            if var in os.environ:
                del os.environ[var]

        # Recarrega o .env
        load_dotenv()
            
        local = os.getenv(tipo)
        if local is None:
            Log(f'Variável de ambiente {tipo} não encontrada.', level='error')
            return False

        # Configura as opções do Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            'download.default_directory': local,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })

        if head:
            if os.getenv('Ver') == '0':
                chrome_options.add_argument('--headless=new')  # Habilita o modo headless

        # Configura o driver do Chrome
        chromedriver_autoinstaller.install()
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_window_size(1800, 1200)
        if not head:
            driver.minimize_window()
            
        url = 'https://sispmjp.joaopessoa.pb.gov.br:8080/sispmjp/login.jsf'
        max_tentativas = 3
        tentativas = 0
        driver.get(url)

        while tentativas < max_tentativas:
            driver.find_element(By.ID, 'j_username').send_keys(login)
            driver.find_element(By.ID, 'j_password').send_keys(senha, Keys.RETURN)
            time.sleep(1)  # Aguarde alguns segundos para a página carregar
            current_url = driver.current_url

            if current_url == url:
                Log("Problema ao logar: Tentativa de login falhou", level='info')
                tentativas += 1
                if tentativas >= max_tentativas:
                    Log("Máximo de tentativas atingido. Login falhou.", level='error')
                    return False
                driver.get(url)
            elif current_url == 'https://sispmjp.joaopessoa.pb.gov.br:8080/sispmjp/paginas/mensagens.jsf':
                driver.find_element(By.ID, 'formMensagens:commandButton_confirmar').click()
                try:
                    WebDriverWait(driver, 10).until(EC.url_to_be('https://sispmjp.joaopessoa.pb.gov.br:8080/sispmjp/paginas/index.jsf'))
                except Exception as e:
                    Log("Falha ao redirecionar para a página inicial.", level='error')
                    return False
                return driver
            elif current_url == 'https://sispmjp.joaopessoa.pb.gov.br:8080/sispmjp/paginas/index.jsf':
                Log('Login bem-sucedido', level='info')
                return driver
            else:
                Log("URL inesperada. Redirecionando para a página inicial.", level='info')
                driver.get('https://sispmjp.joaopessoa.pb.gov.br:8080/sispmjp/paginas/index.jsf')
                return driver
    except Exception as e:
        Log(f'Exceção capturada', level='error')
        return False
    finally:
        if driver:
            driver.quit()      
class MyJSAPI:
    def __init__(self):
        self.dialog_open = False
        wx.Log.SetActiveTarget(None)
    
    def get_logs(self):
        if os.path.exists('./processo.log'):
            with open('./processo.log', 'r') as file:
                return file.read()
        else:
            return 'Arquivo de log não encontrado.'

    def iniciar_threads_e_aguardar(self, linha):
        threads = []
        
        def mesM():
            hoje = datetime.now()
            primeiro_dia_mes_atual = hoje.replace(day=1)
            ultimo_dia_mes_passado = primeiro_dia_mes_atual - timedelta(days=1)
            primeiro_dia_mes_passado = ultimo_dia_mes_passado.replace(day=1)
            primeiro_dia_mes_passado_str = primeiro_dia_mes_passado.strftime('%d/%m/%Y')
            ultimo_dia_mes_passado_str = ultimo_dia_mes_passado.strftime('%d/%m/%Y')
            datas = {
                'inicio': primeiro_dia_mes_passado_str,
                'fim': ultimo_dia_mes_passado_str
            }

            return datas

        def numeromesM():
            hoje = datetime.now()
            primeiro_dia_mes_atual = hoje.replace(day=1)
            primeiro_dia_ultimo_mes = primeiro_dia_mes_atual - timedelta(days=1)
            primeiro_dia_ultimo_mes = primeiro_dia_ultimo_mes.replace(day=1)
            mes = primeiro_dia_ultimo_mes.month
            ano = primeiro_dia_ultimo_mes.year
            resultado = {
                'mes': mes,
                'ano': ano
            }

            return resultado

        
        if linha['RELATORIO NFS-E PRESTADOR'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.RelatorioPrestador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['DECL SERV PRESTADOR'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.DeclaracaoPrestador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['GUIA ISS PRESTADOS'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.GuiaIssPrestador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['XML NFS-E PRESTADOR'].strip().upper() == 'SIM':
            mes = mesM()
            thread = threading.Thread(target=self.NfsPrestador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['inicio'], mes['fim']))
            thread.start()
            threads.append(thread)
        
        if linha['RELATORIO NFSE TOMADOR'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.RelatorioTomador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['DECL SERV TOMADOR'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.DeclaracaoTomador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['GUIA ISS TOMADOR'].strip().upper() == 'SIM':
            mes = numeromesM()
            thread = threading.Thread(target=self.GuiaIssTomador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['mes'], mes['ano']))
            thread.start()
            threads.append(thread)
        
        if linha['XML NFS-E TOMADOR'].strip().upper() == 'SIM':
            mes = mesM()
            thread = threading.Thread(target=self.NfsTomador, args=( str(linha['LOGIN']).strip(), str(linha['SENHA']).strip(), mes['inicio'], mes['fim']))
            thread.start()
            threads.append(thread)
        
        if linha['CERTIDÃO NEGATIVA DE DEBITOS '].strip().upper() == 'SIM':
            login = str(linha['CNPJ']).strip()
            thread = threading.Thread(target=self.PortalContribuinte, args=(login,))
            thread.start()
            threads.append(thread)

        
        # Espera todas as threads terminarem antes de retornar
        for thread in threads:
            thread.join()

    def close(self):
        if webview.windows:
            webview.windows[0].destroy()

    def minimize(self):
        if webview.windows:
            webview.windows[0].minimize()

    def fullscreen(self):
        if webview.windows:
            webview.windows[0].toggle_fullscreen()
            
    def enviar_dados(self, dados):
        global caminho_arquivo_selecionado
        print('Dados recebidos:', dados)
        print('Caminho do arquivo:', caminho_arquivo_selecionado)
        
        # Ler o arquivo Excel
        try:
            df = pd.read_excel(caminho_arquivo_selecionado)
            for item in dados:
                indice = int(item['indice'])
                if indice < len(df):
                    linha = df.iloc[indice]
                    self.iniciar_threads_e_aguardar(linha)
                else:
                    print(f'Índice {indice} está fora do alcance do DataFrame.')
        
        except Exception as e:
            print('Erro ao ler o arquivo:', e)
            traceback.print_exc()
    
        return {"status": "sucesso"}
    
    def click(self,on):
        env_file_path = Path(__file__).parent / '.env'
        var_name = 'Ver'
        new_value = 1 if on else 0

        # Read the .env file
        with open(env_file_path, 'r') as file:
            lines = file.readlines()

        # Update the variable in the lines
        var_found = False
        for i, line in enumerate(lines):
            if line.startswith(var_name + '='):
                lines[i] = f'{var_name}={new_value}\n'
                var_found = True
                break

        # If the variable was not found, add it to the end
        if not var_found:
            lines.append(f'{var_name}={new_value}\n')

        # Write the lines back to the .env file
        with open(env_file_path, 'w') as file:
            file.writelines(lines)

    def open_folder_dialog(self, key):
        if self.dialog_open:
            return  # Exit if a dialog is already open

        def show_dialog():
            try:
                self.dialog_open = True
                app = wx.App(False)  # Create a new application instance
                dialog = wx.DirDialog(None, "Selecione uma pasta", style=wx.DD_DEFAULT_STYLE)
                if dialog.ShowModal() == wx.ID_OK:
                    folder_path = dialog.GetPath()
                    folder_path = folder_path.replace('\\', '\\\\')
                    js_code = f"window.selectedFolder('{key}', '{folder_path}');"
                    if webview.windows:
                        webview.windows[0].evaluate_js(js_code)
                dialog.Destroy()
            finally:
                self.dialog_open = False  # Reset the flag when the dialog is closed

        # Run the dialog on the main thread
        show_dialog()
        
    def confirmConfig(self, dados):
        update_env_file(dados)
        webview.windows[0].evaluate_js(f"window.confirmConfiguration();")
        
    def config(self):
        variaveis = get_env_from_file()
        webview.windows[0].evaluate_js(f"window.configuration({str(variaveis)});")
        
        
    def selecionar_arquivo_excel(self):
        global caminho_arquivo_selecionado

        app = wx.App(False)
        frame = wx.Frame(None, -1, 'Esconder')
        frame.SetSize(0, 0, 200, 50)
        frame.Show(False)
        dialog = wx.FileDialog(
            frame,
            message="Selecione o arquivo Excel",
            wildcard="Arquivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        )
        
        if dialog.ShowModal() == wx.ID_OK:
                caminho_arquivo_selecionado = dialog.GetPath()
                nome_arquivo = os.path.basename(caminho_arquivo_selecionado)
                
                # Lê o arquivo Excel e obtém os CNPJs
                df = pd.read_excel(caminho_arquivo_selecionado)
                cnpjs = df['CNPJ'].dropna().astype(str).tolist()  # Considera a coluna CNPJ
                
                # Lista para armazenar todos os índices
                todos_indices = list(range(len(cnpjs)))  # Todos os índices da lista de CNPJs
                
        else:
                nome_arquivo = None
                cnpjs = []
                todos_indices = []

        dialog.Destroy()
        wx.GetApp().GetTopWindow().Destroy()  # Fecha o frame
        return {"nome_arquivo": nome_arquivo, "cnpjs": cnpjs, 'posicao': todos_indices}

    def apagar_arquivo_excel(self):
        global caminho_arquivo_selecionado
        caminho_arquivo_selecionado = None
        return "Arquivo apagado"
        
    def RelatorioPrestador(self, login, senha, mes, ano):
        fim = False
        driver = logar(login=login, senha=senha, tipo='RelatorioPrestador')
        if not driver:
            Log("Falha ao iniciar o driver no método RelatorioPrestador.", level='error')
            return False

        try:
            wait = WebDriverWait(driver, 10)
            Log("Aguardando o elemento 'Declaração' no método RelatorioPrestador.", level='info')

            # Espera até que o elemento "Declaração" esteja localizado
            declaracao = wait.until(EC.presence_of_element_located((By.XPATH, "//a[span[text()='Declaração']]")))
            
            actions = webdriver.ActionChains(driver)
            actions.move_to_element(declaracao).perform()
            time.sleep(0.1)

            # Move to "Declaração" submenu item
            submenu_item = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[1]/a')
            actions.move_to_element(submenu_item).perform()
            time.sleep(0.1)

            # Click on the specific menu item
            final_item = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[1]/ul/li[1]/a')
            final_item.click()

            Log("Preenchendo os campos de mês e ano.", level='info')

            # Wait for the mes element to be located and send keys
            wait.until(EC.presence_of_element_located((By.ID, "formEntregarDeclaracaoPrestador:mes")))
            time.sleep(0.3)
            
            driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:mes').send_keys(str(mes))
            driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:ano').send_keys(ano, Keys.RETURN)
            time.sleep(0.3)

            try:
                # Check for error message
                driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:mes')
                Log("Erro: A data da competência deve ser igual ou posterior à do campo 'Data de Constituição'.", level='error')
                return {'erro': 'A data da competência deve ser igual ou posterior à do campo "Data de Constituição", informada pelo contribuinte na opção de menu "Dados Cadastrais -> Contribuinte -> Dados Básicos".'}
            except:
                pass

            Log("Gerando relatório...", level='info')
            # Click to proceed with the report generation
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt85"]/span[2]').click()
            time.sleep(0.5)
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt98:gerarRelatorio"]').click()
            fim = True
        except Exception as error:
            Log(f"Erro ao clicar no elemento 'Declaração'", level='error')
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Relatório gerado com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.RelatorioPrestador();")    
            else:
                Log("Falha na geração do relatório.", level='info')
                webview.windows[0].evaluate_js("window.RelatorioPrestador(false);")    
            driver.quit()

    def RelatorioTomador(self, login, senha, mes, ano):
        fim = False
        driver = logar(login=login, senha=senha, tipo='RelatorioTomador')
        if not driver:
            Log("Falha ao iniciar o driver no método RelatorioTomador.", level='error')
            return False

        try:
            wait = WebDriverWait(driver, 10)
            Log("Aguardando o elemento 'Declaração' no método RelatorioTomador.", level='info')

            # Espera até que o elemento "Declaração" esteja localizado
            declaracao = wait.until(EC.presence_of_element_located((By.XPATH, "//a[span[text()='Declaração']]")))

            actions = webdriver.ActionChains(driver)
            actions.move_to_element(declaracao).perform()
            time.sleep(0.1)

            # Move to "Declaração" submenu item
            submenu_item = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[2]/a/span[1]')
            actions.move_to_element(submenu_item).perform()
            time.sleep(0.1)

            # Click on the specific menu item
            final_item = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[2]/ul/li[1]/a/span')
            final_item.click()

            Log("Preenchendo os campos de mês e ano.", level='info')

            # Wait for the mes element to be located and send keys
            wait.until(EC.presence_of_element_located((By.ID, "formEntregarDeclaracaoTomador:mes")))
            driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:mes').send_keys(mes)
            driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:ano').send_keys(ano, Keys.RETURN)
            time.sleep(0.3)

            try:
                # Check for error message
                driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:mes')
                Log("Erro: A data da competência deve ser igual ou posterior à do campo 'Data de Constituição'.", level='error')
                return {'erro': 'A data da competência deve ser igual ou posterior à do campo "Data de Constituição", informada pelo contribuinte na opção de menu "Dados Cadastrais -> Contribuinte -> Dados Básicos".'}
            except:
                pass

            Log("Gerando relatório...", level='info')
            # Click to proceed with the report generation
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt86"]').click()
            time.sleep(0.3)
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt100:gerarRelatorio"]').click()
            fim = True
        except Exception as error:
            Log(f"Erro ao clicar no elemento 'Declaração'", level='error')
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Relatório gerado com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.RelatorioTomador();")    
            else:
                Log("Falha na geração do relatório.", level='info')
                webview.windows[0].evaluate_js("window.RelatorioTomador(false);")    
            driver.quit()
            
            
    def DeclaracaoPrestador(self, login, senha, mes, ano):
        fim = False
        driver = logar(login=login, senha=senha, tipo='DeclaracaoPrestador')
        if not driver:
            Log("Falha ao iniciar o driver no método DeclaracaoPrestador.", level='error')
            return False

        try:
            Log("Aguardando o elemento 'Declaração' no método DeclaracaoPrestador.", level='info')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[span[text()='Declaração']]")))

            actions = ActionChains(driver)
            actions.move_to_element(driver.find_element(By.XPATH, "//a[span[text()='Declaração']]")).perform()
            time.sleep(0.1)
            actions.move_to_element(driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[1]/a')).perform()
            time.sleep(0.1)
            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[1]/ul/li[1]/a').click()

            Log("Preenchendo os campos de mês e ano.", level='info')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formEntregarDeclaracaoPrestador:mes")))
            driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:mes').send_keys(mes)
            driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:ano').send_keys(ano, Keys.RETURN)
            time.sleep(0.3)

            try:
                # Check for error message
                driver.find_element(By.ID, 'formEntregarDeclaracaoPrestador:mes')
                Log("Erro: A data da competência deve ser igual ou posterior à do campo 'Data de Constituição'.", level='error')
                return {'erro': 'A data da competência deve ser igual ou posterior à do campo "Data de Constituição", informada pelo contribuinte na opção de menu "Dados Cadastrais -> Contribuinte -> Dados Básicos".'}
            except:
                pass

            Log("Gerando relatório...", level='info')
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt85"]').click()
            time.sleep(0.5)
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt98:gerarRelatorio"]').click()
            fim = True
        except Exception as e:
            Log(f"Erro ao clicar no elemento 'Declaração'", level='error')
           
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Relatório gerado com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.DeclaracaoPrestador();")
            else:
                Log("Falha na geração do relatório.", level='info')
                webview.windows[0].evaluate_js("window.DeclaracaoPrestador(false);")
            driver.quit()
            
    def DeclaracaoTomador(self, login, senha, mes, ano):
        fim = False
        driver = logar(login=login, senha=senha, tipo='DeclaracaoTomador')
        if not driver:
            Log("Falha ao iniciar o driver no método DeclaracaoTomador.", level='error')
            return False

        try:
            Log("Aguardando o elemento 'Declaração' no método DeclaracaoTomador.", level='info')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[span[text()='Declaração']]")))

            actions = ActionChains(driver)
            actions.move_to_element(driver.find_element(By.XPATH, "//a[span[text()='Declaração']]")).perform()
            time.sleep(0.1)
            actions.move_to_element(driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[2]/a/span[1]')).perform()
            time.sleep(0.1)
            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[1]/ul/li[2]/ul/li[1]/a').click()

            Log("Preenchendo os campos de mês e ano.", level='info')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formEntregarDeclaracaoTomador:mes")))
            driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:mes').send_keys(mes)
            driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:ano').send_keys(ano, Keys.RETURN)
            time.sleep(0.3)

            try:
                # Check for error message
                driver.find_element(By.ID, 'formEntregarDeclaracaoTomador:mes')
                Log("Erro: A data da competência deve ser igual ou posterior à do campo 'Data de Constituição'.", level='error')
                return {'erro': 'A data da competência deve ser igual ou posterior à do campo "Data de Constituição", informada pelo contribuinte na opção de menu "Dados Cadastrais -> Contribuinte -> Dados Básicos".'}
            except:
                pass

            Log("Gerando relatório...", level='info')
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt86"]').click()
            time.sleep(0.5)
            driver.find_element(By.XPATH, '//*[@id="form_lista_notas_declaracoes:j_idt100:gerarRelatorio"]').click()
            fim = True
        except Exception as e:
            Log(f"Erro ao clicar no elemento 'Declaração'", level='error')
           
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Relatório gerado com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.DeclaracaoTomador();")
            else:
                Log("Falha na geração do relatório.", level='info')
                webview.windows[0].evaluate_js("window.DeclaracaoTomador(false);")
            driver.quit()
                
    def PortalContribuinte(self, cnpj):
        fim = False
        driver = None
        try:
            # Limpa variáveis de ambiente específicas
            env_vars_to_clear = [
                'RelatorioPrestador', 
                'RelatorioTomador', 
                'DeclaracaoPrestador', 
                'DeclaracaoTomador',
                'PortalContribuinte',
                'GuiaIssPrestador', 
                'GuiaIssTomador',
                'NfsTomador',  
                'NfsPrestador', 
                'Ver'
            ]

            for var in env_vars_to_clear:
                if var in os.environ:
                    del os.environ[var]

            # Recarrega o .env
            load_dotenv()
                
            local = os.getenv('PortalContribuinte')
            if not local:
                Log("O diretório 'PortalContribuinte' não está definido nas variáveis de ambiente.", level='error')
                return False

            # Configura as opções do Chrome
            chrome_options = Options()
            prefs = {
                'download.default_directory': local + "\\",
                'download.prompt_for_download': False,
                'download.directory_upgrade': True,
                'safebrowsing.enabled': True
            }
            chrome_options.add_experimental_option('prefs', prefs)
            if os.getenv('Ver') == '0':
                chrome_options.add_argument('--headless=new')  # Habilita o modo headless

            driver = webdriver.Chrome(options=chrome_options)
            Log(f"Diretório de download configurado para: {local}", level='info')

            url = 'https://www.joaopessoa.pb.gov.br/pc/certidaoNegativa.xhtml'
            driver.get(url)
            Log(f"Navegando para a URL: {url}", level='info')
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formbody:opt2"]/div[2]')))
            driver.find_element(By.XPATH, '//*[@id="formbody:opt2"]/div[2]').click()
            driver.find_element(By.XPATH, '//*[@id="formbody:cnpj"]').click()
            time.sleep(0.1)
            actions = ActionChains(driver)
            actions.send_keys(cnpj, Keys.RETURN).perform()
            Log(f"CNPJ inserido: {cnpj}", level='info')

            try:
                driver.find_element(By.XPATH, '//*[@id="formbody:j_idt81"]').click()
            except Exception as inner_exception:
                Log(f"Erro ao tentar clicar no elemento 'formbody:j_idt81', CNPJ não localizado", level='warning')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formbody:j_idt102"]/span[2]')))
            driver.find_element(By.XPATH, '//*[@id="formbody:j_idt102"]/span[2]').click()
            fim = True
            Log("Relatório gerado com sucesso.", level='info')
            time.sleep(1)
        except Exception as e:
            Log(f"Erro ao executar a função PortalContribuinte", level='error')
           
            return False
        finally:
            if fim:
                Log("Execução finalizada com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.PortalContribuinte();")    
            else:
                Log("Falha na execução da função PortalContribuinte.", level='info')
                webview.windows[0].evaluate_js("window.PortalContribuinte(false);")
            if driver:
                driver.quit()
                
    def GuiaIssPrestador(self, login, senha, mes, ano):
        fim = True
        driver = logar(login=login, senha=senha, tipo='GuiaIssPrestador', head=False)
        if not driver:
            Log("Falha ao iniciar o driver. Login ou senha incorretos.", level='error')
            return False

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/a/span[1]')))
            actions = ActionChains(driver)

            menu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/a/span[1]')
            actions.move_to_element(menu_elem).perform()
            time.sleep(0.5)

            submenu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/ul/li[2]/a')
            actions.move_to_element(submenu_elem).perform()
            time.sleep(0.5)

            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/ul/li[2]/ul/li[1]/a').click()
            Log("Menu 'Guia ISS Prestador' selecionado.", level='info')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@title="{ano}"]')))
            driver.find_element(By.XPATH, f'//*[@title="{ano}"]').click()
            Log(f"Ano {ano} selecionado.", level='info')

            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            nome_mes = meses[int(mes) - 1]

            time.sleep(0.9)

            row = driver.find_element(By.XPATH, f'//td[contains(text(), "{nome_mes}")]/parent::tr')
            row.find_element(By.XPATH, './/button[contains(@class, "ui-button")]').click()
            Log(f"Mês {nome_mes} selecionado.", level='info')
            time.sleep(0.3)

            try:
                report_button = driver.find_element(By.ID, 'form:tableGuiaDialog:0:j_idt103:relatorioButton')
                if report_button.is_enabled():
                    report_button.click()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="form:j_idt121:relatorioDialogId"]/div[2]/iframe')))
                    
                    driver.execute_script("""
                        var iframe = document.querySelector('iframe');
                        if (iframe) {
                            iframe.style.position = 'fixed';
                            iframe.className = 'iframeBox width-1000 height-1000';
                            iframe.style.width = '100vw';  
                            iframe.style.height = '100vh'; 
                            iframe.style.top = '0';
                            iframe.style.left = '0';
                            iframe.style.border = 'none';  
                            iframe.style.zIndex = '9999';  
                            document.body.style.zoom = "100%"
                        } else {
                            throw new Error('Iframe não encontrado.');
                        }
                    """)
                    driver.maximize_window()
                    Log("Iframe encontrado e redimensionado.", level='info')
                else:
                    Log("Erro: Download não liberado.", level='error')
                    return { 'erro': 'Download não liberado' }
                fim = True
            except Exception as inner_exception:
                Log(f"Erro ao tentar clicar no botão de relatório", level='error')
               
                return { 'erro': 'Erro sistêmico' }
            
        except Exception as e:
            Log(f"Erro ao executar a função GuiaIssPrestador", level='error')
           
        finally:
            if fim:
                Log("Execução da função GuiaIssPrestador finalizada com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.GuiaIssPrestador();")
            else:
                Log("Falha na execução da função GuiaIssPrestador.", level='info')
                webview.windows[0].evaluate_js("window.GuiaIssPrestador(false);")
            if driver:
                driver.quit()

    def GuiaIssTomador(self, login, senha, mes, ano):
        fim = False
        driver = logar(login=login, senha=senha, tipo='GuiaIssTomador', head=False)
        if not driver:
            Log("Falha ao iniciar o driver. Login ou senha incorretos.", level='error')
            return False

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/a/span[1]')))
            actions = ActionChains(driver)

            menu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/a/span[1]')
            actions.move_to_element(menu_elem).perform()
            time.sleep(0.5)

            submenu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/ul/li[2]/a')
            actions.move_to_element(submenu_elem).perform()
            time.sleep(0.5)

            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[2]/ul/li[2]/ul/li[2]/a').click()
            Log("Menu 'Guia ISS Tomador' selecionado.", level='info')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//*[@title="{ano}"]')))
            driver.find_element(By.XPATH, f'//*[@title="{ano}"]').click()
            Log(f"Ano {ano} selecionado.", level='info')

            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            nome_mes = meses[int(mes) - 1]

            time.sleep(0.5)

            row = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f'//td[contains(text(), "{nome_mes}")]/parent::tr'))
            )
            Log(f"Mês {nome_mes} localizado.", level='info')

            # Re-localize o botão dentro dessa linha e clique
            button = WebDriverWait(row, 10).until(
                EC.presence_of_element_located((By.XPATH, './/button[contains(@class, "ui-button")]'))
            )
            button.click()
            Log("Botão para gerar relatório clicado.", level='info')
            time.sleep(0.3)

            try:
                report_button = driver.find_element(By.ID, 'form:tableGuiaDialog:0:j_idt103:relatorioButton')
                if report_button.is_enabled():
                    report_button.click()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="form:j_idt121:relatorioDialogId"]/div[2]/iframe')))
                    
                    driver.execute_script("""
                        var iframe = document.querySelector('iframe');
                        if (iframe) {
                            iframe.style.position = 'fixed';
                            iframe.className = 'iframeBox width-1000 height-1000';
                            iframe.style.width = '100vw';  
                            iframe.style.height = '100vh'; 
                            iframe.style.top = '0';
                            iframe.style.left = '0';
                            iframe.style.border = 'none';  
                            iframe.style.zIndex = '9999';  
                            document.body.style.zoom = "100%"
                        } else {
                            throw new Error('Iframe não encontrado.');
                        }
                    """)
                    driver.maximize_window()
                    Log("Iframe encontrado e redimensionado.", level='info')
                else:
                    Log("Erro: Download não liberado.", level='error')
                    return { 'erro': 'Download não liberado' }
            except Exception as inner_exception:
                Log(f"Erro ao tentar clicar no botão de relatório", level='error')
               
                return { 'erro': 'Erro sistêmico' }
        
            fim = True
        except Exception as e:
            Log(f"Erro ao executar a função GuiaIssTomador", level='error')
           
            return False
        finally:
            if fim:
                Log("Execução da função GuiaIssTomador finalizada com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.GuiaIssTomador();")
            else:
                Log("Falha na execução da função GuiaIssTomador.", level='info')
                webview.windows[0].evaluate_js("window.GuiaIssTomador(false);")
            if driver:
                driver.quit() 
                
    def NfsTomador(self, login, senha, competencia1, competencia2):
        fim = False
        driver = logar(login=login, senha=senha, tipo='NfsTomador')
        if not driver:
            Log("Falha ao iniciar o driver. Login ou senha incorretos.", level='error')
            return False

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/a')))
            actions = ActionChains(driver)

            menu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/a')
            actions.move_to_element(menu_elem).perform()
            time.sleep(0.1)

            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/ul/li[2]/a').click()
            Log("Menu 'NFS Tomador' selecionado.", level='info')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'form:dtComp_input')))
            driver.find_element(By.ID, 'form:dtComp_input').send_keys(competencia1)
            driver.find_element(By.ID, 'form:compFim_input').send_keys(competencia2)
            Log(f"Competência de início: {competencia1}, Competência de fim: {competencia2}.", level='info')

            driver.find_element(By.XPATH, '//*[@id="form:nrNfse"]').click()
            driver.find_element(By.XPATH, '//*[@id="form:j_idt60"]').click()
            time.sleep(0.3)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="commandButton_exportar"]')))
            driver.find_element(By.XPATH, '//*[@id="commandButton_exportar"]').click()
            Log("Botão 'Exportar' clicado.", level='info')

            fim = True
        except Exception as e:
            Log(f"Erro ao clicar no elemento 'NFS Tomador'", level='error')
           
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Execução da função NfsTomador finalizada com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.NfsTomador();")
            else:
                Log("Falha na execução da função NfsTomador.", level='info')
                webview.windows[0].evaluate_js("window.NfsTomador(false);")
            if driver:
                driver.quit()
    
    def NfsPrestador(self, login, senha, competencia1, competencia2):
        fim = False
        driver = logar(login=login, senha=senha, tipo='NfsPrestador', head=False)
        if not driver:
            Log("Falha ao iniciar o driver. Login ou senha incorretos.", level='error')
            return False

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/a')))
            actions = ActionChains(driver)

            menu_elem = driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/a')
            actions.move_to_element(menu_elem).perform()
            time.sleep(0.1)

            driver.find_element(By.XPATH, '//*[@id="formMenuPrincipal:menuPrincipal"]/ul/li[4]/ul/li[1]/a').click()
            Log("Menu 'NFS Prestador' selecionado.", level='info')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'form:dtComp_input')))
            driver.find_element(By.ID, 'form:dtComp_input').send_keys(competencia1)
            driver.find_element(By.ID, 'form:compFim_input').send_keys(competencia2)
            Log(f"Competência de início: {competencia1}, Competência de fim: {competencia2}.", level='info')

            driver.find_element(By.XPATH, '//*[@id="form:nrNfse"]').click()
            driver.find_element(By.XPATH, '//*[@id="form:j_idt60"]').click()
            time.sleep(0.3)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="commandButton_exportar"]')))
            driver.find_element(By.XPATH, '//*[@id="commandButton_exportar"]').click()
            Log("Botão 'Exportar' clicado.", level='info')

            fim = True
        except Exception as e:
            Log(f"Erro ao clicar no elemento 'NFS Prestador'", level='error')
           
            return False
        finally:
            time.sleep(1)
            if fim:
                Log("Execução da função NfsPrestador finalizada com sucesso.", level='info')
                webview.windows[0].evaluate_js("window.NfsPrestador();")
            else:
                Log("Falha na execução da função NfsPrestador.", level='info')
                webview.windows[0].evaluate_js("window.NfsPrestador(false);")
            if driver:
                driver.quit()
                
    def taskList(self):
        tasks = []
        
        with open('./task.tsk', 'r', encoding='utf-8') as file:
            for linha in file:
                linha = linha.strip()
                if '||' in linha:
                    caminho_arquivo, data_hora = linha.split('||')
                    caminho_arquivo = caminho_arquivo.strip()
                    data_hora = data_hora.strip()

                    # Adiciona ao array de tasks como um dicionário
                    tasks.append({
                        'path': caminho_arquivo,
                        'data': data_hora
                        })

        # Serializa o array de tasks para JSON e passa para o JavaScript
        tasks_json = json.dumps(tasks)
        webview.windows[0].evaluate_js(f'window.renderTaskList({tasks_json});')

    def deleteTask(self, path):
        tasks = []

        # Lê as tarefas existentes
        with open('./task.tsk', 'r', encoding='utf-8') as file:
            for linha in file:
                linha = linha.strip()
                if '||' in linha:
                    caminho_arquivo, data_hora = linha.split('||')
                    caminho_arquivo = caminho_arquivo.strip()
                    data_hora = data_hora.strip()

                    if caminho_arquivo != path:
                        tasks.append({
                            'path': caminho_arquivo,
                            'data': data_hora
                        })

        # Escreve as tarefas restantes de volta no arquivo
        with open('./task.tsk', 'w', encoding='utf-8') as file:
            for task in tasks:
                file.write(f"{task['path']}||{task['data']}\n")

        # Serializa o array de tasks para JSON e passa para o JavaScript
        tasks_json = json.dumps(tasks)
        webview.windows[0].evaluate_js(f'window.renderTaskList({tasks_json});')
        
    def addTask(self, path, data):
        tasks = []

        # Lê as tarefas existentes
        with open('./task.tsk', 'r', encoding='utf-8') as file:
            for linha in file:
                linha = linha.strip()
                if '||' in linha:
                    caminho_arquivo, data_hora = linha.split('||')
                    caminho_arquivo = caminho_arquivo.strip()
                    data_hora = data_hora.strip()

                    # Adiciona ao array de tasks
                    tasks.append({
                        'path': caminho_arquivo,
                        'data': data_hora
                    })

        # Adiciona a nova tarefa
        tasks.append({
            'path': path,
            'data': data
        })

        # Escreve todas as tarefas de volta no arquivo
        with open('./task.tsk', 'w', encoding='utf-8') as file:
            for task in tasks:
                file.write(f"{task['path']}||{task['data']}\n")

        # Serializa o array de tasks para JSON e passa para o JavaScript
        tasks_json = json.dumps(tasks)
        webview.windows[0].evaluate_js(f'window.renderTaskList({tasks_json});')