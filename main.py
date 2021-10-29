from sys import path
from RPA.Browser.Selenium import Selenium
from RPA.JSON import JSON
import time
import pandas as pd
import win32com.client as win32
import os
import shutil


# Inicialização de apps
driver = Selenium()
json_lib = JSON()


def click_button_dive_in(timeout=60):
    
    # Loop para realizar retentativa da rotina
    for i in range(timeout):
        
        # Try/Exception para evitar quebra da execução
        try:

            # Clica no Elemento por XPath
            id_button = 'xpath://*[@id="node-23"]/div/div/div/div/div/div/div/a'
            driver.set_focus_to_element(id_button)
            driver.click_element(id_button)

            # Executa a instrução JS para validar se o grid apareceu
            inst_js = ' return document.getElementsByClassName("row top-gutter-20 top-margin-10")[0].firstElementChild.firstElementChild.getAttribute("aria-expanded")'
            validate_execution = driver.execute_javascript(inst_js)

            # Se o aria-expanded == "true", significa que o grid está disponível
            if validate_execution == "true":
                print("click_button_dive_in() - Success on click button DiveIn")
                return True
        except Exception as e:
            log_last_error = str(e)
            pass
        
        # Delay para retentativa
        time.sleep(1)

    # Caso tenha estourado Timeout, retorna False
    print("click_button_dive_in() - Timeout")
    print("click_button_dive_in() - Last error registred: " + log_last_error)
    return False 


def get_list_agencies(timeout=60):

    # Loop para realizar retentativa da rotina
    for i in range(timeout):

        # Declaramos o objeto aqui pois se houver algum erro de execução, inicializamos do zero
        dict_list_agencies = {}

        # Executa a instrução JS para capturar a quantidade de agências a serem capturadas
        inst_js_length = 'return document.getElementById("agency-tiles-widget").getElementsByClassName("wrapper")[0].getElementsByClassName("h4 w200").length'
        quant_agencies = int(driver.execute_javascript(inst_js_length))

        # Caso a quantidade de Agências != 0, realizamos a captura dos dados
        # Nota -> Devido ao delay entre a ação do click e o carregamento das agências, realizamos essa verificação
        if quant_agencies != 0:

            # Try/Exception para evitar quebra da execução
            try:

                # Realizaremos a captura de todas as agências encontradas
                for index in range(quant_agencies):

                    # Executa a instrução JS para capturar a descrição das agências e valores gastos
                    inst_js_desc = 'return document.getElementsByClassName("wrapper")[1].getElementsByClassName("h4 w200")[' + str(index) + '].innerText'
                    inst_js_value = 'return document.getElementsByClassName("wrapper")[1].getElementsByClassName("h1 w900")[' + str(index) + '].innerText'
                    desc = driver.execute_javascript(inst_js_desc)
                    value = driver.execute_javascript(inst_js_value)

                    # Atribuimos os dados ao nosso dicionário
                    dict_list_agencies[desc] = value
                
                # Por fim, caso tenha obtido êxito ao capturar os dados, retorna o dicionário
                print("get_list_agencies() - Success in get list of agencies")
                return dict_list_agencies

            except Exception as e:
                log_last_error = str(e)
                pass

        # Delay para retentativa
        time.sleep(1)

    # Caso tenha estourado Timeout, retorna False
    print("get_list_agencies() - Timeout")
    print("get_list_agencies() - Last error registred: " + log_last_error)
    return False
    

def create_output_folder(path):
    try:
        base_path = path
        if not os.path.exists(base_path):
            os.makedirs(path)
    except Exception as e:
        print("create_output_folder() - Error: " + str(e))


def file_move_download_to_output_folder(path_orig, path_dest):
    try:
        shutil.move(path_orig, path_dest)
    except Exception as e:
        print("move_arquivo_diretorio_contato() - Error on file move")


def read_data_excel(dict_data, folder, filename, method="default"):

    # Nota -> Tive problemas ao inicializar o metodo open_aplication() da classe Application. Portanto, realizei a rotina
    # do Excel através da lib Pandas
    # AttributeError: '<win32com.gen_py.Microsoft Excel 16.0 Object Library._Workbook instance at 0x2096423444640>' object has no attribute '__len__'

    if method == "default":
        df = pd.DataFrame(data=dict_data, index=[0])
    elif (method == "dictionary_of_list"):
        df=pd.DataFrame.from_dict(dict_data)
    
    df = df.T
    print(df)
    path_save_file = folder + filename + ".xlsx"
    df.to_excel(path_save_file, header=None)
    auto_fit_sheet(path=path_save_file)


def auto_fit_sheet(path, sheetname="Sheet1"):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path)
    ws = wb.Worksheets(sheetname)
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


def click_specific_agency_data(param_agency, timeout=60):

    # Variáveis auxiliares a rotina
    execute_click = False

    # Loop para realizar retentativa da rotina
    for i in range(timeout):

        # Try/Exception para evitar quebra da execução
        try:

            # Somente realiza a rotina do click não tenha sido realizada
            if not execute_click:

                # Executa a instrução JS para capturar a quantidade de agências a serem verificadas
                inst_js_length = 'return document.getElementById("agency-tiles-widget").getElementsByClassName("wrapper")[0].getElementsByClassName("h4 w200").length'
                quant_agencies = int(driver.execute_javascript(inst_js_length))

                # Caso a quantidade de Agências != 0, realizamos a captura dos dados
                # Nota -> Devido ao delay entre a ação do click e a carregamento das agências, realizamos essa verificação
                if quant_agencies != 0:

                    # Realizaremos a verificação de todas as agências encontradas
                    for index in range(quant_agencies):

                        # Executa a instrução JS para capturar a descrição das agências e valores gastos
                        base_element_js = 'document.getElementById("agency-tiles-widget").getElementsByClassName("wrapper")[0].getElementsByClassName("h4 w200")[' + str(index) + ']'
                        inst_js_get_desc = 'return ' + base_element_js + '.innerText'
                        desc = driver.execute_javascript(inst_js_get_desc)
                        
                        # Remove espaços e padroniza letras minúsculas da agencia a ser selecionada
                        replace_agency_param = param_agency.replace(" ","")
                        replace_agency_param = replace_agency_param.lower()

                        # Remove espaços e padroniza letras minúsculas da decrição capturada
                        replace_desc = desc.replace(" ","")
                        replace_desc = replace_desc.lower()

                        if replace_agency_param == replace_desc:
                            inst_js_click_desc = base_element_js + '.click()'
                            driver.execute_javascript(inst_js_click_desc)
                            execute_click = True
                            print("click_specific_agency_data() - Success on click in specific agency")
                    
            # Caso já tenha realizado o click, aguarda o carregamento da nova página
            else:

                # Executa a instrução JS para validar o carregamento da paga referente a agência especificada
                inst_js_get_desc_new_page = 'return document.getElementsByClassName("h4 w200 agencyName")[0].innerText'
                desc_new_page = driver.execute_javascript(inst_js_get_desc_new_page)

                # Remove espaços e padroniza letras minúsculas da decrição capturada
                replace_desc_new_page = desc_new_page.replace(" ","")
                replace_desc_new_page = replace_desc_new_page.lower()
                
                # Caso o nome da agência da nova página for o esperado, retorna True
                if replace_desc_new_page == replace_agency_param:
                    print("click_specific_agency_data() - Success in load page of agency: " + str(param_agency))
                    return True
            
        except Exception as e:
            log_last_error = str(e)
            pass

        # Delay para retentativa
        time.sleep(1)

    # Caso tenha estourado Timeout, retorna False
    print("click_specific_agency_data() - Timeout")
    print("click_specific_agency_data() - Last error registred: " + log_last_error)
    return False


def click_select_all_individual_investiments(timeout=60):

    # Loop para realizar retentativa da rotina
    for i in range(timeout):

        # Try/Exception para evitar quebra da execução
        try:
            inst_js_current_filter = 'return document.getElementsByClassName("dataTables_info")[0].innerText'
            current_filter = driver.execute_javascript(inst_js_current_filter)
            current_filter_split = current_filter.split(" ")
            entries_max_showing_on_page = current_filter_split[3]
            entries_max = current_filter_split[5]

            if entries_max_showing_on_page != entries_max:
                inst_js_filter_all_data = 'document.getElementsByClassName("dataTables_length")[0].getElementsByTagName("select")[0].value = "-1"'
                inst_js_change_event = 'document.getElementsByClassName("dataTables_length")[0].getElementsByTagName("select")[0].dispatchEvent(new CustomEvent("change"))'
                driver.execute_javascript(inst_js_filter_all_data)
                driver.execute_javascript(inst_js_change_event)
            else:
                print("click_select_all_individual_investiments() - Success on select all individual investments")
                return True

        except Exception as e:
            log_last_error = str(e)
            pass

        # Delay para retentativa
        time.sleep(1)
    
    # Caso tenha estourado Timeout, retorna False
    print("click_select_all_individual_investiments() - Timeout")
    print("click_select_all_individual_investiments() - Last error registred: " + log_last_error)
    return False


def get_individual_investments(timeout=60):
    
    # Loop para realizar retentativa da rotina
    for i in range(timeout):

        # Declaramos o objeto aqui pois se houver algum erro de execução, inicializamos do zero
        dict_list_individual_investments = {}

        # Executa a instrução JS para capturar a quantidade de agências
        inst_js_length = 'return document.getElementsByClassName("dataTables_scrollBody")[0].getElementsByTagName("tbody")[0].getElementsByTagName("tr").length'
        quant_individual_investments = int(driver.execute_javascript(inst_js_length))

        # Caso a quantidade de Agências != 0, realizamos a captura dos dados
        # Nota -> Devido ao delay entre a ação do click e o carregamento das agências, realizamos essa verificação
        if quant_individual_investments != 0:

            # Try/Exception para evitar quebra da execução
            try:

                # Realizaremos a captura de todas as agências encontradas
                for index in range(quant_individual_investments):

                    # Armazenaremos aqui os dados para cada UII
                    list_items_individual_investments = []
                    
                    # Captura a key referente ao individual investments
                    inst_js_uii = 'return document.getElementsByClassName("dataTables_scrollBody")[0].getElementsByTagName("tbody")[0].getElementsByTagName("tr")[' + str(index) + '].getElementsByTagName("td")[0].innerText'
                    key_individual_investment = driver.execute_javascript(inst_js_uii)

                    # Captura os dados de todas as colunas referentes a linha
                    for td_index in range (7):
                        
                        # Descartamos a primeira coluna pois já capturamos anteriormente
                        if td_index == 0:
                            continue
                        
                        # Executa a instrução JS para capturar os valores de todas as colunas referentes a linha
                        inst_js_value = 'return document.getElementsByClassName("dataTables_scrollBody")[0].getElementsByTagName("tbody")[0].getElementsByTagName("tr")[' + str(index) + '].getElementsByTagName("td")[' + str(td_index) + '].innerText'
                        value = driver.execute_javascript(inst_js_value)
                        list_items_individual_investments.append(value)
                        
                    # Por fim, atribuimos os dados ao nosso dicionário
                    dict_list_individual_investments[str(key_individual_investment)] = list_items_individual_investments
                
                # Por fim, caso tenha obtido êxito ao capturar os dados, retorna o dicionário
                print("get_individual_investments() - Success in get list of agencies")
                return dict_list_individual_investments

            except Exception as e:
                log_last_error = str(e)
                pass

        # Delay para retentativa
        time.sleep(1)

    # Caso tenha estourado Timeout, retorna False
    print("get_individual_investments() - Timeout")
    print("get_individual_investments() - Last error registred: " + log_last_error)
    return False


def get_url_business_case_pdf(timeout=60):
    
    # Loop para realizar retentativa da rotina
    for i in range(timeout):

        # Declaramos o objeto aqui pois se houver algum erro de execução, inicializamos do zero
        dict_url_business_case_pdf = {}

        # Executa a instrução JS para capturar a quantidade de agências
        inst_js_length = 'return document.getElementsByClassName("dataTables_scrollBody")[0].getElementsByTagName("tbody")[0].getElementsByTagName("tr").length'
        quant_individual_investments = int(driver.execute_javascript(inst_js_length))

        # Caso a quantidade de Agências != 0, realizamos a captura dos dados
        # Nota -> Devido ao delay entre a ação do click e o carregamento das agências, realizamos essa verificação
        if quant_individual_investments != 0:

            # Try/Exception para evitar quebra da execução
            try:

                # Realizaremos a captura de todas as agências encontradas
                for index in range(quant_individual_investments):

                    # Elemento base para realizar as rotinas
                    base_element = 'document.getElementsByClassName("dataTables_scrollBody")[0].getElementsByTagName("tbody")[0].getElementsByTagName("tr")[' + str(index) + '].getElementsByTagName("td")[0]'
                    
                    # Captura a key referente ao individual investments
                    inst_js_uii = 'return ' + base_element + '.innerText'
                    key_individual_investment = driver.execute_javascript(inst_js_uii)

                    # Executa a instrução JS para capturar os elementos que possuem links
                    inst_js_child_href = 'return ' + base_element + '.lastElementChild'
                    child_href = driver.execute_javascript(inst_js_child_href)

                    # Caso o elemento obtenha link, realiza a rotina seguinte

                    if child_href != None:

                        # Executa a instrução JS para capturar o link encontrado no elemento
                        inst_js_get_attribute_href = 'return ' + base_element + '.lastElementChild.getAttribute("href")'
                        href = driver.execute_javascript(inst_js_get_attribute_href)
                        
                        # Por fim, atribuimos os dados ao nosso dicionário
                        dict_url_business_case_pdf[str(key_individual_investment)] = href
                
                # Por fim, caso tenha obtido êxito ao capturar os dados, retorna o dicionário
                print("get_individual_investments() - Success in get list of agencies")
                return dict_url_business_case_pdf

            except Exception as e:
                log_last_error = str(e)
                pass

        # Delay para retentativa
        time.sleep(1)

    # Caso tenha estourado Timeout, retorna False
    print("get_individual_investments() - Timeout")
    print("get_individual_investments() - Last error registred: " + log_last_error)
    return False


def download_business_case_pdf(dict_url_pdf, folder_destination, timeout=60):

    # Url base para realizar os respectivos downloads
    base_url = 'https://itdashboard.gov'

    # Try/Except para tentar capturar instância ativa se houver
    try:
        # Redireciona para a url aproveitando instância previamente criada
        driver.go_to(base_url)

    except:
        # Caso não exista nova instância, cria uma nova 
        driver.open_available_browser(base_url)
    
    # Tratamento para caso apareça o alerta
    try:
        driver.handle_alert(action="DISMISS")
    except:
        pass

    # Loop para baixar todos os pdfs
    for key, url in dict_url_pdf.items():

        # Variaveis auxiliares a rotina
        click_download = False

        # Loop para realizar retentativa da rotina
        for i in range(timeout):

            # Try/Exception para evitar quebra da execução
            try:

                # Path base onde o arquivo será salvo
                path_file_download = r'C:\Users\ITGREEN\Downloads\\' + str(key) + ".pdf"
                path_file_destination = r'C:\Users\ITGREEN\Documents\pessoal\codigospy\webScraping\output\\' + str(folder_destination) + "\\" + str(key) + ".pdf"

                # URL onde encontra o respectivo arquivo para Download
                url_download_file = base_url + str(url)

                # Realiza a rotina caso não tenha sido realizada a ação de click no botão Download
                if click_download == False:

                    # Captura a url atual para comparar
                    current_url = driver.get_location()
                    
                    # Caso a url atual seja diferente da desejada, redirecionamos para tal
                    if current_url != url_download_file or current_url + "#" != url_download_file + "#":

                        # Redireciona para página
                        driver.go_to(base_url + str(url))

                        # Tratamento para caso apareça o alerta
                        try:
                            driver.handle_alert(action="DISMISS")
                        except:
                            pass

                    # Realiza o click para baixar o arquivo
                    driver.execute_javascript("document.querySelector('[href=" + '"#"' + "]').click()")
                    click_download = True

                # Se o arquivo existir, quebra o loop de retentativa e baixa o próximo arquivo
                if os.path.exists(path_file_download):
                    print ("download_business_case_pdf() - Success on download file: " + path_file_download)

                    # Move arquivo para pasta de saida
                    file_move_download_to_output_folder(path_orig=path_file_download, path_dest=path_file_destination)
                    break

            except Exception as e:
                log_last_error = str(e)
                print("download_business_case_pdf() - Last error registred: " + log_last_error)
                pass

            # Delay para retentativa
            time.sleep(1)
        
        # Caso tenha estourado Timeout, loga o erro e tenta fazer o restante
        if not os.path.exists(path_file_destination):
            print("download_business_case_pdf() - Error on download file: " + path_file_download)


def get_specific_agency_json():

    # Realiza a leitura da agencia especificada no arquivo JSON para realizar a criação de pastas
    json = json_lib.load_json_from_file('config.json')
    get_specific_agency_json = json_lib.get_value_from_json(json, "$.agency")

    # Retorna a agencia capturada
    return get_specific_agency_json


def prepare_ambient(base_path, name_folder):

    # Cria subpasta referente a agencia especificada no arquivo json
    create_output_folder(path=base_path + "\\output\\" + name_folder)


# Define a main() function that calls the other functions in order:
def main():
    try:
        
        # Loga inicialização
        print("main() - Start automation...")

        # Captura o nome da agencia especificada no arquivo JSON
        get_name_agency_json = get_specific_agency_json()

        # Trata o nome da agencia capturada para criar sua respectiva pasta caso não exista
        name_agency = get_name_agency_json.replace(" ", "_")

        # Cria as pastas de saida
        prepare_ambient(base_path=os.getcwd(), name_folder=name_agency)

        # Redireciona a URL
        driver.open_available_browser("https://itdashboard.gov/")

        # Maximiza o navegador
        driver.maximize_browser_window()

        # Clica no botão para aparecer o grid com os departamentos
        click_button_dive_in()

        # Captura os departamentos
        data_agencies = get_list_agencies()

        # Escreve no Excel
        read_data_excel(dict_data=data_agencies, folder=os.getcwd() + "\\output\\", filename="agencies")

        # Clica na Agencia especificada no Json
        click_specific_agency_data(param_agency=get_name_agency_json)

        # Seleciona todos os investimentos aplicados
        click_select_all_individual_investiments()

        # Captura todos os investimentos
        individual_investments = get_individual_investments()

        # Grava os dados no Excel
        read_data_excel(dict_data=individual_investments, 
                        folder=os.getcwd() + "\\output\\" + str(name_agency) + "\\", 
                        filename=name_agency, 
                        method="dictionary_of_list")
        
        # Captura todos os links que possuem PDF
        urls_pdf_individual_investments = get_url_business_case_pdf()

        # Baixa os respectivos arquivos
        download_business_case_pdf(dict_url_pdf=urls_pdf_individual_investments, folder_destination=name_agency)

        # Loga finalização
        print("main() - End Execution")
        
    finally:
        driver.close_all_browsers()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
    