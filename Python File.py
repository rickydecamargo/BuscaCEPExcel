
#Importante, para funcionar corretamente deve se colocar o programa chromedriver na pasta do projeto

from openpyxl import load_workbook
#importando o OS para abrir o arquivo Excel.
import os

#Indicamos onde serão salvos os dados
nome_arquivo_cep = "C:\\Users\Windows\Desktop\Python Projetos\BuscaCEPExcel\PesquisaEndereco_2.xlsx"
planilhaDadosEndereco = load_workbook(nome_arquivo_cep)


sheet_selecionada = planilhaDadosEndereco["CEP"]


#Parte básica para funcionar o Selenium e poder abrir o Chrome
from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
#importando o elemento By
from selenium.webdriver.common.by import By

#importar a biblioteca pyautogui para controle teclas e delays.
import pyautogui as tempoEspera

#Cria a variável navegador, abre o navegador do Google Chrome e abre o site do Busca CEP
navegador = opcoesSelenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

#aguardar 3 segundos
tempoEspera.sleep(3)

#Apontar no site onde será digitado o CEP e o digita
navegador.find_element(By.NAME, "endereco").send_keys("23548057") #Opcao com Name
#navegador.find_element(By.ID, "endereco").send_keys("05892387") #Opcao com ID:
#navegador.find_element(By.XPATH, '//*[@id="endereco"]').send_keys("05892387") #opção com XPATH

#aguardar 3 segundos
tempoEspera.sleep(3)

#Para clicar no Botão Buscar
#Lembrando que o nome do campo NAME ou ID pode mudar com o tempo.
navegador.find_element(By.NAME, "btn_pesquisar").click()

#aguardar 3 segundos
tempoEspera.sleep(3)

#Para pegar as informações - Inpsecionar o elemento em nome da rua, copiar XPATH

#para
for linha in range(2, len(sheet_selecionada['A']) + 1):

    # aguardar 3 segundos
    tempoEspera.sleep(3)

    #pesquisando pelo elemendo ID para clicar no botão nova busca.
    navegador.find_element(By.XPATH, '//*[@id="btn_nbusca"]').click()

    # aguardar 3 segundos
    tempoEspera.sleep(3)

    cepPesquisa = sheet_selecionada['A%s' % linha].value

    # aguardar 3 segundos
    tempoEspera.sleep(3)

    # Apontar no site onde será digitado o CEP e o digita
    navegador.find_element(By.XPATH, '//*[@id="btn_nbusca"]').send_keys(sheet_selecionada['A%s' % linha])

    # aguardar 3 segundos
    tempoEspera.sleep(3)

    # Para clicar no Botão Buscar
    # Lembrando que o nome do campo NAME ou ID pode mudar com o tempo.
    navegador.find_element(By.XPATH, '//*[@id="btn_nbusca"]').send_keys(sheet_selecionada['A%s' % linha])

    # aguardar 4 segundos
    tempoEspera.sleep(4)

    #Pega os dados da rua no site da busca CEP pelo XPATH
    rua = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[0].text
    print(rua)

    #Pega os dados da rua no site da busca CEP pelo XPATH
    bairro = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[1].text
    print(bairro)

    #Pega os dados da rua no site da busca CEP pelo XPATH
    cidade = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[2].text
    print(cidade)

    #Pega os dados da rua no site da busca CEP pelo XPATH
    cep = navegador.find_elements(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]')[3].text
    print(cep)

    #Seleciona a Sheet de Dados
    sheet_Dados_Para_Imprimir_Endereco = planilhaDadosEndereco["CEP"]
    linhaCorrentePlanilhaCEP = len(sheet_selecionada['A']) + 1

    #Pegamos a última linha preenchida na coluna A e acrescentamos + 1
    colunaA = "A" + str(linhaCorrentePlanilhaCEP) #Criando a variável para juntar A + a última linha ex: A2
    colunaB = "B" + str(linhaCorrentePlanilhaCEP) #Criando a variável para juntar B + a última linha ex: B2
    colunaC = "C" + str(linhaCorrentePlanilhaCEP) #Criando a variável para juntar C + a última linha ex: C2
    colunaD = "D" + str(linhaCorrentePlanilhaCEP) #Criando a variável para juntar D + a última linha ex: D2

    #Imprimindo as informações do site na planilha
    sheet_Dados_Para_Imprimir_Endereco[colunaA] = rua #A2 - A3 - A4
    sheet_Dados_Para_Imprimir_Endereco[colunaB] = bairro #B2 - B3 - C=B 4
    sheet_Dados_Para_Imprimir_Endereco[colunaC] = cidade #C2
    sheet_Dados_Para_Imprimir_Endereco[colunaD] = cep #D2

#Salvando o arquivo do excel com as novas informações
planilhaDadosEndereco.save(filename=nome_arquivo_cep)

#Abrindo e apresentando na tela o arquivo com os dados
os.startfile(nome_arquivo_cep)


