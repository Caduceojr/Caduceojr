from webbrowser import get
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import datetime


def salvar_excel(dt, lst):
    arquivoX = pd.ExcelWriter(f'Monitores_{dt}.xlsx', engine='xlsxwriter')
    df = pd.DataFrame(lst, columns=['Descrição;Preço;url'])
    df.to_excel(arquivoX, sheet_name='Aba1', index=False)
    arquivoX.save()


navegador = webdriver.Chrome(service=Service(r"C:\Users\caduc\Documents\DOCS_VS\chromedriver.exe"))
navegador.get('https://www.kabum.com.br/')
pyautogui.sleep(3)
navegador.find_element(By.ID, 'input-busca').send_keys('Monitor')
pyautogui.sleep(2)
pyautogui.press('Enter')
data = datetime.date.today()
listadf = []
lista_produtos = navegador.find_elements(By.CLASS_NAME, 'productCard')
for item in lista_produtos:
    nomeProduto = ''
    precoProduto = ''
    urlProduto = ''
    if nomeProduto == '':
        try:
            nomeProduto = item.find_element(By.CLASS_NAME, 'kRYNji').text
        except Exception:
            pass
    elif nomeProduto == '':
        try:
            nomeProduto = item.find_element(By.CLASS_NAME, 'sc-d99ca57-0').text
        except Exception:
            pass
    # ----------
    if precoProduto == '':
        try:
            precoProduto = item.find_element(By.CLASS_NAME, 'jTvomc').text
        except Exception:
            pass
    elif precoProduto == '':
        try:
            precoProduto = item.find_element(By.CLASS_NAME, 'sc-3b515ca1-2').text
        except Exception:
            pass
    else:
        precoProduto = '0'
    # ------------
    if urlProduto == '':
        try:
            urlProduto = item.find_element(By.TAG_NAME, 'a').get_attribute('href')
        except Exception:
            pass
    else:
        urlProduto = '-'
    # print(nomeProduto, '-', precoProduto)
    # print(urlProduto)
    dadoslinha = nomeProduto + ';' + precoProduto + ';' + urlProduto
    listadf.append(dadoslinha)
salvar_excel(data, listadf)
navegador.close()
print('programa finalizado')