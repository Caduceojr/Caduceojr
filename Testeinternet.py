import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import datetime

def salvar_excel(dt, lst):
    arquivoX = pd.ExcelWriter(f'Vel_{dt}.xlsx', engine='xlsxwriter')
    df = pd.DataFrame(lst, columns=['DATA;DOWNLOAD;UPLOAD'])
    df.to_excel(arquivoX, sheet_name='Aba1', index=False)
    arquivoX.save()


navegador = webdriver.Chrome(service=Service(r"C:\Users\caduc\Documents\DOCS_VS\chromedriver.exe"))
navegador.get('https://www.speedtest.net/pt')
# pyautogui.click(x=875, y=25)
pyautogui.sleep(7)
navegador.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/div/div/div/div[2]/div[3]/div[1]/a/span[4]').click()
pyautogui.sleep(40)
d = navegador.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/div/div/div/div[2]/div[3]/div[3]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/span').text
pyautogui.sleep(1)
u = navegador.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/div/div/div/div[2]/div[3]/div[3]/div/div[3]/div/div/div[2]/div[1]/div[2]/div/div[2]/span').text
navegador.close()
pyautogui.sleep(1)
data = datetime.date.today()
linha = [data, d, u]
salvar_excel(data, linha)
print(f'Hoje {data} a velocidade de download da internet é de {d} MB e a de upload é de {u} MB')
print('Programa finalizado')



