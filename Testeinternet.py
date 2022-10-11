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



'''options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
#driver = webdriver.Chrome(chrome_options=options)'''

'''
pyautogui.click(x=260, y=50)
pyautogui.sleep(1)
pyautogui.press('backspace')
pyautogui.sleep(2)
pyautogui.write('Teste de internet')
pyautogui.press('enter')'''

'''df = pd.read_excel(fr'C:\\Users\\caduc\\Documents\\DOCS VS\\Velocidade internet.xlsx')
df['UPLOAD'][2] = u
df['DOWNLOAD'][2] = d
df['DATA'][2] = data'''

'''rowData = [data, d, u]
book = openpyxl.load_workbook('Velocidade_internet.xlsx')
writer = pd.ExcelWriter('Velocidade internet.xlsx', engine='openpyxl') 
writer.book = book
writer.sheets = dict(('Sheet1', ws) for ws in book.worksheets)
rowData.to_excel(writer, "Main", cols=['DATA', 'DOWNLOAD', 'UPLOAD'])
writer.save()'''

'''wb = openpyxl.load_workbook('Velocidade internet.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
rowData = [data, d, u]
sheet.append(rowData)
wb.save('Velocidade internet.xlsx')'''