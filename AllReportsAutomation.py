############ DAILY REPORTS AUTOMATION ############
import pyautogui
import win32gui
import win32con
import time
import sys
from selenium.common.exceptions import ElementClickInterceptedException
from datetime import date
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.by import By
#from PIL import Image
import glob, os, shutil
import pandas as pd
from datetime import datetime
import pytz
import win32com.client


print('reading preliminary information...')
                            
#CAMINNHOS PARA OS BOTÕES A SEREM CLICADOS
#BOTÃO EXPORTAR
ButtonExportURL = '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/div/div/report/exploration-container/div/div/section/app-bar/div/div[1]/button[2]/span'
#BOTÃO PDF
ButtonExportPDFURL = '/html/body/div[4]/div[2]/div/div/div/button[2]'
#BOTÃO CONFIRMAR EXPORTAÇÃO
ButtonPDFEnter = '/html/body/div[12]/div[2]/div/div/div/button[2]'
#
ButtonWppURL="/html/body/div[1]/div/div/div[3]/div/div[1]/div/button"
#UsernameURL = "//*[@id='i0116']"
#PasswordURL = "//*[@id='i0118']"
#BOTÃO DE ANEXO NO CHAT DO WHATSAPP
clipURL ='/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/div'
#BOTÃO DE ARQUIVO EMFORMATO DE IMAGEM, VÍDEO ETC
uploadURL = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'

print('variables loaded')



#Entrada automática no whatsapp e pbi

options = webdriver.ChromeOptions()
options.add_argument(r"user-data-dir=C:\Users\Caio Saber\Documents\Rappi\Pyhton\Credentials\chromedriver2")
options.add_argument("--start-maximized")

print('options loaded')
print(options)
#Driver google chrome
driver = webdriver.Chrome(executable_path=r'C:\Users\Caio Saber\Documents\Rappi\Pyhton\Credentials\chromedriver2\chromedriver.exe', options=options)

#_
  #_
    #_
      #_
        #_
          #_
            #_
################# G M V    D A I L Y #######################


print('Power BI GMV Daily Report Loading...')
driver.get("https://app.powerbi.com/groups/98872066-035a-4b3b-b358-e9d84864b249/reports/a0aeb3e6-fd9d-463a-9141-af1e61cf414a/ReportSection6e955e09817caab7289a")


#Botão Power BI Exportar Relatório
ButtonExport = WebDriverWait(driver, 60).until(lambda driver: driver.find_element_by_xpath(ButtonExportURL))
ButtonExport.click()
 


#Botão Exportar em PDF
try:
  # Tenta Clicar no Elemento
     ButtonExportPDF = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath(ButtonPDFEnter))
     ButtonExportPDF.click()
except ElementClickInterceptedException:
  # Usa JavaScript para scrollar até o início da página e evitar Interceptação
     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")



ButtonPDFEnter = WebDriverWait(driver, 30).until(lambda driver: driver.find_element_by_xpath('/html/body/div[13]/div/div/div/div[2]/legacy-scoped-root/ng-transclude/exploration-scoped-services-bridge/host-dialog-container/ng-transclude/export-report-dialog/dialog-frame/div/div[2]/section/dialog-footer/button[1]'))
ButtonPDFEnter.click()
      


#/html/body/div[13]/div/div/div/div[2]/legacy-scoped-root/ng-transclude/exploration-scoped-services-bridge/host-dialog-container/ng-transclude/export-report-dialog/dialog-frame/div/div[2]/section/dialog-footer/button[1]
# Nome do Arquivo
d=date.today() #Armazena data do dia de hoje
day_=str(d.day) #Transforma o dia de hoje em String
month=str(d.month) #Transforma o mês de hoje em String
year=str(d.year) #Transforma o ano de hoje em String
gmv_daily = ' _GMV Daily_.pdf'
date = year+month+day_
FM = 'February 2021 FM'
reportA= date + gmv_daily #concatena strings YYYYMMD criadas com nome do Report


#Diretório de arquivos GMV Daily
pathA = "G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\08. GMV Daily\\2. Regional\\6. June\\"

source_dir = 'C:\\Users\\Caio Saber\\Downloads'


#Deletando arquivo em caso dele já existir, para abrir espaço para a versão mais atualizada 
# for file in os.listdir(pathA):
#     if os.path.exists(path_to_gmv_daily_report_file):
#         os.remove(path_to_gmv_daily_report_file)


file_exists = False

path_to_gmv_daily_downloaded_file = os.path.join(source_dir, "[Daily*")
path_to_gmv_daily_report_file = os.path.join(pathA, reportA)

#condição while para aguardar carregamento do report: 
#verifica se arquivo já existe no diretório de download ou diretório do report
while file_exists == False:
    files = glob.glob(path_to_gmv_daily_downloaded_file)
    for file in os.listdir(pathA):
            if file == reportA:
              file_exists = True
    for file in files:
            file_exists = os.path.exists(file)


#Move PDF recém baixado para o diretório desejado
dst = pathA 
files = glob.glob(path_to_gmv_daily_downloaded_file)
for file in files:
    if os.path.isfile(file):
        #print('testing')
        shutil.move(file, dst)


#Cria número inteiro para o dia de hoje
day_of_today = d.day



for file in os.listdir(pathA):
    #retorna dado da data de arquivo no formato ctime
    date_time_str = time.ctime(os.path.getmtime(pathA+file))
    #formata ctime para datetime
    date_time_obj = datetime.strptime(date_time_str, '%a %b %d %H:%M:%S %Y')
    #condicional para mudar somente o nome do arquivo que foi baixado hoje
    if date_time_obj.day == day_of_today:
        os.rename(f'G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\08. GMV Daily\\2. Regional\\6. June\\{file}', "G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\08. GMV Daily\\2. Regional\\6. June\\{}".format(reportA))
    else:
            None        


#Fecha a janela do power bi
driver.quit()
print('GMV Daily sucessfully loaded!')
ButtonPDFEnter = '/html/body/div[9]/div/div/div/div[2]/legacy-scoped-root/ng-transclude/exploration-scoped-services-bridge/host-dialog-container/ng-transclude/export-report-dialog/dialog-frame/div/div[2]/section/dialog-footer/button[1]'
driver = webdriver.Chrome(executable_path=r'C:\Users\Caio Saber\Documents\Rappi\Pyhton\Credentials\chromedriver2\chromedriver.exe', options=options)

#_
  #_
    #_
      #_
        #_
          #_
            #_
################# M T D   O V E R V I E W #######################
print('Power BI MTD Overview Report Loading...')

driver.get("https://app.powerbi.com/groups/98872066-035a-4b3b-b358-e9d84864b249/reports/d2ead647-a89e-4eb0-915b-67ffe7e17e9e/ReportSection")

#Botão Power BI Exportar Relatório
ButtonExport = WebDriverWait(driver, 60).until(lambda driver: driver.find_element_by_xpath('/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/div/div/report/exploration-container/div/div/section/app-bar/div/div[1]/button[2]'))
ButtonExport.click()


#Botão Exportar em PDF
time.sleep(2)
try:
  # Tenta Clicar no Elemento
     ButtonExportPDF = WebDriverWait(driver, 30).until(lambda driver: driver.find_element_by_xpath('/html/body/div[3]/div[2]/div/div/div/button[2]'))
     ButtonExportPDF.click()
except ElementClickInterceptedException:
  # Usa JavaScript para scrollar até o início da página e evitar Interceptação
     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)
try:
  # Tenta Clicar no Elemento
     ButtonExportPDF = WebDriverWait(driver, 30).until(lambda driver: driver.find_element_by_xpath('/html/body/div[12]/div[2]/div/div/div/button[2]'))
     ButtonExportPDF.click()
except ElementClickInterceptedException:
  # Usa JavaScript para scrollar até o início da página e evitar Interceptação
     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#Botão Enter

ButtonPDFEnter = WebDriverWait(driver, 30).until(lambda driver: driver.find_element_by_xpath('/html/body/div[14]/div/div/div/div[2]/legacy-scoped-root/ng-transclude/exploration-scoped-services-bridge/host-dialog-container/ng-transclude/export-report-dialog/dialog-frame/div/div[2]/section/dialog-footer/button[1]'))
ButtonPDFEnter.click()


# Nome do Arquivo
mtd = ' _MTD Overview_.pdf'
reportB = date+mtd #concatena strings YYYYMMD criadas com nome do Report


#Diretório de arquivos MTD Overview
pathB = "G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\06. Data Overview\\1. S&P Data\\1. Month to Date\\6. June 2021\\"
#Espera do Download
file_exists = False

path_to_mtd_overview_downloaded_file = os.path.join(source_dir, "[Monthly*")
mtd_ovw_dir_file = os.path.join(pathB, reportB)

while file_exists == False:
    files = glob.iglob(path_to_mtd_overview_downloaded_file)
    for file in os.listdir(pathB):
          if file == reportB:
            file_exists = True
    for file in files:
            file_exists = os.path.exists(file)


#Deletando arquivo em caso dele já existir, para abrir espaço para a versão mais atualizada [yyyymmdd]
# if os.path.exists(pathB+reportB):
#     os.remove(pathB+reportB)

#Move PDF recém baixado para o diretório desejado
dst = pathB 
files = glob.iglob(path_to_mtd_overview_downloaded_file)

for file in files:
    if os.path.isfile(file):
        shutil.move(file, dst)

for file in os.listdir(pathB):
    #retorna dado da data de arquivo no formato ctime
    date_time_str = time.ctime(os.path.getmtime(pathB+file))
    #formata ctime para datetime
    date_time_obj = datetime.strptime(date_time_str, '%a %b %d %H:%M:%S %Y')
    #condicional para mudar somente o nome do arquivo que foi baixado hoje
    if date_time_obj.day == day_of_today:
        os.rename(f'G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\06. Data Overview\\1. S&P Data\\1. Month to Date\\6. June 2021\\{file}', "G:\\.shortcut-targets-by-id\\1RMxTvua_hdPwZf9gOIlE1qY_ZAve-p2w\\Strategy & Planning\\02. Routines\\06. Data Overview\\1. S&P Data\\1. Month to Date\\6. June 2021\\{}".format(reportB))
    else:
            None


driver.quit()
print('MTD Overview sucessfully loaded!')

print('Processing WhatsApp Send...')
#Abre o WhatsApp
driver = webdriver.Chrome(executable_path=r'C:\Users\Caio Saber\Documents\Rappi\Pyhton\Credentials\chromedriver2\chromedriver.exe', options=options)
driver.get("https://web.whatsapp.com/")


# Seleciona a caixa de pesquisa de conversa
wait = WebDriverWait(driver, 300) 


# Replace 'Friend's Name' with the name of your friend  
# or the name of a group  


target = '"Reports | Restaurantes"'
try:
  # Tenta Clicar no Elemento
    x_arg = '//span[contains(@title,' + target + ')]'
    group_title = wait.until(EC.presence_of_element_located(( 
    By.XPATH, x_arg))) 
    group_title.click() 
except ElementClickInterceptedException:
  # Usa JavaScript para scrollar até o início da página e evitar Interceptação
     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")


inp_XPATH = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_XPATH))) 


#INSERIR TEXTO QUE DESEJA ENVIAR
if datetime.now().strftime('%p') == 'AM':
   text_to_send = """Bom dia pessoal, seguem os KPIs atualizados:
                    _*Link para GMV Daily*_ https://cutt.ly/chn5n6w
                    _*Link para MTD Overview*_ https://cutt.ly/nhn5WdY"""
else:
   text_to_send = """Boa tarde pessoal, seguem os KPIs atualizados:
                    _*Link para GMV Daily*_ https://cutt.ly/chn5n6w
                    _*Link para MTD Overview*_ https://cutt.ly/nhn5WdY"""


input_box.send_keys( text_to_send + Keys.ENTER ) 
print('Text printed')

#Clica no botão de anexo
clipElement = WebDriverWait(driver, 20).until(lambda driver: driver.find_element_by_xpath(clipURL))
clipElement.click()


# Clica no botão de arquivo
inp_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[3]'
button =  WebDriverWait(driver, 20).until(lambda driver: driver.find_element_by_xpath(inp_XPATH))
button.click() 


hdlg = 0
while hdlg == 0:
    hdlg = win32gui.FindWindow(None, "Abrir")


time.sleep(1)   # second. This pause is needed


# Set filename and press Enter key
hwnd = win32gui.FindWindowEx(hdlg, 0, 'ComboBoxEx32', None)
hwnd = win32gui.FindWindowEx(hwnd, 0, 'ComboBox', None)
hwnd = win32gui.FindWindowEx(hwnd, 0, 'Edit', None)


#NOME DO pathA + ARQUIVO A SER ENVIADO
filename = pathA + reportA
win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, filename)
time.sleep(1)


# Press Save button
hwnd = win32gui.FindWindowEx(hdlg, 0, 'Button', '&Abrir')
win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)

time.sleep(1)

ButtonElement5 = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div/span'))
ButtonElement5.click()

time.sleep(1.5)


#Clica no botão de anexo
clipElement = WebDriverWait(driver, 20).until(lambda driver: driver.find_element_by_xpath(clipURL))
clipElement.click()


# Clica no botão de arquivo
inp_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[3]'
button =  WebDriverWait(driver, 20).until(lambda driver: driver.find_element_by_xpath(inp_XPATH))
button.click() 


hdlg = 0
while hdlg == 0:
    hdlg = win32gui.FindWindow(None, "Abrir")


time.sleep(1)   # second. This pause is needed


# Set filename and press Enter key
hwnd = win32gui.FindWindowEx(hdlg, 0, 'ComboBoxEx32', None)
hwnd = win32gui.FindWindowEx(hwnd, 0, 'ComboBox', None)
hwnd = win32gui.FindWindowEx(hwnd, 0, 'Edit', None)


#NOME DO pathA + ARQUIVO A SER ENVIADO
filename = pathB + reportB
win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, filename)

time.sleep(2)


# Press Save button
hwnd = win32gui.FindWindowEx(hdlg, 0, 'Button', '&Abrir')
win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)

time.sleep(2)

#Pressiona botão de enviar no WhatsApp
ButtonElement5 = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div/span'))
ButtonElement5.click()

print('Files Sent')
print('All tasks successfully executed!')

