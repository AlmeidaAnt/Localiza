import time
from selenium import webdriver
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.support.ui import Select
import openpyxl
from selenium.webdriver import ActionChains
import os
import shutil
import locale

navegador = webdriver.Chrome()
navegador.maximize_window()

navegador.get('https://autenticacao.localiza.com/login?source_system=https:%2F%2Ffrota360.localiza.com%2Fhome&system_code=PORTAL_CLIENTE')

#Acessar Siter (login)
login = navegador.find_element(By.XPATH,'//*[@id="mat-input-0"]')
login.send_keys('gabriel.ramos@endicon.com.br')
time.sleep(2)

senha = navegador.find_element(By.XPATH,'//*[@id="mat-input-1"]')
senha.send_keys('frota2021')
time.sleep(2)

entrar = navegador.find_element(By.XPATH,'/html/body/app-root/div/app-card-login/div/div/form/button/span')
entrar.click()
time.sleep(15)

#Acessar Planilha de Cadastro
workbook = openpyxl.load_workbook('C:\\temp\\crlv_localiza.xlsx')
planilha = workbook.active

#Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]

#Pagina CRLV
    navegador.get('https://frota360.localiza.com/gerenciador-frota/gerenciador-crlv')

#placa
    placa = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mat-input-6"]')))
    placa.send_keys(veiculo)
    placa.send_keys(Keys.ENTER)
    time.sleep(5)

# Localiza o elemento mat-icon
    icone = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="divPortal"]/app-root/app-full-layout/div/div/mat-sidenav-container/mat-sidenav-content/div/app-exportar-crlv/div/div/app-generic-table/mat-table/mat-row/mat-cell[7]')))
    icone.click()
#Pesquisa
    baixar = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="divPortal"]/app-root/app-full-layout/div/div/mat-sidenav-container/mat-sidenav-content/div/app-exportar-crlv/div/div/app-generic-table/mat-table/mat-row/mat-cell[7]/span[1]/span')))
    baixar.click
    time.sleep(5)

    print('Documento',veiculo,'Baixado com Sucesso!')