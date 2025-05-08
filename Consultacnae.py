from selenium import webdriver
import pickle
import time
from PIL import Image
import pytesseract
import os
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyodbc
from selenium.webdriver.support.ui import Select
import unicodedata
import re


# Configuração do banco de dados
server = 'NTBK-03\\SQLEXPRESSLOYAL'
database = 'bdloyal'
username = 'raul'
password = '123456'
connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Conectar ao banco
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()

def pegar_antes_do_traco(texto):
    return texto.split(' - ')[0]

def limpar_cnpj(cnpj):
    return re.sub(r'[^\d]', '', cnpj)

def formatar_municipio(nome):
    # Remove acentos
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join([c for c in nome if not unicodedata.combining(c)])
    
    # Converte para minúsculas
    nome = nome.lower()
    
    # Remove pontuações
    nome = re.sub(r'[^\w\s-]', '', nome)
    
    # Substitui espaços por hífen
    nome = nome.replace(' ', '-')
    
    return nome

def acessarcnae(cnae):
    driver = webdriver.Chrome()
    try:
        driver.get('https://www.contabeis.com.br/ferramentas/simples-nacional/'+cnae+'/')

        # Carrega os cookies salvos
        with open("cookies.pkl", "rb") as f:
            cookies = pickle.load(f)

        for cookie in cookies:
            driver.add_cookie(cookie)
    except:
        print('erro')
        
#/html/body/div[5]/div[1]/div[5]/div[3]/div[1]/ul/li[2]/a
#/html/body/div[6]/div[1]/div[5]/div[3]/div[1]/ul/li[2]/a

# Recarrega a página já logado
    driver.get('https://www.contabeis.com.br/ferramentas/simples-nacional/'+cnae+'/')    
    driver.maximize_window()
    time.sleep(5)
    try:
        elementos = driver.find_elements(By.XPATH, '//*[@target="_blank"]')

        # Verifica se existem pelo menos 4
        if len(elementos) >= 4:
            texto = elementos[3].text  # Índice 3 = 4ª ocorrência
            anexo = texto
            driver.quit()
            return anexo
        else:           
            anexo = driver.find_element(By.XPATH, '/html/body/div[6]/div[1]/div[5]/div[3]/div[1]/ul/li[2]/a')         
            anexo = anexo.text
            driver.quit()
            return anexo
    except: 
               
        anexo = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div[5]/div[3]/div[1]/ul/li[2]/a')
        anexo = anexo.text
        if len(anexo) > 10:            
            anexo = driver.find_element(By.XPATH, '/html/body/div[6]/div[1]/div[5]/div[3]/div[1]/ul/li[1]/a')
            anexo = anexo.text
            anexo = pegar_antes_do_traco(anexo)
        driver.quit()
        return anexo


def acessarcnpj(cnpj,municipio):
    driver = webdriver.Chrome()
    driver.get('https://consultas.plus/lista-de-empresas/sao-paulo/'+municipio+'/'+cnpj+'/')

    time.sleep(8)
    try:
        strongs = driver.find_elements(By.TAG_NAME, "strong")

    # Verifica se encontrou algum
        if strongs:
            texto = strongs[-1].text  # Última ocorrência
            cnae = texto[:9]
            cnae = limpar_cnpj(cnae)       
            driver.quit()
            return cnae
   
    except:
        print('erro')
cursor.execute("SELECT CodSistema, CNPJ, Local FROM Empresas where Status = 'ATIVOS' and Cod = 317 and Local !='BRASILIA'")
empresas = cursor.fetchall()


# Verifica se encontrou resultados
if empresas:    
  for empresa in empresas:
    cod_interno, cnpj, municipio = empresa
    print(cod_interno)
    cnpj = limpar_cnpj(cnpj)
    if municipio == 'SBC':
       municipio = 'sao-bernardo-do-campo'
       cnae = acessarcnpj(cnpj,municipio)
    elif municipio == 'SCS':
       municipio = 'sao-caetano-do-sul'
       cnae = acessarcnpj(cnpj,municipio)
    else:    
        municipio = formatar_municipio(municipio)
        cnae = acessarcnpj(cnpj,municipio)
  
    print(cnae)    
    anexo = acessarcnae(cnae)
    print(anexo)
    if len(anexo) > 15:
        print(cod_interno, ' erro anexo errado'+anexo)
    else:
        cursor.execute("UPDATE Empresas SET Anexos1 = ? WHERE CodSistema = ?", (anexo, cod_interno))
        conn.commit()
        print(f"✅ Anexo atualizado para {anexo} no código {cod_interno}")

# Fechar conexão
cursor.close()
conn.close()   