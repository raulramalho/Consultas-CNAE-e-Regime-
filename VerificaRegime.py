import time
import pyodbc
import os
import sys
import pandas as pd
from datetime import datetime, timedelta
import re
import pyautogui
from PIL import Image
import pytesseract
import os
import cv2
import numpy as np
def acessar(cnpj,cnpjformatado):
    time.sleep(5)
  
    pyautogui.hotkey('tab')
    pyautogui.write(cnpjformatado)
    pyautogui.hotkey('enter')
    time.sleep(5)
    pyautogui.hotkey('enter')
    time.sleep(5)
    pyautogui.hotkey('end')
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('s')
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('s')
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    imagem = Image.open(r'C:\Users\NTBK_03\Desktop\imagem OCR\Screenshot_1.png')
        

        # Usar Tesseract para fazer a leitura do texto
    texto = pytesseract.image_to_string(imagem)
        # Exibir o texto lido
    texto = texto[:7]
    print(texto)
        #Obtém a URL da imagem
    pasta = r'C:\Users\NTBK_03\Desktop\imagem OCR'

        # Extensões de imagem que você deseja apagar    
    extensoes_imagem = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']

        # Percorre todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
            # Verifica se o arquivo tem uma extensão de imagem
        if any(arquivo.endswith(ext) for ext in extensoes_imagem):
            # Cria o caminho completo do arquivo
            caminho_arquivo = os.path.join(pasta, arquivo)
            # Remove o arquivo
            os.remove(caminho_arquivo)
            print(f'Arquivo removido: {caminho_arquivo}')
    #fechando popup         
    time.sleep(5)
    comparacao='Optante'
       
    if texto == comparacao:
        cursor.execute("UPDATE Empresas SET Regime = N'SIMPLES NACIONAL✅' WHERE CNPJ ='"+str(cnpj)+"'")
        conn.commit()
    else:
        cursor.execute("UPDATE Empresas SET Regime = N'SIMPLES NACIONAL❌' WHERE CNPJ ='"+str(cnpj)+"'")
        conn.commit()
    pyautogui.keyDown('alt')
    pyautogui.keyDown('left')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('left')    



        

# Configuração da conexão com o SQL Server
SERVER = 'NTBK-03\SQLEXPRESSLOYAL'
DATABASE = 'bdloyal'
USERNAME = 'raul'
PASSWORD = '123456'

def limpar_string(texto):
    return re.sub(r'[./-]', '', texto)

# Exemplo de uso




# Conectar ao banco de dados
try:
    conn = pyodbc.connect(
        f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
    )
    cursor = conn.cursor()
    
    # Consulta para buscar os CNPJs
    cursor.execute("SELECT CNPJ FROM Empresas where Regime = 'SIMPLES NACIONAL' and Cod >303")
    
    # Percorrer os resultados e realizar ação para cada CNPJ
    for row in cursor.fetchall():
        cnpj = row[0]
        cnpjformatado = limpar_string(cnpj)
        # Exemplo de ação (substitua pela sua lógica)
        print(f'Executando ação para o CNPJ: {cnpj}')
        acessar(cnpj,cnpjformatado)
    
    # Fechar conexão
    cursor.close()
    conn.close()

except Exception as e:
    print(f"Erro ao conectar ao banco: {e}")
