
# import pytesseract
# import cv2

# # links uteis:
# # corrigir instalação windows: https://stackoverflow.com/questions/50951955/pytesseract-tesseractnotfound-error-tesseract-is-not-installed-or-its-not-i
# # instalar outra língua: https://github.com/tesseract-ocr/tessdata
# # pegar linguas: print(pytesseract.get_languages())

# imagem = cv2.imread("curriculo.png")

# resultado = pytesseract.image_to_string(imagem)

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\Tesseract.exe"


# resultado = pytesseract.image_to_string(imagem)

# print(resultado)

import pytesseract
from PIL import Image
import pandas as pd
import re

# Corrige o caminho para o executável Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\Tesseract.exe"

# Função para extrair todos os campos do texto


def extrair_campos(texto):
    campos = {}
    linhas = texto.split('\n')

    chave_atual = None
    valor_atual = ''

    for linha in linhas:
        if linha.strip():
            if linha.isdigit():
                if chave_atual is not None:
                    campos[chave_atual.strip()] = valor_atual.strip()
                chave_atual = linha.strip()
                valor_atual = ''
            else:
                valor_atual += linha + '\n'

    if chave_atual is not None:
        campos[chave_atual.strip()] = valor_atual.strip()

    return campos


# Abre a imagem que você quer processar
imagem = Image.open("scan.png")

# Extrai texto da imagem
resultado = pytesseract.image_to_string(imagem)

print("Texto extraído:")
print(resultado)

# Verificar se o texto extraído está vazio
if not resultado.strip():
    print("O texto extraído está vazio. Verifique a imagem ou o processo de OCR.")
    exit()

# Extrair todos os campos do texto
campos_extraidos = extrair_campos(resultado)

print("Campos extraídos:")
print(campos_extraidos)

# Verificar se campos_extraidos está vazio
if not campos_extraidos:
    print("Nenhum campo foi extraído do texto. Verifique o processo de extração de campos.")
    exit()

# Caminho para o arquivo Excel
excel_file = r'C:\Users\Aux01PCS-CSPFA-CSB\Desktop\scan\dadoscan.xlsx'

# Criar uma lista de dicionários para cada campo extraído
dados = [campos_extraidos]

# Criar o DataFrame a partir da lista de dicionários
df = pd.DataFrame(dados)

# Salvar o DataFrame atualizado de volta no arquivo Excel
df.to_excel(excel_file, index=False)

print("Texto adicionado à planilha Excel com sucesso.")
