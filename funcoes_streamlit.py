import streamlit as st
import zipfile
import os
import openpyxl

# Função para criar um arquivo zip
def create_zip_file(file_paths, zip_name):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for file in file_paths:
            zipf.write(file, os.path.basename(file))



def format_as_text(filename):
    # Carregar o arquivo Excel
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active

    # Formatando cada célula na coluna A como texto
    for cell in sheet['A']:  # 'A' representa a coluna A
        cell.number_format = '@'  # O formato '@' é usado para texto em Excel

    # Salvar o arquivo
    workbook.save(filename)
