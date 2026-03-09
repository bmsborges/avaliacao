import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import zipfile

st.title("Gerador de Avaliações (Suporte a Macros)")

# 1. Upload dos ficheiros
file_importacao = st.file_uploader("1. Ficheiro Importação (Nomes)", type=["xlsx"])
file_modelo = st.file_uploader("2. Modelo com Macros (.xlsm)", type=["xlsm"])

if file_importacao and file_modelo:
    df_nomes = pd.read_excel(file_importacao)
    coluna_nome = st.selectbox("Selecione a coluna dos nomes:", df_nomes.columns)
    nomes = df_nomes[coluna_nome].dropna().tolist()

    if st.button(f"Gerar {len(nomes)} Ficheiros XLSM"):
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for nome in nomes:
                # REINICIAR O PONTEIRO DO FICHEIRO
                file_modelo.seek(0)
                
                # CARREGAR COM SUPORTE A VBA
                # keep_vba=True é crucial para não perder as macros
                wb = load_workbook(file_modelo, keep_vba=True)
                ws = wb.active 

                # PREENCHER O NOME (Ajuste a célula 'B2' para a sua célula real)
                ws['B2'] = nome 

                # GUARDAR NUM BUFFER
                temp_file_buffer = io.BytesIO()
                wb.save(temp_file_buffer)
                
                # ADICIONAR AO ZIP (Mantendo a extensão .xlsm)
                file_name = f"Avaliacao_{nome.replace(' ', '_')}.xlsm"
                zip_file.writestr(file_name, temp_file_buffer.getvalue())

        st.success("Processamento concluído!")

        st.download_button(
            label="Baixar Tudo (.zip)",
            data=zip_buffer.getvalue(),
            file_name="avaliacoes_com_macros.zip",
            mime="application/zip"
        )
