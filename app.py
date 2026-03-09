import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import zipfile

st.title("Gerador de Avaliações Dinâmicas")

# --- CONFIGURAÇÃO MANUAL (Ajuste aqui as coordenadas) ---
CELULA_NOME_MODELO = 'B5'       # Onde entra o nome do formando no modelo
LINHA_INICIO_PARAMETROS = 15     # Linha onde começam os parâmetros no modelo
COLUNA_PARAMETROS_MODELO = 'B'  # Coluna onde estão os nomes dos parâmetros no modelo
COLUNA_VALORES_MODELO = 'C'     # Coluna onde as notas serão inseridas no modelo

# 1. Upload
file_importacao = st.file_uploader("1. Ficheiro 'Importação' (Dados na Coluna K)", type=["xlsx"])
file_modelo = st.file_uploader("2. Ficheiro 'Modelo' (Com Macros)", type=["xlsm"])

if file_importacao and file_modelo:
    try:
        # --- PASSO 1: LER PARÂMETROS DO MODELO ---
        wb_check = load_workbook(file_modelo, data_only=True)
        ws_check = wb_check.active
        
        parametros = []
        # Lê 10 parâmetros (ajuste o range se forem mais)
        for row in range(LINHA_INICIO_PARAMETROS, LINHA_INICIO_PARAMETROS + 10):
            val = ws_check[f"{COLUNA_PARAMETROS_MODELO}{row}"].value
            if val:
                parametros.append(val)
        
        st.write(f"**Parâmetros detetados no modelo:** {', '.join(parametros)}")

        # --- PASSO 2: LER DADOS DA IMPORTAÇÃO ---
        # Lemos a Coluna K (Nomes) e as seguintes (Valores)
        num_colunas_necessarias = len(parametros) + 1
        df_dados = pd.read_excel(file_importacao, skiprows=12, usecols=range(10, 10 + num_colunas_necessarias), header=None)
        
        if st.button(f"Gerar {len(df_dados)} Avaliações Completas"):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for _, linha in df_dados.iterrows():
                    nome_formando = str(linha[10]) # Coluna K
                    if nome_formando == 'nan': continue
                    
                    # Carregar modelo com macros
                    file_modelo.seek(0)
                    wb = load_workbook(file_modelo, keep_vba=True)
                    ws = wb.active
                    
                    # Preencher Nome
                    ws[CELULA_NOME_MODELO] = nome_formando
                    
                    # Preencher Parâmetros (Coluna L, M, N... da importação para a Coluna C do modelo)
                    for i, param in enumerate(parametros):
                        valor_avaliacao = linha[12 + i] # Colunas L, M, N...
                        ws[f"{COLUNA_VALORES_MODELO}{LINHA_INICIO_PARAMETROS + i}"] = valor_avaliacao
                    
                    # Salvar em memória
                    temp_file = io.BytesIO()
                    wb.save(temp_file)
                    zip_file.writestr(f"Avaliacao_{nome_formando.replace(' ', '_')}.xlsm", temp_file.getvalue())
            
            st.success("✅ Todos os ficheiros foram gerados com as macros preservadas!")
            st.download_button("📥 Descarregar ZIP", zip_buffer.getvalue(), "avaliacoes.zip")

    except Exception as e:
        st.error(f"Erro na leitura: {e}")
