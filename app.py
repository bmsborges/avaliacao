import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import zipfile

# Título da Aplicação Web
st.set_page_config(page_title="Automação de Avaliações", layout="centered")
st.title("Excel Auto-Preenchimento (Com Macros)")

# 1. Zona de Upload
st.subheader("📁 Carregar Ficheiros")
file_importacao = st.file_uploader("Carregar ficheiro 'Importação' (.xlsx)", type=["xlsx"])
file_modelo = st.file_uploader("Carregar Ficheiro 'Modelo' (.xlsm)", type=["xlsm"])

# Célula onde o nome será escrito no MODELO (Ajuste aqui, ex: 'B5')
CELULA_DESTINO_NOME = 'B5' 

if file_importacao and file_modelo:
    try:
        # Lógica para ler a coluna K a partir da linha 13 no ficheiro Importação
        # 'usecols="K"' foca na coluna K
        # 'skiprows=12' ignora as primeiras 12 linhas (começa na 13)
        # 'header=None' se a linha 13 já for um nome
        df_nomes = pd.read_excel(file_importacao, usecols="K", skiprows=12, header=None)
        
        # O pandas lê a coluna K como a primeira coluna disponível (índice 0)
        lista_formandos = df_nomes[df_nomes.columns[0]].dropna().astype(str).tolist()
        
        st.info(f"Foram detetados {len(lista_formandos)} formandos na lista.")

        if st.button("🚀 Gerar Avaliações Individuais"):
            # Criar um buffer para o ficheiro ZIP
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                progresso = st.progress(0)
                
                for i, nome in enumerate(lista_formandos):
                    # Reiniciar o ponteiro do modelo para cada novo formando
                    file_modelo.seek(0)
                    
                    # Carregar modelo mantendo as MACROS (keep_vba=True)
                    wb = load_workbook(file_modelo, keep_vba=True)
                    ws = wb.active # Assume a primeira folha aberta
                    
                    # ESCREVER O NOME NO MODELO
                    ws[CELULA_DESTINO_NOME] = nome
                    
                    # Guardar o ficheiro XLSM modificado num buffer temporário
                    temp_file = io.BytesIO()
                    wb.save(temp_file)
                    
                    # Adicionar ao arquivo ZIP com o nome do formando
                    nome_arquivo = f"Avaliacao_{nome.replace(' ', '_')}.xlsm"
                    zip_file.writestr(nome_arquivo, temp_file.getvalue())
                    
                    # Atualizar barra de progresso
                    progresso.progress((i + 1) / len(lista_formandos))

            st.success("✅ Processamento concluído com sucesso!")

            # Botão de Download Final
            st.download_button(
                label="📦 Descarregar Todas as Avaliações (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Avaliacoes_Praticas_Geradas.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Erro ao processar ficheiros: {e}")
        st.warning("Certifique-se que o ficheiro de importação tem dados na Coluna K a partir da linha 13.")

else:
    st.write("Aguardando o upload de ambos os ficheiros para começar...")
