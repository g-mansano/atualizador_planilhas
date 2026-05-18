import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io

st.set_page_config(page_title="Atualizador de Carteira SAP", page_icon="⚙️", layout="wide")

st.title("⚙️ Atualizador de Carteira SAP (Nova Estrutura)")

st.sidebar.header("Upload de Arquivos")
mestre_file = st.sidebar.file_uploader("Upload da Planilha Mestre (.xlsx)", type=['xlsx'])
dados_file = st.sidebar.file_uploader("Upload da Exportação SAP (3 abas)", type=['xlsx'])

def processar_dados(dados_excel):
    # 1. Ler as 3 abas específicas do SAP
    df_prod = pd.read_excel(dados_excel, sheet_name='Em produção')
    df_lib = pd.read_excel(dados_excel, sheet_name='Liberados')
    df_sep = pd.read_excel(dados_excel, sheet_name='Em separação')

    # 2. Tratamento Inicial (ffill nos cabeçalhos)
    colunas_ffill = ['Nº doc.', 'Código do PN', 'Nome do PN']
    for df in [df_prod, df_lib]:
        col_presentes = [col for col in colunas_ffill if col in df.columns]
        if col_presentes:
            df[col_presentes] = df[col_presentes].ffill()
        # Filtra linhas válidas
        if 'Linha do documento' in df.columns:
            df.dropna(subset=['Linha do documento'], inplace=True)

    # 3. Lógica Simplificada de Atrasados (Simulação da Aba 5)
    hoje = datetime.today()
    
    df_atrasados_lista = []
    for nome_origem, df_temp in [('Em Produção', df_prod), ('Liberados', df_lib)]:
        if 'Data de Entrega' in df_temp.columns:
            df_temp['Data de Entrega'] = pd.to_datetime(df_temp['Data de Entrega'], errors='coerce')
            atrasados = df_temp[df_temp['Data de Entrega'].dt.date < hoje.date()].copy()
            atrasados['Origem'] = nome_origem
            df_atrasados_lista.append(atrasados)
            
    if df_atrasados_lista:
        df_atrasados = pd.concat(df_atrasados_lista, ignore_index=True)
    else:
        df_atrasados = pd.DataFrame()

    # 4. Dicionário com os Nomes EXATOS das abas da Planilha Mestre
    dict_dfs = {
        '2-EM PRODUÇÃO': df_prod,
        '3-LIBERADOS': df_lib,
        '4-EM SEPARAÇÃO': df_sep,
        '5-ATRASADOS': df_atrasados
    }
    
    return dict_dfs

def injetar_dados_mestre(mestre_file, dict_dfs):
    wb = openpyxl.load_workbook(mestre_file)
    abas_atualizadas = []
    
    for nome_aba, df in dict_dfs.items():
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            max_row = ws.max_row
            max_col = ws.max_column
            
            # Limpa os dados antigos (linha 2 para baixo)
            if max_row >= 2:
                for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
                    for cell in row:
                        cell.value = None
                        
            # Injeta os novos dados sem quebrar o layout
            if not df.empty:
                for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
                    for c_idx, value in enumerate(row, 1):
                        if pd.isna(value):
                            value = None
                        elif isinstance(value, np.integer):
                            value = int(value)
                        elif isinstance(value, np.floating):
                            value = float(value)
                        ws.cell(row=r_idx, column=c_idx, value=value)
            
            abas_atualizadas.append(nome_aba)
            
    if not abas_atualizadas:
        st.warning(f"Nenhuma aba compatível foi encontrada na Planilha Mestre. Abas presentes no arquivo: {wb.sheetnames}")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if st.button("Processar e Atualizar Planilha", type="primary", use_container_width=True):
    if not mestre_file or not dados_file:
        st.error("Por favor, faça o upload da Planilha Mestre e dos Dados SAP.")
    else:
        try:
            with st.spinner("Lendo Dados Brutos SAP (Nova Estrutura)..."):
                dict_dfs = processar_dados(dados_file)

            with st.spinner("Injetando dados e preservando formatação..."):
                arquivo_processado = injetar_dados_mestre(mestre_file, dict_dfs)
                
            st.success("Planilha processada e atualizada com sucesso! ✅")
            st.download_button(
                label="📥 Baixar Carteira_Atualizada.xlsx",
                data=arquivo_processado,
                file_name="Carteira_Atualizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            st.error(f"Erro inesperado durante o processamento: {str(e)}")
