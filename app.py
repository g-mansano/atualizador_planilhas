import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io

st.set_page_config(page_title="Atualizador de Carteira SAP", page_icon="⚙️", layout="wide")

st.title("⚙️ Atualizador de Carteira SAP (Quimidrol)")

st.sidebar.header("Upload de Arquivos")
mestre_file = st.sidebar.file_uploader("Upload da Planilha Mestre (.xlsx)", type=['xlsx'])
dados_file = st.sidebar.file_uploader("Upload da Exportação SAP (3 abas)", type=['xlsx'])

def processar_dados(dados_excel):
    """
    Processa os dados brutos respeitando a estrutura exata do SAP da Quimidrol
    """
    # 1. Ler as 3 abas estruturadas do SAP (Nomes exatos identificados no arquivo)
    df_prod = pd.read_excel(dados_excel, sheet_name='Em produção')
    df_lib = pd.read_excel(dados_excel, sheet_name='Liberados')
    df_sep = pd.read_excel(dados_excel, sheet_name='Em Separação') # S maiúsculo

    # 2. Tratamento de cabeçalhos vazios (Forward Fill nos campos agrupados do SAP)
    colunas_ffill = ['Nº doc.', 'Código do PN', 'Nome do PN']
    for df in [df_prod, df_lib]:
        col_presentes = [col for col in colunas_ffill if col in df.columns]
        if col_presentes:
            df[col_presentes] = df[col_presentes].ffill()
        
        # Filtrar linhas válidas (Garante que só entram linhas com itens)
        if 'Linha do documento' in df.columns:
            df.dropna(subset=['Linha do documento'], inplace=True)

    # 3. Lógica para gerar a aba de ATRASADOS (Filtro por data menor que hoje)
    hoje = datetime.today().date()
    df_atrasados_lista = []

    for nome_origem, df_temp in [('EM PRODUÇÃO', df_prod), ('3-LIBERADOS', df_lib)]:
        # Coluna real de data identificada no SAP: 'Data entrega/vencimento'
        if 'Data entrega/vencimento' in df_temp.columns:
            df_temp['Data entrega/vencimento'] = pd.to_datetime(df_temp['Data entrega/vencimento'], errors='coerce')
            
            # Filtrar quem está com data menor que hoje e ignorar nulos
            atrasados = df_temp[df_temp['Data entrega/vencimento'].dt.date < hoje].copy()
            if not atrasados.empty:
                atrasados['Origem'] = nome_origem
                df_atrasados_lista.append(atrasados)

    if df_atrasados_lista:
        df_atrasados = pd.concat(df_atrasados_lista, ignore_index=True)
    else:
        df_atrasados = pd.DataFrame()

    # Converter de volta as colunas de data para string formatada para não bugar no Excel
    for df in [df_prod, df_lib]:
        if 'Data entrega/vencimento' in df.columns:
            df['Data entrega/vencimento'] = df['Data entrega/vencimento'].dt.strftime('%d/%m/%Y')
            
    if not df_atrasados.empty and 'Data entrega/vencimento' in df_atrasados.columns:
        df_atrasados['Data entrega/vencimento'] = df_atrasados['Data entrega/vencimento'].dt.strftime('%d/%m/%Y')

    # 4. Mapeamento exato com as abas físicas da sua Planilha Mestre
    dict_dfs = {
        'EM PRODUÇÃO': df_prod,
        '3-LIBERADOS': df_lib,
        'EM SEPARAÇÃO': df_sep,
        '5-ATRASADOS': df_atrasados
    }
    
    return dict_dfs

def injetar_dados_mestre(mestre_file, dict_dfs):
    """
    Limpa os dados antigos a partir da linha 2 e injeta os novos preservando estilos
    """
    wb = openpyxl.load_workbook(mestre_file)
    abas_atualizadas = []
    
    for nome_aba, df in dict_dfs.items():
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            max_row = ws.max_row
            max_col = ws.max_column
            
            # Limpa tudo da linha 2 para baixo
            if max_row >= 2:
                for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
                    for cell in row:
                        cell.value = None
                        
            # Injeta os novos registros linha por linha
            if not df.empty:
                for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
                    for c_idx, value in enumerate(row, 1):
                        # Evita que tipos do numpy ou nulos quebrem o openpyxl
                        if pd.isna(value):
                            value = None
                        elif isinstance(value, np.integer):
                            value = int(value)
                        elif isinstance(value, np.floating):
                            value = float(value)
                        ws.cell(row=r_idx, column=c_idx, value=value)
            
            abas_atualizadas.append(nome_aba)
            
    if not abas_atualizadas:
        st.warning(f"Atenção: Nenhuma aba correspondente foi atualizada. Abas na mestre: {wb.sheetnames}")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Botão de execução da interface
if st.button("Processar e Atualizar Carteira", type="primary", use_container_width=True):
    if not mestre_file or not dados_file:
        st.error("Por favor, faça o upload de ambos os arquivos na barra lateral.")
    else:
        try:
            with st.spinner("Lendo e cruzando os dados do SAP..."):
                dict_dfs = processar_dados(dados_file)

            with st.spinner("Limpando dados antigos e injetando novos sem perder a formatação..."):
                arquivo_processado = injetar_dados_mestre(mestre_file, dict_dfs)
                
            st.success("Sua Carteira de Pedidos foi atualizada com sucesso! ✅")
            st.download_button(
                label="📥 Baixar Carteira_Atualizada.xlsx",
                data=arquivo_processado,
                file_name="Carteira_Atualizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento: {str(e)}")
            st.exception(e)
