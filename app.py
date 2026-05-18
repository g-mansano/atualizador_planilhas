import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io

st.set_page_config(page_title="Atualizador de Carteira SAP", page_icon="⚙️", layout="wide")

st.title("⚙️ Atualizador de Carteira SAP")

st.sidebar.header("Upload de Arquivos")
mestre_file = st.sidebar.file_uploader("Upload da Planilha Mestre (.xlsx)", type=['xlsx'])
dados_file = st.sidebar.file_uploader("Upload dos Dados Brutos SAP (.xlsx ou .csv)", type=['xlsx', 'csv'])

def processar_dados(df):
    """
    Realiza o tratamento de dados conforme as regras de negócio
    """
    df_tratado = df.copy()

    # 1. Fazer forward fill (ffill()) nas colunas de cabeçalho do SAP
    colunas_ffill = ['Nº doc.', 'Código do PN', 'Nome do PN']
    colunas_presentes = [col for col in colunas_ffill if col in df_tratado.columns]
    
    if len(colunas_presentes) < len(colunas_ffill):
        col_faltantes = set(colunas_ffill) - set(colunas_presentes)
        st.warning(f"Aviso: Algumas colunas para ffill não foram encontradas: {', '.join(col_faltantes)}")
    
    if colunas_presentes:
        df_tratado[colunas_presentes] = df_tratado[colunas_presentes].ffill()

    # 2. Filtrar apenas as linhas válidas (onde 'Linha do documento' não for nula)
    if 'Linha do documento' in df_tratado.columns:
        df_tratado = df_tratado[df_tratado['Linha do documento'].notna()]
    else:
        st.warning("Aviso: Coluna 'Linha do documento' não encontrada. O filtro não pôde ser aplicado.")

    # 3. Classificar os status baseando-se na data atual gerando uma coluna ClassItem
    # Como não sabemos a coluna exata de data no arquivo original, vamos usar 'Data de Vencimento' ou 'Data' como tentativa
    hoje = datetime.today()
    
    col_data_encontrada = None
    for col in ['Data de Vencimento', 'Data de Entrega', 'Data', 'Vencimento']:
        if col in df_tratado.columns:
            col_data_encontrada = col
            break
            
    if col_data_encontrada:
        df_tratado[col_data_encontrada] = pd.to_datetime(df_tratado[col_data_encontrada], errors='coerce')
        
        def classificar_data(data):
            if pd.isna(data):
                return 'PENDENTE'
            elif data.date() < hoje.date():
                return 'ATRASADO'
            else:
                return 'ATENDIDO'
                
        df_tratado['ClassItem'] = df_tratado[col_data_encontrada].apply(classificar_data)
    else:
        # Lógica genérica caso não haja coluna de data identificável
        st.info("Nota: Nenhuma coluna de data padrão (ex: 'Data de Vencimento') identificada para classificação. Marcando todos como 'PENDENTE'.")
        df_tratado['ClassItem'] = 'PENDENTE'

    # 4. Separar os DataFrames em abas lógicas (ex: df_producao, df_separacao, df_atrasados)
    # A lógica exata dependerá das suas necessidades. Aqui dividimos usando o ClassItem como um exemplo.
    df_producao = df_tratado[df_tratado['ClassItem'] == 'ATENDIDO']
    df_separacao = df_tratado[df_tratado['ClassItem'] == 'PENDENTE']
    df_atrasados = df_tratado[df_tratado['ClassItem'] == 'ATRASADO']

    dict_dfs = {
        'Producao': df_producao,
        'Separacao': df_separacao,
        'Atrasados': df_atrasados,
        'Base Completa': df_tratado
    }
    
    return dict_dfs

def injetar_dados_mestre(mestre_file, dict_dfs):
    """
    Lê a planilha mestre, limpa os dados antigos e injeta os novos dados célula a célula.
    Preserva a formatação, cores, etc.
    """
    # Carregar o template (regra de ouro)
    wb = openpyxl.load_workbook(mestre_file)
    abas_atualizadas = []
    
    for nome_aba, df in dict_dfs.items():
        # Apenas tenta atualizar se a aba existir na planilha mestre
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            
            # Limpar os dados antigos (da linha 2 para baixo)
            max_row = ws.max_row
            max_col = ws.max_column
            
            if max_row >= 2:
                for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
                    for cell in row:
                        cell.value = None
                        
            # Injetar os novos dados
            if not df.empty:
                # Converter DataFrame em linhas (ignorando header e index)
                for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
                    for c_idx, value in enumerate(row, 1):
                        # Tratamento para valores Nulos/Especiais do Pandas e Numpy (que quebram o openpyxl)
                        if pd.isna(value):
                            value = None
                        elif isinstance(value, np.integer):
                            value = int(value)
                        elif isinstance(value, np.floating):
                            value = float(value)
                            
                        # Acessar célula e atribuir valor (mantém a formatação original da célula)
                        ws.cell(row=r_idx, column=c_idx, value=value)
            
            abas_atualizadas.append(nome_aba)
            
    if not abas_atualizadas:
        st.warning("Nenhuma das abas ('Producao', 'Separacao', 'Atrasados', 'Base Completa') foi encontrada na Planilha Mestre. O arquivo não foi modificado.")
            
    # Salvar em buffer de memória para o Streamlit Download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

if st.button("Processar e Atualizar Planilha", type="primary", use_container_width=True):
    if not mestre_file or not dados_file:
        st.error("Por favor, faça o upload da Planilha Mestre e dos Dados SAP na barra lateral antes de processar.")
    else:
        try:
            with st.spinner("Lendo Dados Brutos SAP..."):
                if dados_file.name.lower().endswith('.csv'):
                    try:
                        # Tenta ler com separador ponto e vírgula e encoding latin1 (comum no SAP no Brasil)
                        df_bruto = pd.read_csv(dados_file, sep=';', encoding='latin1')
                    except Exception:
                        dados_file.seek(0)
                        # Fallback para o padrão
                        df_bruto = pd.read_csv(dados_file)
                else:
                    df_bruto = pd.read_excel(dados_file)

            with st.spinner("Aplicando regras de negócio..."):
                dict_dfs = processar_dados(df_bruto)

            with st.spinner("Injetando dados na Planilha Mestre... (Isso pode demorar um pouco)"):
                arquivo_processado = injetar_dados_mestre(mestre_file, dict_dfs)
                
            st.success("Planilha processada e atualizada com sucesso! ✅")
            
            st.download_button(
                label="📥 Baixar Planilha_Atualizada.xlsx",
                data=arquivo_processado,
                file_name="Planilha_Atualizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"Erro inesperado durante o processamento: {str(e)}")
            st.exception(e)
