import streamlit as st
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO
import seaborn as sns
from unidecode import unidecode
import os

st.set_page_config(page_title="Gerador de Relatórios de Vendas", page_icon="📊", layout="centered")

# === ESTILO PERSONALIZADO ===
st.markdown("""
    <style>
    .main {
        background-color: #111827;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        background-color: #1f2937;
        border-radius: 12px;
        box-shadow: 0px 0px 10px rgba(0,0,0,0.3);
    }
    h1, h2, h3, h4, h5, h6 {
        color: #facc15;
        font-family: 'Segoe UI', sans-serif;
    }
    .stButton button {
        background-color: #10b981;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6em 1.5em;
        margin-top: 1rem;
    }
    .stButton button:hover {
        background-color: #059669;
    }
    .stFileUploader label {
        font-size: 1.1em;
        color: #d1d5db;
        margin-bottom: 0.5rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
# 📊 <span style='color:white;'>Gerador de Relatórios de Vendas</span> <span style='color:#facc15;'>- Bling</span>
""", unsafe_allow_html=True)

arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls ou .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos and st.button("📝 Gerar Relatórios"):
    df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
    df_custo_all = pd.read_excel(arquivo_custos, sheet_name=None)
    df_custo_all = pd.concat(df_custo_all.values(), ignore_index=True)

    df_bling.columns = df_bling.columns.str.strip()
    df_custo_all.columns = df_custo_all.columns.str.strip()

    mapeamento = {
        'SKU': 'SKU',
        'DATA': 'DATA_VENDA',
        'NUMERO': 'NF',
        'PRECO': 'PRECO_UNIT',
        'COMISS': 'COMISSAO',
        'FRETE': 'FRETE',
        'DESCRI': 'DESC_PRODUTO',
        'QUANT': 'QUANTIDADE'
    }

    for original in df_bling.columns:
        col_simplificada = unidecode(original.upper()).strip()
        for chave, novo_nome in mapeamento.items():
            if chave in col_simplificada or (chave == 'PRECO' and 'UNIT' in col_simplificada):
                df_bling.rename(columns={original: novo_nome}, inplace=True)
                break

    for col in df_bling.columns:
        col_clean = unidecode(col.upper().strip())
        if 'NUMERO' in col_clean:
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            break

    for original in df_custo_all.columns:
        col_simplificada = unidecode(original.upper()).strip()
        if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
            df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
        elif 'CUSTO' in col_simplificada:
            df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

    if 'SKU' not in df_custo_all.columns:
        st.error("❌ Não foi possível localizar a coluna de SKU na planilha de custos. Verifique se o cabeçalho está correto.")
        st.stop()

    df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()

    df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
    df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
    df_bling['NF'] = df_bling['NF'].astype(str).str.strip()

    df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

    df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']

    df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
        'DATA_VENDA': 'first',
        'DESC_PRODUTO': 'first',
        'PRECO_UNIT': 'mean',
        'QUANTIDADE': 'sum',
        'TOTAL_ITEM': 'sum',
        'FRETE': 'sum',
        'COMISSAO': 'sum'
    })

    df_agrupado['FRETE_DIV'] = 0.0
    df_agrupado['COMISSAO_DIV'] = 0.0
    df_agrupado['VALOR_RECEBIDO'] = 0.0

    for nf, grupo in df_agrupado.groupby('NF'):
        total_nf = grupo['TOTAL_ITEM'].sum()
        frete_nf = grupo['FRETE'].sum()
        comissao_nf = grupo['COMISSAO'].sum()
        for idx, row in grupo.iterrows():
            proporcao = row['TOTAL_ITEM'] / total_nf if total_nf else 0
            df_agrupado.loc[idx, 'FRETE_DIV'] = frete_nf * proporcao
            df_agrupado.loc[idx, 'COMISSAO_DIV'] = comissao_nf * proporcao
            df_agrupado.loc[idx, 'VALOR_RECEBIDO'] = row['TOTAL_ITEM'] - df_agrupado.loc[idx, 'FRETE_DIV'] - df_agrupado.loc[idx, 'COMISSAO_DIV']

    df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')
    df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
    df['IMPOSTO'] = df['TOTAL_ITEM'] * 0.09
    df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

    st.success("✅ Relatório processado com sucesso!")
    st.write(df)

    lucro_total = df['LUCRO'].sum()
    total_produtos = df['QUANTIDADE'].sum()

    st.markdown(f"### 💰 Lucro Total: R$ {lucro_total:.2f}")
    st.markdown(f"### 📦 Total de Produtos Vendidos: {int(total_produtos)}")

    fig, ax = plt.subplots()
    top_lucro = df.groupby('SKU')['LUCRO'].sum().nlargest(5)
    top_lucro.plot(kind='barh', ax=ax, color='#10b981')
    ax.set_title("Top 5 SKUs por Lucro")
    st.pyplot(fig)

    fig2, ax2 = plt.subplots()
    df.groupby('NF')['VALOR_RECEBIDO'].sum().plot(kind='line', marker='o', ax=ax2, color='#facc15')
    ax2.set_title("Recebimento por Nota Fiscal")
    st.pyplot(fig2)
