import streamlit as st
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from unidecode import unidecode

st.set_page_config(page_title="Gerador de Relatórios de Vendas", page_icon="📊", layout="centered")

# === ESTILO ===
st.markdown("""
    <style>
    .main { background-color: #111827; }
    .block-container { background-color: #1f2937; border-radius: 12px; padding: 2rem; }
    h1, h2, h3, h4 { color: #facc15; font-family: 'Segoe UI', sans-serif; }
    .stButton button {
        background-color: #10b981;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6em 1.5em;
        margin-top: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("## 📊 <span style='color:white;'>Gerador de Relatórios de Vendas</span> <span style='color:#facc15;'>- Bling</span>", unsafe_allow_html=True)

# === Upload dos arquivos ===
arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls ou .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos and st.button("📝 Gerar Relatórios"):
    df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
    df_custo_all = pd.read_excel(arquivo_custos, sheet_name=None)
    df_custo_all = pd.concat(df_custo_all.values(), ignore_index=True)

    df_bling.columns = df_bling.columns.str.strip()
    df_custo_all.columns = df_custo_all.columns.str.strip()

    # === Mapeamento flexível ===
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

    # === Renomeação robusta da coluna NF ===
    possiveis_nomes_nf = ['NF', 'NÚMERO', 'NUMERO', 'NÚMERO NF', 'Nº', 'NO', 'NÚMERO DA NF']
    col_nf_encontrada = False
    for col in df_bling.columns:
        col_normalizado = unidecode(col).strip().upper()
        if any(unidecode(poss).upper() in col_normalizado for poss in possiveis_nomes_nf):
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            col_nf_encontrada = True
            break
    if not col_nf_encontrada or 'NF' not in df_bling.columns:
        st.error("❌ Coluna 'NF' (número da nota fiscal) não encontrada. Verifique o cabeçalho da planilha Bling.")
        st.stop()

    # === Renomeia custo ===
    for col in df_custo_all.columns:
        nome = unidecode(col.upper().strip())
        if 'SKU' in nome or 'CODIGO' in nome:
            df_custo_all.rename(columns={col: 'SKU'}, inplace=True)
        elif 'CUSTO' in nome:
            df_custo_all.rename(columns={col: 'CUSTO'}, inplace=True)

    if 'SKU' not in df_custo_all.columns:
        st.error("❌ Coluna 'SKU' não encontrada na planilha de custos.")
        st.stop()

    # === Limpeza e tipos ===
    df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
    df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
    df_bling['NF'] = df_bling['NF'].astype(str).str.strip()

    df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
    df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

    df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']

    # === Agrupamento NF + SKU (remove duplicados) ===
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
    st.dataframe(df)

    st.metric("💰 Lucro Total", f"R$ {df['LUCRO'].sum():,.2f}")
    st.metric("📦 Total de Produtos", int(df['QUANTIDADE'].sum()))

    fig, ax = plt.subplots()
    df.groupby('SKU')['LUCRO'].sum().nlargest(5).plot(kind='barh', ax=ax, color='#10b981')
    ax.set_title("Top 5 SKUs por Lucro")
    st.pyplot(fig)
