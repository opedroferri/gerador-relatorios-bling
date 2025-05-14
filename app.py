import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO
from datetime import datetime
import seaborn as sns
from unidecode import unidecode

st.set_page_config(page_title="Gerador de Relatórios Bling", layout="centered")
st.title("📊 Gerador de Relatórios de Vendas - Bling")

# === UPLOAD ===
arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls ou .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos and st.button("🚀 Gerar Relatórios"):
    # === LEITURA ===
    df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
    xls = pd.ExcelFile(arquivo_custos)
    df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

    # === RENOMEIA COLUNAS ===
    df_bling.columns = df_bling.columns.str.strip()
    df_custo_all.columns = df_custo_all.columns.str.strip()

    # Mapeamento flexível
    mapeamento = {
        'SKU': 'SKU', 'DATA': 'DATA_VENDA', 'NUMERO': 'NF', 'PRECO': 'PRECO_UNIT',
        'COMISS': 'COMISSAO', 'FRETE': 'FRETE', 'DESCRI': 'DESC_PRODUTO', 'QUANT': 'QUANTIDADE'
    }
    for original in df_bling.columns:
        col_simplificada = unidecode(original.upper()).strip()
        for chave, novo_nome in mapeamento.items():
            if chave in col_simplificada or (chave == 'PRECO' and 'UNIT' in col_simplificada):
                df_bling.rename(columns={original: novo_nome}, inplace=True)
                break

    # Força renomeação da coluna NF
    col_nf_encontrada = False
    for col in df_bling.columns:
        col_nf = unidecode(col.upper().strip())
        if 'NUMERO' in col_nf or col_nf == 'NF':
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            col_nf_encontrada = True
            break
    if not col_nf_encontrada or 'NF' not in df_bling.columns:
        st.error("❌ Coluna 'NF' não encontrada no arquivo do Bling.")
        st.stop()

    for original in df_custo_all.columns:
        col_simplificada = unidecode(original.upper()).strip()
        if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
            df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
        elif 'CUSTO' in col_simplificada:
            df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)
    if 'SKU' not in df_custo_all.columns:
        st.error("❌ Coluna SKU não encontrada na planilha de custos.")
        st.stop()

    # === TRATAMENTO ===
    df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
    df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
    df_bling['NF'] = df_bling['NF'].astype(str).str.strip()

    df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
    df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

    df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']

    df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
        'DATA_VENDA': 'first',
        'DESC_PRODUTO': 'first',
        'PRECO_UNIT': 'mean',
        'QUANTIDADE': 'sum',
        'TOTAL_ITEM': 'sum',
        'COMISSAO': 'sum',
        'FRETE': 'sum'
    })

    df_agrupado['FRETE_DIV'] = 0.0
    df_agrupado['COMISSAO_DIV'] = 0.0
    df_agrupado['VALOR_RECEBIDO'] = 0.0

    for nf, grupo in df_agrupado.groupby('NF'):
        total_nf = grupo['TOTAL_ITEM'].sum()
        frete_nf = grupo['FRETE'].sum()
        comissao_nf = grupo['COMISSAO'].sum()
        for idx, row in grupo.iterrows():
            prop = row['TOTAL_ITEM'] / total_nf if total_nf else 0
            df_agrupado.loc[idx, 'FRETE_DIV'] = prop * frete_nf
            df_agrupado.loc[idx, 'COMISSAO_DIV'] = prop * comissao_nf
            df_agrupado.loc[idx, 'VALOR_RECEBIDO'] = row['TOTAL_ITEM'] - df_agrupado.loc[idx, 'FRETE_DIV'] - df_agrupado.loc[idx, 'COMISSAO_DIV']

    df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')
    df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
    df['IMPOSTO'] = df['VALOR_RECEBIDO'] * 0.09
    df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

    # === GERA EXCEL DE SAÍDA ===
    df_saida = pd.DataFrame({
        'Data da Venda': pd.to_datetime(df['DATA_VENDA'], errors='coerce').dt.strftime('%d/%m/%Y'),
        'NF': df['NF'],
        'SKU': df['SKU'],
        'Descrição do Produto': df['DESC_PRODUTO'],
        'Quantidade': df['QUANTIDADE'],
        'Preço Unitário': df['PRECO_UNIT'],
        'Preço Total': df['TOTAL_ITEM'],
        'Valor Recebido': df['VALOR_RECEBIDO'],
        'Custo': df['CUSTO_TOTAL'],
        'Imposto (9%)': df['IMPOSTO'],
        'Lucro': df['LUCRO']
    })

    excel_buffer = BytesIO()
    df_saida.to_excel(excel_buffer, index=False)
    st.download_button("📥 Baixar Excel Final", data=excel_buffer.getvalue(), file_name="RELATORIO_FINAL.xlsx")

    # === GERA PDF COM GRÁFICOS ===
    pdf_buffer = BytesIO()
    with PdfPages(pdf_buffer) as pdf:
        fig1, ax1 = plt.subplots(figsize=(8.5, 11))
        ax1.axis('off')
        ax1.set_title('Resumo Executivo de Vendas', fontsize=16, fontweight='bold', pad=20)
        texto = f"""
        📅 Data do Relatório: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

        🛒 Quantidade Total de Produtos Vendidos: {int(df['QUANTIDADE'].sum())}
        💰 Valor Total Recebido: R$ {df['VALOR_RECEBIDO'].sum():,.2f}
        📈 Lucro Total: R$ {df['LUCRO'].sum():,.2f}
        🧾 Total de Notas Fiscais: {df['NF'].nunique()}
        🧩 SKUs únicos: {df['SKU'].nunique()}
        """
        ax1.text(0.05, 0.95, texto, fontsize=12, va='top')
        pdf.savefig(fig1)
        plt.close()

        fig2, ax2 = plt.subplots(figsize=(10, 6))
        lucro_por_sku = df.groupby('SKU')['LUCRO'].sum().sort_values(ascending=True)
        lucro_por_sku.plot(kind='barh', ax=ax2, color='#10b981')
        ax2.set_title('Lucro Total por SKU')
        pdf.savefig(fig2)
        plt.close()

        fig3, ax3 = plt.subplots(figsize=(10, 6))
        qtd_por_sku = df.groupby('SKU')['QUANTIDADE'].sum().sort_values(ascending=True)
        qtd_por_sku.plot(kind='barh', ax=ax3, color="#2ca02c")
        ax3.set_title('Quantidade Vendida por SKU')
        pdf.savefig(fig3)
        plt.close()

        fig4, ax4 = plt.subplots(figsize=(10, 5))
        df['Data da Venda'] = pd.to_datetime(df['DATA_VENDA'], errors='coerce')
        lucro_por_dia = df.groupby('Data da Venda')['LUCRO'].sum()
        lucro_por_dia.plot(kind='line', ax=ax4, marker='o', linestyle='-', color="#d62728")
        ax4.set_title('Lucro Total por Data de Venda')
        ax4.tick_params(axis='x', rotation=45)
        pdf.savefig(fig4)
        plt.close()

    st.download_button("📥 Baixar PDF Analítico", data=pdf_buffer.getvalue(), file_name="RELATORIO_ANALITICO.pdf")
