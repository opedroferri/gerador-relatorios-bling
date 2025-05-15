# Versão 100% fiel adaptada para Streamlit
import streamlit as st
import pandas as pd
import os
from unidecode import unidecode
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO

st.set_page_config(page_title="Gerador de Relatórios", layout="centered")
st.title("📊 Gerador de Relatórios de Vendas - Bling")

arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls ou .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos and st.button("🚀 Gerar Relatório"):
    df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
    xls = pd.ExcelFile(arquivo_custos)
    df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

    mapeamento = {
        'SKU': 'SKU', 'DATA': 'DATA_VENDA', 'NUMERO': 'NF', 'PRECO': 'PRECO_UNIT',
        'COMISS': 'COMISSAO', 'FRETE': 'FRETE', 'DESCRI': 'DESC_PRODUTO', 'QUANT': 'QUANTIDADE'
    }

    df_bling.columns = df_bling.columns.str.strip()
    for original in df_bling.columns:
        col_simplificada = unidecode(original.upper()).strip()
        for chave, novo_nome in mapeamento.items():
            if chave in col_simplificada or (chave == 'PRECO' and 'UNIT' in col_simplificada):
                df_bling.rename(columns={original: novo_nome}, inplace=True)
                break

    if 'NF' not in df_bling.columns:
        for col in df_bling.columns:
            col_clean = unidecode(col.upper().strip())
            if 'NUMERO' in col_clean or 'NUMER' in col_clean:
                df_bling.rename(columns={col: 'NF'}, inplace=True)
                break
    if 'NF' not in df_bling.columns:
        for col in df_bling.columns:
            if col.strip() == 'NÃºmero':
                df_bling.rename(columns={col: 'NF'}, inplace=True)
                break
    if 'NF' not in df_bling.columns:
        st.error("❌ Mesmo após renomeação, a coluna 'NF' não foi encontrada.")
        st.stop()

    df_custo_all.columns = df_custo_all.columns.str.strip()
    for original in df_custo_all.columns:
        col_simplificada = unidecode(original.upper()).strip()
        if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
            df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
        elif 'CUSTO' in col_simplificada:
            df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

    colunas_necessarias = ['SKU', 'DATA_VENDA', 'NF', 'PRECO_UNIT', 'COMISSAO', 'FRETE', 'QUANTIDADE']
    for col in colunas_necessarias:
        if col not in df_bling.columns:
            st.error(f"❌ Coluna obrigatória não encontrada no arquivo do Bling: {col}")
            st.stop()

    df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
    df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
    df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
    df_bling['NF'] = df_bling['NF'].astype(str).str.strip().str.lstrip('0')
    df_bling['DESC_PRODUTO'] = df_bling['DESC_PRODUTO'].astype(str).str.upper().str.strip()

    df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
    df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

    df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']
    df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
        'DATA_VENDA': 'first', 'DESC_PRODUTO': 'first', 'PRECO_UNIT': 'mean',
        'QUANTIDADE': 'sum', 'TOTAL_ITEM': 'sum', 'COMISSAO': 'first', 'FRETE': 'first'
    })

    df_agrupado['COMISSAO_DIV'] = 0.0
    df_agrupado['FRETE_DIV'] = 0.0
    df_agrupado['VALOR_RECEBIDO'] = 0.0

    for nf, grupo in df_agrupado.groupby('NF'):
        total_nf = grupo['TOTAL_ITEM'].sum()
        total_comissao = grupo['COMISSAO'].astype(float).iloc[0]
        total_frete = grupo['FRETE'].astype(float).iloc[0]
        for i in grupo.index:
            proporcao = grupo.at[i, 'TOTAL_ITEM'] / total_nf if total_nf else 0
            df_agrupado.at[i, 'COMISSAO_DIV'] = proporcao * total_comissao
            df_agrupado.at[i, 'FRETE_DIV'] = proporcao * total_frete
            df_agrupado.at[i, 'VALOR_RECEBIDO'] = grupo.at[i, 'TOTAL_ITEM'] - df_agrupado.at[i, 'COMISSAO_DIV'] - df_agrupado.at[i, 'FRETE_DIV']

    df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')
    df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
    df['IMPOSTO'] = df['TOTAL_ITEM'] * 0.09
    df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

    colunas_monetarias = ['PRECO_UNIT', 'TOTAL_ITEM', 'COMISSAO_DIV', 'FRETE_DIV', 'VALOR_RECEBIDO', 'CUSTO_TOTAL', 'IMPOSTO', 'LUCRO']
    df[colunas_monetarias] = df[colunas_monetarias].round(2)

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

    df_saida = df_saida.drop_duplicates(subset=['NF', 'SKU'], keep='first')
    buffer_excel = BytesIO()
    df_saida.to_excel(buffer_excel, index=False)

    st.download_button(
        label="📥 Baixar Excel de Saída",
        data=buffer_excel.getvalue(),
        file_name="RELATORIO_FINAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF (resumo simples)
    buffer_pdf = BytesIO()
    with PdfPages(buffer_pdf) as pdf:
        fig, ax = plt.subplots(figsize=(8.5, 11))
        ax.axis('off')
        resumo = f"""
        RELATÓRIO ANALÍTICO DE VENDAS
        Data do Relatório: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

        Total de Produtos Vendidos: {int(df_saida['Quantidade'].sum())}
        Valor Total Recebido: R$ {df_saida['Valor Recebido'].sum():,.2f}
        Lucro Total: R$ {df_saida['Lucro'].sum():,.2f}
        Total de Notas Fiscais Emitidas: {df_saida['NF'].nunique()}
        Produtos distintos vendidos: {df_saida['SKU'].nunique()}
        """
        ax.text(0.1, 0.8, resumo, fontsize=12, va='top')
        pdf.savefig(fig)
        plt.close()

    st.download_button(
        label="📥 Baixar PDF Analítico",
        data=buffer_pdf.getvalue(),
        file_name="RELATORIO_ANALITICO.pdf",
        mime="application/pdf"
    )
