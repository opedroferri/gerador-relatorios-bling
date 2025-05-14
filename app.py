import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO
from datetime import datetime
import seaborn as sns
import os
from unidecode import unidecode

st.set_page_config(page_title="Gerador de Relatórios Bling", layout="centered")
st.title("📊 Gerador de Relatórios de Vendas - Bling")

# === UPLOAD ===
arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls ou .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos:
    if st.button("🚀 Gerar Relatórios"):
        # === LEITURA ===
        df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
        xls = pd.ExcelFile(arquivo_custos)
        df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

        # === RENOMEIA ===
        mapeamento = {
            'SKU': 'SKU', 'DATA': 'DATA_VENDA', 'NUMERO': 'NF', 'PRECO': 'PRECO_UNIT',
            'COMISS': 'COMISSAO', 'FRETE': 'FRETE', 'DESCRI': 'DESC_PRODUTO', 'QUANT': 'QUANTIDADE'}

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
                if 'NUMERO' in col_clean or col.strip() == 'NÃºmero':
                    df_bling.rename(columns={col: 'NF'}, inplace=True)
                    break

        df_custo_all.columns = df_custo_all.columns.str.strip()
        for original in df_custo_all.columns:
            col_simplificada = unidecode(original.upper()).strip()
            if 'SKU' in col_simplificada:
                df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
            elif 'CUSTO' in col_simplificada:
                df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

        # === LIMPEZA ===
        df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
        df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
        df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
        df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)

        df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
        df_bling['NF'] = df_bling['NF'].astype(str).str.strip().str.lstrip('0')
        df_bling['DESC_PRODUTO'] = df_bling['DESC_PRODUTO'].astype(str).str.upper().str.strip()

        df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
        df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

        # === AGRUPAMENTO ===
        df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']
        df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
            'DATA_VENDA': 'first',
            'DESC_PRODUTO': 'first',
            'PRECO_UNIT': 'mean',
            'QUANTIDADE': 'sum',
            'TOTAL_ITEM': 'sum',
            'COMISSAO': 'first',
            'FRETE': 'first'
        })

        # === RATEIO ===
        for col in ['COMISSAO_DIV', 'FRETE_DIV', 'VALOR_RECEBIDO']:
            df_agrupado[col] = 0.0

        for nf, grupo in df_agrupado.groupby('NF'):
            total_nf = grupo['TOTAL_ITEM'].sum()
            total_comissao = grupo['COMISSAO'].astype(float).iloc[0]
            total_frete = grupo['FRETE'].astype(float).iloc[0]
            for i in grupo.index:
                prop = grupo.at[i, 'TOTAL_ITEM'] / total_nf if total_nf else 0
                df_agrupado.at[i, 'COMISSAO_DIV'] = prop * total_comissao
                df_agrupado.at[i, 'FRETE_DIV'] = prop * total_frete
                df_agrupado.at[i, 'VALOR_RECEBIDO'] = grupo.at[i, 'TOTAL_ITEM'] - \
                    df_agrupado.at[i, 'COMISSAO_DIV'] - df_agrupado.at[i, 'FRETE_DIV']

        df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')
        df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
        df['IMPOSTO'] = df['VALOR_RECEBIDO'] * 0.09
        df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

        colunas_monetarias = ['PRECO_UNIT', 'TOTAL_ITEM', 'COMISSAO_DIV', 'FRETE_DIV',
                              'VALOR_RECEBIDO', 'CUSTO_TOTAL', 'IMPOSTO', 'LUCRO']
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

        # === Excel ===
        excel_buffer = BytesIO()
        df_saida.to_excel(excel_buffer, index=False)
        st.success("✅ Excel gerado com sucesso!")
        st.download_button("📥 Baixar Excel", data=excel_buffer.getvalue(), file_name="RELATORIO_FINAL.xlsx")

        # === PDF ===
        pdf_buffer = BytesIO()
        with PdfPages(pdf_buffer) as pdf:
            fig1, ax1 = plt.subplots(figsize=(8.5, 11))
            ax1.axis('off')
            ax1.set_title('Resumo Executivo de Vendas', fontsize=18, fontweight='bold', pad=30)
            resumo = f"""
            Data do Relatório: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

            Total Produtos Vendidos: {int(df_saida['Quantidade'].sum())}
            Valor Recebido: R$ {df_saida['Valor Recebido'].sum():,.2f}
            Lucro Total: R$ {df_saida['Lucro'].sum():,.2f}
            Notas Fiscais: {df_saida['NF'].nunique()}
            Produtos distintos: {df_saida['SKU'].nunique()}
            """
            ax1.text(0.05, 0.95, resumo, fontsize=12, va='top')
            pdf.savefig(fig1)
            plt.close()

            fig2, ax2 = plt.subplots(figsize=(10, 6))
            lucro_por_sku = df_saida.groupby('SKU')['Lucro'].sum().sort_values(ascending=True)
            lucro_por_sku.plot(kind='barh', ax=ax2, color="#1f77b4")
            ax2.set_title('Lucro Total por SKU')
            pdf.savefig(fig2)
            plt.close()

            fig3, ax3 = plt.subplots(figsize=(10, 6))
            qtd_por_sku = df_saida.groupby('SKU')['Quantidade'].sum().sort_values(ascending=True)
            qtd_por_sku.plot(kind='barh', ax=ax3, color="#2ca02c")
            ax3.set_title('Quantidade Vendida por SKU')
            pdf.savefig(fig3)
            plt.close()

            df_saida['Data da Venda'] = pd.to_datetime(df_saida['Data da Venda'], errors='coerce')
            lucro_por_dia = df_saida.groupby('Data da Venda')['Lucro'].sum()
            fig4, ax4 = plt.subplots(figsize=(10, 5))
            lucro_por_dia.plot(kind='line', ax=ax4, marker='o', linestyle='-', color="#d62728")
            ax4.set_title('Lucro Total por Data de Venda')
            ax4.tick_params(axis='x', rotation=45)
            pdf.savefig(fig4)
            plt.close()

        st.success("✅ PDF analítico gerado!")
        st.download_button("📥 Baixar PDF Analítico", data=pdf_buffer.getvalue(), file_name="RELATORIO_ANALITICO.pdf")

