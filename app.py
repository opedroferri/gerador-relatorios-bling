import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import seaborn as sns
from io import BytesIO
from unidecode import unidecode

st.set_page_config(page_title="Gerador de Relatórios Bling", layout="wide")

st.title("📊 Gerador de Relatórios Bling com Rateio Inteligente de Frete")

arquivo_bling = st.file_uploader("Selecione o arquivo CSV exportado do Bling", type="csv")
arquivo_custos = st.file_uploader("Selecione a planilha de Custos Finais (.xls, .xlsx)", type=["xls", "xlsx"])

if arquivo_bling and arquivo_custos:
    df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
    xls = pd.ExcelFile(arquivo_custos)
    df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

    # Renomeia colunas do Bling
    df_bling.columns = df_bling.columns.str.strip()
    mapeamento = {
        'SKU': 'SKU', 'DATA': 'DATA_VENDA', 'NUMERO': 'NF',
        'PRECO': 'PRECO_UNIT', 'COMISS': 'COMISSAO',
        'FRETE': 'FRETE', 'DESCRI': 'DESC_PRODUTO', 'QUANT': 'QUANTIDADE'
    }
    for col in df_bling.columns:
        col_simpl = unidecode(col.upper())
        for chave, novo in mapeamento.items():
            if chave in col_simpl or (chave == 'PRECO' and 'UNIT' in col_simpl):
                df_bling.rename(columns={col: novo}, inplace=True)
                break
    # Renomeação emergência
    if 'NF' not in df_bling.columns:
        for col in df_bling.columns:
            if 'NUMERO' in unidecode(col.upper()) or col == 'NÃºmero':
                df_bling.rename(columns={col: 'NF'}, inplace=True)
                break
    if 'NF' not in df_bling.columns:
        st.error("❌ A coluna 'NF' não foi encontrada.")
        st.stop()

    # Renomeia planilha de custo
    df_custo_all.columns = df_custo_all.columns.str.strip()
    for col in df_custo_all.columns:
        col_simpl = unidecode(col.upper())
        if 'SKU' in col_simpl or 'CODIGO' in col_simpl:
            df_custo_all.rename(columns={col: 'SKU'}, inplace=True)
        elif 'CUSTO' in col_simpl:
            df_custo_all.rename(columns={col: 'CUSTO'}, inplace=True)

    # Conversões e limpeza
    col_necessarias = ['SKU', 'DATA_VENDA', 'NF', 'PRECO_UNIT', 'COMISSAO', 'FRETE', 'QUANTIDADE']
    for col in col_necessarias:
        if col not in df_bling.columns:
            st.error(f"❌ Coluna obrigatória ausente: {col}")
            st.stop()

    df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '').str.replace(',', '.').astype(float)
    df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '').str.replace(',', '.').astype(float)
    df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '').str.replace(',', '.').astype(float)
    df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
    df_bling['SKU'] = df_bling['SKU'].str.strip().str.upper()
    df_bling['NF'] = df_bling['NF'].str.strip().str.lstrip('0')
    df_bling['DESC_PRODUTO'] = df_bling['DESC_PRODUTO'].str.upper().str.strip()
    df_custo_all['SKU'] = df_custo_all['SKU'].str.strip().str.upper()
    df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '').str.replace(',', '.').astype(float)

    # Cálculo e agrupamento
    df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']
    df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
        'DATA_VENDA': 'first', 'DESC_PRODUTO': 'first', 'PRECO_UNIT': 'mean',
        'QUANTIDADE': 'sum', 'TOTAL_ITEM': 'sum', 'COMISSAO': 'first', 'FRETE': 'first'
    })

    df_agrupado['COMISSAO_DIV'] = 0.0
    df_agrupado['FRETE_DIV'] = 0.0
    df_agrupado['VALOR_RECEBIDO'] = 0.0

    # NOVA REGRA: rateia frete apenas entre produtos com preço ≥ 79.90
    for nf, grupo in df_agrupado.groupby('NF'):
        total_nf = grupo['TOTAL_ITEM'].sum()
        total_comissao = grupo['COMISSAO'].iloc[0]
        total_frete = grupo['FRETE'].iloc[0]

        grupo_eligible = grupo[grupo['PRECO_UNIT'] >= 79.9]
        total_eligible = grupo_eligible['TOTAL_ITEM'].sum()

        for i in grupo.index:
            prop_comissao = grupo.at[i, 'TOTAL_ITEM'] / total_nf if total_nf else 0
            df_agrupado.at[i, 'COMISSAO_DIV'] = prop_comissao * total_comissao

            if grupo.at[i, 'PRECO_UNIT'] >= 79.9 and total_eligible:
                prop_frete = grupo.at[i, 'TOTAL_ITEM'] / total_eligible
                df_agrupado.at[i, 'FRETE_DIV'] = prop_frete * total_frete

            df_agrupado.at[i, 'VALOR_RECEBIDO'] = grupo.at[i, 'TOTAL_ITEM'] - df_agrupado.at[i, 'COMISSAO_DIV'] - df_agrupado.at[i, 'FRETE_DIV']

    # Merge e finais
    df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')
    df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
    df['IMPOSTO'] = df['TOTAL_ITEM'] * 0.09
    df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

    col_monetarias = ['PRECO_UNIT', 'TOTAL_ITEM', 'COMISSAO_DIV', 'FRETE_DIV', 'VALOR_RECEBIDO', 'CUSTO_TOTAL', 'IMPOSTO', 'LUCRO']
    df[col_monetarias] = df[col_monetarias].round(2)

    df_saida = pd.DataFrame({
        'Data da Venda': pd.to_datetime(df['DATA_VENDA'], errors='coerce').dt.strftime('%d/%m/%Y'),
        'NF': df['NF'], 'SKU': df['SKU'], 'Descrição do Produto': df['DESC_PRODUTO'],
        'Quantidade': df['QUANTIDADE'], 'Preço Unitário': df['PRECO_UNIT'],
        'Preço Total': df['TOTAL_ITEM'], 'Valor Recebido': df['VALOR_RECEBIDO'],
        'Custo': df['CUSTO_TOTAL'], 'Imposto (9%)': df['IMPOSTO'], 'Lucro': df['LUCRO']
    })

    df_saida = df_saida.drop_duplicates(subset=['NF', 'SKU'])

    # Baixar Excel
    buffer_xlsx = BytesIO()
    df_saida.to_excel(buffer_xlsx, index=False)
    st.download_button("⬇️ Baixar Excel Final", buffer_xlsx.getvalue(), file_name="RELATORIO_FINAL.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # PDF gráfico
    buffer_pdf = BytesIO()
    with PdfPages(buffer_pdf) as pdf:
        fig1, ax1 = plt.subplots(figsize=(8.5, 11))
        ax1.axis('off')
        ax1.set_title('Resumo Executivo de Vendas', fontsize=18, fontweight='bold', pad=30)
        resumo = f"""
        Data do Relatório: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

        Total de Produtos Vendidos: {int(df_saida['Quantidade'].sum())}
        Valor Total Recebido: R$ {df_saida['Valor Recebido'].sum():,.2f}
        Lucro Total: R$ {df_saida['Lucro'].sum():,.2f}
        Total de Notas Fiscais Emitidas: {df_saida['NF'].nunique()}
        Produtos distintos vendidos: {df_saida['SKU'].nunique()}
        """
        ax1.text(0.05, 0.95, resumo, fontsize=12, va='top')
        pdf.savefig(fig1)
        plt.close()

        fig2, ax2 = plt.subplots(figsize=(10, 6))
        df_saida.groupby('SKU')['Lucro'].sum().sort_values().plot(kind='barh', ax=ax2)
        ax2.set_title('Lucro Total por SKU')
        pdf.savefig(fig2)
        plt.close()

        fig3, ax3 = plt.subplots(figsize=(10, 6))
        df_saida.groupby('SKU')['Quantidade'].sum().sort_values().plot(kind='barh', color="#2ca02c", ax=ax3)
        ax3.set_title('Quantidade Vendida por SKU')
        pdf.savefig(fig3)
        plt.close()

        df_saida['Data da Venda'] = pd.to_datetime(df_saida['Data da Venda'], errors='coerce')
        fig4, ax4 = plt.subplots(figsize=(10, 5))
        df_saida.groupby('Data da Venda')['Lucro'].sum().sort_index().plot(ax=ax4, marker='o', color='orange')
        ax4.set_title('Lucro Total por Data de Venda')
        ax4.tick_params(axis='x', rotation=45)
        pdf.savefig(fig4)
        plt.close()

    st.download_button("📄 Baixar PDF Analítico", buffer_pdf.getvalue(), file_name="RELATORIO_ANALITICO_FINAL.pdf", mime="application/pdf")

