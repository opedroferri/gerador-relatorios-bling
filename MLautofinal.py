import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from unidecode import unidecode
from datetime import datetime

# === INTERFACE ===
root = tk.Tk()
root.withdraw()

arquivo_bling = filedialog.askopenfilename(title="Selecione o arquivo CSV exportado do Bling")
arquivo_custos = filedialog.askopenfilename(title="Selecione a planilha de Custos Finais (.xls, .xlsx)")

# === LEITURA DOS DADOS ===
df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)
xls = pd.ExcelFile(arquivo_custos)
df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

# === RENOMEIA COLUNAS DO BLING ===
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

df_bling.columns = df_bling.columns.str.strip()
for original in df_bling.columns:
    col_simplificada = unidecode(original.upper()).strip()
    for chave, novo_nome in mapeamento.items():
        if chave in col_simplificada or (chave == 'PRECO' and 'UNIT' in col_simplificada):
            df_bling.rename(columns={original: novo_nome}, inplace=True)
            break

# === RENOMEAÇÃO DE EMERGÊNCIA PARA NF ===
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
    raise Exception("❌ Mesmo após renomeação, a coluna 'NF' não foi encontrada.")

# === RENOMEIA PLANILHA DE CUSTO ===
df_custo_all.columns = df_custo_all.columns.str.strip()
for original in df_custo_all.columns:
    col_simplificada = unidecode(original.upper()).strip()
    if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
        df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
    elif 'CUSTO' in col_simplificada:
        df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

# === VERIFICAÇÃO DAS COLUNAS ===
colunas_necessarias = ['SKU', 'DATA_VENDA', 'NF', 'PRECO_UNIT', 'COMISSAO', 'FRETE', 'QUANTIDADE']
for col in colunas_necessarias:
    if col not in df_bling.columns:
        raise Exception(f"❌ Coluna obrigatória não encontrada no arquivo do Bling: {col}")

# === CONVERSÃO E LIMPEZA ===
df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)

df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip().str.upper()
df_bling['NF'] = df_bling['NF'].astype(str).str.strip().str.lstrip('0')
df_bling['DESC_PRODUTO'] = df_bling['DESC_PRODUTO'].astype(str).str.upper().str.strip()

df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

# === CÁLCULO TOTAL ITEM ===
df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']

# === AGRUPAMENTO POR NF + SKU ===
df_agrupado = df_bling.groupby(['NF', 'SKU'], as_index=False).agg({
    'DATA_VENDA': 'first',
    'DESC_PRODUTO': 'first',
    'PRECO_UNIT': 'mean',
    'QUANTIDADE': 'sum',
    'TOTAL_ITEM': 'sum',
    'COMISSAO': 'first',
    'FRETE': 'first'
})

# === RATEIO PROPORCIONAL ===
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
        df_agrupado.at[i, 'VALOR_RECEBIDO'] = grupo.at[i, 'TOTAL_ITEM'] - \
                                               df_agrupado.at[i, 'COMISSAO_DIV'] - \
                                               df_agrupado.at[i, 'FRETE_DIV']

# === MERGE COM CUSTOS ===
df = pd.merge(df_agrupado, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')

# === CÁLCULOS FINAIS ===
df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
df['IMPOSTO'] = df['TOTAL_ITEM'] * 0.09
df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

# === ARREDONDAMENTO ===
colunas_monetarias = ['PRECO_UNIT', 'TOTAL_ITEM', 'COMISSAO_DIV', 'FRETE_DIV',
                      'VALOR_RECEBIDO', 'CUSTO_TOTAL', 'IMPOSTO', 'LUCRO']
df[colunas_monetarias] = df[colunas_monetarias].round(2)

# === PLANILHA FINAL ===
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

# === ELIMINA DUPLICATAS POR NF + SKU ===
df_saida = df_saida.drop_duplicates(subset=['NF', 'SKU'], keep='first')

# === SALVA COM NOME ÚNICO ===
nome_arquivo = f"RELATORIO_FINAL_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
saida_path = os.path.join(os.path.dirname(arquivo_bling), nome_arquivo)
df_saida.to_excel(saida_path, index=False)
print(f"\n✅ Relatório gerado com sucesso: {saida_path}")

# === GERA RELATÓRIO ANALÍTICO VISUAL EM PDF ===
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import seaborn as sns

with PdfPages(os.path.join(os.path.dirname(arquivo_bling), "RELATORIO_ANALITICO_FINAL.pdf")) as pdf:
    cor_primaria = "#1f77b4"
    fonte_titulo = {'fontsize': 18, 'fontweight': 'bold'}
    fonte_sub = {'fontsize': 12}

    # Página 1 - Resumo geral
    fig1, ax1 = plt.subplots(figsize=(8.5, 11))
    ax1.axis('off')
    ax1.set_title('Resumo Executivo de Vendas', **fonte_titulo, pad=30)

    resumo = f"""
    Data do Relatório: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

    Total de Produtos Vendidos: {int(df_saida['Quantidade'].sum())}
    Valor Total Recebido: R$ {df_saida['Valor Recebido'].sum():,.2f}
    Lucro Total: R$ {df_saida['Lucro'].sum():,.2f}
    Total de Notas Fiscais Emitidas: {df_saida['NF'].nunique()}
    Produtos distintos vendidos: {df_saida['SKU'].nunique()}
    """

    ax1.text(0.05, 0.95, resumo, fontsize=12, va='top')
    ax1.text(0.05, 0.4, "Destaques:", fontsize=13, fontweight='bold')
    top_lucro = df_saida.loc[df_saida['Lucro'].idxmax()]
    ax1.text(0.05, 0.36,
             f"• Produto com maior lucro: {top_lucro['SKU']} - R$ {top_lucro['Lucro']:.2f}", fontsize=11)
    pdf.savefig(fig1)
    plt.close()

    # Página 2 - Lucro por SKU
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    lucro_por_sku = df_saida.groupby('SKU')['Lucro'].sum().sort_values(ascending=True)
    lucro_por_sku.plot(kind='barh', ax=ax2, color=cor_primaria)
    ax2.set_title('Lucro Total por SKU', **fonte_sub)
    ax2.set_xlabel('Lucro (R$)')
    ax2.set_ylabel('SKU')
    ax2.grid(axis='x', linestyle='--', alpha=0.7)
    pdf.savefig(fig2)
    plt.close()

    # Página 3 - Quantidade por SKU
    fig3, ax3 = plt.subplots(figsize=(10, 6))
    qtd_por_sku = df_saida.groupby('SKU')['Quantidade'].sum().sort_values(ascending=True)
    qtd_por_sku.plot(kind='barh', ax=ax3, color="#2ca02c")
    ax3.set_title('Quantidade Vendida por SKU', **fonte_sub)
    ax3.set_xlabel('Unidades')
    ax3.set_ylabel('SKU')
    ax3.grid(axis='x', linestyle='--', alpha=0.7)
    pdf.savefig(fig3)
    plt.close()

    # Página 4 - Lucro por data
    df_saida['Data da Venda'] = pd.to_datetime(df_saida['Data da Venda'], errors='coerce')
    lucro_por_dia = df_saida.groupby('Data da Venda')['Lucro'].sum().sort_index()
    fig4, ax4 = plt.subplots(figsize=(10, 5))
    lucro_por_dia.plot(kind='line', ax=ax4, marker='o', linestyle='-', color="#d62728")
    ax4.set_title('Lucro Total por Data de Venda', **fonte_sub)
    ax4.set_ylabel('Lucro (R$)')
    ax4.set_xlabel('Data')
    ax4.grid(True, linestyle='--', alpha=0.6)
    ax4.tick_params(axis='x', rotation=45)
    pdf.savefig(fig4)
    plt.close()

print("✅ PDF analítico gerado junto ao relatório Excel.")
