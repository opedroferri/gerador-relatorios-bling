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
            print(f"⚠️ Renomeando coluna '{col}' → 'NF'")
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            break

if 'NF' not in df_bling.columns:
    for col in df_bling.columns:
        if col.strip() == 'NÃºmero':
            print(f"⚠️ Renomeando coluna corrompida '{col}' → 'NF'")
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            break

if 'NF' not in df_bling.columns:
    print("\n❌ Colunas atuais:")
    print(df_bling.columns.tolist())
    raise Exception("❌ Mesmo após a tentativa de renomeação, a coluna 'NF' não foi encontrada no arquivo do Bling.")

# === RENOMEIA PLANILHA DE CUSTO ===
df_custo_all.columns = df_custo_all.columns.str.strip()
for original in df_custo_all.columns:
    col_simplificada = unidecode(original.upper()).strip()
    if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
        df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
    elif 'CUSTO' in col_simplificada:
        df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

# === VERIFICAÇÃO DAS COLUNAS OBRIGATÓRIAS ===
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

df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip().str.upper()
df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

# === CÁLCULO DE TOTAL ITEM ===
df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']
df_bling['FRETE_DIV'] = 0.0
df_bling['COMISSAO_DIV'] = 0.0
df_bling['VALOR_RECEBIDO'] = 0.0

# === RATEIO PROPORCIONAL COM BASE NO TOTAL DA NF (corrigido) ===
for nf, grupo in df_bling.groupby('NF'):
    soma_total_nf = grupo['TOTAL_ITEM'].sum()
    valor_comissao_nf = grupo['COMISSAO'].astype(float).iloc[0]
    valor_frete_nf = grupo['FRETE'].astype(float).iloc[0]

    for i in grupo.index:
        proporcao = df_bling.at[i, 'TOTAL_ITEM'] / soma_total_nf if soma_total_nf else 0
        df_bling.at[i, 'COMISSAO_DIV'] = proporcao * valor_comissao_nf
        df_bling.at[i, 'FRETE_DIV'] = proporcao * valor_frete_nf
        df_bling.at[i, 'VALOR_RECEBIDO'] = df_bling.at[i, 'TOTAL_ITEM'] - df_bling.at[i, 'COMISSAO_DIV'] - df_bling.at[i, 'FRETE_DIV']

# === MERGE COM PLANILHA DE CUSTOS ===
df = pd.merge(df_bling, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')

# === CÁLCULOS FINAIS ===
df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
df['IMPOSTO'] = df['VALOR_RECEBIDO'] * 0.09
df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

# === ARREDONDAMENTO DE VALORES MONETÁRIOS ===
colunas_monetarias = ['PRECO_UNIT', 'TOTAL_ITEM', 'COMISSAO_DIV', 'FRETE_DIV',
                      'VALOR_RECEBIDO', 'CUSTO_TOTAL', 'IMPOSTO', 'LUCRO']
df[colunas_monetarias] = df[colunas_monetarias].round(2)

# === PLANILHA FINAL ===
df_saida = pd.DataFrame({
    'Data da Venda': pd.to_datetime(df['DATA_VENDA'], errors='coerce').dt.strftime('%d/%m/%Y'),
    'NF': df['NF'],
    'SKU': df['SKU'],
    'Descrição do Produto': df.get('DESC_PRODUTO', ''),
    'Quantidade': df['QUANTIDADE'],
    'Preço Unitário': df['PRECO_UNIT'],
    'Preço Total': df['TOTAL_ITEM'],
    'Valor Recebido': df['VALOR_RECEBIDO'],
    'Custo': df['CUSTO_TOTAL'],
    'Imposto (9%)': df['IMPOSTO'],
    'Lucro': df['LUCRO']
})

# === SALVAR COM NOME ÚNICO ===
nome_arquivo = f"RELATORIO_FINAL_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
saida_path = os.path.join(os.path.dirname(arquivo_bling), nome_arquivo)
df_saida.to_excel(saida_path, index=False)
print(f"\n✅ Relatório gerado com sucesso: {saida_path}")
