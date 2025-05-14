import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from unidecode import unidecode

# === INTERFACE ===
root = tk.Tk()
root.withdraw()

arquivo_bling = filedialog.askopenfilename(title="Selecione o arquivo CSV exportado do Bling")
arquivo_custos = filedialog.askopenfilename(title="Selecione a planilha de Custos Finais (.xls, .xlsx)")

# === LEITURA DOS DADOS ===
df_bling = pd.read_csv(arquivo_bling, sep=';', encoding='latin1', dtype=str)

# === MOSTRA AS COLUNAS ORIGINAIS ===
print("\n" + "="*60)
print("🔎 Colunas encontradas no arquivo do Bling:\n")
for i, col in enumerate(df_bling.columns, start=1):
    print(f"{i}. '{col}'")
print("="*60 + "\n")

# === LEITURA PLANILHA DE CUSTOS ===
xls = pd.ExcelFile(arquivo_custos)
df_custo_all = pd.concat([xls.parse(aba) for aba in xls.sheet_names], ignore_index=True)

# === PADRONIZAÇÃO NOMES DE COLUNAS ===
df_bling.columns = df_bling.columns.str.strip()
df_custo_all.columns = df_custo_all.columns.str.strip()

# === MAPEAMENTO FLEXÍVEL DAS COLUNAS DO BLING ===
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

# === RENOMEIA AUTOMATICAMENTE COM UNIDECODE ===
print("\n🔁 Tentando renomear colunas do Bling:\n")
for original in df_bling.columns:
    col_simplificada = unidecode(original.upper()).strip()
    print(f"🔍 Analisando: '{original}' → '{col_simplificada}'")
    for chave, novo_nome in mapeamento.items():
        if chave in col_simplificada or (chave == 'PRECO' and 'UNIT' in col_simplificada):
            print(f"✔️ Renomeando '{original}' → '{novo_nome}' (via '{chave}')")
            df_bling.rename(columns={original: novo_nome}, inplace=True)
            break

# === RENOMEAÇÃO DE EMERGÊNCIA: NF (força total)
if 'NF' not in df_bling.columns:
    for col in df_bling.columns:
        col_limpo = unidecode(col.upper()).strip()
        if 'NUMERO' in col_limpo or col.strip().upper() == 'NÚMERO' or col.strip() == 'NÃºmero':
            print(f"⚠️ Renomeação forçada: '{col}' → 'NF'")
            df_bling.rename(columns={col: 'NF'}, inplace=True)
            break

# === RENOMEIA PLANILHA DE CUSTO ===
print("\n🧾 Renomeando colunas da planilha de custos:")
for original in df_custo_all.columns:
    col_simplificada = unidecode(original.upper()).strip()
    print(f"🔍 Analisando: '{original}' → '{col_simplificada}'")
    if 'SKU' in col_simplificada or 'CODIGO' in col_simplificada:
        print(f"✔️ Renomeando '{original}' → 'SKU'")
        df_custo_all.rename(columns={original: 'SKU'}, inplace=True)
    elif 'CUSTO' in col_simplificada:
        print(f"✔️ Renomeando '{original}' → 'CUSTO'")
        df_custo_all.rename(columns={original: 'CUSTO'}, inplace=True)

# === VERIFICAÇÃO DAS COLUNAS ===
print("\n📋 Colunas finais do Bling após renomeação:")
print(df_bling.columns.tolist())

colunas_necessarias = ['SKU', 'DATA_VENDA', 'NF', 'PRECO_UNIT', 'COMISSAO', 'FRETE', 'QUANTIDADE']
for col in colunas_necessarias:
    if col not in df_bling.columns:
        raise Exception(f"❌ Coluna obrigatória não encontrada no arquivo do Bling: {col}")

# === CONVERSÃO E LIMPEZA ===
df_bling['PRECO_UNIT'] = df_bling['PRECO_UNIT'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['COMISSAO'] = df_bling['COMISSAO'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['FRETE'] = df_bling['FRETE'].str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)
df_bling['QUANTIDADE'] = df_bling['QUANTIDADE'].str.replace(',', '.').astype(float)
df_bling['SKU'] = df_bling['SKU'].astype(str).str.strip()
df_bling['NF'] = df_bling['NF'].astype(str).str.strip()

df_custo_all['SKU'] = df_custo_all['SKU'].astype(str).str.strip()
df_custo_all['CUSTO'] = df_custo_all['CUSTO'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.').astype(float)

# === CÁLCULO TOTAL E RATEIO POR NF ===
df_bling['TOTAL_ITEM'] = df_bling['PRECO_UNIT'] * df_bling['QUANTIDADE']
df_bling['FRETE_DIV'] = 0.0
df_bling['COMISSAO_DIV'] = 0.0

for nf, grupo in df_bling.groupby('NF'):
    soma_total = grupo['TOTAL_ITEM'].sum()
    for i, row in grupo.iterrows():
        proporcao = row['TOTAL_ITEM'] / soma_total if soma_total != 0 else 0
        df_bling.at[i, 'FRETE_DIV'] = row['FRETE'] * proporcao
        df_bling.at[i, 'COMISSAO_DIV'] = row['COMISSAO'] * proporcao

df_bling['VALOR_RECEBIDO'] = df_bling['TOTAL_ITEM'] - df_bling['FRETE_DIV'] - df_bling['COMISSAO_DIV']

# === MERGE COM PLANILHA DE CUSTO ===
df = pd.merge(df_bling, df_custo_all[['SKU', 'CUSTO']], on='SKU', how='left')

# === CÁLCULOS FINAIS ===
df['CUSTO_TOTAL'] = df['CUSTO'] * df['QUANTIDADE']
df['IMPOSTO'] = df['VALOR_RECEBIDO'] * 0.09
df['LUCRO'] = df['VALOR_RECEBIDO'] - df['CUSTO_TOTAL'] - df['IMPOSTO']

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

# === SALVA PLANILHA ===
saida_path = os.path.join(os.path.dirname(arquivo_bling), 'RELATORIO_FINAL.xlsx')
df_saida.to_excel(saida_path, index=False)
print(f"\n✅ Relatório gerado com sucesso: {saida_path}")
