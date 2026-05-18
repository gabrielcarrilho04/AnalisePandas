import pandas as pd

# ==================================================
# IMPORTAÇÃO DA PLANILHA
# ==================================================

df = pd.read_excel(
    r"C:\EstudosPython\PandasToGiftHub\projeto_analise_vendas.xlsx"
)

# ==================================================
# EXPLORAÇÃO DOS DADOS
# ==================================================

print("========== PRIMEIRAS 5 LINHAS ==========")
print(df.head())

print("\n========== INFORMAÇÕES DO DATAFRAME ==========")
print(df.info())

print("\n========== TAMANHO DO DATAFRAME ==========")
print(df.shape)

print("\n========== VALORES NULOS ==========")
print(df.isnull().sum())

# ==================================================
# LIMPEZA DOS DADOS
# ==================================================

# ------------------------
# CLIENTE
# ------------------------

df["cliente"] = (
    df["cliente"]
    .str.strip()
    .str.title()
)

df["cliente"] = df["cliente"].fillna(
    "Desconhecido"
)

# ------------------------
# VALOR
# ------------------------

df["valor"] = pd.to_numeric(
    df["valor"],
    errors="coerce"
)

media_valor = df["valor"].mean()

df["valor"] = df["valor"].fillna(
    media_valor
)

# ------------------------
# QUANTIDADE
# ------------------------

df["quantidade"] = pd.to_numeric(
    df["quantidade"],
    errors="coerce"
)

df["quantidade"] = df["quantidade"].fillna(1)

# ------------------------
# DATA
# ------------------------

df["data_venda"] = pd.to_datetime(
    df["data_venda"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

# ==================================================
# CRIAÇÃO DE COLUNAS
# ==================================================

df["mes"] = df["data_venda"].dt.month

df["dia"] = df["data_venda"].dt.day

df[["quantidade", "mes", "dia"]] = (
    df[["quantidade", "mes", "dia"]]
    .astype("Int64")
)

# ------------------------
# VALOR TOTAL
# ------------------------

df["valor_total"] = (
    df["valor"] * df["quantidade"]
)

# ==================================================
# ANÁLISES GERAIS
# ==================================================

# ------------------------
# FATURAMENTO TOTAL
# ------------------------

faturamento_total = df["valor_total"].sum()

print("\n========== FATURAMENTO TOTAL ==========")
print(f"R$ {faturamento_total:.2f}")

# ------------------------
# TICKET MÉDIO
# ------------------------

ticket_medio = df["valor_total"].mean()

print("\n========== TICKET MÉDIO ==========")
print(f"R$ {ticket_medio:.2f}")

# ------------------------
# PRODUTO MAIS VENDIDO
# ------------------------

produto_quantidade = (
    df.groupby("produto")["quantidade"]
    .sum()
)

produto_mais_vendido = (
    produto_quantidade.idxmax()
)

print("\n========== PRODUTO MAIS VENDIDO ==========")
print(produto_mais_vendido)

# ------------------------
# PRODUTO COM MAIOR FATURAMENTO
# ------------------------

produto_faturamento = (
    df.groupby("produto")["valor_total"]
    .sum()
)

produto_mais_rentavel = (
    produto_faturamento.idxmax()
)

print("\n========== PRODUTO COM MAIOR FATURAMENTO ==========")
print(produto_mais_rentavel)

# ------------------------
# CLIENTE QUE MAIS COMPROU
# ------------------------

cliente_faturamento = (
    df.groupby("cliente")["valor_total"]
    .sum()
)

cliente_mais_comprou = (
    cliente_faturamento.idxmax()
)

print("\n========== CLIENTE QUE MAIS COMPROU ==========")
print(cliente_mais_comprou)

# ------------------------
# CATEGORIA MAIS RENTÁVEL
# ------------------------

categoria_faturamento = (
    df.groupby("categoria")["valor_total"]
    .sum()
)

categoria_mais_rentavel = (
    categoria_faturamento.idxmax()
)

print("\n========== CATEGORIA MAIS RENTÁVEL ==========")
print(categoria_mais_rentavel)

# ==================================================
# RANKINGS E AGRUPAMENTOS
# ==================================================

# ------------------------
# RANKING DE CLIENTES
# ------------------------

ranking_clientes = (
    cliente_faturamento
    .sort_values(ascending=False)
)

print("\n========== RANKING DE CLIENTES ==========")
print(ranking_clientes)

# ------------------------
# FATURAMENTO POR CATEGORIA
# ------------------------

print("\n========== FATURAMENTO POR CATEGORIA ==========")
print(categoria_faturamento)

# ------------------------
# FATURAMENTO POR MÊS
# ------------------------

mes_faturamento = (
    df.groupby("mes")["valor_total"]
    .sum()
)

print("\n========== FATURAMENTO POR MÊS ==========")
print(mes_faturamento)

# ------------------------
# QUANTIDADE VENDIDA POR PRODUTO
# ------------------------

print("\n========== QUANTIDADE VENDIDA POR PRODUTO ==========")
print(produto_quantidade)

# ==================================================
# FILTROS
# ==================================================

# ------------------------
# VENDAS ACIMA DE 1000
# ------------------------

vendas_1000 = df[
    df["valor_total"] > 1000
]

print("\n========== VENDAS ACIMA DE 1000 ==========")
print(vendas_1000)

# ------------------------
# PRODUTOS ELETRÔNICOS
# ------------------------

eletronicos = df[
    df["categoria"] == "Eletrônicos"
]

print("\n========== PRODUTOS ELETRÔNICOS ==========")
print(eletronicos)

# ------------------------
# CLIENTES ACIMA DA MÉDIA
# ------------------------

clientes_acima_media = df[
    df["valor_total"] > ticket_medio
]["cliente"].unique()

print("\n========== CLIENTES ACIMA DA MÉDIA ==========")
print(clientes_acima_media)

# ==================================================
# IMPORTAÇÃO DA ABA CLIENTES
# ==================================================

df_clientes = pd.read_excel(
    r"C:\EstudosPython\PandasToGiftHub\projeto_analise_vendas.xlsx",
    sheet_name="clientes"
)

# ==================================================
# LIMPEZA DA TABELA CLIENTES
# ==================================================

df_clientes["cliente"] = (
    df_clientes["cliente"]
    .str.strip()
    .str.title()
)

# ==================================================
# MERGE
# ==================================================

df_merge = pd.merge(
    df,
    df_clientes,
    on="cliente",
    how="left"
)

print("\n========== DATAFRAME COM MERGE ==========")
print(df_merge.head())

# ==================================================
# ANÁLISES COM MERGE
# ==================================================

# ------------------------
# FATURAMENTO POR CIDADE
# ------------------------

cidade_faturamento = (
    df_merge.groupby("cidade")["valor_total"]
    .sum()
)

print("\n========== FATURAMENTO POR CIDADE ==========")
print(cidade_faturamento)

# ------------------------
# CIDADE COM MAIOR FATURAMENTO
# ------------------------

cidade_mais_rica = (
    cidade_faturamento.idxmax()
)

print("\n========== CIDADE COM MAIOR FATURAMENTO ==========")
print(cidade_mais_rica)

# ------------------------
# MÉDIA DE IDADE
# ------------------------

media_idade = (
    df_merge["idade"].mean()
)

print("\n========== MÉDIA DE IDADE ==========")
print(f"{media_idade:.1f}")

# ==================================================
# EXPORTAÇÃO
# ==================================================

df_merge.to_excel(
    "resultado_final.xlsx",
    index=False
)

print("\n========== ARQUIVO EXPORTADO ==========")
print("resultado_final.xlsx salvo com sucesso!")
