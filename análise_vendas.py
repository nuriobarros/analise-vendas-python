# 1-Importação e exploração

import pandas as pd

xls = pd.ExcelFile("Exercicios.xlsx")

print(xls.sheet_names)

df_clientes = pd.read_excel("Exercicios.xlsx", sheet_name= "Clientes")
print(df_clientes.head())

df_vendas = pd.read_excel("Exercicios.xlsx", sheet_name= "Vendas")
print(df_vendas.head())

df_produto = pd.read_excel("Exercicios.xlsx", sheet_name= "Produto")
print(df_produto.head())

df_item = pd.read_excel("Exercicios.xlsx", sheet_name= "Item")
print(df_item.head())

df_pagamentos = pd.read_excel("Exercicios.xlsx", sheet_name= "Pagamentos")
print(df_pagamentos.head())

# 2-Limpeza

print(df_clientes.isnull().sum())
df_clientes = df_clientes.dropna(how="all")

print(df_clientes.duplicated().count())
df = df_clientes.drop_duplicates()

df_clientes["DataCadastro"] = pd.to_datetime(df_clientes["DataCadastro"], errors = "coerce")
df_vendas["DataVenda"] = pd.to_datetime(df_vendas["DataVenda"], errors = "coerce")
df_pagamentos["DataPagamento"] = pd.to_datetime(df_pagamentos["DataPagamento"], errors = "coerce")
df_vendas["Total"] = pd.to_numeric(df_vendas["Total"], errors= "coerce")

# 3-KPIs 

print("Total Vendido:", df_vendas["Total"].sum())
print("Quantidade Vendida:", df_item["Quantidade"].sum())
print("Receita Total:", (df_item["Quantidade"]*df_item["PreçoUnitario"]).sum())
print("Ticket Médio:", df_vendas["Total"].mean())
preco_medio = (df_item["Quantidade"]*df_item["PreçoUnitario"]).sum()/df_item["Quantidade"].sum()
print("Preço Médio Ponderado:", preco_medio)

# 4-Agrupamentos:

# 4.1-Vendas por Cliente

vendas_clientes = df_vendas.groupby("ClienteID")["Total"].sum().sort_values(ascending = False)
print("Vendas por Cliente:", vendas_clientes)

# 4.2-Vendas por Produto 

df_item["ValorTotalLinha"] = df_item["Quantidade"]*df_item["PreçoUnitario"]
vendas_produto = df_item.groupby("ProdutoID")["ValorTotalLinha"].sum().sort_values(ascending = False)
print("Vendas por Produto:", vendas_produto)

# 4.3-Vendas por Mês

df_vendas["AnoMes"] = df_vendas["DataVenda"].dt.to_period("M")
vendas_mes = df_vendas.groupby("AnoMes")["Total"].sum()
print("Vendas por Mês:", vendas_mes)

# 5-Visualização

import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 5.1-Gráfico Evolução das Vendas por Ano-Mês

plt.figure(figsize = (10,5))
sns.lineplot(x=vendas_mes.index.astype(str), y=vendas_mes.values)
plt.xticks(rotation=45)
plt.title("Evolução das Vendas por Ano-Mês")
plt.show()

# 5.2-Gráfico Top 10 Clientes

vendas_clientes = df_vendas.groupby("ClienteID")["Total"].sum()
df_vendas_clientes = vendas_clientes.reset_index().merge(df_clientes, on="ClienteID")
df_vendas_clientes = df_vendas_clientes.sort_values("Total", ascending=False)

top_clientes = df_vendas_clientes.head(10).iloc[::-1]
cores = plt.cm.Blues(np.linspace(0.4, 1, len(top_clientes)))

plt.figure(figsize=(10,5))
bars = plt.barh(top_clientes["Nome"], top_clientes["Total"], color = cores)
plt.xlabel("Total Compras")
plt.ylabel("ClienteID")
plt.title("Top 10 Clientes")
for bar in bars:
    width = bar.get_width()  
    plt.text(width + (width*0.01),             
             bar.get_y() + bar.get_height()/2, 
             f"{width:,.2f}",               
             va="center")    
plt.show()

# 5.3-Gráfico Distribuição por Forma de Pagamento

pagamentos = df_pagamentos.groupby("FormaPagamento")["ValorPago"].sum()

fig, ax = plt.subplots(figsize=(6,6))
wedges, texts, autotexts = ax.pie(
    pagamentos,
    labels=pagamentos.index,
    autopct="%.1f%%",
    startangle=90,
    pctdistance=0.8
)

centre_circle = plt.Circle((0,0), 0.70, fc='white')
fig.gca().add_artist(centre_circle)

plt.title("Distribuição por Forma de Pagamento")
plt.tight_layout()
plt.show()

# 6-Insights

# 6.1-Clientes que compraram pouco

cliente_total = df_vendas.groupby("ClienteID")["Total"].sum().reset_index()
cliente_total = cliente_total.sort_values("Total", ascending=True)
print("Top 10 clientes que compraram menos:", cliente_total.head(10))

# 6.2-Sazonalidade por mês

df_vendas["DataVenda"] = pd.to_datetime(df_vendas["DataVenda"], errors="coerce")
df_vendas["Mes"] = df_vendas["DataVenda"].dt.to_period("M")
vendas_mes_total = df_vendas.groupby("Mes")["Total"].sum().reset_index()
print(vendas_mes_total)

plt.bar(vendas_mes_total["Mes"].astype(str), vendas_mes_total["Total"])
plt.title("Sazonalidades por Mês")
plt.xlabel("Mês")
plt.ylabel("Total de Vendas")
plt.xticks(rotation=45)
plt.show()

# 6.3-% do top 10 clientes

clientes_total = df_vendas.groupby("ClienteID")["Total"].sum().reset_index()
clientes_total = clientes_total.sort_values("Total", ascending=False)
top10 = clientes_total.head(10)
percent_top10 = top10["Total"].sum() / clientes_total["Total"].sum() * 100
print(f"Os 10 maiores clientes representam {percent_top10:.2f}% do total das vendas.")


