import pandas as pd

df = pd.read_excel("mercadin.xlsx")  # inserindo o excel que vamos ler

df = df[["Gender", "Product line", "Total"]]  # separando as colunas que queremos pegar as informações
# caso queira todas não precisa colocar em um dataframe só seguir

pivot_table = df.pivot_table(index="Gender", columns="Product line", values="Total", aggfunc="sum")
# separando como sendo as linhas o genero, as colunas o produto, e os valores dentro da tabela.

pivot_table.to_excel("tabela_produtos.xlsx", "Report", startrow=4)  # exportando como arquivo excel e a linha em que ele começa

