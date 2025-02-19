import pandas as pd
from tkinter import messagebox

referencia = pd.read_excel('./ref.xlsx')
produtos = pd.read_excel('./lista.xlsx')

lista_produtos = referencia['codigo'].tolist()

for i in lista_produtos:
  fator = int(referencia['fator'].loc[referencia['codigo'] == i].values[0])
  produtos.loc[produtos['codigo'] == i, 'fator'] = fator
  produtos.loc[produtos['codigo'] == i, 'pr_custo_mktplace'] = produtos['fator'] * produtos['pr_custo_mktplace']
  produtos.loc[produtos['codigo'] == i, 'pr_atacado'] = produtos['fator'] * produtos['pr_atacado']
  produtos.loc[produtos['codigo'] == i, 'pr_varejo'] = produtos['fator'] * produtos['pr_varejo']

produtos.to_excel('lista.xlsx', index=False)

messagebox.showinfo('Concluído', 'Arquivo processado!')