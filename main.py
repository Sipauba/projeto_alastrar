import pandas as pd
from tkinter import messagebox
import os
import hashlib


executa = True

# Caminho do arquivo no disco C:
caminho_arquivo = r"C:\Automação_Sipauba\Controle_de_processamento.txt"

# Verifica se o arquivo já existe
if os.path.exists(caminho_arquivo):
    print("O arquivo já existe!")
else:
    os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
    with open(caminho_arquivo, "w") as arquivo:
        arquivo.write("")

try:
  referencia = pd.read_excel('./ref.xlsx')
except FileNotFoundError:
  messagebox.showerror('ERRO!!!', 'Arquivo de referência (ref.xlsx) não encontrado!')

try:
  produtos = pd.read_excel('./lista.xlsx')
except FileNotFoundError:
  messagebox.showerror('ERRO!!!', 'Arquivo a ser processado (lista.xlsx) não encontrado!')

hash_atual = str(hashlib.md5(pd.util.hash_pandas_object(produtos).values.tobytes()).hexdigest())
print('atual: ',hash_atual)

with open(caminho_arquivo, 'r') as arquivo:
  hashes = arquivo.read().splitlines() 
  print(hashes)
  if hash_atual in hashes:
    messagebox.showwarning('ATENÇÃO!!!','Arquivo já foi processado.')
    executa = False

lista_produtos = referencia['codigo'].tolist()

if executa == True:
  try:
    for i in lista_produtos:
      fator = int(referencia['fator'].loc[referencia['codigo'] == i].values[0])
      produtos.loc[produtos['codigo'] == i, 'fator'] = fator
      produtos.loc[produtos['codigo'] == i, 'pr_custo_mktplace'] = produtos['fator'] * produtos['pr_custo_mktplace']
      produtos.loc[produtos['codigo'] == i, 'pr_atacado'] = produtos['fator'] * produtos['pr_atacado']
      produtos.loc[produtos['codigo'] == i, 'pr_varejo'] = produtos['fator'] * produtos['pr_varejo']

    produtos.to_excel('lista.xlsx', index=False)
    produtos_alterados = pd.read_excel('lista.xlsx')
    hash_novo = hashlib.md5(pd.util.hash_pandas_object(produtos_alterados).values.tobytes()).hexdigest()
    print('Novo: ', hash_novo)
    with open(caminho_arquivo, "a") as arquivo:
          arquivo.write(f"{hash_novo}\n")

  except Exception as e:
    messagebox.showerror('ERRO!!!', f'Ocorreu um erro do tipo {type(e).__name__}: {e}')
  
  messagebox.showinfo('Concluído', 'Arquivo processado!')

print(executa)
