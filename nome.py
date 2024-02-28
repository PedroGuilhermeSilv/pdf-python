import pandas as pd
df = pd.read_excel('db.xlsx')
nomes=[]


for index, row in df.iterrows():
    textos = row.to_dict()
    superior_imediato = row["setor"]    
    if superior_imediato not in nomes:
        nomes.append(superior_imediato)

with open('setor.txt', 'w') as file:
    for nome in nomes:
        file.write(nome + '\n')

