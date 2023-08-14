import pandas as pd
from array import array as arry

df = pd.read_excel("C:/Users/marcio.souza/OneDrive - Partage/Padronização/DataExtract01932.xlsx")
df['sh'] = ""
Categorias = df[['Área','Código da Área']].apply(pd.Series.unique)

sh = pd.read_excel('C:/Users/marcio.souza/OneDrive - Partage/Padronização/Shoppings.xlsx')
tabCat = {}

for c in sh['CS']: tabCat[c]=arry('i',[0]*len(Categorias['Área']))
tabCat['Código da Área']=Categorias['Código da Área'].values

def Separar(texto):
    nt = []
    sep = [' ','/','(',')']
    for c in sep:
        texto = texto.replace(c,'-')
    return texto

for c in range(len(Categorias['Área'])):
    df['sh'][c] = []
    tabCat['Código da Área'][c] = "C"+Categorias['Código da Área'][c]
    for d in (Separar(Categorias['Área'][c].replace("Partage ADM","Partage_ADM")).split('-')):
        if (d in sh['CS'].values):
            df['sh'][c].append(d)
            tabCat[d][c] = 1


tabCat['Código da Área'] = tabCat['Código da Área']
pd.DataFrame(tabCat).to_csv('Categorias.csv',sep=';',index=False)

df['Código da Área'] = "C"+df['Código da Área']
df.to_csv('Dados.csv',sep=';')
print("Fim!")


