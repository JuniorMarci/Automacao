import pandas as pd
from array import array as arry

# Configurações para os dados do relatório de Valores das Metas
df = pd.read_excel("C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/ValoresDasMetas.xlsx")
df['sh'] = ""
Categorias = df[['Área','Código da Área']].apply(pd.Series.unique)
df['CódigoÚnico'] = "C"+df['Código da Área']+"::"+df['Código da Meta']+"::"+df['Data de referência']

# Configurações para os dados do relatório das Notas Individuais
dne = pd.read_excel("C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/NotasEmpreendimento.xlsx",dtype={'mês':'str'})
dnp = pd.read_excel("C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/NotasPresidencia.xlsx",dtype={'mês':'str'})
dn = pd.concat([dne, dnp], axis=0, ignore_index=True)
dn['CódigoÚnico'] = "C"+dn['Código da Área']+"::"+dn['Código da Meta']+"::2023/"+dn['mês']

# Configurações para os dados do relatório das Notas Consolidadas
nge = pd.read_excel("C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/NG_Empreendimento.xlsx")
ngp = pd.read_excel("C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/NG_Presidencia.xlsx")
ng = pd.concat([nge, ngp], axis=0, ignore_index=True)

# Carregando a tabela de informações dos shoppings
sh = pd.read_excel('C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Shoppings.xlsx')

tabCat = [] # Criada variável da nova tabela de categorias

def Separar(texto):
    nt = []
    sep = [' ','/','(',')']
    for c in sep:
        texto = texto.replace(c,'-')
    return texto

for c in range(len(df['CódigoÚnico'])): # Identificação dos códigos únicos
    df['sh'][c] = [] # Criação da base da coluna

    for d in (Separar(df['Área'][c].replace("Partage ADM","Partage_ADM")).split('-')): # Iteração nas áreas já formatadas
        if (d in sh['CS'].values): # Comparação com cada código de shopping
            df['sh'][c].append(d) # Acréscimo na coluna. Tentativa de fazer tudo numa só tabela
            tabCat.append([df['CódigoÚnico'][c],d]) # Alimentação da tabela de categorias






pd.DataFrame(tabCat).to_csv('Categorias.csv',sep=';',index=False)


dn.to_csv('Notas.csv',sep=';')
ng.to_csv('NotasConsolidadas.csv',sep=';')
df.to_csv('Dados.csv',sep=';')
print("Fim!")


