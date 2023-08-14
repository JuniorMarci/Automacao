import pandas as pd

# Carregando a tabela de informações consolidadas
ConsPres = 'C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Correções (V2)/Consolidada_Presidencia.xlsx'
ConsEmp = 'C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Correções (V2)/Consolidada_Empreendimentos.xlsx'
ConsPres = pd.read_excel(ConsPres)
ConsEmp = pd.read_excel(ConsEmp)

Consolidada = pd.concat([ConsPres, ConsEmp], axis=0, ignore_index=True)
Consolidada['Cod.Unico'] ="C"+Consolidada['Código da Área']+"::"+Consolidada['Login']

# Carregando a tabela de informações separados por metas
MetPres = 'C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Correções (V2)/Meta_Presidencia.xlsx'
MetEmp = 'C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Correções (V2)/Meta_Empreendimentos.xlsx'
MetPres = pd.read_excel(MetPres)
MetEmp = pd.read_excel(MetEmp)

Metas = pd.concat([MetPres, MetEmp], axis=0, ignore_index=True)
Metas['Cod.Unico'] ="C"+Metas['Código da Área']+"::"+Metas['Login']

# Carregando a tabela de informações dos shoppings
sh = pd.read_excel('C:/Users/marcio.souza/Partage/#Sistemas e Qualidade - Documentos/Análise de Metas - BI/Correções (V2)/Shoppings.xlsx')

tabCat = [] # Criada variável da nova tabela de categorias
tabCons = [] # Criada variável dos dados consolidados separados por mês


def Separar(texto):
    nt = []
    sep = [' ','/','(',')']
    for c in sep:
        texto = texto.replace(c,'-')
    return texto

for c in range(len(Consolidada['Cod.Unico'])): # Identificação dos códigos únicos
    #df['sh'][c] = [] # Criação da base da coluna

    for d in (Separar(Consolidada['Descrição da Área'][c].replace("Partage ADM","Partage_ADM")).split('-')): # Iteração nas áreas já formatadas
        if (d in sh['CS'].values): # Comparação com cara código de shopping
            #df['sh'][c].append(d) # Acréscimo na coluna. Tentativa de fazer tudo numa só tabela
            tabCat.append([Consolidada['Cod.Unico'][c],d]) # Alimentação da tabela de categorias

    for m in range(1,9):
        actual = "mês "+str(m)
        tabCons.append([Consolidada['Cod.Unico'][c],m,Consolidada[actual][c]])



pd.DataFrame(tabCons).to_csv('DConsolidada.csv',sep=';',index=False)
pd.DataFrame(tabCat).to_csv('Categorias.csv',sep=';',index=False)
Consolidada.to_csv('Consolidada.csv',sep=';')
Metas.to_csv('Metas.csv',sep=';')
#df.to_csv('Dados.csv',sep=';')

#----------------------------------------------

# Carregando a tabela de Planos de Ações
PlA = r'C:\Users\marcio.souza\Partage\#Sistemas e Qualidade - Documentos\Análise de Metas - BI\Correções (V2)\Planos de Ação\Plano.xlsx'
PlA = pd.read_excel(PlA)
Planos = [] # Criada tabela para correlacionar os meses com os planos
for c in range (len(PlA['Código da Ação'])):
    for d in range (int(PlA['Data início planejada'][c][3:5]),int(PlA['Data fim planejada'][c][3:5])+1):
        Planos.append([PlA['Código da Ação'][c],d])

pd.DataFrame(Planos).to_csv('Planos de Ação\Planos-Mês.csv',sep=';',index=False)

print("Fim!")


