import pandas as pd

# Nome do arquivo original
nome_planilha = "Cópia de PLANILHA DE MONITORAMENTO - ALTA COMPLEXIDADE - EIXO ADULTO - ALBERGUE MARTIN LUTHER KING JR - 29 de dezembro de 2023, 11_18 (2).xlsx"

# Carregar a planilha original ignorando as primeiras 5 linhas
ppa = pd.read_excel(nome_planilha, skiprows=5)

# Remover duplicados da coluna de nomes
nomes_unicos = ppa['NOME COMPLETO'].drop_duplicates()

# Criar nova planilha com os nomes únicos
nova_planilha = pd.DataFrame({'NOMES UNICOS': nomes_unicos})

# Função para obter datas únicas e contar a quantidade
def obter_datas_e_quantidade(nome):
    datas = ppa.loc[ppa['NOME COMPLETO'] == nome, 'DATA DO ACOLHIMENTO ATUAL DD/MM/AAAA'].drop_duplicates()
    datas = datas.dropna().astype(str)
    
    # Converter as datas para o formato DD/MM/AAAA
    datas_formatadas = pd.to_datetime(datas, errors='coerce').dt.strftime('%d/%m/%Y').dropna()
    
    return datas_formatadas, len(datas_formatadas)

# Aplicar a função para cada nome
nova_planilha['DATAS DE ACOLHIMENTO'], nova_planilha['QUANTIDADE DE DATAS'] = zip(
    *nova_planilha['NOMES UNICOS'].apply(obter_datas_e_quantidade)
)

# Transformar a lista de datas em string para salvar no Excel
nova_planilha['DATAS DE ACOLHIMENTO'] = nova_planilha['DATAS DE ACOLHIMENTO'].apply(lambda x: ', '.join(x))

# Remover linhas com nomes vazios ou células nulas
nova_planilha = nova_planilha.dropna()

# Salvar a nova planilha no Excel
novo_arquivo = 'nomes_unicos_com_datas_e_quantidade.xlsx'
nova_planilha.to_excel(novo_arquivo, index=False)

print(f"Planilha com nomes únicos, datas de acolhimento formatadas e quantidade salva em: {novo_arquivo}")
