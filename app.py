import pandas as pd

# Função para ler a planilha e retornar uma lista de matrículas
def extrair_matriculas(planilha):
    # Leitura da planilha
    df = pd.read_excel(planilha)
    
    # Verificação se a coluna "Matrícula" existe na planilha
    if 'Matrícula' not in df.columns:
        raise ValueError("A coluna 'Matrícula' não foi encontrada na planilha.")
    
    # Extrair matrículas
    matriculas = df['Matrícula'].astype(str).tolist()
    
    return matriculas

# Função para criar a string SQL com as matrículas
def criar_sql_delete(matriculas):
    if not matriculas:
        raise ValueError("Nenhuma matrícula encontrada na planilha.")
    
    # Criação da string SQL
    sql = "delete from esocial.s2299 where matricula in ({})".format(','.join(f"'{matricula}'" for matricula in matriculas))
    
    return sql

# Caminho para a planilha
caminho_planilha = "C:/Users/itarg/Downloads/DEMITIDOS.xlsx"

try:
    # Extrair matrículas da planilha
    matriculas = extrair_matriculas(caminho_planilha)
    
    # Criar a string SQL
    sql_delete = criar_sql_delete(matriculas)
    
    print("Comando SQL gerado:")
    print(sql_delete)

except Exception as e:
    print("Ocorreu um erro:", e)
