import os
from datetime import datetime
from planilha import planilhaContagem  # Certifique-se de que o nome do seu arquivo é planilha.py

# Função para criar e testar a planilha
def test_planilha():
    # Dados de teste
    test_data = {
        "Cliente": "Rota Sorocabana",
        "Código": "SOR2507",
        "Ponto": "P002",
        "Data": "13-05-2025",
        "Localização": "Rodovia SP-123",
        "Num_Movimentos": "2",
        "Duração em dias": "2",
        "Duração em horas": "24",
        "Periodo_Inicio": "08:00",
        "Periodo_Fim": "08:00",
        "Movimentos": ["a", "b"]
    }

    # Criar uma instância da planilha
    planilha = planilhaContagem(codigo=test_data["Código"], ponto=test_data["Ponto"])

    # Adicionar os dados à planilha
    print("Adicionando dados à planilha...")
    planilha.add_data(test_data)

    # Criar o diretório 'output' se não existir
    if not os.path.exists("output"):
        os.makedirs("output")

    # Salvar a planilha
    print("Salvando a planilha...")
    planilha.save()

    # Verificar se o arquivo foi criado
    output_file = f"output/{test_data['Código']}_{test_data['Ponto']}.xlsx"
    if os.path.exists(output_file):
        print(f"Sucesso: Planilha '{output_file}' foi criada com sucesso!")
    else:
        print(f"Erro: Planilha '{output_file}' não foi criada.")

# Executar o teste
if __name__ == "__main__":
    try:
        test_planilha()
    except Exception as e:
        print(f"Erro durante o teste: {str(e)}")