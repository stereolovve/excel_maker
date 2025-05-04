## Documentação do Projeto Interface_Excel
# Visão Geral
Este projeto consiste em uma aplicação para criação e gerenciamento de planilhas de contagem, com interface gráfica desenvolvida em Flet. A aplicação permite inserir dados de contagem, incluindo informações como ponto, data, localização e movimentos, e gera automaticamente uma planilha Excel formatada.

Estrutura do Projeto
Interface_Excel/
├── main.py            # Ponto de entrada da aplicação
├── src/
│   └── planilha.py    # Lógica de criação e formatação da planilha
└── ui/
    └── entrada.py     # Interface de usuário para entrada de dados
# Componentes Principais
1. Interface de Usuário (ui/entrada.py)
   - A classe DataEntryForm implementa um formulário para entrada de dados com os seguintes campos:
     - Ponto
     - Data
     - Localização
     - Número de Movimentos (gera campos dinâmicos)
     - Duração em dias e horas
     - Período (início e fim)
     - Campos dinâmicos para movimentos

2. Geração de Planilhas (src/planilha.py)
   - O módulo planilha.py é responsável por criar e manipular planilhas Excel com base nos dados fornecidos:
     - Leitura de dados de entrada
     - Criação de tabelas duplicadas para cada movimento
     - Configuração de estilos e formatação
     - Salvamento do arquivo final

3. Aplicação Principal (main.py)
Integra os componentes e inicia a aplicação:

Configura a janela da aplicação
Instancia o formulário de entrada
Gerencia o fluxo de dados do formulário para a planilha
Requisitos
Python 3.7+
Flet
openpyxl
### Instalação

bash
Copy Code
# Clone o repositório
git clone https://github.com/seu-usuario/Interface_Excel.git
cd Interface_Excel

# Instale as dependências
pip install -r requirements.txt
Uso
Execute o arquivo principal:

bash
Copy Code
python main.py
Preencha os campos do formulário
Defina o número de movimentos para gerar campos adicionais
Clique em "Salvar" para gerar a planilha Excel
Fluxo de Dados
O usuário insere dados no formulário
Ao salvar, os dados são coletados em um dicionário
O dicionário é passado para a classe planilhaContagem
A planilha é gerada com duas abas (Entrada e Resumo)
O arquivo Excel é salvo no diretório atual