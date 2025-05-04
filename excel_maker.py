import argparse
from src.config_reader import ConfigReader
from src.table_generator import TableGenerator
import ui, src



def main():
    parser = argparse.ArgumentParser(description='Gerador de relatórios Excel')
    parser.add_argument('--config', default='config/settings.json', help='Caminho do arquivo de configuração')
    parser.add_argument('--input', help='Caminho do arquivo Excel de entrada (sobrescreve config)')
    parser.add_argument('--output', help='Caminho do arquivo de saída (sobrescreve config)')
    args = parser.parse_args()

    # Carregar configurações
    config = ConfigReader(args.config)

    # Sobrescrever com argumentos da linha de comando, se fornecidos
    input_file = args.input or config.get_input_file()
    output_file = args.output or config.get_output_file()

    # Gerar relatório
    generator = TableGenerator(input_file, config)
    generator.generate_tables()
    generator.save_workbook(output_file)

    print(f"Relatório gerado com sucesso: {output_file}")

if __name__ == "__main__":
    main()