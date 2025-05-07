#src/planilha.py
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from datetime import datetime, timedelta

# Create a new workbook
class planilhaContagem:
    def __init__(self, codigo="",ponto=""):
        self.filename = f"{codigo}_{ponto}.xlsx"
        self.wb = Workbook()
        self.entrada = self.abaEntrada(self.wb)
        self.relatorio = self.abaRelatorio(self.wb)
        self.hr = self.abaHr(self.wb)

    # Create first sheet
    class abaEntrada:
        def __init__(self, wb):
            self.wb = wb
            self.sheet1 = self.wb.active
            self.sheet1.title = "Entrada"

            # Definir estilos
            self.header_font = Font(bold=True)
            self.title_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            self.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            self.center_align = Alignment(horizontal='left')

        def add_data(self, data=None):
            # Cabeçalhos com formatação
            headers = [
                ('B1', "Ponto"), ('B2', "Data inicial"), ('B3', "Num_Movimentos"),
                ('B4', "Localização"), ('B5', "Duração_dias"), ('B6', "Duração_horas"),
                ('B7', "Hora_início"), ('B8', "Hora_fim"), ('E1', "Movimento")
            ]

            for cell_pos, value in headers:
                cell = self.sheet1[cell_pos]
                cell.value = value
                cell.font = self.header_font
                cell.fill = self.title_fill
                cell.border = self.border
                cell.alignment = self.center_align

            # Preencher de acordo com os dados fornecidos na interface 'data'
            if data:
                # Mapeamento de células e valores
                data_cells = [
                    
                    ('C1', data.get("Ponto", "")),
                    ('C2', data.get("Data", "")),
                    ('C3', data.get("Num_Movimentos")),
                    ('C4', data.get("Localização")),
                    ('C5', data.get("Duração em dias")),
                    ('C6', data.get("Duração em horas")),
                    ('C7', data.get("Periodo_Inicio", "")),
                    ('C8', data.get("Periodo_Fim", ""))
                ]

                # Aplicar valores e formatação
                for cell_pos, value in data_cells:
                    cell = self.sheet1[cell_pos]
                    cell.value = value
                    cell.border = self.border

                # Preencher movimentos dinamicamente com formatação
                movimentos = data.get("Movimentos", [])
                for i, movimento in enumerate(movimentos, start=2):
                    cell = self.sheet1[f'E{i}']
                    cell.value = movimento
                    cell.border = self.border

    # Create second sheet
    class abaRelatorio:
        def __init__(self, wb):
            self.wb = wb
            self.sheet2 = self.wb.create_sheet(title="Relatório")

            # Definir estilos para a aba de resumo
            """
            Criar NamedStyles para melhor organização
            """
            # Criar uma formatação NamedStyle
            self.header_style = NamedStyle(name="header_style")
            self.header_style.alignment = Alignment(horizontal='left', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            self.header_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)
            
        def create_movement_table(self, start_row, data, movement, movement_index):
            """
            Definir cabeçalho da tabela de contagem.
            """
            ponto = data.get("Ponto", "")
            movimento_concatenado = f"{ponto}{movement}" if ponto and movement else movement
            self.sheet2['B1'] = "Data:"
            self.sheet2['C1'] = data.get("Data", "")
            self.sheet2['B2'] = "Movimento:"
            self.sheet2['C2'] = movimento_concatenado
            
            
            for col in  ['B', 'C']:
                for row in range(start_row, start_row + len(movement)):
                    cell = self.sheet2[f'{col}{row}']
                    cell.style = self.header_style
            
            """
            Definir coluna de veiculos
            """
            table_columns = [
                'B3:C3',  # Hora
                'D3:D4',  # Leves
                'E3:G3',  # Carretinha
                'H3:H4',  # VUC
                'I3:K3',  # Caminhões
                'L3:R3',  # Carretas
                'S3:T3',  # Ônibus
                'U3:U4',  # Motos
                'V3:AC3',  # Pesados
                'AD3:AD4'  # Veiculos Totais
            ]
            # Aplicar mesclagem nas áreas definidas
            for header_info in table_columns:
                self.sheet2.merge_cells(header_info)
            
            # Dar nome para as colunas mescladas
            headers = [
                ('B3', "Horas"),
                ('D3', "Leves"),
                ('E3', "Carretinha"),
                ('H3', "VUC"),
                ('I3', "Caminhões"),
                ('L3', "Carreta"),
                ('S3', "Ônibus"),
                ('U3', "Motos"),
                ('V3', "Pesados"),
                ('AD3', "Veículos Totais"),
            ]

            for header_info in headers:
                cell = self.sheet2[header_info[0]]
                cell.value = header_info[1]

            # Estilo para cabeçalhos principais
            header_style = Font(bold=True, size=11)
            header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            self.header_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin'),
                diagonal=Side(style='thin'),

            )
            center_align = Alignment(horizontal='center', vertical='center')

            # Aplicar cabeçalhos principais
            for header_info in table_columns:
                for row in self.sheet2[header_info]:
                    for cell in row:
                        cell.border = self.header_border
                        cell.alignment = center_align

            # Adicionar subcategorias na linha 5
            subcategories = [
                ('B4', "das"), ('C4', "as"),
                ('E4', "1 Eixo"), ('F4', "2 Eixos"), ('G4', "3 Eixos"), 
                ('I4', "2 Eixos"), ('J4', "3 Eixos"), ('K4', "4 Eixos"),
                ('L4', "2 E"), ('M4', "3 E"), ('N4', "4 E"), ('O4', "5 E"), ('P4', "6 E"), ('Q4', "7 E"), ('R4', "8 E"),
                ('S4', "2 E"), ('T4', "3 E ou +"),
                ('V4', "% Cam"), ('w4', "Caminhões"), ('x4', "% Carr"), ('Y4', "Carretas"), ('Z4', "% Ônib"), ('AA4', "Ônibus"), ('AB4', "% Pes"), ('AC4', "Total")
            ]

            # Estilo para subcategorias
            subcat_style = Font(size=10)
            subcat_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            subcat_align = Alignment(horizontal='center', vertical='center')

            # Aplicar subcategorias
            for cell_pos, value in subcategories:
                cell = self.sheet2[cell_pos]
                cell.value = value
                cell.font = subcat_style
                cell.border = subcat_border
                cell.alignment = subcat_align

            
            """
            Definir o laço de 15min das colunas das e as
            """
            # Criar uma formatação NamedStyle
            time_style = NamedStyle(name="time")
            time_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            time_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)

            # Criar laço para preencher a coluna "das" de 15 em 15 minutos até dar 23:59
            das_inicio = datetime.strptime("00:00", "%H:%M")
            as_inicio = datetime.strptime("00:15","%H:%M")
            das_fim = datetime.strptime("23:45", "%H:%M")
            as_fim = datetime.strptime("23:59", "%H:%M")
            cell_das_inicio = self.sheet2['B5']
            cell_as_inicio = self.sheet2['C5']

            while das_inicio <= das_fim and as_fim:
                cell_das_inicio.value = das_inicio.strftime("%H:%M")
                cell_das_inicio.style = time_style
                cell_as_inicio.value = as_inicio.strftime("%H:%M")
                cell_as_inicio.style = time_style
                cell_das_inicio = self.sheet2.cell(row=cell_das_inicio.row + 1, column=cell_das_inicio.column)
                cell_as_inicio = self.sheet2.cell(row=cell_as_inicio.row + 1, column=cell_as_inicio.column)
                das_inicio += timedelta(minutes=15)
                as_inicio += timedelta(minutes=15)
        
            # Criar uma formatação com NamedStyled

            """
            Definir formulas de total de veículos e suas porcentagens.
            """
            start_row = 5
            end_row = 100

            # Loop para iterar entre start e end
            for row in range(start_row, end_row + 1):
                # Formulas de total 
                formula_caminhoes = f"=SUM(I{row}:K{row})"
                formula_carretas = f"=SUM(L{row}:R{row})"
                formula_onibus = f"=SUM(S{row}:T{row})"
                formula_total_pesados = f"=SUM(W{row},Y{row},AA{row})"
                formula_total_vehicles = f"=SUM(D{row}:H{row},AC{row})"

                # Aplicas as formulas em cada celula
                self.sheet2[f'W{row}'].value = formula_caminhoes
                self.sheet2[f'Y{row}'].value = formula_carretas
                self.sheet2[f'AA{row}'].value = formula_onibus
                self.sheet2[f'AC{row}'].value = formula_total_pesados
                self.sheet2[f'AD{row}'].value = formula_total_vehicles

                # Formula de percentual
                formula_perc_caminhoes = f"=IFERROR(W{row}/AD{row}, 0)"
                formula_perc_carretas = f"=IFERROR(Y{row}/AD{row}, 0)"
                formula_perc_onibus = f"=IFERROR(AA{row}/AD{row}, 0)"
                formula_perc_pesados = f"=IFERROR(AC{row}/AD{row}, 0)"

                # Aplicar formulas de porcentagem
                self.sheet2[f'V{row}'].value = formula_perc_caminhoes
                self.sheet2[f'V{row}'].number_format = '0.0%'
                self.sheet2[f'X{row}'].value = formula_perc_carretas
                self.sheet2[f'X{row}'].number_format = '0.0%'
                self.sheet2[f'Z{row}'].value = formula_perc_onibus
                self.sheet2[f'Z{row}'].number_format = '0.0%'
                self.sheet2[f'AB{row}'].value = formula_perc_pesados
                self.sheet2[f'AB{row}'].number_format = '0.0%'

            """
            Definir a linha final da tabela, onde exibe o total do dia.
            """
            footer_style = NamedStyle(name="footer_style", font=Font(bold=True, size=11))
            footer_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            footer_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)

            footer_row = 101

            # Mesclar celula total
            table_columns = [
                f'B{footer_row}:C{footer_row}'
            ]
            # Total
            self.sheet2['B101'] = "Total"

            # Aplicar mesclagem nas áreas definidas
            for footer_info in table_columns:
                self.sheet2.merge_cells(footer_info)

            # Aplicar fórmula de soma para as colunas D até U na linha do rodapé
            for col in range(ord('D'), ord('U') + 1): # percorre do D  até U
                col_letter = chr(col)
                formula = f"=SUM({col_letter}4:{col_letter}100)"
                self.sheet2[f"{col_letter}{footer_row}"].value = formula
            
            # Criar e aplicar as formulas de percentual
            columns_perc = ['V', 'X', 'Z', 'AB']
            for col_letter in columns_perc:
                formula = f"=IFERROR(({col_letter}4:{col_letter}100), 0)"
                self.sheet2[f"{col_letter}{footer_row}"].value = formula

            # Criar e aplicar as formulas de total pesados
            self.sheet2[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"

            self.sheet2[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"

            self.sheet2[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"

            self.sheet2[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{footer_row},AA{footer_row})"

            # Formula veiculos totais
            self.sheet2[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"

            # Adicionando 2 linhas para espaçamento entre tabelas.
            return footer_row + 2

        def add_data(self, data):
            """
            Função para encapsular tudo e criar tabelas de acordo com os movimentos
            """
            movimentos = data.get("Movimentos", [])
            start_row = 1
            for i, movimento in enumerate(movimentos):
                start_row = self.create_movement_table(start_row, data, movimento, i)

    class abaHr:
        def __init__(self, wb):
            self.wb = wb
            self.sheet3 = self.wb.create_sheet(title="Hr")

        def create_movement_table(self, start_row, data, movement, movement_index):
            """
            Definir cabeçalho da tabela de contagem.
            """
            ponto = data.get("Ponto", "")
            movimento_concatenado = f"{ponto}{movement}" if ponto and movement else movement
            
            # Criar uma formatação NamedStyle
            hr_header_style = NamedStyle(name="hr_header_style", font=Font(bold=True, size=12, color="FF0000"), alignment=Alignment(horizontal='center', vertical='center'))
            hr_header_style.font = Font(bold=True, size=12, color="FF0000")
            hr_header_style.alignment = Alignment(horizontal='center', vertical='center')

            # Criar o cabeçalho, definir seu valor e aplicar o estilo
            self.sheet3['B1'] = "Data:"
            self.sheet3['C1'].value = data.get("Data","")
            self.sheet3['C1'].style = hr_header_style

            self.sheet3['B2'] = "Movimento:"
            self.sheet3['C2'].value = movimento_concatenado
            self.sheet3['C2'].style = hr_header_style
            
            """
            Definir as colunas de veiculos
            """
            
            # Criar uma formatação com NamedStyled
            vehicle_col_style = NamedStyle(name="vehicle_columns")
            vehicle_col_style.font = Font(bold=True, size=12)
            vehicle_col_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            vehicle_col_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)

            # Lista para mesclar celulas
            table_columns = [
                'B3:C3',  # Hora
                'D3:D4',  # Leves
                'E3:G3',  # Carretinha
                'H3:H4',  # VUC
                'I3:K3',  # Caminhões
                'L3:R3',  # Carretas
                'S3:T3',  # Ônibus
                'U3:U4',  # Motos
                'V3:AC3',  # Pesados
                'AD3:AD4'  # Veiculos Totais
            ]
            # Aplicar mesclagem nas áreas definidas
            for header_info in table_columns:
                self.sheet3.merge_cells(header_info)
            
            # Dar nome para as colunas mescladas
            headers = [
                ('B3', "Horas"),
                ('D3', "Leves"),
                ('E3', "Carretinha"),
                ('H3', "VUC"),
                ('I3', "Caminhões"),
                ('L3', "Carreta"),
                ('S3', "Ônibus"),
                ('U3', "Motos"),
                ('V3', "Pesados"),
                ('AD3', "Veículos Totais"),
            ]

            for header_info in headers:
                cell = self.sheet3[header_info[0]]
                cell.value = header_info[1]

            # Estilo para cabeçalhos principais
            center_align = Alignment(horizontal='center', vertical='center')

            # Aplicar cabeçalhos principais
            for header_info in table_columns:
                for row in self.sheet3[header_info]:
                    for cell in row:
                        cell.alignment = center_align

            # Adicionar subcategorias na linha 5
            subcategories = [
                ('B4', "das"), ('C4', "as"),
                ('E4', "1 Eixo"), ('F4', "2 Eixos"), ('G4', "3 Eixos"), 
                ('I4', "2 Eixos"), ('J4', "3 Eixos"), ('K4', "4 Eixos"),
                ('L4', "2 E"), ('M4', "3 E"), ('N4', "4 E"), ('O4', "5 E"), ('P4', "6 E"), ('Q4', "7 E"), ('R4', "8 E"),
                ('S4', "2 E"), ('T4', "3 E ou +"),
                ('V4', "% Cam"), ('w4', "Caminhões"), ('x4', "% Carr"), ('Y4', "Carretas"), ('Z4', "% Ônib"), ('AA4', "Ônibus"), ('AB4', "% Pes"), ('AC4', "Total")
            ]

            # Estilo para subcategorias
            subcat_style = Font(size=10)
            subcat_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            subcat_align = Alignment(horizontal='center', vertical='center')

            # Aplicar subcategorias
            for cell_pos, value in subcategories:
                cell = self.sheet3[cell_pos]
                cell.value = value
                cell.font = subcat_style
                cell.border = subcat_border
                cell.alignment = subcat_align

            """
            Definir o intervalo de 1 hora
            """

            # Criar uma formatação NamedStyle
            time_style = NamedStyle(name="time")
            time_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            time_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)
            # Criar laço para preencher a coluna "das" de 15 em 15 minutos até dar 23:59
            das_inicio = datetime.strptime("00:00", "%H:%M")
            as_inicio = datetime.strptime("01:00","%H:%M")
            das_fim = datetime.strptime("23:00", "%H:%M")
            as_fim = datetime.strptime("23:59", "%H:%M")
            cell_das_inicio = self.sheet3['B5']
            cell_as_inicio = self.sheet3['C5']

            while das_inicio <= das_fim and as_fim:
                cell_das_inicio.value = das_inicio.strftime("%H:%M")
                cell_das_inicio.style = time_style
                cell_as_inicio.value = as_inicio.strftime("%H:%M")
                cell_as_inicio.style = time_style
                cell_das_inicio = self.sheet3.cell(row=cell_das_inicio.row + 1, column=cell_das_inicio.column)
                cell_as_inicio = self.sheet3.cell(row=cell_as_inicio.row + 1, column=cell_as_inicio.column)
                das_inicio += timedelta(hours=1)
                as_inicio += timedelta(hours=1)
        
            """
            Total veiculos e porcentagem
            """
            # Criar uma formatação com NamedStyled
            total_vehicles_style = NamedStyle(name="total_vehicles")
            total_vehicles_style.font = Font(bold=True, size=12)
            total_vehicles_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            total_vehicles_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)

            start_row = 5
            end_row = 29

            # Loop para iterar entre star e end
            for row in range(start_row, end_row + 1):
                # Formulas de total 
                formula_caminhoes = f"=SUM(I{row}:K{row})"
                formula_carretas = f"=SUM(L{row}:R{row})"
                formula_onibus = f"=SUM(S{row}:T{row})"
                formula_total_pesados = f"=SUM(W{row},Y{row},AA{row})"
                formula_total_vehicles = f"=SUM(D{row}:H{row},AC{row})"

                # Aplicas as formulas em cada celula
                self.sheet3[f'W{row}'].value = formula_caminhoes
                self.sheet3[f'Y{row}'].value = formula_carretas
                self.sheet3[f'AA{row}'].value = formula_onibus
                self.sheet3[f'AC{row}'].value = formula_total_pesados
                self.sheet3[f'AD{row}'].value = formula_total_vehicles

                # Formula de percentual
                formula_perc_caminhoes = f"=IFERROR(W{row}/AD{row}, 0)"
                formula_perc_carretas = f"=IFERROR(Y{row}/AD{row}, 0)"
                formula_perc_onibus = f"=IFERROR(AA{row}/AD{row}, 0)"
                formula_perc_pesados = f"=IFERROR(AC{row}/AD{row}, 0)"

                # Aplicar formulas de porcentagem
                self.sheet3[f'V{row}'].value = formula_perc_caminhoes
                self.sheet3[f'V{row}'].number_format = '0.0%'
                self.sheet3[f'X{row}'].value = formula_perc_carretas
                self.sheet3[f'X{row}'].number_format = '0.0%'
                self.sheet3[f'Z{row}'].value = formula_perc_onibus
                self.sheet3[f'Z{row}'].number_format = '0.0%'
                self.sheet3[f'AB{row}'].value = formula_perc_pesados
                self.sheet3[f'AB{row}'].number_format = '0.0%'

            """
            Rodapé
            """
            footer_row = 29

            # Mesclar celula total
            table_columns = [
                f'B{footer_row}:C{footer_row}'
            ]

            # Aplicar mesclagem nas áreas definidas
            for footer_info in table_columns:
                self.sheet3.merge_cells(footer_info)

            # Aplicar fórmula de soma para as colunas D até U na linha do rodapé
            for col in range(ord('D'), ord('U') + 1): # percorre do D  até U
                col_letter = chr(col)
                formula = f"=SUM({col_letter}5:{col_letter}28)"
                self.sheet3[f"{col_letter}{footer_row}"].value = formula

            # Criar e aplicar as formulas de percentual
            columns_perc = ['V', 'X', 'Z', 'AB']

            for col_letter in columns_perc:
                formula = f"=IFERROR(({col_letter}5:{col_letter}28), 0)"
                self.sheet3[f"{col_letter}{footer_row}"].value = formula

            # Criar e aplicar as formulas de total pesados
            self.sheet3[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"
            self.sheet3[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"
            self.sheet3[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"
            self.sheet3[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{footer_row},AA{footer_row})"

            # Formula veiculos totais
            self.sheet3[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"

        def add_data(self, data):
            movimentos = data.get("Movimentos", [])
            start_row = 1
            for i, movimento in enumerate(movimentos):
                start_row = self.create_movement_table(start_row, data, movimento, i)

    def add_data(self, data):
        """
        Encapsular todas abas para lidar com criação das abas
        """
        self.data = data
        self.entrada.add_data(data)
        self.relatorio.add_data(data)
        self.hr.add_data(data)

        # Tratar a duplicação do relatorio e hr com base na duração de dias
        duration_days_str = data.get("Duração em dias", 1)
        # Transformar str em int
        try:
            duration_days = int(duration_days_str)  # Convert to integer
        except (ValueError, TypeError):
            duration_days = 1  # Fallback to 1 if conversion fails
            print(f"Warning: Invalid 'Duração em dias' value '{duration_days_str}', defaulting to 1.")
        
        if duration_days > 1:
            initial_date = datetime.strptime(data.get("Data", ""), "%d-%m-%Y")
            for day in range(1, duration_days):
                # Criar as novas abas
                relatorio_copy = self.abaRelatorio(self.wb)
                relatorio_copy.sheet2.title = f"Relatório ({day})"
                hr_copy = self.abaHr(self.wb)
                hr_copy.sheet3.title = f"Hr ({day})"

                # Atualizar as datas novas
                copy_data = data.copy()
                new_date = initial_date + timedelta(days=day)
                copy_data["Data"] = new_date.strftime("%d-%m-%Y")
                relatorio_copy.add_data(copy_data)
                hr_copy.add_data(copy_data)

    def save(self):
        # Ajustar largura das colunas automaticamente
        for sheet in self.wb.worksheets:
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min((max_length + 2),100)
                sheet.column_dimensions[column].width = adjusted_width


        # Salvar na pasta output/
        self.wb.save(f"output/{self.filename}")
        print(f"Planilha salva como {self.filename}")