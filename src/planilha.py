#src/planilha.py
from tracemalloc import start
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
            # Header
            ponto = data.get("Ponto", "")
            movimento_concatenado = f"{ponto}{movement}" if ponto and movement else movement
            self.sheet2[f'B{start_row}'] = "Data:"
            self.sheet2[f'C{start_row}'] = data.get("Data", "")
            self.sheet2[f'B{start_row + 1}'] = "Movimento:"
            self.sheet2[f'C{start_row + 1}'] = movimento_concatenado
            for col in ['B', 'C']:
                for row in range(start_row, start_row + 2):
                    cell = self.sheet2[f'{col}{row}']
                    cell.style = self.header_style

            # Vehicle columns
            header_row = start_row + 2
            subcat_row = start_row + 3
            table_columns = [
                f'B{header_row}:C{header_row}',  # Hora
                f'D{header_row}:D{subcat_row}',  # Leves
                f'E{header_row}:G{header_row}',  # Carretinha
                f'H{header_row}:H{subcat_row}',  # VUC
                f'I{header_row}:K{header_row}',  # Caminhões
                f'L{header_row}:R{header_row}',  # Carretas
                f'S{header_row}:T{header_row}',  # Ônibus
                f'U{header_row}:U{subcat_row}',  # Motos
                f'V{header_row}:AC{header_row}',  # Pesados
                f'AD{header_row}:AD{subcat_row}'  # Veículos Totais
            ]
            for header_info in table_columns:
                self.sheet2.merge_cells(header_info)

            headers = [
                (f'B{header_row}', "Horas"),
                (f'D{header_row}', "Leves"),
                (f'E{header_row}', "Carretinha"),
                (f'H{header_row}', "VUC"),
                (f'I{header_row}', "Caminhões"),
                (f'L{header_row}', "Carreta"),
                (f'S{header_row}', "Ônibus"),
                (f'U{header_row}', "Motos"),
                (f'V{header_row}', "Pesados"),
                (f'AD{header_row}', "Veículos Totais"),
            ]
            for cell_pos, value in headers:
                cell = self.sheet2[cell_pos]
                cell.value = value
                cell.font = Font(bold=True, size=11)
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.header_style.border

            # Subcategories
            subcategories = [
                (f'B{subcat_row}', "das"), (f'C{subcat_row}', "as"),
                (f'E{subcat_row}', "1 Eixo"), (f'F{subcat_row}', "2 Eixos"), (f'G{subcat_row}', "3 Eixos"),
                (f'I{subcat_row}', "2 Eixos"), (f'J{subcat_row}', "3 Eixos"), (f'K{subcat_row}', "4 Eixos"),
                (f'L{subcat_row}', "2 E"), (f'M{subcat_row}', "3 E"), (f'N{subcat_row}', "4 E"),
                (f'O{subcat_row}', "5 E"), (f'P{subcat_row}', "6 E"), (f'Q{subcat_row}', "7 E"),
                (f'R{subcat_row}', "8 E"),
                (f'S{subcat_row}', "2 E"), (f'T{subcat_row}', "3 E ou +"),
                (f'V{subcat_row}', "% Cam"), (f'W{subcat_row}', "Caminhões"),
                (f'X{subcat_row}', "% Carr"), (f'Y{subcat_row}', "Carretas"),
                (f'Z{subcat_row}', "% Ônib"), (f'AA{subcat_row}', "Ônibus"),
                (f'AB{subcat_row}', "% Pes"), (f'AC{subcat_row}', "Total")
            ]
            for cell_pos, value in subcategories:
                cell = self.sheet2[cell_pos]
                cell.value = value
                cell.font = Font(size=10)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Time intervals (15-minute intervals)
            time_style = NamedStyle(name=f"time_{movement_index}_relatorio")
            time_style.alignment = Alignment(horizontal='center', vertical='center')
            time_style.border = Border(top=Side(style='thin'), left=Side(style='thick'), right=Side(style='thick'), bottom=Side(style='thin'))
            das_inicio = datetime.strptime("00:00", "%H:%M")
            as_inicio = datetime.strptime("00:15", "%H:%M")
            das_fim = datetime.strptime("23:45", "%H:%M")
            row = start_row + 4
            while das_inicio <= das_fim:
                self.sheet2[f'B{row}'].value = das_inicio.strftime("%H:%M")
                self.sheet2[f'B{row}'].style = time_style
                self.sheet2[f'C{row}'].value = as_inicio.strftime("%H:%M")
                self.sheet2[f'C{row}'].style = time_style
                das_inicio += timedelta(minutes=15)
                as_inicio += timedelta(minutes=15)
                row += 1

            # Total vehicles and percentages
            table_start_row = start_row + 4
            table_end_row = start_row + 99  # 96 intervals (24 hours * 4)
            for row in range(table_start_row, table_end_row + 1):
                self.sheet2[f'W{row}'].value = f"=SUM(I{row}:K{row})"
                self.sheet2[f'Y{row}'].value = f"=SUM(L{row}:R{row})"
                self.sheet2[f'AA{row}'].value = f"=SUM(S{row}:T{row})"
                self.sheet2[f'AC{row}'].value = f"=SUM(W{row},Y{row},AA{row})"
                if row >= table_start_row:  # Skip merged rows for AD
                    self.sheet2[f'AD{row}'].value = f"=SUM(D{row}:H{row},AC{row})"
                self.sheet2[f'V{row}'].value = f"=IFERROR(W{row}/AD{row}, 0)"
                self.sheet2[f'V{row}'].number_format = '0.0%'
                self.sheet2[f'X{row}'].value = f"=IFERROR(Y{row}/AD{row}, 0)"
                self.sheet2[f'X{row}'].number_format = '0.0%'
                self.sheet2[f'Z{row}'].value = f"=IFERROR(AA{row}/AD{row}, 0)"
                self.sheet2[f'Z{row}'].number_format = '0.0%'
                self.sheet2[f'AB{row}'].value = f"=IFERROR(AC{row}/AD{row}, 0)"
                self.sheet2[f'AB{row}'].number_format = '0.0%'

            # Footer
            footer_row = table_end_row + 1
            footer_style = NamedStyle(name="footer_style", font=Font(bold=True, size=11))
            footer_style.alignment = Alignment(horizontal='center', vertical='center')
            footer_style.border = Border(top=Side(style='thin'), left=Side(style='thick'), right=Side(style='thick'), bottom=Side(style='thin'))
            self.sheet2[f'B{footer_row}'] = "Total"
            self.sheet2.merge_cells(f'B{footer_row}:C{footer_row}')
            for col in range(ord('D'), ord('U') + 1):
                col_letter = chr(col)
                self.sheet2[f'{col_letter}{footer_row}'].value = f"=SUM({col_letter}{table_start_row}:{col_letter}{table_end_row})"
                self.sheet2[f'{col_letter}{footer_row}'].style = footer_style
            for col in ['V', 'X', 'Z', 'AB']:
                self.sheet2[f'{col}{footer_row}'].value = f"=IFERROR(SUM({col}{table_start_row}:{col}{table_end_row}), 0)"
                self.sheet2[f'{col}{footer_row}'].style = footer_style
                self.sheet2[f'{col}{footer_row}'].number_format = '0.0%'
            self.sheet2[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"
            self.sheet2[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"
            self.sheet2[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"
            self.sheet2[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{row},AA{row})"
            self.sheet2[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"
            for col in ['W', 'Y', 'AA', 'AC', 'AD']:
                self.sheet2[f'{col}{footer_row}'].style = footer_style

            return footer_row + 3  # Space for next table

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
            self.header_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

        def create_movement_table(self, start_row, data, movement, movement_index):
            # Header
            ponto = data.get("Ponto", "")
            movimento_concatenado = f"{ponto}{movement}" if ponto and movement else movement
            hr_header_style = NamedStyle(name=f"hr_header_{movement_index}", font=Font(bold=True, size=12, color="FF0000"))
            hr_header_style.alignment = Alignment(horizontal='center', vertical='center')
            self.sheet3[f'B{start_row}'] = "Data:"
            self.sheet3[f'C{start_row}'].value = data.get("Data", "")
            self.sheet3[f'C{start_row}'].style = hr_header_style
            self.sheet3[f'B{start_row + 1}'] = "Movimento:"
            self.sheet3[f'C{start_row + 1}'].value = movimento_concatenado
            self.sheet3[f'C{start_row + 1}'].style = hr_header_style

            # Vehicle columns
            header_row = start_row + 2
            subcat_row = start_row + 3
            table_columns = [
                f'B{header_row}:C{header_row}',  # Hora
                f'D{header_row}:D{subcat_row}',  # Leves
                f'E{header_row}:G{header_row}',  # Carretinha
                f'H{header_row}:H{subcat_row}',  # VUC
                f'I{header_row}:K{header_row}',  # Caminhões
                f'L{header_row}:R{header_row}',  # Carretas
                f'S{header_row}:T{header_row}',  # Ônibus
                f'U{header_row}:U{subcat_row}',  # Motos
                f'V{header_row}:AC{header_row}',  # Pesados
                f'AD{header_row}:AD{subcat_row}'  # Veículos Totais
            ]
            for header_info in table_columns:
                self.sheet3.merge_cells(header_info)

            headers = [
                (f'B{header_row}', "Horas"),
                (f'D{header_row}', "Leves"),
                (f'E{header_row}', "Carretinha"),
                (f'H{header_row}', "VUC"),
                (f'I{header_row}', "Caminhões"),
                (f'L{header_row}', "Carreta"),
                (f'S{header_row}', "Ônibus"),
                (f'U{header_row}', "Motos"),
                (f'V{header_row}', "Pesados"),
                (f'AD{header_row}', "Veículos Totais"),
            ]
            for cell_pos, value in headers:
                cell = self.sheet3[cell_pos]
                cell.value = value
                cell.font = Font(bold=True, size=11)
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.header_border

            # Subcategories
            subcategories = [
                (f'B{subcat_row}', "das"), (f'C{subcat_row}', "as"),
                (f'E{subcat_row}', "1 Eixo"), (f'F{subcat_row}', "2 Eixos"), (f'G{subcat_row}', "3 Eixos"),
                (f'I{subcat_row}', "2 Eixos"), (f'J{subcat_row}', "3 Eixos"), (f'K{subcat_row}', "4 Eixos"),
                (f'L{subcat_row}', "2 E"), (f'M{subcat_row}', "3 E"), (f'N{subcat_row}', "4 E"),
                (f'O{subcat_row}', "5 E"), (f'P{subcat_row}', "6 E"), (f'Q{subcat_row}', "7 E"),
                (f'R{subcat_row}', "8 E"),
                (f'S{subcat_row}', "2 E"), (f'T{subcat_row}', "3 E ou +"),
                (f'V{subcat_row}', "% Cam"), (f'W{subcat_row}', "Caminhões"),
                (f'X{subcat_row}', "% Carr"), (f'Y{subcat_row}', "Carretas"),
                (f'Z{subcat_row}', "% Ônib"), (f'AA{subcat_row}', "Ônibus"),
                (f'AB{subcat_row}', "% Pes"), (f'AC{subcat_row}', "Total")
            ]
            for cell_pos, value in subcategories:
                cell = self.sheet3[cell_pos]
                cell.value = value
                cell.font = Font(size=10)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Time intervals (hourly)
            time_style = NamedStyle(name=f"time_{movement_index}_hr")
            time_style.alignment = Alignment(horizontal='center', vertical='center')
            time_style.border = Border(top=Side(style='thin'), left=Side(style='thick'), right=Side(style='thick'), bottom=Side(style='thin'))
            das_inicio = datetime.strptime("00:00", "%H:%M")
            as_inicio = datetime.strptime("01:00", "%H:%M")
            das_fim = datetime.strptime("23:00", "%H:%M")
            row = start_row + 4
            while das_inicio <= das_fim:
                self.sheet3[f'B{row}'].value = das_inicio.strftime("%H:%M")
                self.sheet3[f'B{row}'].style = time_style
                self.sheet3[f'C{row}'].value = as_inicio.strftime("%H:%M")
                self.sheet3[f'C{row}'].style = time_style
                das_inicio += timedelta(hours=1)
                as_inicio += timedelta(hours=1)
                row += 1

            # Total vehicles and percentages
            table_start_row = start_row + 4
            table_end_row = start_row + 27  # 24 hourly intervals
            for row in range(table_start_row, table_end_row + 1):
                self.sheet3[f'W{row}'].value = f"=SUM(I{row}:K{row})"
                self.sheet3[f'Y{row}'].value = f"=SUM(L{row}:R{row})"
                self.sheet3[f'AA{row}'].value = f"=SUM(S{row}:T{row})"
                self.sheet3[f'AC{row}'].value = f"=SUM(W{row},Y{row},AA{row})"
                if row >= table_start_row:  # Skip merged rows for AD
                    self.sheet3[f'AD{row}'].value = f"=SUM(D{row}:H{row},AC{row})"
                self.sheet3[f'V{row}'].value = f"=IFERROR(W{row}/AD{row}, 0)"
                self.sheet3[f'V{row}'].number_format = '0.0%'
                self.sheet3[f'X{row}'].value = f"=IFERROR(Y{row}/AD{row}, 0)"
                self.sheet3[f'X{row}'].number_format = '0.0%'
                self.sheet3[f'Z{row}'].value = f"=IFERROR(AA{row}/AD{row}, 0)"
                self.sheet3[f'Z{row}'].number_format = '0.0%'
                self.sheet3[f'AB{row}'].value = f"=IFERROR(AC{row}/AD{row}, 0)"
                self.sheet3[f'AB{row}'].number_format = '0.0%'

            # Footer
            footer_row = table_end_row + 1
            footer_style = NamedStyle(name="footer_style", font=Font(bold=True, size=11))
            footer_style.alignment = Alignment(horizontal='center', vertical='center')
            footer_style.border = Border(top=Side(style='thin'), left=Side(style='thick'), right=Side(style='thick'), bottom=Side(style='thin'))
            self.sheet3[f'B{footer_row}'] = "Total"
            self.sheet3.merge_cells(f'B{footer_row}:C{footer_row}')
            for col in range(ord('D'), ord('U') + 1):
                col_letter = chr(col)
                self.sheet3[f'{col_letter}{footer_row}'].value = f"=SUM({col_letter}{table_start_row}:{col_letter}{table_end_row})"
                self.sheet3[f'{col_letter}{footer_row}'].style = footer_style
            for col in ['V', 'X', 'Z', 'AB']:
                self.sheet3[f'{col}{footer_row}'].value = f"=IFERROR(SUM({col}{table_start_row}:{col}{table_end_row}), 0)"
                self.sheet3[f'{col}{footer_row}'].style = footer_style
                self.sheet3[f'{col}{footer_row}'].number_format = '0.0%'
            self.sheet3[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"
            self.sheet3[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"
            self.sheet3[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"
            self.sheet3[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{footer_row},AA{footer_row})"
            self.sheet3[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"
            for col in ['W', 'Y', 'AA', 'AC', 'AD']:
                self.sheet3[f'{col}{footer_row}'].style = footer_style

            return footer_row + 5  # Space Presse Enter twice to execute the code block for next table

        def add_data(self, data):
            movimentos = data.get("Movimentos", [])
            start_row = 1
            for i, movimento in enumerate(movimentos):
                start_row = self.create_movement_table(start_row, data, movimento, i)

    def add_data(self, data):
        self.data = data
        self.entrada.add_data(data)
        self.relatorio.add_data(data)
        self.hr.add_data(data)

        duration_days_str = data.get("Duração em dias", 1)
        try:
            duration_days = int(duration_days_str)
        except (ValueError, TypeError):
            duration_days = 1
            print(f"Warning: Invalid 'Duração em dias' value '{duration_days_str}', defaulting to 1.")

        if duration_days > 1:
            try:
                initial_date = datetime.strptime(data.get("Data", ""), "%d-%m-%Y")
            except ValueError:
                initial_date = datetime.now()
                print(f"Warning: Invalid date format for '{data.get('Data', '')}', using current date.")
            for day in range(1, duration_days):
                relatorio_copy = self.abaRelatorio(self.wb)
                relatorio_copy.sheet2.title = f"Relatório ({day})"
                hr_copy = self.abaHr(self.wb)
                hr_copy.sheet3.title = f"Hr ({day})"
                copy_data = data.copy()
                new_date = initial_date + timedelta(days=day)
                copy_data["Data"] = new_date.strftime("%d-%m-%Y")
                relatorio_copy.add_data(copy_data)
                hr_copy.add_data(copy_data)

    def save(self):
        for sheet in self.wb.worksheets:
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min((max_length + 2), 100)
                sheet.column_dimensions[column].width = adjusted_width
        self.wb.save(f"output/{self.filename}")
        print(f"Planilha salva como {self.filename}")