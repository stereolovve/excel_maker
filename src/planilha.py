from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from datetime import datetime, timedelta
import os

class planilhaContagem:
    def __init__(self, codigo="", ponto=""):
        self.filename = f"{codigo}_{ponto}.xlsx"
        self.wb = Workbook()
        self.vehicle_data = []  # Store data for summary
        self.entrada = self.abaEntrada(self.wb)
        self.resumo = self.abaResumo(self.wb)
        self.relatorio = self.abaRelatorio(self.wb, self)  # Pass parent
        self.hr = self.abaHr(self.wb)

    class abaEntrada:
        def __init__(self, wb):
            self.wb = wb
            self.sheet1 = self.wb.active
            self.sheet1.title = "Entrada"
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

            if data:
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
                for cell_pos, value in data_cells:
                    cell = self.sheet1[cell_pos]
                    cell.value = value
                    cell.border = self.border

                movimentos = data.get("Movimentos", [])
                for i, movimento in enumerate(movimentos, start=2):
                    cell = self.sheet1[f'E{i}']
                    cell.value = movimento
                    cell.border = self.border

    class abaResumo:
        def __init__(self, wb):
            self.wb = wb
            self.sheet = self.wb.create_sheet(title="Resumo")
            self.header_font = Font(bold=True)
            self.header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            self.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            self.center_align = Alignment(horizontal='center', vertical='center')

        def add_data(self, vehicle_data, days):
            # Get data from Entrada sheet
            entrada_sheet = self.wb['Entrada']
            ponto = entrada_sheet['C1'].value or "Unknown"
            num_movimentos = entrada_sheet['C3'].value or 0
            duracao_horas = entrada_sheet['C6'].value or 0  # Use hours from C6

            # Convert num_movimentos to integer with error handling
            try:
                num_movimentos = int(num_movimentos)
            except (ValueError, TypeError):
                num_movimentos = 0
                print(f"Warning: Invalid 'Num_Movimentos' value '{entrada_sheet['C3'].value}', defaulting to 0.")

            # Fetch movements safely
            movimentos = []
            if num_movimentos > 0:
                movimentos = [
                    entrada_sheet[f'E{i}'].value
                    for i in range(2, 2 + num_movimentos)
                    if entrada_sheet[f'E{i}'].value
                ]

            # Section 1: Summary by Category
            self.sheet['B2'] = "Resumo por Categoria"
            self.sheet['B2'].font = self.header_font
            self.sheet['B2'].fill = self.header_fill
            headers = ['Movimento', 'Quantidade', 'Duração da Contagem (horas)']
            for col, header in enumerate(headers, 2):
                cell = self.sheet.cell(row=3, column=col)
                cell.value = header
                cell.font = self.header_font
                cell.fill = self.header_fill
                cell.border = self.border
                cell.alignment = self.center_align

            # Populate movements and duration
            if movimentos:
                for row, movimento in enumerate(movimentos, 4):
                    self.sheet.cell(row=row, column=2).value = movimento
                    self.sheet.cell(row=row, column=3).value = 1  # Each movement listed individually
                    self.sheet.cell(row=row, column=4).value = duracao_horas
                    for col in range(2, 5):
                        self.sheet.cell(row=row, column=col).border = self.border
            else:
                self.sheet.cell(row=4, column=2).value = "Nenhum movimento registrado"
                self.sheet.cell(row=4, column=2).border = self.border

            # Section 2: Morning Peak Hour (7h–8h)
            self.sheet['B10'] = "Horário Pico Manhã (7h–8h)"
            self.sheet['B10'].font = self.header_font
            self.sheet['B10'].fill = self.header_fill
            vehicle_types = ['Leves', 'VUC', 'Caminhões', 'Carretas', 'Ônibus', 'Pesados', 'Motos', 'Total', 'Total s/VUC']
            headers = ['Movimento'] + vehicle_types
            for col, header in enumerate(headers, 2):
                cell = self.sheet.cell(row=11, column=col)
                cell.value = header
                cell.font = self.header_font
                cell.fill = self.header_fill
                cell.border = self.border
                cell.alignment = self.center_align

            # Fetch data from Hr sheet for 7h–8h
            try:
                hr_sheet = self.wb['Hr']
                if movimentos:
                    for row, movimento in enumerate(movimentos, 12):
                        self.sheet.cell(row=row, column=2).value = movimento
                        # Find the 7h–8h row in Hr sheet (assuming it starts at row 5 and each row is an hour)
                        hr_row = 5 + 7  # 7th hour (7:00–8:00)
                        col_mapping = {
                            'Leves': 'D', 'VUC': 'H', 'Caminhões': 'W', 'Carretas': 'Y',
                            'Ônibus': 'AA', 'Motos': 'U'
                        }
                        for col_idx, vt in enumerate(vehicle_types, 3):
                            col_letter = col_mapping.get(vt)
                            if col_letter:
                                self.sheet.cell(row=row, column=col_idx).value = f"='Hr'!{col_letter}{hr_row}"
                            elif vt == 'Pesados':
                                # Sum Caminhões, Carretas, Ônibus (columns E, F, G in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(E{row},F{row},G{row})"
                            elif vt == 'Total':
                                # Sum Leves, VUC, Motos (columns C, D, I in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(C{row},D{row},I{row})"
                            elif vt == 'Total s/VUC':
                                # Sum Leves, Motos (columns C, I in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(C{row},I{row})"
                            self.sheet.cell(row=row, column=col_idx).border = self.border
                else:
                    self.sheet.cell(row=12, column=2).value = "Nenhum movimento registrado"
                    self.sheet.cell(row=12, column=2).border = self.border
            except KeyError:
                self.sheet.cell(row=12, column=2).value = "Folha 'Hr' não encontrada"
                self.sheet.cell(row=12, column=2).border = self.border
                print("Warning: 'Hr' sheet not found in workbook.")

            # Bar Chart: Morning Peak Hour
            if movimentos:
                chart = BarChart()
                chart.title = "Morning Peak Hour (7h–8h) Vehicle Distribution"
                chart.x_axis.title = "Movimento"
                chart.y_axis.title = "Number of Vehicles"
                data = Reference(self.sheet, min_col=3, max_col=11, min_row=11, max_row=11+len(movimentos))
                cats = Reference(self.sheet, min_col=2, min_row=12, max_row=12+len(movimentos)-1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                self.sheet.add_chart(chart, "M5")

            # Section 3: Afternoon Peak Hour (17h–18h)
            self.sheet['B20'] = "Horário Pico Tarde (17h–18h)"
            self.sheet['B20'].font = self.header_font
            self.sheet['B20'].fill = self.header_fill
            for col, header in enumerate(headers, 2):
                cell = self.sheet.cell(row=21, column=col)
                cell.value = header
                cell.font = self.header_font
                cell.fill = self.header_fill
                cell.border = self.border
                cell.alignment = self.center_align

            # Fetch data from Hr sheet for 17h–18h
            try:
                hr_sheet = self.wb['Hr']
                if movimentos:
                    for row, movimento in enumerate(movimentos, 22):
                        self.sheet.cell(row=row, column=2).value = movimento
                        hr_row = 5 + 17  # 17th hour (17:00–18:00)
                        for col_idx, vt in enumerate(vehicle_types, 3):
                            col_letter = col_mapping.get(vt)
                            if col_letter:
                                self.sheet.cell(row=row, column=col_idx).value = f"='Hr'!{col_letter}{hr_row}"
                            elif vt == 'Pesados':
                                # Sum Caminhões, Carretas, Ônibus (columns E, F, G in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(E{row},F{row},G{row})"
                            elif vt == 'Total':
                                # Sum Leves, VUC, Motos (columns C, D, I in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(C{row},D{row},H{row})"
                            elif vt == 'Total s/VUC':
                                # Sum Leves, Motos (columns C, I in Resumo)
                                self.sheet.cell(row=row, column=col_idx).value = f"=SUM(C{row},H{row})"
                            self.sheet.cell(row=row, column=col_idx).border = self.border
                else:
                    self.sheet.cell(row=22, column=2).value = "Nenhum movimento registrado"
                    self.sheet.cell(row=22, column=2).border = self.border
            except KeyError:
                self.sheet.cell(row=22, column=2).value = "Folha 'Hr' não encontrada"
                self.sheet.cell(row=22, column=2).border = self.border
                print("Warning: 'Hr' sheet not found in workbook.")

            # Bar Chart: Afternoon Peak Hour
            if movimentos:
                chart = BarChart()
                chart.title = "Afternoon Peak Hour (17h–18h) Vehicle Distribution"
                chart.x_axis.title = "Movimento"
                chart.y_axis.title = "Number of Vehicles"
                data = Reference(self.sheet, min_col=3, max_col=11, min_row=21, max_row=21+len(movimentos))
                cats = Reference(self.sheet, min_col=2, min_row=22, max_row=22+len(movimentos)-1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                self.sheet.add_chart(chart, "M20")

            # Section 4: Fluxo Total do Dia
            start_row = 30 + len(movimentos)  # Place after Afternoon Peak Hour table
            self.sheet[f'B{start_row}'] = "Fluxo Total do Dia"
            self.sheet[f'B{start_row}'].font = self.header_font
            self.sheet[f'B{start_row}'].fill = self.header_fill
            vehicle_types_daily = ['Leves', 'VUC', 'Caminhões', 'Carretas', 'Ônibus', 'Pesados', 'Motos', 'Total']
            headers_daily = ['Movimento'] + vehicle_types_daily
            for col, header in enumerate(headers_daily, 2):
                cell = self.sheet.cell(row=start_row+1, column=col)
                cell.value = header
                cell.font = self.header_font
                cell.fill = self.header_fill
                cell.border = self.border
                cell.alignment = self.center_align

            # Fetch totals from Hr sheet footer
            try:
                hr_sheet = self.wb['Hr']
                if movimentos:
                    for idx, movimento in enumerate(movimentos, 0):
                        row = start_row + 2 + idx
                        self.sheet.cell(row=row, column=2).value = movimento
                        # Hr footer row for each movement (28 for first, 34 for second, etc.)
                        hr_footer_row = 28 + (idx * 6)  # Each movement table in Hr is 6 rows apart
                        col_mapping_daily = {
                            'Leves': 'D', 'VUC': 'H', 'Caminhões': 'W', 'Carretas': 'Y',
                            'Ônibus': 'AA', 'Pesados': 'AC', 'Motos': 'U', 'Total': 'AD'
                        }
                        for col_idx, vt in enumerate(vehicle_types_daily, 3):
                            col_letter = col_mapping_daily.get(vt)
                            if col_letter:
                                self.sheet.cell(row=row, column=col_idx).value = f"='Hr'!{col_letter}{hr_footer_row}"
                            self.sheet.cell(row=row, column=col_idx).border = self.border
                else:
                    self.sheet.cell(row=start_row+2, column=2).value = "Nenhum movimento registrado"
                    self.sheet.cell(row=start_row+2, column=2).border = self.border
            except KeyError:
                self.sheet.cell(row=start_row+2, column=2).value = "Folha 'Hr' não encontrada"
                self.sheet.cell(row=start_row+2, column=2).border = self.border
                print("Warning: 'Hr' sheet not found in workbook.")

            # Bar Chart: Fluxo Total do Dia
            if movimentos:
                chart = BarChart()
                chart.title = "Fluxo Total do Dia por Movimento"
                chart.x_axis.title = "Movimento"
                chart.y_axis.title = "Number of Vehicles"
                data = Reference(self.sheet, min_col=3, max_col=10, min_row=start_row+1, max_row=start_row+1+len(movimentos))
                cats = Reference(self.sheet, min_col=2, min_row=start_row+2, max_row=start_row+2+len(movimentos)-1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                self.sheet.add_chart(chart, f"M{start_row}")

            # Adjust column widths
            for col in self.sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min((max_length + 2), 50)
                self.sheet.column_dimensions[column].width = adjusted_width
    class abaRelatorio:
        def __init__(self, wb, parent):
            self.wb = wb
            self.parent = parent  # Reference to planilhaContagem
            self.sheet2 = self.wb.create_sheet(title="Relatório")
            self.header_style = NamedStyle(name="header_style")
            self.header_style.alignment = Alignment(horizontal='left', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            self.header_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)
            self.header_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

        def create_movement_table(self, start_row, data, movement, movement_index):
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

            header_row = start_row + 2
            subcat_row = start_row + 3
            table_columns = [
                f'B{header_row}:C{header_row}', f'D{header_row}:D{subcat_row}', f'E{header_row}:G{header_row}',
                f'H{header_row}:H{subcat_row}', f'I{header_row}:K{header_row}', f'L{header_row}:R{header_row}',
                f'S{header_row}:T{header_row}', f'U{header_row}:U{subcat_row}', f'V{header_row}:AC{header_row}',
                f'AD{header_row}:AD{subcat_row}'
            ]
            for header_info in table_columns:
                self.sheet2.merge_cells(header_info)

            headers = [
                (f'B{header_row}', "Horas"), (f'D{header_row}', "Leves"), (f'E{header_row}', "Carretinha"),
                (f'H{header_row}', "VUC"), (f'I{header_row}', "Caminhões"), (f'L{header_row}', "Carreta"),
                (f'S{header_row}', "Ônibus"), (f'U{header_row}', "Motos"), (f'V{header_row}', "Pesados"),
                (f'AD{header_row}', "Veículos Totais"),
            ]
            for cell_pos, value in headers:
                cell = self.sheet2[cell_pos]
                cell.value = value
                cell.font = Font(bold=True, size=11)
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.header_border

            subcategories = [
                (f'B{subcat_row}', "das"), (f'C{subcat_row}', "as"),
                (f'E{subcat_row}', "1 Eixo"), (f'F{subcat_row}', "2 Eixos"), (f'G{subcat_row}', "3 Eixos"),
                (f'I{subcat_row}', "2 Eixos"), (f'J{subcat_row}', "3 Eixos"), (f'K{subcat_row}', "4 Eixos"),
                (f'L{subcat_row}', "2 E"), (f'M{subcat_row}', "3 E"), (f'N{subcat_row}', "4 E"),
                (f'O{subcat_row}', "5 E"), (f'P{subcat_row}', "6 E"), (f'Q{subcat_row}', "7 E"),
                (f'R{subcat_row}', "8 E"), (f'S{subcat_row}', "2 E"), (f'T{subcat_row}', "3 E ou +"),
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

            table_start_row = start_row + 4
            table_end_row = start_row + 99
            for row in range(table_start_row, table_end_row + 1):
                self.sheet2[f'W{row}'].value = f"=SUM(I{row}:K{row})"
                self.sheet2[f'Y{row}'].value = f"=SUM(L{row}:R{row})"
                self.sheet2[f'AA{row}'].value = f"=SUM(S{row}:T{row})"
                self.sheet2[f'AC{row}'].value = f"=SUM(W{row},Y{row},AA{row})"
                if row >= table_start_row:
                    self.sheet2[f'AD{row}'].value = f"=SUM(D{row}:H{row},AC{row})"
                self.sheet2[f'V{row}'].value = f"=IFERROR(W{row}/AD{row}, 0)"
                self.sheet2[f'V{row}'].number_format = '0.0%'
                self.sheet2[f'X{row}'].value = f"=IFERROR(Y{row}/AD{row}, 0)"
                self.sheet2[f'X{row}'].number_format = '0.0%'
                self.sheet2[f'Z{row}'].value = f"=IFERROR(AA{row}/AD{row}, 0)"
                self.sheet2[f'Z{row}'].number_format = '0.0%'
                self.sheet2[f'AB{row}'].value = f"=IFERROR(AC{row}/AD{row}, 0)"
                self.sheet2[f'AB{row}'].number_format = '0.0%'

            footer_row = table_end_row + 1
            footer_style = NamedStyle(name="footer_style", font=Font(bold=True, size=11))
            footer_style.alignment = Alignment(horizontal='center', vertical='center')
            footer_style.border = Border(top=Side(style='thin'), left=Side(style='thick'), right=Side(style='thick'), bottom=Side(style='thin'))
            self.sheet2[f'B{footer_row}'] = "Total"
            self.sheet2.merge_cells(f'B{footer_row}:C{footer_row}')
            vehicle_totals = {}
            vehicle_types = ['Leves', 'Carretinha 1E', 'Carretinha 2E', 'Carretinha 3E', 'VUC',
                             'Caminhões 2E', 'Caminhões 3E', 'Caminhões 4E', 'Carreta 2E', 'Carreta 3E',
                             'Carreta 4E', 'Carreta 5E', 'Carreta 6E', 'Carreta 7E', 'Carreta 8E',
                             'Ônibus 2E', 'Ônibus 3E+', 'Motos']
            for col, vt in zip(range(ord('D'), ord('U') + 1), vehicle_types):
                col_letter = chr(col)
                self.sheet2[f'{col_letter}{footer_row}'].value = f"=SUM({col_letter}{table_start_row}:{col_letter}{table_end_row})"
                self.sheet2[f'{col_letter}{footer_row}'].style = footer_style
                vehicle_totals[vt] = 0  # Placeholder for actual values
            for col in ['V', 'X', 'Z', 'AB']:
                self.sheet2[f'{col}{footer_row}'].value = f"=IFERROR(SUM({col}{table_start_row}:{col}{table_end_row}), 0)"
                self.sheet2[f'{col}{footer_row}'].style = footer_style
                self.sheet2[f'{col}{footer_row}'].number_format = '0.0%'
            self.sheet2[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"
            self.sheet2[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"
            self.sheet2[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"
            self.sheet2[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{footer_row},AA{footer_row})"
            self.sheet2[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"
            vehicle_totals['Pesados'] = 0  # Placeholder
            vehicle_totals['Total'] = 0  # Placeholder
            for col in ['W', 'Y', 'AA', 'AC', 'AD']:
                self.sheet2[f'{col}{footer_row}'].style = footer_style

            return footer_row + 5, movimento_concatenado, vehicle_totals, data.get("Data", "")

        def add_data(self, data):
            movimentos = data.get("Movimentos", [])
            start_row = 1
            for i, movimento in enumerate(movimentos):
                start_row, movement_name, vehicle_totals, date = self.create_movement_table(start_row, data, movimento, i)
                self.parent.vehicle_data.append((date, movement_name, vehicle_totals))

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

            header_row = start_row + 2
            subcat_row = start_row + 3
            table_columns = [
                f'B{header_row}:C{header_row}', f'D{header_row}:D{subcat_row}', f'E{header_row}:G{header_row}',
                f'H{header_row}:H{subcat_row}', f'I{header_row}:K{header_row}', f'L{header_row}:R{header_row}',
                f'S{header_row}:T{header_row}', f'U{header_row}:U{subcat_row}', f'V{header_row}:AC{header_row}',
                f'AD{header_row}:AD{subcat_row}'
            ]
            for header_info in table_columns:
                self.sheet3.merge_cells(header_info)

            headers = [
                (f'B{header_row}', "Horas"), (f'D{header_row}', "Leves"), (f'E{header_row}', "Carretinha"),
                (f'H{header_row}', "VUC"), (f'I{header_row}', "Caminhões"), (f'L{header_row}', "Carreta"),
                (f'S{header_row}', "Ônibus"), (f'U{header_row}', "Motos"), (f'V{header_row}', "Pesados"),
                (f'AD{header_row}', "Veículos Totais"),
            ]
            for cell_pos, value in headers:
                cell = self.sheet3[cell_pos]
                cell.value = value
                cell.font = Font(bold=True, size=11)
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.header_border

            subcategories = [
                (f'B{subcat_row}', "das"), (f'C{subcat_row}', "as"),
                (f'E{subcat_row}', "1 Eixo"), (f'F{subcat_row}', "2 Eixos"), (f'G{subcat_row}', "3 Eixos"),
                (f'I{subcat_row}', "2 Eixos"), (f'J{subcat_row}', "3 Eixos"), (f'K{subcat_row}', "4 Eixos"),
                (f'L{subcat_row}', "2 E"), (f'M{subcat_row}', "3 E"), (f'N{subcat_row}', "4 E"),
                (f'O{subcat_row}', "5 E"), (f'P{subcat_row}', "6 E"), (f'Q{subcat_row}', "7 E"),
                (f'R{subcat_row}', "8 E"), (f'S{subcat_row}', "2 E"), (f'T{subcat_row}', "3 E ou +"),
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

            # Populate vehicle counts from Relatório sheet
            try:
                relatorio_sheet = self.wb['Relatório']
                table_start_row = start_row + 4
                table_end_row = start_row + 27
                for hr_row, hour in enumerate(range(24), table_start_row):
                    # Calculate the starting row in Relatório for this hour (4 rows per hour, 15-min intervals)
                    rel_row_start = 5 + (hour * 4)  # Relatório starts at row 5
                    # Map Hr columns to Relatório columns
                    col_mapping = {
                        'D': 'D',  # Leves
                        'H': 'H',  # VUC
                        'U': 'U',  # Motos
                        'E': 'E', 'F': 'F', 'G': 'G',  # Carretinha 1E, 2E, 3E
                        'I': 'I', 'J': 'J', 'K': 'K',  # Caminhões 2E, 3E, 4E
                        'L': 'L', 'M': 'M', 'N': 'N', 'O': 'O', 'P': 'P', 'Q': 'Q', 'R': 'R',  # Carreta 2E to 8E
                        'S': 'S', 'T': 'T',  # Ônibus 2E, 3E+
                        'W': 'W', 'Y': 'Y', 'AA': 'AA', 'AC': 'AC', 'AD': 'AD'  # Aggregates and totals
                    }
                    for col in col_mapping:
                        # Sum 4 rows (15-min intervals) for this hour
                        self.sheet3[f'{col}{hr_row}'].value = (
                            f"=SUM('Relatório'!{col}{rel_row_start}:'Relatório'!{col}{rel_row_start+3})"
                        )
            except KeyError:
                print("Warning: 'Relatório' sheet not found in workbook.")
                self.sheet3[f'B{table_start_row}'].value = "Folha 'Relatório' não encontrada"

            # Set up formulas for aggregates
            for row in range(table_start_row, table_end_row + 1):
                self.sheet3[f'W{row}'].value = f"=SUM(I{row}:K{row})"  # Caminhões
                self.sheet3[f'Y{row}'].value = f"=SUM(L{row}:R{row})"  # Carretas
                self.sheet3[f'AA{row}'].value = f"=SUM(S{row}:T{row})"  # Ônibus
                self.sheet3[f'AC{row}'].value = f"=SUM(W{row},Y{row},AA{row})"  # Pesados
                self.sheet3[f'AD{row}'].value = f"=SUM(D{row}:H{row},AC{row})"  # Total
                self.sheet3[f'V{row}'].value = f"=IFERROR(W{row}/AD{row}, 0)"  # % Caminhões
                self.sheet3[f'V{row}'].number_format = '0.0%'
                self.sheet3[f'X{row}'].value = f"=IFERROR(Y{row}/AD{row}, 0)"  # % Carretas
                self.sheet3[f'X{row}'].number_format = '0.0%'
                self.sheet3[f'Z{row}'].value = f"=IFERROR(AA{row}/AD{row}, 0)"  # % Ônibus
                self.sheet3[f'Z{row}'].number_format = '0.0%'
                self.sheet3[f'AB{row}'].value = f"=IFERROR(AC{row}/AD{row}, 0)"  # % Pesados
                self.sheet3[f'AB{row}'].number_format = '0.0%'

            # Footer row with totals
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

            return footer_row + 5

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

        days = []
        try:
            initial_date = datetime.strptime(data.get("Data", ""), "%d-%m-%Y")
        except ValueError:
            try:
                initial_date = datetime.strptime(data.get("Data", ""), "%Y-%m-%d")
            except ValueError:
                initial_date = datetime.now()
                print(f"Warning: Invalid date format, using current date.")
        days.append(initial_date.strftime("%Y-%m-%d"))

        if duration_days > 1:
            for day in range(1, duration_days):
                relatorio_copy = self.abaRelatorio(self.wb, self)  # Pass parent
                relatorio_copy.sheet2.title = f"Relatório ({day})"
                hr_copy = self.abaHr(self.wb)
                hr_copy.sheet3.title = f"Hr ({day})"
                copy_data = data.copy()
                new_date = initial_date + timedelta(days=day)
                copy_data["Data"] = new_date.strftime("%Y-%m-%d")
                days.append(new_date.strftime("%Y-%m-%d"))
                relatorio_copy.add_data(copy_data)
                hr_copy.add_data(copy_data)

        self.resumo.add_data(self.vehicle_data, days)

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