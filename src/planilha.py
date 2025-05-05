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
        self.resumo = self.abaRelatorio(self.wb)
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
            self.header_font = Font(bold=True, size=12, color="FF0000")
            self.header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            self.border = Border(
                left=Side(style='medium'),
                right=Side(style='medium'),
                top=Side(style='medium'),
                bottom=Side(style='medium')
            )

        def add_header(self):
            self.sheet2['B1'] = "Data:"
            self.sheet2['B2'] = "Movimento:"

        def add_header_value(self, data=None):
            if data:
                mov_nomes = data.get("Movimentos", [])[0] if data.get("Movimentos", []) else ""
                ponto = data.get("Ponto", "")
                movimento_concatenado = f"{ponto}{mov_nomes}" if ponto and mov_nomes else mov_nomes
                header_values = [
                    ('C1', data.get("Data", "")),
                    ('C2', movimento_concatenado)
                ]
                for header_pos, value in header_values:
                    header_value = self.sheet2[header_pos]
                    header_value.value = value
                    header_value.font = self.header_font
        
        def add_vehicle_columns(self):
            # Lista para mesclar celulas
            merged_areas = [
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
            for header_info in merged_areas:
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
            header_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin'),
                diagonal=Side(style='thin'),

            )
            center_align = Alignment(horizontal='center', vertical='center')

            # Aplicar cabeçalhos principais
            for header_info in merged_areas:
                for row in self.sheet2[header_info]:
                    for cell in row:
                        cell.border = self.border
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

        def add_time_intervals(self):
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
        
        def add_total_vehicles(self):
            # Criar uma formatação com NamedStyled


            start_row = 5
            end_row = 100

            # Loop para iterar entre star e end
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

        def add_footer_vehicles(self):
            # Linha do rodapé
            footer_row = 101

            # Mesclar celula total
            merged_areas = [
                f'B{footer_row}:C{footer_row}'
            ]

            # Aplicar mesclagem nas áreas definidas
            for footer_info in merged_areas:
                self.sheet2.merge_cells(footer_info)
                for row in self.sheet2[footer_info]:
                    for cell in row:
                        cell.border = self.border

            # Aplicar fórmula de soma para as colunas D até U na linha do rodapé
            for col in range(ord('D'), ord('U') + 1): # percorre do D  até U
                col_letter = chr(col)
                formula = f"=SUM({col_letter}4:{col_letter}100)"
                self.sheet2[f"{col_letter}{footer_row}"].value = formula
                self.sheet2[f"{col_letter}{footer_row}"].border = self.border

        def add_footer_total(self):
            footer_row = 101

            

            # Criar e aplicar as formulas de percentual
            columns_perc = ['V', 'X', 'Z', 'AB']
            for col_letter in columns_perc:
                formula = f"=IFERROR(({col_letter}4:{col_letter}100), 0)"
                self.sheet2[f"{col_letter}{footer_row}"].value = formula
                self.sheet2[f"{col_letter}{footer_row}"].border = self.border

            # Criar e aplicar as formulas de total pesados
            self.sheet2[f'W{footer_row}'].value = f"=SUM(I{footer_row}:K{footer_row})"
            self.sheet2[f'W{footer_row}'].border = self.border

            self.sheet2[f'Y{footer_row}'].value = f"=SUM(L{footer_row}:R{footer_row})"
            self.sheet2[f'Y{footer_row}'].border = self.border

            self.sheet2[f'AA{footer_row}'].value = f"=SUM(S{footer_row}:T{footer_row})"
            self.sheet2[f'AA{footer_row}'].border = self.border

            self.sheet2[f'AC{footer_row}'].value = f"=SUM(W{footer_row},Y{footer_row},AA{footer_row})"
            self.sheet2[f'AC{footer_row}'].border = self.border

            # Formula veiculos totais
            self.sheet2[f'AD{footer_row}'].value = f"=SUM(D{footer_row}:H{footer_row},AC{footer_row})"
            self.sheet2[f'AD{footer_row}'].border = self.border

    class abaHr:
        def __init__(self, wb):
            self.wb = wb
            self.sheet3 = self.wb.create_sheet(title="Hr")

        def add_header(self, data=None):
            # Criar uma formatação NamedStyle
            header_style = NamedStyle(name="header_style")
            header_style.font = Font(bold=True, size=12, color="FF0000")
            header_style.alignment = Alignment(horizontal='center', vertical='center')

            # Criar cabeçalho
            self.sheet3['B1'] = "Data:"
            self.sheet3['C1'].value = "=Relatório!C1"
            # Formatar com NamedStyle
            self.sheet3['C1'].style = header_style

            self.sheet3['B2'] = "Movimento:"
            self.sheet3['C2'].value = '=Relatório!C2'
            self.sheet3['C2'].style = header_style

        def add_vehicle_columns(self):
            # Criar uma formatação com NamedStyled
            vehicle_col_style = NamedStyle(name="vehicle_columns")
            vehicle_col_style.font = Font(bold=True, size=12)
            vehicle_col_style.alignment = Alignment(horizontal='center', vertical='center')
            thin = Side(border_style="thin")
            thick = Side(border_style="thick")
            vehicle_col_style.border = Border(top=thin, left=thick, right=thick, bottom=thin)

            # Lista para mesclar celulas
            merged_areas = [
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
            for header_info in merged_areas:
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
            for header_info in merged_areas:
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

        def add_time_intervals(self):
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
        
        def add_total_vehicles(self):
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

        def add_footer_vehicles(self):
            # Linha do rodapé
            footer_row = 29

            # Mesclar celula total
            merged_areas = [
                f'B{footer_row}:C{footer_row}'
            ]

            # Aplicar mesclagem nas áreas definidas
            for footer_info in merged_areas:
                self.sheet3.merge_cells(footer_info)


            # Aplicar fórmula de soma para as colunas D até U na linha do rodapé
            for col in range(ord('D'), ord('U') + 1): # percorre do D  até U
                col_letter = chr(col)
                formula = f"=SUM({col_letter}5:{col_letter}28)"
                self.sheet3[f"{col_letter}{footer_row}"].value = formula

        def add_footer_total(self):
            footer_row = 29
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

        
    def save(self):
        # Ajustar largura das colunas automaticamente
        for sheet in [self.entrada.sheet1, self.resumo.sheet2, self.hr.sheet3]:
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter
                has_data = False
                for cell in col:
                    if cell.value:
                        has_data = True
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min((max_length + 2),100)
                sheet.column_dimensions[column].width = adjusted_width


        # Salvar na pasta output/
        self.wb.save(f"output/{self.filename}")
        print(f"Planilha salva como {self.filename}")