#src/planilha.py
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
# from ui.entrada import DataEntryForm  # Comentado pois não é necessário para o exemplo

# Create a new workbook
class planilhaContagem:
    def __init__(self, filename="Planilha_Contagem.xlsx"):
        self.filename = filename
        self.wb = Workbook()
        self.entrada = self.abaEntrada(self.wb)
        self.resumo = self.abaResumo(self.wb)

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
                    ('C3', data.get("Num_Movimentos", "0")),
                    ('C4', data.get("Localização", "0")),
                    ('C5', data.get("Duração em dias", "0")),
                    ('C6', data.get("Duração em horas", "0")),
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
    class abaResumo:
        def __init__(self, wb):
            self.wb = wb
            self.sheet2 = self.wb.create_sheet(title="Resumo")

            # Definir estilos para a aba de resumo
            self.header_font = Font(bold=True, size=14)
            self.header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            self.border = Border(
                left=Side(style='medium'),
                right=Side(style='medium'),
                top=Side(style='medium'),
                bottom=Side(style='medium')
            )

        def add_data(self):
            # Aplicar formatação ao cabeçalho
            cell = self.sheet2['A1']
            cell.value = "Total"
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.border

            # Aplicar formatação à célula com fórmula
            cell = self.sheet2['A2']
            cell.value = "=SUM(Entrada!B2:B3)"  # Formula referencing Sheet1
            cell.border = self.border

    def save(self):
        # Ajustar largura das colunas automaticamente
        for sheet in [self.entrada.sheet1, self.resumo.sheet2]:
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

        self.wb.save(self.filename)
        print(f"Planilha salva como {self.filename}")