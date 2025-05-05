# main.py
import flet as ft
from ui.entrada import DataEntryForm
from src.planilha import planilhaContagem


def main(page: ft.Page):
    page.title = "Entrada de Dados"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    page.window.width = 400
    page.window.height = 600

    def save_to_excel(data):
        ponto = data.get("Ponto", "")
        codigo = data.get("Código", "")
        wb = planilhaContagem(ponto=ponto, codigo=codigo)  # Passa o valor do campo Ponto e Código
        wb.entrada.add_data(data)
        wb.resumo.add_header()
        wb.resumo.add_header_value(data)
        wb.resumo.add_vehicle_columns()
        wb.resumo.add_time_intervals()
        wb.resumo.add_total_vehicles()
        wb.resumo.add_footer_vehicles()
        wb.resumo.add_footer_total()
        wb.hr.add_header()
        wb.hr.add_vehicle_columns()
        wb.hr.add_time_intervals()
        wb.hr.add_total_vehicles()
        wb.hr.add_footer_vehicles()
        wb.hr.add_footer_total()
        wb.save()
        page.snack_bar = ft.SnackBar(content=ft.Text("Dados salvos com sucesso na planilha!"))
        page.snack_bar.open = True
        page.update()

    data_entry_form = DataEntryForm(page, on_save_callback=save_to_excel)
    page.add(data_entry_form)

if __name__ == "__main__":
    ft.app(target=main)