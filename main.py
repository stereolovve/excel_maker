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
        wb = planilhaContagem(ponto=ponto, codigo=codigo)
        wb.add_data(data)
        wb.save()
        page.snack_bar = ft.SnackBar(content=ft.Text("Dados salvos com sucesso na planilha!"))
        page.snack_bar.open = True
        page.update()

    data_entry_form = DataEntryForm(page, on_save_callback=save_to_excel)
    page.add(data_entry_form)

if __name__ == "__main__":
    ft.app(target=main)