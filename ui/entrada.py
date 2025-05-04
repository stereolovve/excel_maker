#ui/entrada.py

import flet as ft
from datetime import datetime

class DataEntryForm(ft.Column):
    def __init__(self, page, on_save_callback=None):
        self.on_save_callback = on_save_callback
        self.page = page
        super().__init__()
        

        # Create fixed input fields
        self.ponto = ft.TextField(label="Ponto", width=400)
        self.data_inicio = ft.TextField(
            label="Data (DD-MM-YYYY)",
            value=datetime.now().strftime("%d-%m-%Y")
        )
        self.localizacao = ft.TextField(label="Localização", width=400)
        self.num_movimento = ft.TextField(
            label="Numero Movimentos",
            value="0",
            keyboard_type=ft.KeyboardType.NUMBER,
            width=400,
            on_change=self.dynamic_movement_field
        )
        self.duracao_dias = ft.TextField(
            label="Duração em dias",
            value="0",
            keyboard_type=ft.KeyboardType.NUMBER,
            width=400
        )
        self.duracao_horas = ft.TextField(
            label="Duração em horas",
            value="0",
            keyboard_type=ft.KeyboardType.NUMBER,
            width=400
        )
        self.hora_inicio = ft.TextField(label="Periodo_Inicio", width=400)
        self.hora_fim = ft.TextField(label="Periodo_Fim", width=400)

        self.movement_container = ft.Column()
        self.movement_fields = []

        self.save_button = ft.ElevatedButton(text="Salvar", on_click=self.save_data, width=float('inf'))

        self.controls = [
            self.ponto,
            self.data_inicio,
            self.localizacao,
            self.num_movimento,
            self.duracao_dias,
            self.duracao_horas,
            self.hora_inicio,
            self.hora_fim,
            self.movement_container,
            self.save_button
        ]   

    def dynamic_movement_field(self, e):
        num = int(self.num_movimento.value)
        if num < 0:
            raise ValueError("O número de movimentos não pode ser negativo.")

        self.movement_container.controls.clear()
        self.movement_fields.clear()

        for i in range(num):
            movement_input = ft.TextField(label=f"Movimento {i + 1}", width=400)
            self.movement_fields.append(movement_input)
            self.movement_container.controls.append(movement_input)
        self.update()
        
        
    def save_data(self, e):
        data = {
            "Ponto": self.ponto.value,
            "Data": self.data_inicio.value,
            "Localização": self.localizacao.value,
            "Numero Movimentos": self.num_movimento.value,
            "Duração em dias": self.duracao_dias.value,
            "Duração em horas": self.duracao_horas.value,
            "Periodo_Inicio": self.hora_inicio.value,
            "Periodo_Fim": self.hora_fim.value,
            "Movimentos": [movement_input.value for movement_input in self.movement_fields]
        }   
        print("Dados salvos:", data)
        self.on_save_callback(data)
        self.page.snack_bar = ft.SnackBar(content=ft.Text("Dados salvos com sucesso!"))
        self.page.snack_bar.open = True
        self.page.update()
        

