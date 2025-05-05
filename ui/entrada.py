#ui/entrada.py

import flet as ft
from datetime import datetime
import os
import requests
from config.config import API_URL

class DataEntryForm(ft.Column):
    def __init__(self, page, on_save_callback=None):
        super().__init__()
        self.on_save_callback = on_save_callback
        self.base_api = API_URL.rstrip("/") + "/trabalhos/api/"
        self.page = page
        
        
    

        # Create fixed input fields
        self.cliente = ft.Dropdown(label="Cliente", width=400, options=[], on_change=self.on_cliente_change)
        self.codigo = ft.Dropdown(label="Código", width=400, options=[], on_change=self.on_codigo_change)
        self.ponto = ft.Dropdown(label="Ponto", width=400, options=[])
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
            self.cliente,
            self.codigo,
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
        self.load_clientes()
        # Buscar os clientes
    def load_clientes(self):
        url = self.base_api + "clientes/"
        # Implement API call to fetch clients
        response = requests.get(url)
            
        clientes = response.json()
        self.cliente.options = [
            ft.dropdown.Option(text=c["nome"], key=str(c["id"]))
            for c in clientes
        ]
        self.page.update()

    def on_cliente_change(self, e):
        """
        Quando usuário seleciona um cliente,
        busca os códigos relacionados e atualiza o dropdown.
        """
        cliente_id = e.control.value

        url = self.base_api + "codigos/"

        self.codigo.options = []
        self.codigo.value = None
        self.ponto.options = []
        self.ponto.value = None

        response = requests.get(url)
        response.raise_for_status()
        codigos = [
            cod for cod in response.json()
            if str(cod["id"]) == cliente_id
        ]
        self.codigo.options = [
            ft.dropdown.Option(text=c["codigo"], key=str(c["id"])) 
            for c in codigos
        ]

        self.page.update()

    def on_codigo_change(self, e):
        codigo_id = e.control.value
        self.ponto.options = []
        self.ponto.value = None

        url = f"{self.base_api}pontos/?codigo={codigo_id}"
        response = requests.get(url)
        pontos = response.json()
        self.ponto.options = [
            ft.dropdown.Option(text=p["nome"], key=str(p["id"])) for p in pontos
        ]
        self.page.update()
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
        cliente_id = self.cliente.value
        codigo_id = self.codigo.value
        ponto_id = self.ponto.value

        cliente_name = next(
            opt.text for opt in self.cliente.options if opt.key == cliente_id
        )
        codigo_name = next(
            opt.text for opt in self.codigo.options if opt.key == codigo_id
        )
        ponto_name = next(
            opt.text for opt in self.ponto.options if opt.key == ponto_id
        )

        data = {
            "Cliente": cliente_name,
            "Código": codigo_name,
            "Ponto": ponto_name,
            "Data": self.data_inicio.value,
            "Localização": self.localizacao.value,
            "Num_Movimentos": self.num_movimento.value,
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
        

