"""
REGISTRO DE HORAS TRABAJADAS - Con fecha actual automática y modificable
Campos separados para horas y minutos, guardado en Excel
"""

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle
from kivy.utils import platform
from kivy.uix.widget import Widget
from kivy.uix.dropdown import DropDown
from kivy.uix.spinner import Spinner
from datetime import datetime
import re
import pandas as pd
from openpyxl import Workbook
import os

# Configurar tamaño para móvil
if platform in ('android', 'ios'):
    Window.size = (360, 640)
else:
    Window.size = (400, 820)

class HorasApp(App):
    def build(self):
        self.title = 'Registro de Horas'
        return MainScreen()

class MainScreen(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = [15, 10, 15, 10]
        self.spacing = 6
        
        # Lista de registros
        self.registros = []
        self.total_horas = 0.0
        self.total_precio = 0.0
        self.contador = 0
        
        # Colores del tema
        self.color_principal = (0.1, 0.4, 0.8, 1)
        self.color_secundario = (0.9, 0.9, 0.9, 1)
        self.color_fondo = (0.95, 0.95, 0.95, 1)
        
        self.build_ui()
        
    def build_ui(self):
        # Título
        header = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.04),
            padding=[10, 5],
            spacing=10
        )
        

        
        titulo = Label(
            text='[b]REGISTRO DE HORAS[/b]',
            markup=True,
            font_size=20,
            size_hint=(0.7, 1),
            color=(0.1, 0.4, 0.8, 1),
            halign='center',
            valign='middle'
        )
        titulo.bind(size=titulo.setter('text_size'))
        header.add_widget(titulo)
        
        # Mostrar fecha actual en el header
        fecha_actual = datetime.now().strftime('%d/%m/%y')
        fecha_header = Label(
            text=fecha_actual,
            font_size=14,
            size_hint=(0.15, 1),
            color=(1, 1, 1, 1),
            halign='right',
            valign='middle'
        )
        fecha_header.bind(size=fecha_header.setter('text_size'))
        header.add_widget(fecha_header)
        
        self.add_widget(header)
        self.add_widget(Widget(size_hint=(1, 0.005)))
        
        # Sección: DATOS DEL CLIENTE
        label_seccion = Label(
            text='[b]DATOS DEL CLIENTE[/b]',
            markup=True,
            font_size=16,
            size_hint=(1, 0.025),
            color=(0.1, 0.4, 0.8, 1),
            halign='left',
            valign='middle'
        )
        label_seccion.bind(size=label_seccion.setter('text_size'))
        self.add_widget(label_seccion)
        
        # Nombre del cliente
        cliente_box = BoxLayout(
            orientation='vertical',
            size_hint=(1, 0.045),
            spacing=1
        )
        
        label_cliente = Label(
            text='Nombre del Cliente',
            font_size=13,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='left',
            valign='middle'
        )
        label_cliente.bind(size=label_cliente.setter('text_size'))
        cliente_box.add_widget(label_cliente)
        
        self.cliente_input = TextInput(
            hint_text='Ej: Juan Pérez',
            multiline=False,
            font_size=15,
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5]
        )
        cliente_box.add_widget(self.cliente_input)
        self.add_widget(cliente_box)
        
        # Fecha con campos separados (Día, Mes, Año)
        label_fecha = Label(
            text='FECHA',
            font_size=13,
            size_hint=(1, 0.02),
            color=(1, 1, 1, 1),
            halign='left',
            valign='middle'
        )
        label_fecha.bind(size=label_fecha.setter('text_size'))
        self.add_widget(label_fecha)
        
        # Campos de fecha (Día, Mes, Año)
        fecha_container = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.045),
            spacing=5
        )
        
        # Día
        dia_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.33, 1),
            spacing=1
        )
        
        label_dia = Label(
            text='Día',
            font_size=11,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_dia.bind(size=label_dia.setter('text_size'))
        dia_box.add_widget(label_dia)
        
        # Obtener día actual
        dia_actual = datetime.now().strftime('%d')
        self.dia_input = TextInput(
            text=dia_actual,
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        dia_box.add_widget(self.dia_input)
        fecha_container.add_widget(dia_box)
        
        # Mes
        mes_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.33, 1),
            spacing=1
        )
        
        label_mes = Label(
            text='Mes',
            font_size=11,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_mes.bind(size=label_mes.setter('text_size'))
        mes_box.add_widget(label_mes)
        
        # Obtener mes actual
        mes_actual = datetime.now().strftime('%m')
        self.mes_input = TextInput(
            text=mes_actual,
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        mes_box.add_widget(self.mes_input)
        fecha_container.add_widget(mes_box)
        
        # Año
        año_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.34, 1),
            spacing=1
        )
        
        label_año = Label(
            text='Año',
            font_size=11,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_año.bind(size=label_año.setter('text_size'))
        año_box.add_widget(label_año)
        
        # Obtener año actual
        año_actual = datetime.now().strftime('%Y')
        self.año_input = TextInput(
            text=año_actual,
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        año_box.add_widget(self.año_input)
        fecha_container.add_widget(año_box)
        
        self.add_widget(fecha_container)
        
        # Botón para restablecer fecha actual
        btn_fecha_actual = Button(
            text='[b]📅 RESTABLECER FECHA ACTUAL[/b]',
            markup=True,
            font_size=13,
            size_hint=(1, 0.03),
            background_color=(0.6, 0.6, 0.6, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_fecha_actual.bind(on_press=self.establecer_fecha_actual)
        self.add_widget(btn_fecha_actual)
        

        
        # Sección: AGREGAR HORARIO
        label_seccion2 = Label(
            text='[b]AGREGAR HORARIO[/b]',
            markup=True,
            font_size=16,
            size_hint=(1, 0.025),
            color=(0.1, 0.4, 0.8, 1),
            halign='left',
            valign='middle'
        )
        label_seccion2.bind(size=label_seccion2.setter('text_size'))
        self.add_widget(label_seccion2)
        
        # Hora de entrada
        label_entrada = Label(
            text='HORA DE ENTRADA',
            font_size=13,
            size_hint=(1, 0.02),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_entrada.bind(size=label_entrada.setter('text_size'))
        self.add_widget(label_entrada)
        
        # Campos de entrada (horas y minutos)
        entrada_container = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.05),
            spacing=5
        )
        
        # Hora entrada
        hora_entrada_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.5, 1),
            spacing=1
        )
        
        label_hora_entrada = Label(
            text='Horas',
            font_size=12,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_hora_entrada.bind(size=label_hora_entrada.setter('text_size'))
        hora_entrada_box.add_widget(label_hora_entrada)
        
        self.hora_entrada_input = TextInput(
            hint_text='09',
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        hora_entrada_box.add_widget(self.hora_entrada_input)
        entrada_container.add_widget(hora_entrada_box)
        
        # Minutos entrada
        min_entrada_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.5, 1),
            spacing=1
        )
        
        label_min_entrada = Label(
            text='Minutos',
            font_size=12,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_min_entrada.bind(size=label_min_entrada.setter('text_size'))
        min_entrada_box.add_widget(label_min_entrada)
        
        self.min_entrada_input = TextInput(
            hint_text='00',
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        min_entrada_box.add_widget(self.min_entrada_input)
        entrada_container.add_widget(min_entrada_box)
        
        self.add_widget(entrada_container)
        
        # Hora de salida
        label_salida = Label(
            text='HORA DE SALIDA',
            font_size=13,
            size_hint=(1, 0.02),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_salida.bind(size=label_salida.setter('text_size'))
        self.add_widget(label_salida)
        
        # Campos de salida (horas y minutos)
        salida_container = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.05),
            spacing=5
        )
        
        # Hora salida
        hora_salida_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.5, 1),
            spacing=1
        )
        
        label_hora_salida = Label(
            text='Horas',
            font_size=12,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_hora_salida.bind(size=label_hora_salida.setter('text_size'))
        hora_salida_box.add_widget(label_hora_salida)
        
        self.hora_salida_input = TextInput(
            hint_text='13',
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        hora_salida_box.add_widget(self.hora_salida_input)
        salida_container.add_widget(hora_salida_box)
        
        # Minutos salida
        min_salida_box = BoxLayout(
            orientation='vertical',
            size_hint=(0.5, 1),
            spacing=1
        )
        
        label_min_salida = Label(
            text='Minutos',
            font_size=12,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='center',
            valign='middle'
        )
        label_min_salida.bind(size=label_min_salida.setter('text_size'))
        min_salida_box.add_widget(label_min_salida)
        
        self.min_salida_input = TextInput(
            hint_text='30',
            multiline=False,
            font_size=16,
            input_filter='int',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5],
            halign='center'
        )
        min_salida_box.add_widget(self.min_salida_input)
        salida_container.add_widget(min_salida_box)
        
        self.add_widget(salida_container)
        
        
        # Precio por hora
        precio_box = BoxLayout(
            orientation='vertical',
            size_hint=(1, 0.045),
            spacing=1
        )
        
        label_precio = Label(
            text='Precio por Hora ($)',
            font_size=13,
            size_hint=(1, 0.3),
            color=(1, 1, 1, 1),
            halign='left',
            valign='middle'
        )
        label_precio.bind(size=label_precio.setter('text_size'))
        precio_box.add_widget(label_precio)
        
        self.precio_input = TextInput(
            hint_text='Ej: 15.50',
            multiline=False,
            font_size=15,
            input_filter='float',
            size_hint=(1, 0.7),
            background_color=(1, 1, 1, 1),
            foreground_color=(0.1, 0.1, 0.1, 1),
            cursor_color=(0.1, 0.4, 0.8, 1),
            padding=[10, 5]
        )
        precio_box.add_widget(self.precio_input)
        self.add_widget(precio_box)
        
        # Botón AGREGAR HORARIO
        btn_agregar = Button(
            text='[b]AGREGAR HORARIO[/b]',
            markup=True,
            font_size=16,
            size_hint=(1, 0.045),
            background_color=(0.1, 0.4, 0.8, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_agregar.bind(on_press=self.agregar_horario)
        self.add_widget(btn_agregar)
        
        # Botón para calcular precio total
        btn_calcular = Button(
            text='[b]💰 CALCULAR PRECIO TOTAL[/b]',
            markup=True,
            font_size=15,
            size_hint=(1, 0.04),
            background_color=(0, 0.6, 0, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_calcular.bind(on_press=self.calcular_precio_total)
        self.add_widget(btn_calcular)
        
        self.add_widget(Widget(size_hint=(1, 0.005)))
        
        # Sección: REGISTROS
        label_lista = Label(
            text='[b]HORARIOS AGREGADOS[/b]',
            markup=True,
            font_size=15,
            size_hint=(1, 0.025),
            color=(0.1, 0.4, 0.8, 1),
            halign='left',
            valign='middle'
        )
        label_lista.bind(size=label_lista.setter('text_size'))
        self.add_widget(label_lista)
        
        # Lista de registros
        scroll = ScrollView(
            size_hint=(1, 0.12),
            bar_width=5
        )
        self.lista_layout = GridLayout(
            cols=1,
            spacing=2,
            size_hint_y=None
        )
        self.lista_layout.bind(minimum_height=self.lista_layout.setter('height'))
        scroll.add_widget(self.lista_layout)
        self.add_widget(scroll)
        
        self.add_widget(Widget(size_hint=(1, 0.005)))
        
        # Sección: RESULTADOS
        resultados_box = BoxLayout(
            orientation='vertical',
            size_hint=(1, 0.1),
            padding=[5, 5],
            spacing=2
        )
        
        # Total horas
        horas_total_box = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.45),
            spacing=10
        )
        
        label_horas_total = Label(
            text='[b]HORAS TOTALES[/b]',
            markup=True,
            font_size=15,
            size_hint=(0.5, 1),
            color=(0.1, 0.4, 0.8, 1),
            halign='left',
            valign='middle'
        )
        label_horas_total.bind(size=label_horas_total.setter('text_size'))
        horas_total_box.add_widget(label_horas_total)
        
        self.horas_total_label = Label(
            text='[b]0.0 h[/b]',
            markup=True,
            font_size=17,
            size_hint=(0.5, 1),
            color=(0, 0.6, 0, 1),
            halign='right',
            valign='middle'
        )
        self.horas_total_label.bind(size=self.horas_total_label.setter('text_size'))
        horas_total_box.add_widget(self.horas_total_label)
        resultados_box.add_widget(horas_total_box)
        
        # Total precio
        precio_total_box = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.45),
            spacing=10
        )
        
        label_precio_total = Label(
            text='[b]TOTAL A COBRAR[/b]',
            markup=True,
            font_size=15,
            size_hint=(0.5, 1),
            color=(0.1, 0.4, 0.8, 1),
            halign='left',
            valign='middle'
        )
        label_precio_total.bind(size=label_precio_total.setter('text_size'))
        precio_total_box.add_widget(label_precio_total)
        
        self.precio_total_label = Label(
            text='[b]$0.00[/b]',
            markup=True,
            font_size=18,
            size_hint=(0.5, 1),
            color=(0, 0.6, 0, 1),
            halign='right',
            valign='middle'
        )
        self.precio_total_label.bind(size=self.precio_total_label.setter('text_size'))
        precio_total_box.add_widget(self.precio_total_label)
        resultados_box.add_widget(precio_total_box)
        
        self.add_widget(resultados_box)
        
        # Botones de acción
        action_box = BoxLayout(
            orientation='horizontal',
            size_hint=(1, 0.06),
            spacing=5
        )
        
        btn_limpiar = Button(
            text='[b]LIMPIAR[/b]',
            markup=True,
            font_size=14,
            background_color=(0.8, 0.2, 0.2, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_limpiar.bind(on_press=self.limpiar_todo)
        action_box.add_widget(btn_limpiar)
        
        btn_reporte = Button(
            text='[b]REPORTE[/b]',
            markup=True,
            font_size=14,
            background_color=(0.1, 0.4, 0.8, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_reporte.bind(on_press=self.generar_reporte)
        action_box.add_widget(btn_reporte)
        
        btn_excel = Button(
            text='[b]📊 EXCEL[/b]',
            markup=True,
            font_size=14,
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        btn_excel.bind(on_press=self.guardar_excel)
        action_box.add_widget(btn_excel)
        
        self.add_widget(action_box)
        
        # Footer
        footer = Label(
            text=' Sistema de Registro de Horas v4.0',
            font_size=11,
            size_hint=(1, 0.02),
            color=(0.5, 0.5, 0.5, 1),
            halign='center',
            valign='middle'
        )
        footer.bind(size=footer.setter('text_size'))
        self.add_widget(footer)
        
        # Vincular eventos para calcular horas en tiempo real
        self.hora_entrada_input.bind(text=self.verificar_horario)
        self.min_entrada_input.bind(text=self.verificar_horario)
        self.hora_salida_input.bind(text=self.verificar_horario)
        self.min_salida_input.bind(text=self.verificar_horario)
        
        # Vincular eventos para actualizar fecha formateada
        self.dia_input.bind(text=self.actualizar_fecha_formateada)
        self.mes_input.bind(text=self.actualizar_fecha_formateada)
        self.año_input.bind(text=self.actualizar_fecha_formateada)
    
    def actualizar_fecha_formateada(self, instance=None, value=None):
        """Actualiza la fecha formateada en tiempo real"""
        dia = self.dia_input.text.strip().zfill(2)
        mes = self.mes_input.text.strip().zfill(2)
        año = self.año_input.text.strip()
        
        if dia and mes and año:
            try:
                # Validar fecha
                fecha_str = f"{dia}/{mes}/{año}"
                datetime.strptime(fecha_str, '%d/%m/%Y')
                self.fecha_formateada_label.text = f'📅 Fecha: {fecha_str}'
                self.fecha_formateada_label.color = (0.1, 0.4, 0.8, 1)
            except:
                self.fecha_formateada_label.text = '📅 Fecha inválida'
                self.fecha_formateada_label.color = (0.8, 0.2, 0.2, 1)
    
    def calcular_horas(self, h_entrada, m_entrada, h_salida, m_salida):
        """Calcula las horas trabajadas a partir de horas y minutos"""
        try:
            h1 = int(h_entrada) if h_entrada else 0
            m1 = int(m_entrada) if m_entrada else 0
            h2 = int(h_salida) if h_salida else 0
            m2 = int(m_salida) if m_salida else 0
            
            # Validar rangos
            if not (0 <= h1 <= 23 and 0 <= m1 <= 59 and 0 <= h2 <= 23 and 0 <= m2 <= 59):
                return None
            
            # Convertir a minutos totales
            minutos_inicio = h1 * 60 + m1
            minutos_fin = h2 * 60 + m2
            
            # Calcular diferencia en horas
            if minutos_fin > minutos_inicio:
                diferencia = minutos_fin - minutos_inicio
            else:
                # Si es menor, asumimos que pasa al día siguiente
                diferencia = (24 * 60 - minutos_inicio) + minutos_fin
            
            horas = diferencia / 60.0
            return round(horas, 2)
            
        except:
            return None
    
    def verificar_horario(self, instance=None, value=None):
        """Verifica el horario ingresado y muestra las horas calculadas"""
        h_entrada = self.hora_entrada_input.text.strip()
        m_entrada = self.min_entrada_input.text.strip()
        h_salida = self.hora_salida_input.text.strip()
        m_salida = self.min_salida_input.text.strip()
        
        if h_entrada and m_entrada and h_salida and m_salida:
            horas = self.calcular_horas(h_entrada, m_entrada, h_salida, m_salida)
            if horas is not None and horas > 0:
                self.horas_calculadas_label.text = f'⏱ Horas calculadas: {horas:.1f}h'
                self.horas_calculadas_label.color = (0, 0.6, 0, 1)
                return horas
            else:
                self.horas_calculadas_label.text = '⏱ Formato inválido'
                self.horas_calculadas_label.color = (0.8, 0.2, 0.2, 1)
                return None
        else:
            self.horas_calculadas_label.text = '⏱ Horas calculadas: 0.0h'
            self.horas_calculadas_label.color = (0, 0.6, 0, 1)
            return None
    
    def establecer_fecha_actual(self, instance):
        """Establece la fecha actual en los campos de fecha"""
        ahora = datetime.now()
        self.dia_input.text = ahora.strftime('%d')
        self.mes_input.text = ahora.strftime('%m')
        self.año_input.text = ahora.strftime('%Y')
        self.fecha_formateada_label.text = f'📅 Fecha: {ahora.strftime("%d/%m/%Y")}'
        self.fecha_formateada_label.color = (0.1, 0.4, 0.8, 1)
        
        self.mostrar_mensaje('Actualizado', f'📅 Fecha actualizada a:\n{ahora.strftime("%d/%m/%Y")}')
    
    def obtener_fecha_completa(self):
        """Obtiene la fecha completa formateada"""
        dia = self.dia_input.text.strip().zfill(2)
        mes = self.mes_input.text.strip().zfill(2)
        año = self.año_input.text.strip()
        
        if dia and mes and año:
            try:
                fecha_str = f"{dia}/{mes}/{año}"
                datetime.strptime(fecha_str, '%d/%m/%Y')
                return fecha_str
            except:
                return None
        return None
    
    def agregar_horario(self, instance):
        """Agrega un registro de horario trabajado"""
        cliente = self.cliente_input.text.strip()
        fecha = self.obtener_fecha_completa()
        h_entrada = self.hora_entrada_input.text.strip()
        m_entrada = self.min_entrada_input.text.strip()
        h_salida = self.hora_salida_input.text.strip()
        m_salida = self.min_salida_input.text.strip()
        precio_text = self.precio_input.text.strip()
        
        # Validaciones
        if not cliente:
            self.mostrar_mensaje('Error', '⚠️ Ingresa el nombre del cliente')
            return
        
        if not fecha:
            self.mostrar_mensaje('Error', '⚠️ Fecha inválida. Verifica día, mes y año')
            return
            
        if not h_entrada or not m_entrada or not h_salida or not m_salida:
            self.mostrar_mensaje('Error', '⚠️ Ingresa todos los campos de horario')
            return
        
        if not precio_text:
            self.mostrar_mensaje('Error', '⚠️ Ingresa el precio por hora')
            return
        
        # Calcular horas
        horas = self.calcular_horas(h_entrada, m_entrada, h_salida, m_salida)
        if horas is None or horas <= 0:
            self.mostrar_mensaje('Error', '⚠️ Horario inválido. Verifica las horas y minutos')
            return
        
        try:
            precio = float(precio_text)
            if precio <= 0:
                self.mostrar_mensaje('Error', '⚠️ El precio debe ser mayor a 0')
                return
        except ValueError:
            self.mostrar_mensaje('Error', '⚠️ Precio inválido')
            return
        
        # Formatear horario
        horario = f"{h_entrada.zfill(2)}:{m_entrada.zfill(2)} - {h_salida.zfill(2)}:{m_salida.zfill(2)}"
        
        # Calcular subtotal
        subtotal = horas * precio
        
        # Agregar registro
        self.contador += 1
        self.registros.append({
            'id': self.contador,
            'cliente': cliente,
            'fecha': fecha,
            'hora_entrada': f"{h_entrada.zfill(2)}:{m_entrada.zfill(2)}",
            'hora_salida': f"{h_salida.zfill(2)}:{m_salida.zfill(2)}",
            'horario': horario,
            'horas': horas,
            'precio': precio,
            'subtotal': subtotal
        })
        
        # Actualizar totales
        self.total_horas += horas
        self.total_precio += subtotal
        
        self.actualizar_lista()
        self.actualizar_totales()
        
        # Limpiar campos de horario y precio
        self.hora_entrada_input.text = ''
        self.min_entrada_input.text = ''
        self.hora_salida_input.text = ''
        self.min_salida_input.text = ''
        self.precio_input.text = ''
        self.horas_calculadas_label.text = '⏱ Horas calculadas: 0.0h'
        self.horas_calculadas_label.color = (0, 0.6, 0, 1)
        self.hora_entrada_input.focus = True
        
        self.mostrar_mensaje('Éxito', f'✅ Horario agregado:\n{cliente}\n{horario}\n{horas}h a ${precio:.2f}/h')
    
    def calcular_precio_total(self, instance):
        """Calcula el precio total basado en todas las horas agregadas"""
        if not self.registros:
            self.mostrar_mensaje('Info', 'ℹ️ No hay horarios agregados para calcular')
            return
        
        # Verificar si todos los registros tienen el mismo precio por hora
        precios = [reg['precio'] for reg in self.registros]
        precios_unicos = list(set(precios))
        
        if len(precios_unicos) == 1:
            # Todos tienen el mismo precio
            precio_por_hora = precios_unicos[0]
            total = self.total_horas * precio_por_hora
            self.total_precio = total
            self.actualizar_totales()
            self.mostrar_mensaje('Cálculo Completado', 
                               f'✅ Precio total calculado:\n\n'
                               f'⏰ Horas totales: {self.total_horas:.1f}h\n'
                               f'💰 Precio por hora: ${precio_por_hora:.2f}\n'
                               f'💰 TOTAL: ${total:.2f}')
        else:
            # Precios diferentes, usar la suma de subtotales
            self.total_precio = sum(reg['subtotal'] for reg in self.registros)
            self.actualizar_totales()
            self.mostrar_mensaje('Cálculo Completado', 
                               f'✅ Precio total calculado:\n\n'
                               f'⏰ Horas totales: {self.total_horas:.1f}h\n'
                               f'💰 TOTAL: ${self.total_precio:.2f}\n\n'
                               f'ℹ️ Se usaron {len(precios_unicos)} precios diferentes')
    
    def eliminar_registro(self, index):
        """Elimina un registro de la lista"""
        if 0 <= index < len(self.registros):
            registro = self.registros[index]
            self.total_horas -= registro['horas']
            self.total_precio -= registro['subtotal']
            del self.registros[index]
            self.actualizar_lista()
            self.actualizar_totales()
    
    def actualizar_lista(self):
        """Actualiza la lista de registros en la interfaz"""
        self.lista_layout.clear_widgets()
        
        for registro in self.registros:
            card = BoxLayout(
                orientation='horizontal',
                size_hint_y=None,
                height=45,
                padding=[6, 2],
                spacing=4
            )
            
            with card.canvas.before:
                Color(0.95, 0.95, 0.95, 1)
                self.rect = RoundedRectangle(pos=card.pos, size=card.size, radius=[5])
            card.bind(pos=self._update_rect, size=self._update_rect)
            
            # Información del registro
            info_text = f'#{registro["id"]} {registro["cliente"]}\n{registro["fecha"]}  {registro["hora_entrada"]}-{registro["hora_salida"]}'
            label_info = Label(
                text=info_text,
                font_size=10,
                size_hint=(0.45, 1),
                halign='left',
                valign='middle',
                color=(0.1, 0.1, 0.1, 1)
            )
            label_info.bind(size=label_info.setter('text_size'))
            card.add_widget(label_info)
            
            # Horas y precio
            info_precio = f'{registro["horas"]}h\n${registro["subtotal"]:.2f}'
            label_precio = Label(
                text=info_precio,
                font_size=12,
                size_hint=(0.35, 1),
                halign='right',
                valign='middle',
                color=(0, 0.6, 0, 1)
            )
            label_precio.bind(size=label_precio.setter('text_size'))
            card.add_widget(label_precio)
            
            # Botón eliminar
            btn_elim = Button(
                text='✕',
                font_size=11,
                size_hint=(0.2, 0.8),
                background_color=(0.8, 0.2, 0.2, 1),
                color=(1, 1, 1, 1),
                background_normal='',
                pos_hint={'center_y': 0.5}
            )
            idx = self.registros.index(registro)
            btn_elim.bind(on_press=lambda btn, i=idx: self.eliminar_registro(i))
            card.add_widget(btn_elim)
            
            self.lista_layout.add_widget(card)
    
    def _update_rect(self, instance, value):
        """Actualiza el rectángulo de fondo de las tarjetas"""
        if hasattr(instance, 'rect'):
            instance.rect.pos = instance.pos
            instance.rect.size = instance.size
    
    def actualizar_totales(self):
        """Actualiza los totales en la interfaz"""
        self.horas_total_label.text = f'[b]{self.total_horas:.1f} h[/b]'
        self.precio_total_label.text = f'[b]${self.total_precio:.2f}[/b]'
    
    def guardar_excel(self, instance):
        """Guarda los registros en un archivo CSV (más ligero para Android)"""
        if not self.registros:
            self.mostrar_mensaje('Info', 'ℹ️ No hay registros para guardar')
            return
        
        try:
            import csv
            fecha_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
            nombre_archivo = f'registro_horas_{fecha_actual}.csv'
            
            with open(nombre_archivo, 'w', newline='', encoding='utf-8') as archivo:
                writer = csv.writer(archivo)
                # Encabezados
                writer.writerow(['ID', 'Cliente', 'Fecha', 'Hora Entrada', 'Hora Salida', 
                            'Horas', 'Precio', 'Subtotal'])
                
                # Datos
                for reg in self.registros:
                    writer.writerow([
                        reg['id'], reg['cliente'], reg['fecha'], 
                        reg['hora_entrada'], reg['hora_salida'], 
                        reg['horas'], reg['precio'], reg['subtotal']
                    ])
            
            self.mostrar_mensaje('Éxito', f'✅ Guardado: {nombre_archivo}')
        except Exception as e:
            self.mostrar_mensaje('Error', f'❌ Error al guardar:\n{str(e)}')
            
        """Guarda los registros en un archivo CSV (más ligero para Android)"""
        if not self.registros:
            self.mostrar_mensaje('Info', 'ℹ️ No hay registros para guardar')
            return
        
        try:
            import csv
            fecha_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
            nombre_archivo = f'registro_horas_{fecha_actual}.csv'
            
            with open(nombre_archivo, 'w', newline='', encoding='utf-8') as archivo:
                writer = csv.writer(archivo)
                # Encabezados
                writer.writerow(['ID', 'Cliente', 'Fecha', 'Hora Entrada', 'Hora Salida', 
                            'Horas', 'Precio', 'Subtotal'])
                
                # Datos
                for reg in self.registros:
                    writer.writerow([
                        reg['id'], reg['cliente'], reg['fecha'], 
                        reg['hora_entrada'], reg['hora_salida'], 
                        reg['horas'], reg['precio'], reg['subtotal']
                    ])
            
            self.mostrar_mensaje('Éxito', f'✅ Guardado: {nombre_archivo}')
        except Exception as e:
            self.mostrar_mensaje('Error', f'❌ Error al guardar:\n{str(e)}')


    def generar_reporte(self, instance):
        """Genera un reporte detallado de todos los registros"""
        if not self.registros:
            self.mostrar_mensaje('Info', 'ℹ️ No hay registros para generar reporte')
            return
        
        resumen = '📋 REPORTE DE HORAS\n'
        resumen += '═' * 35 + '\n\n'
        
        # Agrupar por cliente
        clientes = {}
        for reg in self.registros:
            if reg['cliente'] not in clientes:
                clientes[reg['cliente']] = []
            clientes[reg['cliente']].append(reg)
        
        for cliente, registros in clientes.items():
            resumen += f'👤 {cliente}\n'
            resumen += '─' * 30 + '\n'
            total_cliente_horas = 0
            total_cliente_precio = 0
            
            for reg in registros:
                resumen += f'  📅 {reg["fecha"]}\n'
                resumen += f'     ⏰ {reg["hora_entrada"]} - {reg["hora_salida"]}\n'
                resumen += f'     {reg["horas"]}h × ${reg["precio"]:.2f} = ${reg["subtotal"]:.2f}\n\n'
                total_cliente_horas += reg['horas']
                total_cliente_precio += reg['subtotal']
            
            resumen += f'  📊 Total: {total_cliente_horas:.1f}h = ${total_cliente_precio:.2f}\n\n'
        
        resumen += '═' * 35 + '\n'
        resumen += f'💰 TOTAL GENERAL: ${self.total_precio:.2f}\n'
        resumen += f'⏰ HORAS TOTALES: {self.total_horas:.1f}h\n'
        resumen += '═' * 35 + '\n'
        resumen += '✅ Reporte generado exitosamente!'
        
        self.mostrar_mensaje('📊 Reporte Detallado', resumen)
    
    def limpiar_todo(self, instance):
        """Limpia todos los registros y campos"""
        if self.registros:
            self.registros.clear()
            self.total_horas = 0.0
            self.total_precio = 0.0
            self.contador = 0
            self.actualizar_lista()
            self.actualizar_totales()
            
            # Limpiar campos de entrada
            self.cliente_input.text = ''
            # No limpiar la fecha, mantener la actual
            self.hora_entrada_input.text = ''
            self.min_entrada_input.text = ''
            self.hora_salida_input.text = ''
            self.min_salida_input.text = ''
            self.precio_input.text = ''
            self.horas_calculadas_label.text = '⏱ Horas calculadas: 0.0h'
            self.horas_calculadas_label.color = (0, 0.6, 0, 1)
            
            self.mostrar_mensaje('Limpiado', '🔄 Todos los registros han sido eliminados')
        else:
            self.mostrar_mensaje('Info', 'ℹ️ No hay registros para limpiar')
    
    def mostrar_mensaje(self, titulo, mensaje):
        """Muestra un mensaje emergente"""
        content = BoxLayout(
            orientation='vertical',
            padding=[20, 15],
            spacing=15
        )
        
        msg_label = Label(
            text=mensaje,
            font_size=15,
            halign='center',
            valign='middle',
            color=(0.1, 0.1, 0.1, 1)
        )
        msg_label.bind(size=msg_label.setter('text_size'))
        content.add_widget(msg_label)
        
        btn_cerrar = Button(
            text='ACEPTAR',
            font_size=16,
            size_hint=(1, 0.35),
            background_color=(0.1, 0.4, 0.8, 1),
            color=(1, 1, 1, 1),
            background_normal=''
        )
        
        popup = Popup(
            title=titulo,
            content=content,
            size_hint=(0.85, 0.5),
            title_size=18,
            title_color=(0.1, 0.4, 0.8, 1)
        )
        
        btn_cerrar.bind(on_press=popup.dismiss)
        content.add_widget(btn_cerrar)
        popup.open()

if __name__ == '__main__':
    HorasApp().run()