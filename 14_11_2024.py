import os
import tempfile
import shutil
import time  
from datetime import datetime, timedelta
from collections import Counter

import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from scipy import stats
from scipy.stats import norm
from sklearn.linear_model import LinearRegression, HuberRegressor
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor, IsolationForest, GradientBoostingRegressor
from sklearn.preprocessing import LabelEncoder, StandardScaler
from sklearn.model_selection import train_test_split, GridSearchCV, TimeSeriesSplit
from sklearn.metrics import accuracy_score, mean_squared_error, r2_score
import joblib

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

from statsmodels.tsa.seasonal import seasonal_decompose
import statsmodels.api as sm

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
#from tkinterweb import HtmlFrame

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk




# Variables globales
turno_combobox = None
name_combobox = None
cargo_entry = None
rotation_letter_entry = None
cal = None
jornada_var = None
btn_mañana = None
btn_tarde = None
btn_noche = None

global_variables = {
    'turno_combobox': None,
    'name_combobox': None,
    'cargo_entry': None,
    'rotation_letter_entry': None,
    'cal': None,
    'jornada_var': None,
    'btn_mañana': None,
    'btn_tarde': None,
    'btn_noche': None,
    'data_display': None,
    'data_display_corrective': None,
    'data_display_preventive': None,
    'range_display': None,
    'warning_label': None,
    'table_frame': None,
    'cal_inicio_resumen': None,
    'cal_fin_resumen': None,
    'rotation_result_frame': None,
    'rotation_text': None,
    'prediction_result_frame': None,
    'prediction_text': None,
    'metrics_result_frame': None,
    'metrics_figure': None,
    'metrics_canvas': None,
    'absence_frame': None,
    'absence_text': None,
    'team_frame': None,
    'team_figure': None,
    'team_canvas': None,
    'resumen_tab': None,
    'ai_text': None,
    'ai_status_label': None,
    'ai_config': None
}



# Constantes
DIAS_SEMANA = {
    0: 'Lunes',
    1: 'Martes',
    2: 'Miércoles',
    3: 'Jueves',
    4: 'Viernes',
    5: 'Sábado',
    6: 'Domingo'
}

MESES = {
    1: 'Enero',
    2: 'Febrero',
    3: 'Marzo',
    4: 'Abril',
    5: 'Mayo',
    6: 'Junio',
    7: 'Julio',
    8: 'Agosto',
    9: 'Septiembre',
    10: 'Octubre',
    11: 'Noviembre',
    12: 'Diciembre'
}

# Ruta del archivo Excel
#excel_file_path = r'C:\Users\mayko\Downloads\DB.xlsx'
excel_file_path = r'C:\Users\mayko\Downloads\DB.xlsx'
#excel_file_path = r'\\nt_metro\Metro_Shared\Mantenimiento\Sistemas y Energia Electrica\SEÑALIZACION Y PILOTAJE AUTOMATICO\Hans\Pruebas_app\DB.xlsx'

# Directorio temporal
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'turnos_graphs')

# Variables globales para widgets
# UI Elements
data_display = None
data_display_corrective = None
data_display_preventive = None
range_display = None
warning_label = None
table_frame = None
# Variables globales
data_display = None
data_display_corrective = None
data_display_preventive = None
range_display = None
warning_label = None
table_frame = None

# Calendar widgets
cal = None
cal_inicio_resumen = None
cal_fin_resumen = None

# Comboboxes and entries
turno_combobox = None
name_combobox = None
cargo_entry = None
rotation_letter_entry = None
jornada_var = None

# Buttons
btn_mañana = None
btn_tarde = None
btn_noche = None
jornada_var = None

# Analysis tabs
rotation_result_frame = None
rotation_text = None
prediction_result_frame = None
prediction_text = None
metrics_result_frame = None
metrics_figure = None
metrics_canvas = None
absence_frame = None
absence_text = None
team_frame = None
team_figure = None
team_canvas = None
resumen_tab = None

# AI related
ai_text = None
ai_status_label = None
ai_config = None
ai_text = None
ai_status_label = None
ai_config = None

# Variables globales adicionales a agregar al inicio del archivo:
rotation_result_frame = None
rotation_text = None
prediction_result_frame = None
prediction_text = None
metrics_result_frame = None
metrics_figure = None
metrics_canvas = None

# Variables globales para widgets
absence_frame = None
absence_text = None
team_frame = None
team_figure = None
team_canvas = None

resumen_tab = None  # Para guardar referencia a la pestaña de resumen

# Definición del directorio temporal
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'turnos_graphs')


def pruebas():
    print("Función pruebas ejecutada")  


def setup_log_tab(notebook):
    """Configura la pestaña de logs en el centro de administración"""
    tab = ttk.Frame(notebook)
    notebook.add(tab, text="Log del Sistema")
    
    # Configurar el grid
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(1, weight=1)
    
    # Frame para filtros
    filter_frame = ttk.LabelFrame(tab, text="Filtros", padding=10)
    filter_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=5)
    
    # Filtro por usuario
    ttk.Label(filter_frame, text="Usuario:").pack(side='left', padx=5)
    user_var = tk.StringVar()
    user_combo = ttk.Combobox(filter_frame, textvariable=user_var, width=15)
    user_combo.pack(side='left', padx=5)
    
    # Filtro por fecha
    ttk.Label(filter_frame, text="Desde:").pack(side='left', padx=5)
    date_from = ttk.Entry(filter_frame, width=12)
    date_from.pack(side='left', padx=5)
    
    ttk.Label(filter_frame, text="Hasta:").pack(side='left', padx=5)
    date_to = ttk.Entry(filter_frame, width=12)
    date_to.pack(side='left', padx=5)
    
    # Botón de filtrar
    ttk.Button(
        filter_frame,
        text="Filtrar",
        command=lambda: refresh_log_table(tree, user_var.get(), date_from.get(), date_to.get()),
        style='Primary.TButton'
    ).pack(side='left', padx=5)
    
    # Frame para la tabla
    table_frame = ttk.Frame(tab)
    table_frame.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)
    
    # Crear tabla
    tree = ttk.Treeview(
        table_frame,
        columns=('usuario', 'fecha', 'accion'),
        show='headings',
        height=20
    )
    
    # Configurar columnas
    tree.heading('usuario', text='Usuario')
    tree.heading('fecha', text='Fecha y Hora')
    tree.heading('accion', text='Acción')
    
    tree.column('usuario', width=150)
    tree.column('fecha', width=200)
    tree.column('accion', width=200)
    
    # Scrollbars
    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    # Grid
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    
    # Configurar grid del frame de la tabla
    table_frame.grid_columnconfigure(0, weight=1)
    table_frame.grid_rowconfigure(0, weight=1)
    
    # Frame para botones
    button_frame = ttk.Frame(tab)
    button_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=5)
    
    ttk.Button(
        button_frame,
        text="Exportar Log",
        command=lambda: export_log(tree),
        style='Primary.TButton'
    ).pack(side='left', padx=5)
    
    ttk.Button(
        button_frame,
        text="Actualizar",
        command=lambda: refresh_log_table(tree),
        style='Primary.TButton'
    ).pack(side='left', padx=5)
    
    # Cargar datos iniciales
    refresh_log_table(tree)
    update_user_filter(user_combo)
    
    return tab

def log_action(username, action):
    """Registra una acción en el log"""
    try:
        # Leer el Excel existente
        df = pd.read_excel(excel_file_path)
        
        # Crear nuevo registro
        new_log = pd.DataFrame({
            'nombre_log': [username],
            'fecha_log': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            'accion_log': [action]
        })
        
        # Concatenar el nuevo registro
        df = pd.concat([df, new_log], ignore_index=True)
        
        # Guardar el Excel actualizado
        df.to_excel(excel_file_path, index=False)
        
    except Exception as e:
        print(f"Error al registrar en log: {str(e)}")

def refresh_log_table(tree, user_filter=None, date_from=None, date_to=None):
    """Actualiza la tabla de logs con los filtros aplicados"""
    # Limpiar tabla actual
    for item in tree.get_children():
        tree.delete(item)
    
    try:
        # Leer el Excel
        df = pd.read_excel(excel_file_path)
        
        # Aplicar filtros si existen
        mask = pd.notnull(df['nombre_log'])  # Filtro base para registros de log
        
        if user_filter:
            mask &= (df['nombre_log'] == user_filter)
            
        if date_from:
            try:
                date_from = pd.to_datetime(date_from)
                mask &= (pd.to_datetime(df['fecha_log']) >= date_from)
            except:
                pass
                
        if date_to:
            try:
                date_to = pd.to_datetime(date_to)
                mask &= (pd.to_datetime(df['fecha_log']) <= date_to)
            except:
                pass
                
        # Aplicar filtros
        df_filtered = df[mask].sort_values('fecha_log', ascending=False)
        
        # Insertar datos filtrados en la tabla
        for i, row in df_filtered.iterrows():
            if pd.notnull(row['nombre_log']):  # Asegurar que es un registro de log
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                tree.insert('', 'end', values=(
                    row['nombre_log'],
                    row['fecha_log'],
                    row['accion_log']
                ), tags=(tag,))
        
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar logs: {str(e)}")

def export_log(tree):
    """Exporta los logs a un nuevo Excel"""
    try:
        # Obtener ruta para guardar
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Exportar Logs"
        )
        
        if file_path:
            # Crear DataFrame con los datos del tree
            data = []
            for item in tree.get_children():
                data.append(tree.item(item)['values'])
            
            df_export = pd.DataFrame(data, columns=['Usuario', 'Fecha y Hora', 'Acción'])
            
            # Exportar a Excel
            df_export.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "Logs exportados correctamente")
    
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar logs: {str(e)}")

def update_user_filter(combo):
    """Actualiza la lista de usuarios en el filtro"""
    try:
        # Leer el Excel
        df = pd.read_excel(excel_file_path)
        
        # Obtener usuarios únicos que tienen registros en el log
        users = df['nombre_log'].dropna().unique().tolist()
        
        # Actualizar valores del combo
        combo['values'] = [''] + users
    
    except Exception as e:
        print(f"Error al actualizar filtro de usuarios: {str(e)}")

def add_comparative_analysis(df, text_widget):
    """Agrega análisis comparativo entre jornadas"""
    text_widget.insert(tk.END, "\nANÁLISIS COMPARATIVO ENTRE JORNADAS:\n", "title")
   
    # Agrupar por jornada
    jornada_stats = {}
    for jornada in ['Mañana', 'Tarde', 'Noche']:
        df_jornada = df[df['Jornada'] == jornada]
        daily_counts = df_jornada.groupby('fecha_asignacion').size()
       
        jornada_stats[jornada] = {
            'promedio': daily_counts.mean(),
            'max': daily_counts.max(),
            'min': daily_counts.min(),
            'std': daily_counts.std()
        }
   
    # Mostrar comparativa
    for jornada, stats in jornada_stats.items():
        text_widget.insert(tk.END, f"\nJornada {jornada}:\n", "subtitle")
        text_widget.insert(tk.END, f"• Promedio: {stats['promedio']:.1f} técnicos\n")
        text_widget.insert(tk.END, f"• Máximo: {int(stats['max'])} técnicos\n")
        text_widget.insert(tk.END, f"• Mínimo: {int(stats['min'])} técnicos\n")
        text_widget.insert(tk.END, f"• Variabilidad: {stats['std']:.2f}\n")
       
def generate_detailed_report(file_path, results):
    """Genera un reporte PDF detallado con procesamiento de datos mejorado"""
    try:
        # Leer datos frescos del Excel para asegurar información actualizada
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        # Configurar el documento
        doc = SimpleDocTemplate(
            file_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
       
        elements = []
        styles = getSampleStyleSheet()
       
        # Título
        elements.append(Paragraph("Análisis Avanzado de Turnos", styles['Title']))
        elements.append(Spacer(1, 20))
       
        # Fecha del reporte
        elements.append(Paragraph(
            f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            styles['Normal']
        ))
        elements.append(Spacer(1, 20))
       
        # Vista General
        elements.append(Paragraph("Vista General", styles['Heading1']))
        elements.append(Spacer(1, 10))
       
        # Procesar datos para vista general
        daily_counts = df.groupby('fecha_asignacion').size()
       
        if not daily_counts.empty:
            stats_text = [
                f"Total de registros analizados: {len(df)}",
                f"Días totales analizados: {len(daily_counts)}",
                f"Promedio diario: {daily_counts.mean():.2f} técnicos",
                f"Máximo diario: {int(daily_counts.max())} técnicos",
                f"Mínimo diario: {int(daily_counts.min())} técnicos",
                f"Desviación estándar: {daily_counts.std():.2f}"
            ]
           
            # Agregar estadísticas por jornada
            for jornada in ['Mañana', 'Tarde', 'Noche']:
                df_jornada = df[df['Jornada'] == jornada]
                if not df_jornada.empty:
                    count_jornada = len(df_jornada)
                    stats_text.append(f"\nJornada {jornada}:")
                    stats_text.append(f"• Total técnicos: {count_jornada}")
                    stats_text.append(f"• Promedio: {count_jornada/len(daily_counts):.2f} técnicos/día")
           
            for text in stats_text:
                elements.append(Paragraph(text, styles['Normal']))
        else:
            elements.append(Paragraph("No hay datos disponibles para el análisis general", styles['Normal']))
       
        elements.append(Spacer(1, 20))
       
        # Predicciones
        elements.append(Paragraph("Predicciones", styles['Heading1']))
        elements.append(Spacer(1, 10))
       
        if results and 'predictions' in results and results['predictions']:
            pred = results['predictions']
            if 'dates' in pred and 'values' in pred:
                # Tabla de predicciones
                data = [['Fecha', 'Día', 'Técnicos Previstos']]
               
                for date, value in zip(pred['dates'], pred['values']):
                    dia_semana = {
                        'Monday': 'Lunes',
                        'Tuesday': 'Martes',
                        'Wednesday': 'Miércoles',
                        'Thursday': 'Jueves',
                        'Friday': 'Viernes',
                        'Saturday': 'Sábado',
                        'Sunday': 'Domingo'
                    }[date.strftime('%A')]
                   
                    data.append([
                        date.strftime('%d/%m/%Y'),
                        dia_semana,
                        f"{int(value)}"
                    ])
               
                table = Table(data, colWidths=[100, 100, 100])
                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ])
                table.setStyle(style)
                elements.append(table)
               
                # Agregar resumen estadístico de predicciones
                elements.append(Spacer(1, 10))
                pred_stats = [
                    f"\nResumen de Predicciones:",
                    f"• Promedio previsto: {np.mean(pred['values']):.2f} técnicos",
                    f"• Máximo previsto: {int(np.max(pred['values']))} técnicos",
                    f"• Mínimo previsto: {int(np.min(pred['values']))} técnicos",
                    f"• Variabilidad prevista: {np.std(pred['values']):.2f}"
                ]
                for text in pred_stats:
                    elements.append(Paragraph(text, styles['Normal']))
            else:
                elements.append(Paragraph("No hay datos de predicciones disponibles", styles['Normal']))
       
        elements.append(Spacer(1, 20))
       
        # Anomalías
        elements.append(Paragraph("Anomalías Detectadas", styles['Heading1']))
        elements.append(Spacer(1, 10))
       
        if results and 'anomalies' in results and results['anomalies']:
            anom = results['anomalies']
            if 'dates' in anom and 'values' in anom and len(anom['dates']) > 0:
                # Tabla de anomalías
                data = [['Fecha', 'Día', 'Cantidad', 'Tipo']]
                mean_value = np.mean(daily_counts)
               
                for date, value in zip(anom['dates'], anom['values']):
                    dia_semana = {
                        'Monday': 'Lunes',
                        'Tuesday': 'Martes',
                        'Wednesday': 'Miércoles',
                        'Thursday': 'Jueves',
                        'Friday': 'Viernes',
                        'Saturday': 'Sábado',
                        'Sunday': 'Domingo'
                    }[date.strftime('%A')]
                   
                    tipo = "Exceso" if value > mean_value else "Déficit"
                    data.append([
                        date.strftime('%d/%m/%Y'),
                        dia_semana,
                        f"{int(value)}",
                        tipo
                    ])
               
                table = Table(data, colWidths=[100, 100, 100, 100])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                elements.append(table)
               
                # Agregar resumen de anomalías
                elements.append(Spacer(1, 10))
                anom_stats = [
                    f"\nResumen de Anomalías:",
                    f"• Total anomalías detectadas: {len(anom['dates'])}",
                    f"• Promedio en anomalías: {np.mean(anom['values']):.2f}",
                    f"• Desviación en anomalías: {np.std(anom['values']):.2f}"
                ]
                for text in anom_stats:
                    elements.append(Paragraph(text, styles['Normal']))
            else:
                elements.append(Paragraph("No se detectaron anomalías en los datos", styles['Normal']))
        else:
            elements.append(Paragraph("No hay datos de anomalías disponibles", styles['Normal']))
       
        # Generar PDF
        doc.build(elements)
       
    except Exception as e:
        print(f"Error detallado: {str(e)}")  # Para debugging
        raise Exception(f"Error al generar el reporte PDF: {str(e)}")
   
   
def generate_excel_summary(file_path, results):
    """Genera un resumen en Excel del análisis"""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Vista General
            if 'patterns' in results and results['patterns']:
                patterns = results['patterns']
                if 'daily_counts' in patterns:
                    daily_df = pd.DataFrame(patterns['daily_counts'])
                    daily_df.to_excel(writer, sheet_name='Vista General', index=True)
           
            # Predicciones
            if 'predictions' in results and results['predictions']:
                pred = results['predictions']
                if 'dates' in pred and 'values' in pred:
                    pred_df = pd.DataFrame({
                        'Fecha': pred['dates'],
                        'Predicción': pred['values']
                    })
                    pred_df.to_excel(writer, sheet_name='Predicciones', index=False)
           
            # Anomalías
            if 'anomalies' in results and results['anomalies']:
                anom = results['anomalies']
                if 'dates' in anom and 'values' in anom:
                    anom_df = pd.DataFrame({
                        'Fecha': anom['dates'],
                        'Cantidad': anom['values']
                    })
                    anom_df.to_excel(writer, sheet_name='Anomalías', index=False)
                   
    except Exception as e:
        raise Exception(f"Error al generar el resumen Excel: {str(e)}")        


def add_efficiency_metrics(df, text_widget):
    """Calcula y muestra métricas de eficiencia"""
    text_widget.insert(tk.END, "\nMÉTRICAS DE EFICIENCIA:\n", "title")
   
    # Cálculo de métricas
    total_tecnicos = len(df['nombre'].unique())
    dias_totales = len(df['fecha_asignacion'].unique())
    utilizacion = df.groupby('fecha_asignacion').size().mean() / total_tecnicos
   
    text_widget.insert(tk.END, f"• Técnicos totales: {total_tecnicos}\n")
    text_widget.insert(tk.END, f"• Días analizados: {dias_totales}\n")
    text_widget.insert(tk.END, f"• Tasa de utilización: {utilizacion:.1%}\n")

def add_optimization_suggestions(df, text_widget):
    """Genera sugerencias de optimización basadas en el análisis"""
    text_widget.insert(tk.END, "\nSUGERENCIAS DE OPTIMIZACIÓN:\n", "title")
   
    # Análisis por día de la semana
    df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek
    carga_por_dia = df.groupby(['dia_semana', 'Jornada']).size().unstack(fill_value=0)
   
    dias = {
        0: 'Lunes', 1: 'Martes', 2: 'Miércoles',
        3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'
    }
   
    for dia in range(7):
        if dia in carga_por_dia.index:
            text_widget.insert(tk.END, f"\n{dias[dia]}:\n", "subtitle")
            for jornada in carga_por_dia.columns:
                carga = carga_por_dia.loc[dia, jornada]
                if carga > 0:
                    text_widget.insert(tk.END, f"• Jornada {jornada}: {int(carga)} técnicos\n")
                   
                    # Sugerencias específicas
                    if carga > carga_por_dia[jornada].mean() * 1.2:
                        text_widget.insert(tk.END, "  ⚠️ Considerar redistribuir carga\n", "warning")
                    elif carga < carga_por_dia[jornada].mean() * 0.8:
                        text_widget.insert(tk.END, "  ⚠️ Posible subfacturación\n", "warning")

def add_export_options(parent, results):
    """Agrega opciones de exportación avanzadas"""
    # Verificar si ya existe un frame de exportación y eliminarlo
    for widget in parent.winfo_children():
        if isinstance(widget, ttk.LabelFrame) and widget.cget("text") == "Exportar Análisis":
            widget.destroy()
           
    export_frame = ttk.LabelFrame(parent, text="Exportar Análisis")
    export_frame.pack(fill='x', padx=5, pady=5)
   
    def export_detailed_pdf():
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            if file_path:
                generate_detailed_report(file_path, results)
                messagebox.showinfo("Éxito", "Reporte PDF generado exitosamente")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar PDF: {str(e)}")
   
    def export_excel_summary():
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                generate_excel_summary(file_path, results)
                messagebox.showinfo("Éxito", "Resumen Excel generado exitosamente")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar Excel: {str(e)}")
   
    ttk.Button(
        export_frame,
        text="Exportar Reporte PDF",
        command=export_detailed_pdf,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)
   
    ttk.Button(
        export_frame,
        text="Exportar Excel",
        command=export_excel_summary,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

def add_search_functionality(parent, df):
    """Agrega funcionalidad de búsqueda avanzada"""
    search_frame = ttk.LabelFrame(parent, text="Búsqueda Avanzada")
    search_frame.pack(fill='x', padx=5, pady=5)
   
    # Variables de búsqueda
    search_var = tk.StringVar()
    jornada_var = tk.StringVar(value="todas")
   
    # Campo de búsqueda
    ttk.Entry(
        search_frame,
        textvariable=search_var
    ).pack(fill='x', pady=2)
   
    # Filtros
    filter_frame = ttk.Frame(search_frame)
    filter_frame.pack(fill='x', pady=2)
   
    ttk.Radiobutton(
        filter_frame,
        text="Todas",
        variable=jornada_var,
        value="todas"
    ).pack(side='left')
   
    for jornada in ['Mañana', 'Tarde', 'Noche']:
        ttk.Radiobutton(
            filter_frame,
            text=jornada,
            variable=jornada_var,
            value=jornada
        ).pack(side='left')

def run_analysis(config_vars, results_notebook):
    """Ejecuta el análisis completo con visualizaciones"""
    try:
        # Leer y preparar datos
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        # Filtrar por jornada si está especificada
        jornada = config_vars['jornada'].get()
        if jornada != "todas":
            df = df[df['Jornada'] == jornada]
           
        if df.empty:
            messagebox.showinfo("Información", "No hay datos para analizar")
            return

        # Obtener parámetros de configuración
        sensitivity = float(config_vars['sensitivity'].get())
        confidence = float(config_vars['confidence'].get())
        horizon = int(config_vars['horizon'].get())

        # 1. Análisis de patrones temporales
        daily_counts = df.groupby('fecha_asignacion').size()
        weekly_pattern = df.groupby(df['fecha_asignacion'].dt.dayofweek).size()
        monthly_pattern = df.groupby(df['fecha_asignacion'].dt.month).size()

        patterns_results = {
            'daily_counts': daily_counts,
            'weekly_pattern': weekly_pattern,
            'monthly_pattern': monthly_pattern
        }

        # 2. Generar predicciones
        predictions_results = generate_predictions_data(df, horizon, confidence)

        # 3. Detectar anomalías
        anomalies_results = detect_anomalies_data(df, sensitivity)

        # Agrupar resultados
        results = {
            'patterns': patterns_results,
            'predictions': predictions_results,
            'anomalies': anomalies_results
        }

        # Mostrar resultados en el notebook
        display_results(results_notebook, results)
       
        return results

    except Exception as e:
        messagebox.showerror("Error", f"Error en el análisis: {str(e)}")
        print(f"Error detallado: {str(e)}")
        return None

def generate_predictions_data(df, horizon, confidence):
    """Genera predicciones usando Random Forest con intervalos de confianza"""
    try:
        # Asegurar que tenemos suficientes datos
        if len(df) < 2:
            print("Insuficientes datos para predicciones")
            return None

        # Preparar datos
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        daily_counts = df.groupby('fecha_asignacion').size().reset_index()
        daily_counts.columns = ['fecha', 'cantidad']
       
        # Crear características
        daily_counts['dia_semana'] = daily_counts['fecha'].dt.dayofweek
        daily_counts['mes'] = daily_counts['fecha'].dt.month
        daily_counts['dia_mes'] = daily_counts['fecha'].dt.day
        daily_counts['semana'] = daily_counts['fecha'].dt.isocalendar().week
       
        # Preparar datos para el modelo
        X = daily_counts[['dia_semana', 'mes', 'dia_mes', 'semana']].values
        y = daily_counts['cantidad'].values
       
        if len(X) < 2:
            print("Insuficientes datos después del procesamiento")
            return None
       
        # Entrenar modelo con más árboles para mejor estabilidad
        model = RandomForestRegressor(
            n_estimators=200,  # Aumentado de 100 a 200
            max_depth=None,
            min_samples_split=2,
            min_samples_leaf=1,
            bootstrap=True,
            random_state=42,
            n_jobs=-1
        )
        model.fit(X, y)
       
        # Generar fechas futuras
        last_date = df['fecha_asignacion'].max()
        future_dates = pd.date_range(start=last_date + timedelta(days=1),
                                   periods=horizon,
                                   freq='D')
       
        # Preparar datos futuros
        future_X = np.array([
            [date.dayofweek,
             date.month,
             date.day,
             date.isocalendar().week]
            for date in future_dates
        ])
       
        # Generar predicciones base
        predictions = model.predict(future_X)
       
        # Generar predicciones de todos los árboles para calcular intervalos
        tree_predictions = np.array([
            tree.predict(future_X)
            for tree in model.estimators_
        ])
       
        # Calcular intervalos de confianza
        predictions_std = np.std(tree_predictions, axis=0)
        confidence_factor = norm.ppf((1 + confidence) / 2)
        confidence_interval = predictions_std * confidence_factor
       
        # Asegurar predicciones no negativas
        predictions = np.maximum(predictions, 0)
        lower_bound = np.maximum(predictions - confidence_interval, 0)
        upper_bound = predictions + confidence_interval
       
        # Redondear predicciones a enteros
        predictions = np.round(predictions).astype(int)
        lower_bound = np.round(lower_bound).astype(int)
        upper_bound = np.round(upper_bound).astype(int)
       
        # Agregar métricas de calidad
        prediction_quality = {
            'r2_score': model.score(X, y),
            'mean_std': np.mean(predictions_std),
            'feature_importance': dict(zip(
                ['dia_semana', 'mes', 'dia_mes', 'semana'],
                model.feature_importances_
            ))
        }
       
        print(f"Predicciones generadas para {len(future_dates)} días")
        print(f"R2 Score: {prediction_quality['r2_score']:.3f}")
       
        return {
            'dates': future_dates,
            'values': predictions,
            'lower_bound': lower_bound,
            'upper_bound': upper_bound,
            'confidence': confidence,
            'quality': prediction_quality
        }
       
    except Exception as e:
        print(f"Error detallado en predicciones: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def detect_anomalies_data(df, sensitivity):
    """Detecta anomalías en los datos usando IsolationForest"""
    try:
        # Agrupar datos por fecha
        daily_counts = df.groupby('fecha_asignacion').size().reset_index()
        daily_counts.columns = ['fecha', 'cantidad']
       
        if len(daily_counts) < 2:
            return None
           
        # Preparar datos para el modelo
        X = daily_counts[['cantidad']].values
       
        # Detectar anomalías
        iso_forest = IsolationForest(
            contamination=sensitivity,
            random_state=42
        )
       
        # Fit y predict
        anomalies = iso_forest.fit_predict(X)
        scores = iso_forest.score_samples(X)
       
        # Identificar días anómalos
        anomaly_mask = anomalies == -1
        anomaly_dates = daily_counts.loc[anomaly_mask, 'fecha']
        anomaly_values = daily_counts.loc[anomaly_mask, 'cantidad']
        anomaly_scores = scores[anomaly_mask]
       
        # Calcular estadísticas para clasificación
        mean_value = daily_counts['cantidad'].mean()
        std_value = daily_counts['cantidad'].std()
       
        # Clasificar anomalías
        types = []
        for value in anomaly_values:
            if value > mean_value + std_value:
                types.append('Exceso significativo')
            elif value < mean_value - std_value:
                types.append('Déficit significativo')
            else:
                types.append('Patrón inusual')
       
        return {
            'dates': anomaly_dates,
            'values': anomaly_values,
            'scores': anomaly_scores,
            'types': types,
            'mean': mean_value,
            'std': std_value
        }
       
    except Exception as e:
        print(f"Error en detección de anomalías: {str(e)}")
        return None

def plot_patterns(ax, patterns_data):
    """Visualiza los patrones detectados"""
    if not patterns_data or 'daily_counts' not in patterns_data:
        ax.text(0.5, 0.5, 'No hay datos suficientes para mostrar patrones',
                ha='center', va='center')
        return

    daily_counts = patterns_data['daily_counts']
    dates = daily_counts.index
    values = daily_counts.values

    # Gráfico principal de datos diarios
    ax.plot(dates, values, 'b-', label='Datos diarios', alpha=0.6)
    ax.scatter(dates, values, c='blue', s=20, alpha=0.5)
   
    # Calcular y mostrar tendencia
    z = np.polyfit(range(len(dates)), values, 1)
    p = np.poly1d(z)
    ax.plot(dates, p(range(len(dates))), "r--", label='Tendencia', alpha=0.8)

    # Calcular y mostrar media móvil
    window_size = 7
    if len(values) >= window_size:
        moving_avg = pd.Series(values).rolling(window=window_size).mean()
        ax.plot(dates, moving_avg, 'g-', label=f'Media móvil ({window_size} días)',
                linewidth=2, alpha=0.7)

    ax.set_title('Análisis de Patrones Temporales')
    ax.set_xlabel('Fecha')
    ax.set_ylabel('Cantidad de Técnicos')
    ax.grid(True, alpha=0.3)
    ax.legend()

    # Formatear eje x para fechas
    ax.figure.autofmt_xdate()
   
   
def show_predictions_metrics(text_widget, predictions):
    """Muestra métricas y explicación detallada de las predicciones"""
    text_widget.configure(state='normal')
    text_widget.delete(1.0, tk.END)
   
    if not predictions or 'dates' not in predictions:
        text_widget.insert(tk.END, "No hay datos suficientes para generar predicciones.\n")
        text_widget.configure(state='disabled')
        return
   
    # Título
    text_widget.insert(tk.END, "ANÁLISIS PREDICTIVO\n\n", "title")
   
    # Explicación del gráfico
    text_widget.insert(tk.END, "INTERPRETACIÓN DEL GRÁFICO:\n", "subtitle")
    text_widget.insert(tk.END, "• La línea azul muestra la predicción de técnicos necesarios\n")
    text_widget.insert(tk.END, "• Los puntos indican valores específicos por día\n")
    text_widget.insert(tk.END, "• El sombreado indica el intervalo de confianza\n\n")
   
    # Predicciones detalladas
    text_widget.insert(tk.END, "PREDICCIONES POR DÍA:\n", "subtitle")
    dates = predictions['dates']
    values = predictions['values']
   
    for date, value in zip(dates, values):
        date_str = date.strftime('%d/%m/%Y')
        day_name = {
            'Monday': 'Lunes',
            'Tuesday': 'Martes',
            'Wednesday': 'Miércoles',
            'Thursday': 'Jueves',
            'Friday': 'Viernes',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }[date.strftime('%A')]
       
        text_widget.insert(tk.END, f"• {date_str} ({day_name}): ")
        text_widget.insert(tk.END, f"{int(value)} técnicos\n")
   
    # Estadísticas
    text_widget.insert(tk.END, "\nESTADÍSTICAS:\n", "subtitle")
    mean_value = np.mean(values)
    max_value = np.max(values)
    min_value = np.min(values)
   
    text_widget.insert(tk.END, f"• Promedio predicho: {mean_value:.1f} técnicos\n")
    text_widget.insert(tk.END, f"• Máximo predicho: {int(max_value)} técnicos\n")
    text_widget.insert(tk.END, f"• Mínimo predicho: {int(min_value)} técnicos\n")
   
    # Recomendaciones
    text_widget.insert(tk.END, "\nRECOMENDACIONES:\n", "subtitle")
   
    # Variabilidad en predicciones
    variability = max_value - min_value
    if variability > 5:
        text_widget.insert(tk.END, "• Alta variabilidad en las necesidades previstas\n", "warning")
        text_widget.insert(tk.END, "• Se recomienda mantener personal de respaldo\n")
        text_widget.insert(tk.END, f"• Diferencia de {int(variability)} técnicos entre día máximo y mínimo\n")
    else:
        text_widget.insert(tk.END, "• Necesidades de personal estables\n", "important")
        text_widget.insert(tk.END, "• Se recomienda mantener la dotación actual\n")
   
    text_widget.configure(state='disabled')
   
   
def show_overview_metrics(text_widget, results):
    """Muestra métricas generales y resumen del análisis"""
    text_widget.insert(tk.END, "=== RESUMEN GENERAL DEL ANÁLISIS ===\n\n", "title")
   
    # Obtener datos de patrones
    patterns = results.get('patterns', {})
    if not patterns:
        text_widget.insert(tk.END, "No hay datos suficientes para el análisis.\n")
        return
       
    daily_counts = patterns.get('daily_counts')
    if daily_counts is not None:
        # Estadísticas generales
        text_widget.insert(tk.END, "ESTADÍSTICAS GENERALES:\n", "subtitle")
        total_days = len(daily_counts)
        mean_value = daily_counts.mean()
        max_value = daily_counts.max()
        min_value = daily_counts.min()
        std_value = daily_counts.std()
       
        text_widget.insert(tk.END, f"• Días analizados: {total_days}\n")
        text_widget.insert(tk.END, f"• Promedio diario: {mean_value:.1f} técnicos\n")
        text_widget.insert(tk.END, f"• Máximo diario: {int(max_value)} técnicos\n")
        text_widget.insert(tk.END, f"• Mínimo diario: {int(min_value)} técnicos\n")
        text_widget.insert(tk.END, f"• Desviación estándar: {std_value:.2f}\n\n")
       
        # Análisis semanal
        text_widget.insert(tk.END, "ANÁLISIS SEMANAL:\n", "subtitle")
        weekly_pattern = patterns.get('weekly_pattern')
        if weekly_pattern is not None:
            days = {
                0: 'Lunes',
                1: 'Martes',
                2: 'Miércoles',
                3: 'Jueves',
                4: 'Viernes',
                5: 'Sábado',
                6: 'Domingo'
            }
            for day, count in weekly_pattern.items():
                text_widget.insert(tk.END, f"• {days[day]}: {count:.1f} promedio\n")
       
        # Análisis mensual
        text_widget.insert(tk.END, "\nANÁLISIS MENSUAL:\n", "subtitle")
        monthly_pattern = patterns.get('monthly_pattern')
        if monthly_pattern is not None:
            months = {
                1: 'Enero', 2: 'Febrero', 3: 'Marzo',
                4: 'Abril', 5: 'Mayo', 6: 'Junio',
                7: 'Julio', 8: 'Agosto', 9: 'Septiembre',
                10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
            }
            for month, count in monthly_pattern.items():
                text_widget.insert(tk.END, f"• {months[month]}: {count:.1f} promedio\n")
       
        # Recomendaciones
        text_widget.insert(tk.END, "\nRECOMENDACIONES:\n", "subtitle")
        cv = std_value / mean_value if mean_value > 0 else 0
       
        if cv > 0.3:
            text_widget.insert(tk.END, "• Alta variabilidad en la dotación diaria\n")
            text_widget.insert(tk.END, "• Se recomienda establecer dotación más estable\n")
        else:
            text_widget.insert(tk.END, "• Buena estabilidad en la dotación diaria\n")
            text_widget.insert(tk.END, "• Mantener esquema actual de distribución\n")
           
        if max_value - min_value > 5:
            text_widget.insert(tk.END, "• Diferencia significativa entre máximos y mínimos\n")
            text_widget.insert(tk.END, "• Revisar causas de variaciones extremas\n")
   
    # Configurar estilos de texto
    text_widget.tag_configure("title", font=('Segoe UI', 12, 'bold'))
    text_widget.tag_configure("subtitle", font=('Segoe UI', 10, 'bold'))
   
    # Hacer el texto de solo lectura
    text_widget.configure(state='disabled')

def plot_overview(ax, results):
    """Visualiza el resumen general de los datos"""
    patterns = results.get('patterns')
    if not patterns or 'daily_counts' not in patterns:
        ax.text(0.5, 0.5, 'No hay datos suficientes para mostrar',
                ha='center', va='center')
        return

    daily_counts = patterns['daily_counts']
    dates = daily_counts.index
    values = daily_counts.values

    # Gráfico de línea principal
    ax.plot(dates, values, 'b-', label='Datos diarios', alpha=0.6, linewidth=1)
   
    # Agregar puntos para mejor visualización
    ax.scatter(dates, values, c='blue', s=20, alpha=0.5)
   
    # Calcular y mostrar tendencia
    z = np.polyfit(range(len(dates)), values, 1)
    p = np.poly1d(z)
    ax.plot(dates, p(range(len(dates))), "r--", label='Tendencia', alpha=0.8)

    # Calcular y mostrar media móvil
    window_size = 7
    if len(values) >= window_size:
        moving_avg = pd.Series(values).rolling(window=window_size).mean()
        ax.plot(dates, moving_avg, 'g-', label=f'Media móvil ({window_size} días)',
                linewidth=2, alpha=0.7)

    ax.set_title('Vista General de Datos')
    ax.set_xlabel('Fecha')
    ax.set_ylabel('Cantidad de Técnicos')
    ax.grid(True, alpha=0.3)
    ax.legend()

    # Formatear eje x para fechas
    ax.figure.autofmt_xdate()
   

def show_patterns_metrics(text_widget, patterns):
    """Muestra análisis detallado de los patrones encontrados"""
    if not patterns:
        text_widget.insert(tk.END, "No hay datos suficientes para analizar patrones.\n")
        return

    text_widget.insert(tk.END, "=== ANÁLISIS DE PATRONES ===\n\n", "title")

    # Patrones diarios
    daily_counts = patterns.get('daily_counts')
    if daily_counts is not None:
        text_widget.insert(tk.END, "PATRONES DIARIOS:\n", "subtitle")
        mean_daily = daily_counts.mean()
        std_daily = daily_counts.std()
        cv = std_daily / mean_daily if mean_daily > 0 else 0

        text_widget.insert(tk.END, f"• Promedio diario: {mean_daily:.1f} técnicos\n")
        text_widget.insert(tk.END, f"• Desviación estándar: {std_daily:.2f}\n")
        text_widget.insert(tk.END, f"• Coeficiente de variación: {cv:.2f}\n\n")

        # Interpretación de variabilidad
        text_widget.insert(tk.END, "ANÁLISIS DE VARIABILIDAD:\n", "subtitle")
        if cv < 0.1:
            text_widget.insert(tk.END, "• Muy baja variabilidad - Dotación muy estable\n")
        elif cv < 0.2:
            text_widget.insert(tk.END, "• Baja variabilidad - Dotación estable\n")
        elif cv < 0.3:
            text_widget.insert(tk.END, "• Variabilidad moderada - Dotación relativamente estable\n")
        else:
            text_widget.insert(tk.END, "• Alta variabilidad - Dotación inestable\n")

    # Patrones semanales
    text_widget.insert(tk.END, "\nPATRONES SEMANALES:\n", "subtitle")
    weekly_pattern = patterns.get('weekly_pattern')
    if weekly_pattern is not None:
        days = {
            0: 'Lunes',
            1: 'Martes',
            2: 'Miércoles',
            3: 'Jueves',
            4: 'Viernes',
            5: 'Sábado',
            6: 'Domingo'
        }
       
        # Encontrar días con mayor y menor carga
        max_day = weekly_pattern.idxmax()
        min_day = weekly_pattern.idxmin()
       
        text_widget.insert(tk.END, f"Día con mayor carga: {days[max_day]} ({weekly_pattern[max_day]:.1f} técnicos)\n")
        text_widget.insert(tk.END, f"Día con menor carga: {days[min_day]} ({weekly_pattern[min_day]:.1f} técnicos)\n\n")
       
        text_widget.insert(tk.END, "Distribución por día:\n")
        for day, count in weekly_pattern.items():
            text_widget.insert(tk.END, f"• {days[day]}: {count:.1f} técnicos promedio\n")

    # Patrones mensuales
    text_widget.insert(tk.END, "\nPATRONES MENSUALES:\n", "subtitle")
    monthly_pattern = patterns.get('monthly_pattern')
    if monthly_pattern is not None:
        months = {
            1: 'Enero', 2: 'Febrero', 3: 'Marzo',
            4: 'Abril', 5: 'Mayo', 6: 'Junio',
            7: 'Julio', 8: 'Agosto', 9: 'Septiembre',
            10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
        }
       
        # Encontrar meses con mayor y menor carga
        max_month = monthly_pattern.idxmax()
        min_month = monthly_pattern.idxmin()
       
        text_widget.insert(tk.END, f"Mes con mayor carga: {months[max_month]} ({monthly_pattern[max_month]:.1f} técnicos)\n")
        text_widget.insert(tk.END, f"Mes con menor carga: {months[min_month]} ({monthly_pattern[min_month]:.1f} técnicos)\n\n")

    # Recomendaciones
    text_widget.insert(tk.END, "\nRECOMENDACIONES:\n", "subtitle")
   
    # Basadas en variabilidad diaria
    if cv > 0.3:
        text_widget.insert(tk.END, "• Se recomienda establecer una dotación más estable\n")
        text_widget.insert(tk.END, "• Investigar causas de alta variabilidad\n")
    else:
        text_widget.insert(tk.END, "• Mantener el esquema actual de distribución\n")
   
    # Basadas en patrones semanales
    if weekly_pattern is not None:
        weekly_cv = weekly_pattern.std() / weekly_pattern.mean() if weekly_pattern.mean() > 0 else 0
        if weekly_cv > 0.2:
            text_widget.insert(tk.END, "• Equilibrar mejor la carga entre días de la semana\n")
            text_widget.insert(tk.END, f"• Especial atención a la diferencia entre {days[max_day]} y {days[min_day]}\n")
   
    # Basadas en patrones mensuales
    if monthly_pattern is not None:
        monthly_cv = monthly_pattern.std() / monthly_pattern.mean() if monthly_pattern.mean() > 0 else 0
        if monthly_cv > 0.2:
            text_widget.insert(tk.END, "• Considerar ajustes estacionales en la dotación\n")
            text_widget.insert(tk.END, f"• Planificar con anticipación los meses de {months[max_month]} y {months[min_month]}\n")

    # Configurar estilos de texto
    text_widget.tag_configure("title", font=('Segoe UI', 12, 'bold'))
    text_widget.tag_configure("subtitle", font=('Segoe UI', 10, 'bold'))
   
    # Hacer el texto de solo lectura
    text_widget.configure(state='disabled')    


def plot_anomalies(ax, anomalies):
    """Visualiza las anomalías detectadas"""
    if not anomalies or 'dates' not in anomalies:
        ax.text(0.5, 0.5, 'No se detectaron anomalías en los datos',
                ha='center', va='center')
        return

    # Leer datos completos para contexto
    df = pd.read_excel(excel_file_path)
    df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
    daily_counts = df.groupby('fecha_asignacion').size()

    # Graficar serie temporal completa
    ax.plot(daily_counts.index, daily_counts.values, 'b-',
            label='Datos normales', alpha=0.6)

    # Resaltar anomalías
    ax.scatter(anomalies['dates'], anomalies['values'],
              color='red', s=100, label='Anomalías',
              zorder=5)

    # Agregar etiquetas a las anomalías
    for date, value, tipo in zip(anomalies['dates'],
                                anomalies['values'],
                                anomalies['types']):
        ax.annotate(f'{int(value)}\n({tipo})',
                   (date, value),
                   xytext=(0, 15),
                   textcoords='offset points',
                   ha='center',
                   bbox=dict(boxstyle='round,pad=0.5',
                           fc='red',
                           alpha=0.1),
                   color='red')

   
def plot_predictions(ax, predictions):
    """Visualiza las predicciones con intervalos de confianza"""
    if not predictions or 'dates' not in predictions or 'values' not in predictions:
        ax.text(0.5, 0.5, 'No hay datos suficientes para mostrar predicciones',
                ha='center', va='center')
        return

    dates = predictions['dates']
    values = predictions['values']
    lower_bound = predictions['lower_bound']
    upper_bound = predictions['upper_bound']
    confidence = predictions['confidence']

    # Línea principal de predicción
    ax.plot(dates, values, 'b-', label='Predicción', linewidth=2)
   
    # Intervalo de confianza
    ax.fill_between(dates, lower_bound, upper_bound, color='b', alpha=0.2,
                   label=f'Intervalo de confianza {confidence*100:.0f}%')
   
    # Puntos de predicción
    ax.scatter(dates, values, c='blue', s=30)
   
    # Etiquetas en los puntos
    for date, value, lb, ub in zip(dates, values, lower_bound, upper_bound):
        ax.annotate(f'{value}\n({lb}-{ub})',
                   (date, value),
                   xytext=(0, 10),
                   textcoords='offset points',
                   ha='center',
                   bbox=dict(boxstyle='round,pad=0.5',
                           fc='white',
                           ec='gray',
                           alpha=0.7))

    # Agregar información de calidad si está disponible
    if 'quality' in predictions:
        quality = predictions['quality']
        r2_score = quality['r2_score']
        ax.text(0.02, 0.98, f'R² Score: {r2_score:.3f}',
                transform=ax.transAxes,
                bbox=dict(facecolor='white', alpha=0.8))

    ax.set_title('Predicción de Técnicos Necesarios')
    ax.set_xlabel('Fecha')
    ax.set_ylabel('Cantidad de Técnicos')
    ax.grid(True, alpha=0.3)
    ax.legend()

    # Formatear eje x para fechas
    ax.figure.autofmt_xdate()
   
def show_anomalies_metrics(text_widget, anomalies):
    """Muestra análisis detallado de las anomalías"""
    text_widget.configure(state='normal')
    text_widget.delete(1.0, tk.END)
   
    if not anomalies or 'dates' not in anomalies:
        text_widget.insert(tk.END, "No se detectaron anomalías en los datos.\n")
        text_widget.configure(state='disabled')
        return
   
    # Título
    text_widget.insert(tk.END, "ANÁLISIS DE ANOMALÍAS\n\n", "title")
   
    # Explicación del gráfico
    text_widget.insert(tk.END, "INTERPRETACIÓN DEL GRÁFICO:\n", "subtitle")
    text_widget.insert(tk.END, "• Línea azul: datos normales\n")
    text_widget.insert(tk.END, "• Puntos rojos: anomalías detectadas\n")
    text_widget.insert(tk.END, "• Líneas punteadas: límites estadísticos\n\n")
   
    # Resumen de anomalías
    dates = anomalies['dates']
    values = anomalies['values']
    mean_value = np.mean(values)
   
    text_widget.insert(tk.END, f"Se detectaron {len(dates)} anomalías:\n\n", "important")
   
    # Detalle de cada anomalía
    for date, value in zip(dates, values):
        date_str = date.strftime('%d/%m/%Y')
        day_name = {
            'Monday': 'Lunes',
            'Tuesday': 'Martes',
            'Wednesday': 'Miércoles',
            'Thursday': 'Jueves',
            'Friday': 'Viernes',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }[date.strftime('%A')]
       
        text_widget.insert(tk.END, f"• {date_str} ({day_name}):\n")
        text_widget.insert(tk.END, f"  {int(value)} técnicos - ")
       
        # Clasificar anomalía
        if value > mean_value:
            text_widget.insert(tk.END, "Exceso de personal\n", "warning")
        else:
            text_widget.insert(tk.END, "Déficit de personal\n", "warning")
   
    # Recomendaciones
    text_widget.insert(tk.END, "\nRECOMENDACIONES:\n", "subtitle")
    text_widget.insert(tk.END, "• Investigar las causas de cada anomalía\n")
    text_widget.insert(tk.END, "• Revisar la planificación en las fechas señaladas\n")
    text_widget.insert(tk.END, "• Establecer protocolos para evitar situaciones similares\n")
   
    text_widget.configure(state='disabled')

def display_results(notebook, results):
    """Muestra los resultados del análisis con explicaciones detalladas"""
    try:
        for tab_id in notebook.tabs():
            tab_name = notebook.tab(tab_id, "text").lower()
            tab = notebook.nametowidget(tab_id)

            # Limpiar el tab
            for widget in tab.winfo_children():
                widget.destroy()

            # Crear frame principal con PanedWindow
            paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
            paned.pack(fill='both', expand=True, padx=5, pady=5)

            # Panel izquierdo para el gráfico
            graph_frame = ttk.LabelFrame(paned, text="Visualización")
           
            # Panel derecho para explicaciones
            text_frame = ttk.LabelFrame(paned, text="Análisis Detallado")

            # Crear figura para el gráfico
            fig = Figure(figsize=(8, 6), dpi=100)
            ax = fig.add_subplot(111)

            # Crear widget de texto con scroll para explicaciones
            text_container = ttk.Frame(text_frame)
            text_container.pack(fill='both', expand=True, padx=5, pady=5)

            text_widget = tk.Text(text_container, wrap=tk.WORD, width=50, font=('Segoe UI', 10))
            scrollbar = ttk.Scrollbar(text_container, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)

            scrollbar.pack(side='right', fill='y')
            text_widget.pack(side='left', fill='both', expand=True)

            # Actualizar contenido según el tipo de pestaña
            if tab_name == "vista general":
                plot_overview(ax, results)
                show_overview_metrics(text_widget, results)
            elif tab_name == "patrones":
                plot_patterns(ax, results.get('patterns'))
                show_patterns_metrics(text_widget, results.get('patterns'))
            elif tab_name == "predicciones":
                plot_predictions(ax, results.get('predictions'))
                show_predictions_metrics(text_widget, results.get('predictions'))
            elif tab_name == "anomalías":
                plot_anomalies(ax, results.get('anomalies'))
                show_anomalies_metrics(text_widget, results.get('anomalies'))

            # Mostrar gráfico
            canvas = FigureCanvasTkAgg(fig, master=graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True, padx=5, pady=5)

            # Agregar barra de herramientas para el gráfico
            toolbar = NavigationToolbar2Tk(canvas, graph_frame)
            toolbar.update()
            toolbar.pack(fill='x', padx=5)

            # Agregar ambos paneles al PanedWindow
            paned.add(graph_frame, weight=1)
            paned.add(text_frame, weight=1)

            # Configurar estilos de texto
            text_widget.tag_configure("title", font=('Segoe UI', 12, 'bold'))
            text_widget.tag_configure("subtitle", font=('Segoe UI', 10, 'bold'))
            text_widget.tag_configure("important", font=('Segoe UI', 10, 'bold'), foreground='#007bff')
            text_widget.tag_configure("warning", font=('Segoe UI', 10, 'bold'), foreground='#dc3545')

            # Hacer el texto de solo lectura
            text_widget.configure(state='disabled')

    except Exception as e:
        print(f"Error en display_results: {str(e)}")
        messagebox.showerror("Error", f"Error al mostrar resultados: {str(e)}")

def plot_overview(ax, results):
    """Genera el gráfico de vista general"""
    if 'patterns' in results and results['patterns']:
        data = results['patterns'].get('daily_counts')
        if data is not None:
            dates = data.index
            values = data.values
            ax.plot(dates, values, 'b-', label='Datos diarios')
            ax.set_title('Vista General de Datos')
            ax.set_xlabel('Fecha')
            ax.set_ylabel('Cantidad')
            ax.grid(True)
            ax.legend()
            fig = ax.figure
            fig.autofmt_xdate()
           
def show_metrics(frame, tab_name, results):
    """Muestra métricas específicas para cada tipo de análisis"""
    if tab_name == "vista general":
        if 'patterns' in results and results['patterns']:
            data = results['patterns'].get('daily_counts')
            if data is not None:
                ttk.Label(frame, text=f"Total de registros: {len(data)}").pack(anchor='w')
                ttk.Label(frame, text=f"Promedio diario: {data.mean():.2f}").pack(anchor='w')
                ttk.Label(frame, text=f"Máximo diario: {data.max()}").pack(anchor='w')
                ttk.Label(frame, text=f"Mínimo diario: {data.min()}").pack(anchor='w')

def setup_ai_controls(control_frame):
    """Configura los controles para el análisis avanzado"""
    # Variables de control
    controls = {
        'date_range': {
            'start': None,
            'end': None
        },
        'config_vars': {
            'confidence': tk.DoubleVar(value=0.95),
            'horizon': tk.IntVar(value=14),
            'sensitivity': tk.DoubleVar(value=0.1)
        }
    }

    # Frame para selección de fechas
    date_frame = ttk.LabelFrame(control_frame, text="Rango de Análisis", padding=10)
    date_frame.pack(fill='x', pady=5, padx=5)

    # Fecha inicial
    ttk.Label(date_frame, text="Fecha Inicial:", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 5))
    controls['date_range']['start'] = Calendar(
        date_frame,
        selectmode='day',
        date_pattern='dd/mm/yyyy',
        background='white',
        foreground='black',
        selectbackground='#007bff'
    )
    controls['date_range']['start'].pack(fill='x', pady=(0, 10))

    # Fecha final
    ttk.Label(date_frame, text="Fecha Final:", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 5))
    controls['date_range']['end'] = Calendar(
        date_frame,
        selectmode='day',
        date_pattern='dd/mm/yyyy',
        background='white',
        foreground='black',
        selectbackground='#007bff'
    )
    controls['date_range']['end'].pack(fill='x', pady=(0, 10))

    # Frame para parámetros de análisis
    params_frame = ttk.LabelFrame(control_frame, text="Parámetros de Análisis", padding=10)
    params_frame.pack(fill='x', pady=5, padx=5)

    # Nivel de confianza
    ttk.Label(params_frame, text="Nivel de Confianza:", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 5))
    confidence_scale = ttk.Scale(
        params_frame,
        from_=0.8,
        to=0.99,
        variable=controls['config_vars']['confidence'],
        orient='horizontal'
    )
    confidence_scale.pack(fill='x', pady=(0, 10))
    ttk.Label(
        params_frame,
        textvariable=controls['config_vars']['confidence'],
        font=('Segoe UI', 9)
    ).pack(anchor='e')

    # Horizonte de predicción
    ttk.Label(params_frame, text="Horizonte de Predicción (días):", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 5))
    horizon_spin = ttk.Spinbox(
        params_frame,
        from_=7,
        to=90,
        textvariable=controls['config_vars']['horizon'],
        width=10
    )
    horizon_spin.pack(fill='x', pady=(0, 10))

    # Sensibilidad de detección
    ttk.Label(params_frame, text="Sensibilidad de Detección:", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 5))
    sensitivity_scale = ttk.Scale(
        params_frame,
        from_=0.01,
        to=0.2,
        variable=controls['config_vars']['sensitivity'],
        orient='horizontal'
    )
    sensitivity_scale.pack(fill='x', pady=(0, 10))
    ttk.Label(
        params_frame,
        textvariable=controls['config_vars']['sensitivity'],
        font=('Segoe UI', 9)
    ).pack(anchor='e')

    # Frame para botones
    button_frame = ttk.Frame(control_frame)
    button_frame.pack(fill='x', pady=10, padx=5)

    ttk.Button(
        button_frame,
        text="Ejecutar Análisis",
        command=lambda: run_analysis(
            controls['date_range']['start'],
            controls['date_range']['end'],
            controls['config_vars'],
            results_notebook
        ),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        button_frame,
        text="Limpiar Resultados",
        command=lambda: clear_results_from_notebook(results_notebook),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        button_frame,
        text="Exportar Análisis",
        command=lambda: export_analysis_results(
            controls['date_range']['start'],
            controls['date_range']['end'],
            controls['config_vars']
        ),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    return controls


def clear_results_from_notebook(notebook):
    """Limpia los resultados anteriores del notebook"""
    if notebook:
        for tab_id in notebook.tabs():
            tab = notebook.nametowidget(tab_id)
            # Limpiar cada frame dentro de la pestaña
            for child in tab.winfo_children():
                if isinstance(child, ttk.LabelFrame):
                    # Limpiar contenido de los frames de visualización y métricas
                    for widget in child.winfo_children():
                        widget.destroy()



def export_analysis_results(start_date, end_date, config_vars):
    """Exporta los resultados del análisis"""
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("PDF files", "*.pdf"),
                ("HTML files", "*.html")
            ],
            title="Exportar Análisis"
        )
       
        if not file_path:
            return

        # Ejecutar análisis
        analyzer = AdvancedAnalyzer(excel_file_path)
        results = analyzer.run_analysis(
            start_date.get_date(),
            end_date.get_date(),
            config_vars
        )

        # Exportar según el formato
        if file_path.endswith('.xlsx'):
            export_to_excel(results, file_path)
        elif file_path.endswith('.pdf'):
            export_to_pdf(results, file_path)
        else:
            export_to_html(results, file_path)

        messagebox.showinfo("Éxito", "Análisis exportado correctamente")

    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar análisis: {str(e)}")

class AdvancedAnalysisDisplay:
    def __init__(self, tab):
        self.tab = tab
        self.setup_display()
       
    def setup_display(self):
        """Configura los elementos de visualización"""
        # Frame para resultados
        self.results_frame = ttk.LabelFrame(self.tab, text="Resultados del Análisis")
        self.results_frame.pack(fill='both', expand=True, padx=10, pady=5)
       
        # Notebook para diferentes vistas
        self.notebook = ttk.Notebook(self.results_frame)
        self.notebook.pack(fill='both', expand=True)
       
        # Crear pestañas
        self.tabs = {
            'overview': self.create_tab("Vista General"),
            'patterns': self.create_tab("Patrones"),
            'predictions': self.create_tab("Predicciones"),
            'anomalies': self.create_tab("Anomalías"),
            'optimization': self.create_tab("Optimización")
        }
       
    def create_tab(self, name):
        """Crea una pestaña con estructura básica"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=name)
       
        # Frame para gráfico
        graph_frame = ttk.LabelFrame(tab, text="Visualización")
        graph_frame.pack(fill='both', expand=True, padx=5, pady=5)
       
        # Frame para métricas
        metrics_frame = ttk.LabelFrame(tab, text="Métricas")
        metrics_frame.pack(fill='x', padx=5, pady=5)
       
        return {
            'tab': tab,
            'graph_frame': graph_frame,
            'metrics_frame': metrics_frame,
            'figure': None,
            'canvas': None
        }
       
    def update_display(self, analyzer_results):
        """Actualiza todas las visualizaciones con nuevos datos"""
        if not analyzer_results:
            self.show_error("No hay datos para mostrar")
            return
           
        # Actualizar cada pestaña
        self.update_overview(analyzer_results)
        self.update_patterns(analyzer_results['patterns'])
        self.update_predictions(analyzer_results['predictions'])
        self.update_anomalies(analyzer_results['anomalies'])
        self.update_optimization(analyzer_results['optimization'])
       
    def update_overview(self, results):
        """Actualiza la vista general"""
        tab = self.tabs['overview']
       
        # Limpiar frame anterior
        for widget in tab['graph_frame'].winfo_children():
            widget.destroy()
           
        # Crear figura con subplots
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=(
                'Tendencia General',
                'Distribución por Jornada',
                'Anomalías Detectadas',
                'Predicción a Futuro'
            )
        )
       
        # Agregar gráficos
        self.add_trend_plot(fig, results['patterns'], 1, 1)
        self.add_distribution_plot(fig, results['patterns'], 1, 2)
        self.add_anomalies_plot(fig, results['anomalies'], 2, 1)
        self.add_prediction_plot(fig, results['predictions'], 2, 2)
       
        # Actualizar layout
        fig.update_layout(
            height=800,
            showlegend=True,
            title_text="Resumen del Análisis"
        )
       
        # Mostrar figura
        self.show_figure(fig, tab)
       
        # Actualizar métricas
        self.update_overview_metrics(results, tab['metrics_frame'])
       
    def show_figure(self, fig, tab):
        """Muestra una figura en la pestaña especificada"""
        canvas = FigureCanvasTkAgg(fig, master=tab['graph_frame'])
        canvas.draw()
       
        # Agregar barra de herramientas
        toolbar = NavigationToolbar2Tk(canvas, tab['graph_frame'])
        toolbar.update()
       
        # Empaquetar widgets
        canvas.get_tk_widget().pack(fill='both', expand=True)
        toolbar.pack(fill='x')
       
        # Guardar referencias
        tab['figure'] = fig
        tab['canvas'] = canvas
       
    def update_overview_metrics(self, results, frame):
        """Actualiza las métricas generales"""
        # Limpiar frame anterior
        for widget in frame.winfo_children():
            widget.destroy()
           
        metrics = {
            'Total de registros': len(results['data']),
            'Periodo analizado': f"{results['start_date']} - {results['end_date']}",
            'Anomalías detectadas': len(results['anomalies']['dates']),
            'Precisión del modelo': f"{results['predictions']['accuracy']:.2f}%"
        }
       
        for name, value in metrics.items():
            ttk.Label(
                frame,
                text=f"{name}: {value}",
                font=('Segoe UI', 10)
            ).pack(anchor='w', padx=5, pady=2)
           
    def show_error(self, message):
        """Muestra un mensaje de error"""
        for tab_data in self.tabs.values():
            for widget in tab_data['graph_frame'].winfo_children():
                widget.destroy()
               
            ttk.Label(
                tab_data['graph_frame'],
                text=message,
                font=('Segoe UI', 12),
                foreground='red'
            ).pack(expand=True)

def run_full_analysis(self, start_date, end_date, config_vars):
    try:
        # Crear analizador
        analyzer = AdvancedAnalyzer(excel_file_path)
       
        # Convertir fechas
        start_dt = datetime.strptime(start_date, '%d/%m/%Y')
        end_dt = datetime.strptime(end_date, '%d/%m/%Y')
       
        # Recopilar resultados
        results = {
            'data': analyzer.df,
            'start_date': start_dt,
            'end_date': end_dt,
            'patterns': analyzer.analyze_patterns(start_dt, end_dt),
            'predictions': analyzer.generate_predictions(),  # Llamada actualizada
            'anomalies': analyzer.detect_anomalies(config_vars['sensitivity']),
            'optimization': analyzer.optimize_distribution(start_dt, end_dt)
        }
       
        # Resto del código
       
        # Actualizar visualización
        display = AdvancedAnalysisDisplay(ai_tab)
        display.update_display(results)
       
        return True
       
    except Exception as e:
        messagebox.showerror("Error", f"Error en el análisis: {str(e)}")
        print(f"Error detallado: {str(e)}")
        return False

class AdvancedAnalyzer:
    def __init__(self, excel_path):
        """Inicializa el analizador con la ruta del archivo Excel"""
        self.excel_path = excel_path
        self.df = None
        self.load_data()

    def load_data(self):
        """Carga y preprocesa los datos"""
        try:
            self.df = pd.read_excel(self.excel_path)
            self.df['fecha_asignacion'] = pd.to_datetime(self.df['fecha_asignacion'])
           
            # Asegurar que las columnas necesarias existen
            required_columns = ['nombre', 'cargo', 'letra_rotacion', 'Jornada']
            for col in required_columns:
                if col not in self.df.columns:
                    self.df[col] = ''
                   
        except Exception as e:
            print(f"Error al cargar datos: {str(e)}")
            raise

    def prepare_features(self, df):
        """Prepara las características para el modelo predictivo"""
        try:
            # Agrupar datos por fecha
            daily_counts = df.groupby('fecha_asignacion').size().reset_index()
            daily_counts.columns = ['fecha_asignacion', 'count']

            # Crear características temporales
            features = pd.DataFrame()
            features['dia_semana'] = daily_counts['fecha_asignacion'].dt.dayofweek
            features['mes'] = daily_counts['fecha_asignacion'].dt.month
            features['dia_mes'] = daily_counts['fecha_asignacion'].dt.day
            features['dia_año'] = daily_counts['fecha_asignacion'].dt.dayofyear
            features['semana'] = daily_counts['fecha_asignacion'].dt.isocalendar().week

            # Asegurar que todas las características tienen el mismo número de muestras
            X = features[['dia_semana', 'mes', 'dia_mes', 'dia_año', 'semana']].values
            y = daily_counts['count'].values

            # Verificar que X e y tienen el mismo número de muestras
            if len(X) != len(y):
                raise ValueError(f"Inconsistencia en número de muestras: X={len(X)}, y={len(y)}")

            return X, y

        except Exception as e:
            print(f"Error en prepare_features: {str(e)}")
            raise

    def analyze_patterns(df):
        """Analiza patrones en los datos"""
        try:
            # Análisis temporal
            temporal_results = analyze_temporal_patterns(df)
       
            # Análisis por día de la semana
            df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek
            weekly_pattern = df.groupby('dia_semana').size()
       
            # Análisis por mes
            df['mes'] = df['fecha_asignacion'].dt.month
            monthly_pattern = df.groupby('mes').size()
       
            return {
                'temporal': temporal_results,
                'weekly': weekly_pattern,
                'monthly': monthly_pattern,
                'daily_counts': temporal_results['daily_counts']
            }
       
        except Exception as e:
            print(f"Error en analyze_patterns: {str(e)}")
            return None

    def generate_predictions(self):
        """Genera predicciones usando Random Forest para todas las fechas"""
        try:
            # Leer todos los datos sin filtrar por fecha
            df_hist = self.df.copy()
       
            if df_hist.empty:
                raise ValueError("No hay datos suficientes para generar predicciones")

            # Preparar características y target
            X, y = self.prepare_features(df_hist)
       
            if len(X) < 2:
                raise ValueError("Insuficientes datos para entrenar el modelo")

            # Entrenar modelo
            model = RandomForestRegressor(n_estimators=100, random_state=42)
            model.fit(X, y)

            # Obtener el rango de fechas futuras
            end_date = df_hist['fecha_asignacion'].max()
            future_dates = pd.date_range(start=end_date + timedelta(days=1), periods=14, freq='D')
       
            # Preparar características para predicción
            future_features = pd.DataFrame()
            future_features['dia_semana'] = future_dates.dayofweek
            future_features['mes'] = future_dates.month
            future_features['dia_mes'] = future_dates.day
            future_features['dia_año'] = future_dates.dayofyear
            future_features['semana'] = future_dates.isocalendar().week
       
            # Realizar predicciones
            predictions = model.predict(future_features.values)

            # Calcular intervalos de confianza
            predictions_std = np.std([tree.predict(future_features.values)
                                    for tree in model.estimators_], axis=0)
       
            confidence_interval = 1.96 * predictions_std

            return {
                'dates': future_dates,
                'values': predictions,
                'confidence_lower': predictions - confidence_interval,
                'confidence_upper': predictions + confidence_interval
            }

        except Exception as e:
            print(f"Error en generate_predictions: {str(e)}")
            raise

    def detect_anomalies(self, sensitivity=0.1):
        """Detecta anomalías usando Isolation Forest"""
        try:
            if self.df.empty:
                return None

            # Asegurar que sensitivity sea un número válido
            try:
                sensitivity = float(sensitivity.get()) if hasattr(sensitivity, 'get') else float(sensitivity)
                # Asegurar que sensitivity está en el rango correcto (0-1)
                sensitivity = max(0.01, min(0.99, sensitivity))
            except (ValueError, AttributeError):
                # Si hay algún error, usar valor por defecto
                sensitivity = 0.1
                print("Usando valor de sensibilidad por defecto: 0.1")

            # Preparar datos para detección de anomalías
            daily_counts = self.df.groupby('fecha_asignacion').size().reset_index()
            daily_counts.columns = ['fecha', 'count']
       
            if len(daily_counts) < 2:
                return None

            # Preparar datos para el modelo
            X = daily_counts[['count']].values
       
            # Detectar anomalías
            iso_forest = IsolationForest(
                contamination=sensitivity,
                random_state=42
            )
       
            anomalies = iso_forest.fit_predict(X)
            scores = iso_forest.score_samples(X)
       
            # Identificar las anomalías
            anomaly_indices = np.where(anomalies == -1)[0]
       
            return {
                'dates': daily_counts.iloc[anomaly_indices]['fecha'],
                'counts': daily_counts.iloc[anomaly_indices]['count'],
                'scores': scores[anomaly_indices]
            }

        except Exception as e:
            print(f"Error en detect_anomalies: {str(e)}")
            return None

    def calculate_metrics(self, data):
        """Calcula métricas estadísticas básicas"""
        try:
            if isinstance(data, pd.Series):
                return {
                    'mean': data.mean(),
                    'std': data.std(),
                    'min': data.min(),
                    'max': data.max(),
                    'count': len(data)
                }
            return None
        except Exception as e:
            print(f"Error en calculate_metrics: {str(e)}")
            return None
   
def update_results_display(fig):
    # Limpiar visualizaciones anteriores
    for widget in results_frame.winfo_children():
        widget.destroy()
   
    # Mostrar nueva visualización
    canvas = FigureCanvasTkAgg(fig, master=results_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)

def get_total_technicians():
    """Obtiene el total de técnicos"""
    df = pd.read_excel(excel_file_path)
    return len(df['nombre'].unique())

def get_daily_average():
    """Obtiene el promedio diario de técnicos"""
    df = pd.read_excel(excel_file_path)
    return df.groupby('fecha_asignacion').size().mean()

def get_daily_maximum():
    """Obtiene el máximo diario de técnicos"""
    df = pd.read_excel(excel_file_path)
    return df.groupby('fecha_asignacion').size().max()

def get_daily_minimum():
    """Obtiene el mínimo diario de técnicos"""
    df = pd.read_excel(excel_file_path)
    return df.groupby('fecha_asignacion').size().min()

def get_standard_deviation():
    """Obtiene la desviación estándar"""
    df = pd.read_excel(excel_file_path)
    return df.groupby('fecha_asignacion').size().std()

def get_prediction_dates():
    """Obtiene las fechas de predicción"""
    last_date = pd.read_excel(excel_file_path)['fecha_asignacion'].max()
    return pd.date_range(start=last_date, periods=14, freq='D')

def get_prediction_values():
    """Obtiene los valores predichos"""
    # Implementar lógica de predicción aquí
    return np.random.normal(10, 2, 14)  # Ejemplo

def get_lower_bounds():
    """Obtiene límites inferiores de predicción"""
    predictions = get_prediction_values()
    return predictions - 1.96 * np.std(predictions)

def get_upper_bounds():
    """Obtiene límites superiores de predicción"""
    predictions = get_prediction_values()
    return predictions + 1.96 * np.std(predictions)

def get_confidence_levels():
    """Obtiene niveles de confianza"""
    return np.full(14, 0.95)  # Ejemplo

def get_anomaly_dates():
    """Obtiene fechas de anomalías"""
    df = pd.read_excel(excel_file_path)
    # Implementar detección de anomalías aquí
    return df['fecha_asignacion'].sample(5)  # Ejemplo

def get_anomaly_types():
    """Obtiene tipos de anomalías"""
    return ['Sobrecarga', 'Subcarga', 'Patrón Inusual', 'Desviación', 'Outlier']

def get_anomaly_severities():
    """Obtiene severidades de anomalías"""
    return ['Alta', 'Media', 'Alta', 'Baja', 'Media']

def get_anomaly_descriptions():
    """Obtiene descripciones de anomalías"""
    return [
        'Exceso significativo de personal',
        'Personal insuficiente',
        'Distribución irregular',
        'Variación menor',
        'Valor atípico detectado'
    ]

def get_metric_names():
    """Obtiene nombres de métricas"""
    return [
        'Eficiencia Global',
        'Balance de Carga',
        'Estabilidad',
        'Rotación',
        'Cobertura'
    ]

def get_metric_values():
    """Obtiene valores de métricas"""
    return [0.85, 0.76, 0.92, 0.68, 0.88]

def get_metric_descriptions():
    """Obtiene descripciones de métricas"""
    return [
        'Indicador de eficiencia general',
        'Balance en la distribución de carga',
        'Estabilidad en asignaciones',
        'Tasa de rotación de personal',
        'Nivel de cobertura de turnos'
    ]

def get_metric_recommendations():
    """Obtiene recomendaciones basadas en métricas"""
    return [
        'Mantener nivel actual',
        'Mejorar distribución',
        'Continuar estrategia actual',
        'Reducir rotación',
        'Incrementar cobertura'
    ]






def export_to_excel(file_path):
    """Exporta los resultados del análisis a Excel"""
    try:
        # Crear Excel writer
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Hoja de Resumen
            summary_df = create_summary_sheet()
            summary_df.to_excel(writer, sheet_name='Resumen', index=False)
           
            # Hoja de Predicciones
            predictions_df = create_predictions_sheet()
            predictions_df.to_excel(writer, sheet_name='Predicciones', index=False)
           
            # Hoja de Anomalías
            anomalies_df = create_anomalies_sheet()
            anomalies_df.to_excel(writer, sheet_name='Anomalías', index=False)
           
            # Hoja de Métricas
            metrics_df = create_metrics_sheet()
            metrics_df.to_excel(writer, sheet_name='Métricas', index=False)
           
            # Ajustar anchos de columna
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        pass  # Suponiendo que la exportación se realiza correctamente
    except Exception as e:
        print(f"Error al exportar a Excel: {e}")
    finally:
        print(f"Excel exportado a: {file_path}")

def export_to_pdf(file_path):
    """Exporta los resultados del análisis a PDF"""
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
       
        # Crear documento
        doc = SimpleDocTemplate(
            file_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
       
        # Lista para elementos del documento
        elements = []
        styles = getSampleStyleSheet()
       
        # Título
        title = Paragraph("Reporte de Análisis Avanzado", styles['Title'])
        elements.append(title)
        elements.append(Spacer(1, 12))
       
        # Resumen Ejecutivo
        elements.append(Paragraph("Resumen Ejecutivo", styles['Heading1']))
        elements.append(Spacer(1, 12))
        summary_text = create_summary_text()
        elements.append(Paragraph(summary_text, styles['Normal']))
        elements.append(Spacer(1, 12))
       
        # Predicciones
        elements.append(Paragraph("Predicciones", styles['Heading1']))
        elements.append(Spacer(1, 12))
        predictions_table = create_predictions_table()
        elements.append(predictions_table)
        elements.append(Spacer(1, 12))
       
        # Anomalías
        elements.append(Paragraph("Anomalías Detectadas", styles['Heading1']))
        elements.append(Spacer(1, 12))
        anomalies_table = create_anomalies_table()
        elements.append(anomalies_table)
        elements.append(Spacer(1, 12))
       
        # Métricas
        elements.append(Paragraph("Métricas y KPIs", styles['Heading1']))
        elements.append(Spacer(1, 12))
        metrics_table = create_metrics_table()
        elements.append(metrics_table)
       
        # Construir documento
        doc.build(elements)
       
        pass  # Suponiendo que la exportación se realiza correctamente
    except Exception as e:
        print(f"Error al exportar a PDF: {e}")
    finally:
        print(f"PDF exportado a: {file_path}")
   

def update_results_tabs(metrics, predictions, anomalies, optimization, fig):
    """Actualiza todas las pestañas de resultados"""
    # Actualizar pestaña de vista general
    update_overview_tab(metrics)
   
    # Actualizar pestaña de predicciones
    update_predictions_tab(predictions)
   
    # Actualizar pestaña de anomalías
    update_anomalies_tab(anomalies)
   
    # Actualizar pestaña de optimización
    update_optimization_tab(optimization)



def export_to_html(file_path):
    """Exporta los resultados del análisis a HTML"""
    try:
        # Crear contenido HTML
        html_content = f"""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Reporte de Análisis Avanzado</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    padding: 20px;
                    max-width: 1200px;
                    margin: 0 auto;
                }}
                .section {{
                    margin-bottom: 30px;
                    padding: 20px;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 15px 0;
                }}
                th, td {{
                    padding: 12px;
                    text-align: left;
                    border-bottom: 1px solid #ddd;
                }}
                th {{
                    background-color: #f5f5f5;
                }}
                .chart {{
                    width: 100%;
                    max-width: 800px;
                    margin: 20px auto;
                }}
            </style>
        </head>
        <body>
            <h1>Reporte de Análisis Avanzado</h1>
           
            <div class="section">
                <h2>Resumen Ejecutivo</h2>
                {create_summary_html()}
            </div>
           
            <div class="section">
                <h2>Predicciones</h2>
                {create_predictions_html()}
            </div>
           
            <div class="section">
                <h2>Anomalías Detectadas</h2>
                {create_anomalies_html()}
            </div>
           
            <div class="section">
                <h2>Métricas y KPIs</h2>
                {create_metrics_html()}
            </div>
        </body>
        </html>
        """
       
        # Guardar archivo HTML
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
       
        pass  # Suponiendo que la exportación se realiza correctamente
    except Exception as e:
        print(f"Error al exportar a HTML: {e}")

def create_summary_sheet():
    """Crea DataFrame con resumen del análisis"""
    summary_data = {
        'Métrica': [
            'Total de técnicos',
            'Promedio diario',
            'Máximo diario',
            'Mínimo diario',
            'Desviación estándar'
        ],
        'Valor': [
            get_total_technicians(),
            get_daily_average(),
            get_daily_maximum(),
            get_daily_minimum(),
            get_standard_deviation()
        ],
        'Tendencia': ['↑', '→', '↓', '→', '↔']
    }
    return pd.DataFrame(summary_data)

def create_predictions_sheet():
    """Crea DataFrame con predicciones"""
    predictions_data = {
        'Fecha': get_prediction_dates(),
        'Predicción': get_prediction_values(),
        'Intervalo Inferior': get_lower_bounds(),
        'Intervalo Superior': get_upper_bounds(),
        'Confianza': get_confidence_levels()
    }
    return pd.DataFrame(predictions_data)

def create_anomalies_sheet():
    """Crea DataFrame con anomalías detectadas"""
    anomalies_data = {
        'Fecha': get_anomaly_dates(),
        'Tipo': get_anomaly_types(),
        'Severidad': get_anomaly_severities(),
        'Descripción': get_anomaly_descriptions()
    }
    return pd.DataFrame(anomalies_data)

def create_metrics_sheet():
    """Crea DataFrame con métricas detalladas"""
    metrics_data = {
        'Métrica': get_metric_names(),
        'Valor': get_metric_values(),
        'Descripción': get_metric_descriptions(),
        'Recomendación': get_metric_recommendations()
    }
    return pd.DataFrame(metrics_data)





def run_advanced_analysis(controls):
    """Ejecuta el análisis avanzado con los parámetros seleccionados"""
    try:
        # Obtener parámetros
        start_date = controls['date_range']['var']['start'].get_date()
        end_date = controls['date_range']['var']['end'].get_date()
        confidence = controls['confidence_level']['var'].get()
        horizon = controls['forecast_horizon']['var'].get()
       
        # Leer datos
        df = pd.read_excel(excel_file_path)
       
        # Calcular métricas
        metrics = analyze_advanced_metrics(df)
       
        # Generar visualizaciones
        fig = visualize_advanced_metrics(metrics)
       
        # Actualizar interfaz
        update_results_display(fig, metrics)
       
        messagebox.showinfo("Éxito", "Análisis completado exitosamente")
       
    except Exception as e:
        messagebox.showerror("Error", f"Error en el análisis: {str(e)}")
       
       
def create_overview_plot(ax, results):
    """Crea el gráfico de vista general"""
    if 'patterns' in results and results['patterns'] is not None:
        data = results['patterns'].get('daily_counts')
        if data is not None:
            ax.plot(data.index, data.values, 'b-', label='Datos diarios')
            ax.set_title('Vista General de Datos')
            ax.set_xlabel('Fecha')
            ax.set_ylabel('Cantidad')
            ax.grid(True)
            ax.legend()

def create_patterns_plot(ax, patterns):
    """Crea el gráfico de patrones"""
    if patterns and 'daily_counts' in patterns:
        data = patterns['daily_counts']
        ax.plot(data.index, data.values, 'g-', label='Patrones detectados')
        if 'trend' in patterns:
            ax.plot(data.index, patterns['trend'], 'r--', label='Tendencia')
        ax.set_title('Análisis de Patrones')
        ax.set_xlabel('Fecha')
        ax.set_ylabel('Cantidad')
        ax.grid(True)
        ax.legend()

def create_predictions_plot(ax, predictions):
    """Crea el gráfico de predicciones"""
    if predictions and 'dates' in predictions and 'values' in predictions:
        ax.plot(predictions['dates'], predictions['values'], 'b-', label='Predicciones')
        ax.set_title('Predicciones')
        ax.set_xlabel('Fecha')
        ax.set_ylabel('Cantidad Predicha')
        ax.grid(True)
        ax.legend()

def create_anomalies_plot(ax, anomalies):
    """Crea el gráfico de anomalías"""
    if anomalies and 'dates' in anomalies and 'counts' in anomalies:
        ax.scatter(anomalies['dates'], anomalies['counts'], c='r', label='Anomalías')
        ax.set_title('Detección de Anomalías')
        ax.set_xlabel('Fecha')
        ax.set_ylabel('Cantidad')
        ax.grid(True)
        ax.legend()

def show_metrics(frame, tab_name, results):
    """Muestra las métricas relevantes para cada pestaña"""
    if tab_name == "vista general":
        if 'patterns' in results and results['patterns'] is not None:
            data = results['patterns'].get('daily_counts')
            if data is not None:
                ttk.Label(frame, text=f"Total de registros: {len(data)}").pack(anchor='w', padx=5, pady=2)
                ttk.Label(frame, text=f"Promedio diario: {data.mean():.2f}").pack(anchor='w', padx=5, pady=2)
                ttk.Label(frame, text=f"Máximo diario: {data.max()}").pack(anchor='w', padx=5, pady=2)
                ttk.Label(frame, text=f"Mínimo diario: {data.min()}").pack(anchor='w', padx=5, pady=2)
   
    elif tab_name == "anomalías" and 'anomalies' in results:
        anomalies = results['anomalies']
        if anomalies and 'dates' in anomalies:
            ttk.Label(frame, text=f"Anomalías detectadas: {len(anomalies['dates'])}").pack(anchor='w', padx=5, pady=2)

def export_analysis_results(controls):
    """Exporta los resultados del análisis"""
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("PDF files", "*.pdf"),
                ("HTML files", "*.html")
            ],
            title="Exportar Resultados"
        )
       
        if file_path:
            if file_path.endswith('.xlsx'):
                export_to_excel(file_path)
            elif file_path.endswith('.pdf'):
                export_to_pdf(file_path)
            else:
                export_to_html(file_path)
               
            messagebox.showinfo("Éxito", "Resultados exportados correctamente")
           
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar: {str(e)}")
       



def update_trend_plot(fig, metrics):
    """Actualiza el gráfico de tendencias"""
    for jornada in metrics:
        # Obtener datos de tendencia
        trend_data = metrics[jornada]['trend']
        dates = pd.to_datetime(trend_data['dates'])
        values = trend_data['values']
       
        # Actualizar trazo existente o crear uno nuevo
        fig.update_traces(
            selector=dict(name=f'Tendencia {jornada}'),
            x=dates,
            y=values,
            mode='lines+markers',
            line=dict(width=2),
            marker=dict(size=6)
        )

def update_correlation_matrix(fig, metrics):
    """Actualiza la matriz de correlación"""
    # Calcular matriz de correlación
    corr_matrix = np.array([
        [metrics[j1]['correlation'][j2]
         for j2 in metrics]
        for j1 in metrics
    ])
   
    # Actualizar heatmap
    fig.update_traces(
        selector=dict(type='heatmap'),
        z=corr_matrix,
        x=list(metrics.keys()),
        y=list(metrics.keys()),
        text=np.round(corr_matrix, 2)
    )

def update_workload_distribution(fig, metrics):
    """Actualiza la distribución de carga de trabajo"""
    for jornada in metrics:
        workload_data = metrics[jornada]['workload']
        dates = pd.to_datetime(workload_data['dates'])
        values = workload_data['values']
       
        fig.update_traces(
            selector=dict(name=jornada),
            x=dates,
            y=values,
            text=values
        )

def update_3d_analysis(fig, metrics):
    """Actualiza el análisis 3D"""
    # Extraer datos 3D
    x_vals = []
    y_vals = []
    z_vals = []
    colors = []
    labels = []
   
    for jornada in metrics:
        data_3d = metrics[jornada]['analysis_3d']
        x_vals.extend(data_3d['x'])
        y_vals.extend(data_3d['y'])
        z_vals.extend(data_3d['z'])
        colors.extend(data_3d['colors'])
        labels.extend(data_3d['labels'])
   
    fig.update_traces(
        selector=dict(type='scatter3d'),
        x=x_vals,
        y=y_vals,
        z=z_vals,
        marker=dict(
            color=colors,
            size=6,
            opacity=0.8
        ),
        text=labels
    )

def update_metrics_table(fig, metrics):
    """Actualiza la tabla de métricas"""
    # Preparar datos de la tabla
    names = []
    values = []
    trends = []
   
    for jornada in metrics:
        metric_data = metrics[jornada]['metrics']
        names.extend(metric_data['names'])
        values.extend([f"{v:.2f}" for v in metric_data['values']])
        trends.extend(metric_data['trends'])
   
    fig.update_traces(
        selector=dict(type='table'),
        cells=dict(
            values=[names, values, trends],
            align='left',
            font=dict(size=11),
            height=30
        )
    )

def update_kpi_indicators(fig, metrics):
    """Actualiza los indicadores KPI"""
    for jornada in metrics:
        kpi_data = metrics[jornada]['kpi']
       
        fig.update_traces(
            selector=dict(type='indicator'),
            value=kpi_data['current'],
            delta={
                'reference': kpi_data['previous'],
                'relative': True,
                'position': "top"
            }
        )

def prepare_metrics_data(df):
    """Prepara los datos de métricas para visualización"""
    metrics = {}
   
    for jornada in df['Jornada'].unique():
        df_jornada = df[df['Jornada'] == jornada]
       
        # Datos de tendencia
        trend = calculate_trend_data(df_jornada)
       
        # Datos de carga de trabajo
        workload = calculate_workload_data(df_jornada)
       
        # Datos 3D
        analysis_3d = calculate_3d_analysis(df_jornada)
       
        # Métricas y KPIs
        metrics_data = calculate_metrics(df_jornada)
        kpi_data = calculate_kpis(df_jornada)
       
        metrics[jornada] = {
            'trend': trend,
            'workload': workload,
            'analysis_3d': analysis_3d,
            'metrics': metrics_data,
            'kpi': kpi_data,
            'correlation': calculate_correlations(df_jornada)
        }
   
    return metrics

def setup_admin_tab(tab):
    """Configura la pestaña de centro de administración"""
    # Notebook para las pestañas de administración
    admin_notebook = ttk.Notebook(tab)
    admin_notebook.pack(fill='both', expand=True, padx=10, pady=10)

    # === Pestaña de Gestión de Personal ===
    personnel_tab = ttk.Frame(admin_notebook)
    admin_notebook.add(personnel_tab, text="Gestión de Personal")

    # Frame principal con PanedWindow
    main_paned = ttk.PanedWindow(personnel_tab, orient=tk.HORIZONTAL)
    main_paned.pack(fill='both', expand=True, padx=5, pady=5)

    # Panel izquierdo: Lista de personal
    left_frame = ttk.LabelFrame(main_paned, text="Personal")
   
    # Toolbar
    toolbar = ttk.Frame(left_frame)
    toolbar.pack(fill='x', padx=5, pady=5)

    ttk.Button(
        toolbar,
        text="Agregar Técnico",
        command=lambda: show_technician_dialog(tree),
        style='Primary.TButton'
    ).pack(side='left', padx=2)

    ttk.Button(
        toolbar,
        text="Eliminar",
        command=lambda: delete_technician_record(tree),
        style='Danger.TButton'
    ).pack(side='left', padx=2)

    # Búsqueda
    search_frame = ttk.Frame(toolbar)
    search_frame.pack(side='right', padx=5)
   
    ttk.Label(search_frame, text="Buscar:").pack(side='left')
    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var)
    search_entry.pack(side='left', padx=2)

    # TreeView para lista de personal
    tree = ttk.Treeview(
        left_frame,
        columns=('nombre', 'cargo', 'turno', 'letra'),
        show='headings',
        height=20
    )
   
    # Configurar columnas
    tree.heading('nombre', text='Nombre')
    tree.heading('cargo', text='Cargo')
    tree.heading('turno', text='Turno')
    tree.heading('letra', text='Letra')
   
    for col in ('nombre', 'cargo', 'turno', 'letra'):
        tree.column(col, width=100)
   
    # Scrollbars
    vsb = ttk.Scrollbar(left_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(left_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
   
    # Grid
    tree.pack(side='left', fill='both', expand=True, pady=5)
    vsb.pack(side='right', fill='y', pady=5)
    hsb.pack(side='bottom', fill='x')

    # Panel derecho: Detalles y edición
    right_frame = ttk.LabelFrame(main_paned, text="Detalles del Técnico")
   
    # Formulario de detalles
    details_frame = ttk.Frame(right_frame)
    details_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # Campos de edición
    fields = ['Nombre:', 'Cargo:', 'Turno:', 'Letra de Rotación:']
    details_vars = {}
   
    for i, field in enumerate(fields):
        ttk.Label(details_frame, text=field).grid(row=i, column=0, sticky='w', pady=5)
        var = tk.StringVar()
        details_vars[field] = var
        ttk.Entry(details_frame, textvariable=var).grid(row=i, column=1, sticky='ew', padx=5, pady=5)
   
    # Botones de acción
    button_frame = ttk.Frame(details_frame)
    button_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)

    ttk.Button(
        button_frame,
        text="Guardar Cambios",
        command=lambda: save_technician_changes(tree, details_vars),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
        button_frame,
        text="Cancelar",
        command=lambda: clear_details(details_vars)
    ).pack(side='left', padx=5)

    # Agregar paneles al PanedWindow
    main_paned.add(left_frame, weight=2)
    main_paned.add(right_frame, weight=1)

    # === Pestaña de Base de Datos ===
    db_tab = ttk.Frame(admin_notebook)
    admin_notebook.add(db_tab, text="Base de Datos")

    # Frame para mantenimiento
    maintenance_frame = ttk.LabelFrame(db_tab, text="Mantenimiento de Base de Datos", padding=10)
    maintenance_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        maintenance_frame,
        text="Verificar Integridad",
        command=verify_data_integrity,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        maintenance_frame,
        text="Limpiar Registros Antiguos",
        command=clean_old_records,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # Frame para respaldos
    backup_frame = ttk.LabelFrame(db_tab, text="Gestión de Respaldos", padding=10)
    backup_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        backup_frame,
        text="Crear Respaldo",
        command=backup_database,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        backup_frame,
        text="Restaurar desde Respaldo",
        command=restore_database,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # === Pestaña de Logs ===
    setup_log_tab(admin_notebook)

    # Cargar datos iniciales
    load_technicians_list(tree)
   
    # Configurar búsqueda en tiempo real
    search_var.trace('w', lambda *args: filter_technicians_list(tree, search_var.get()))
   
    return tab

def save_configuration(config_vars):
    """Guarda la configuración del sistema"""
    try:
        config = {
            'auto_backup': config_vars['auto_backup'].get(),
            'data_retention': config_vars['data_retention'].get(),
            'min_staff': config_vars['min_staff'].get()
        }
       
        # Aquí guardarías la configuración en un archivo o base de datos
        messagebox.showinfo("Éxito", "Configuración guardada correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

def calculate_trend_data(df):
    """Calcula datos de tendencia temporal"""
    dates = df['fecha_asignacion'].unique()
    counts = df.groupby('fecha_asignacion').size()
   
    return {
        'dates': dates,
        'values': counts.values
    }

def calculate_workload_data(df):
    """Calcula datos de carga de trabajo"""
    return {
        'dates': df['fecha_asignacion'].unique(),
        'values': df.groupby('fecha_asignacion').size().values
    }

def calculate_3d_analysis(df):
    """Calcula datos para análisis 3D"""
    return {
        'x': df.groupby('fecha_asignacion').size().values,
        'y': df.groupby('cargo').size().values,
        'z': df.groupby('letra_rotacion').size().values,
        'colors': np.random.rand(len(df)),
        'labels': df['nombre'].unique()
    }

def calculate_metrics(df):
    """Calcula métricas básicas"""
    daily_counts = df.groupby('fecha_asignacion').size()
   
    return {
        'names': ['Promedio', 'Máximo', 'Mínimo', 'Desviación'],
        'values': [
            daily_counts.mean(),
            daily_counts.max(),
            daily_counts.min(),
            daily_counts.std()
        ],
        'trends': ['↑', '↓', '→', '↔']
    }

def calculate_kpis(df):
    """Calcula KPIs"""
    current = len(df)
    previous = len(df[df['fecha_asignacion'] < df['fecha_asignacion'].max()])
   
    return {
        'current': current,
        'previous': previous
    }

def calculate_correlations(df):
    """Calcula correlaciones entre diferentes métricas"""
    # Ejemplo simple de correlación entre fechas y cantidad
    dates_numeric = (df['fecha_asignacion'] - df['fecha_asignacion'].min()).dt.days
    daily_counts = df.groupby('fecha_asignacion').size()
   
    return np.corrcoef(dates_numeric, daily_counts)[0, 1]





def setup_advanced_analysis_interface(tab):
    """Configura la interfaz avanzada de análisis"""
    # Frame principal con dos columnas
    main_frame = ttk.Frame(tab)
    main_frame.pack(fill='both', expand=True, padx=10, pady=5)
    main_frame.grid_columnconfigure(1, weight=1)

    # Panel de Control (Izquierda)
    control_frame = ttk.LabelFrame(main_frame, text="Controles de Análisis Avanzado")
    control_frame.grid(row=0, column=0, sticky='ns', padx=(0, 10), pady=5)

    # Fecha inicial
    date_frame = ttk.LabelFrame(control_frame, text="Rango de Fechas", padding=5)
    date_frame.pack(fill='x', pady=5)

    date_start = Calendar(
        date_frame,
        selectmode='day',
        date_pattern='dd/mm/yyyy',
        background='white',
        foreground='black',
        selectbackground='#007bff',
        width=300
    )
    date_start.pack(pady=5)

    # Fecha final
    date_end = Calendar(
        date_frame,
        selectmode='day',
        date_pattern='dd/mm/yyyy',
        background='white',
        foreground='black',
        selectbackground='#007bff',
        width=300
    )
    date_end.pack(pady=5)

    # Controles de análisis
    analysis_controls = {
        'date_range': {
            'label': "Rango de Fechas",
            'type': 'daterange',
            'var': None
        },
        'confidence_level': {
            'label': "Nivel de Confianza",
            'type': 'scale',
            'var': tk.DoubleVar(value=0.95),
            'range': (0.8, 0.99)
        },
        'forecast_horizon': {
            'label': "Horizonte de Predicción (días)",
            'type': 'spinbox',
            'var': tk.IntVar(value=14),
            'range': (7, 90)
        }
    }

    # Crear controles
    for name, config in analysis_controls.items():
        frame = ttk.LabelFrame(control_frame, text=config['label'], padding=5)
        frame.pack(fill='x', pady=5, padx=5)
       
        if config['type'] == 'daterange':
            config['var'] = create_date_range_selector(frame)
        elif config['type'] == 'scale':
            create_scale_control(frame, config)
        elif config['type'] == 'spinbox':
            create_spinbox_control(frame, config)

    # Resto del código...

    # Botones de acción
    ttk.Button(
        control_frame,
        text="Ejecutar Análisis",
        command=lambda: run_advanced_analysis(analysis_controls),
        style='Primary.TButton'
    ).pack(fill='x', pady=5, padx=5)

    ttk.Button(
        control_frame,
        text="Exportar Resultados",
        command=lambda: export_analysis_results(analysis_controls),
        style='Primary.TButton'
    ).pack(fill='x', pady=5, padx=5)

    # Panel de Resultados (Derecha)
    results_frame = ttk.LabelFrame(main_frame, text="Resultados del Análisis")
    results_frame.grid(row=0, column=1, sticky='nsew', pady=5)

    # Notebook para diferentes vistas de resultados
    results_notebook = ttk.Notebook(results_frame)
    results_notebook.pack(fill='both', expand=True)

    # Pestañas de resultados
    tabs = {
        'overview': {'text': "Vista General", 'frame': None},
        'predictions': {'text': "Predicciones", 'frame': None},
        'patterns': {'text': "Patrones", 'frame': None},
        'metrics': {'text': "Métricas", 'frame': None}
    }

    for name, config in tabs.items():
        config['frame'] = ttk.Frame(results_notebook)
        results_notebook.add(config['frame'], text=config['text'])

    return main_frame, analysis_controls, tabs

def create_date_range_selector(parent):
    """Crea selector de rango de fechas"""
    frame = ttk.Frame(parent)
    frame.pack(fill='x', pady=5)
   
    start_date = Calendar(frame, selectmode='day', date_pattern='dd/mm/yyyy')
    end_date = Calendar(frame, selectmode='day', date_pattern='dd/mm/yyyy')
   
    start_date.pack(side='left', padx=5)
    end_date.pack(side='left', padx=5)
   
    return {'start': start_date, 'end': end_date}

def create_scale_control(parent, config):
    """Crea control de escala"""
    scale = ttk.Scale(
        parent,
        from_=config['range'][0],
        to=config['range'][1],
        variable=config['var'],
        orient='horizontal'
    )
    scale.pack(fill='x', pady=5)
   
    value_label = ttk.Label(parent, textvariable=config['var'])
    value_label.pack()

def create_spinbox_control(parent, config):
    """Crea control spinbox"""
    spinbox = ttk.Spinbox(
        parent,
        from_=config['range'][0],
        to=config['range'][1],
        textvariable=config['var']
    )
    spinbox.pack(fill='x', pady=5)


def add_trend_plot(fig, metrics, row, col):
    """Agrega gráfico de tendencias temporales"""
    for jornada, data in metrics.items():
        fig.add_trace(
            go.Scatter(
                x=data['dates'],
                y=data['trend']['values'],
                name=f'Tendencia {jornada}',
                mode='lines+markers',
                line=dict(width=2),
                marker=dict(size=6),
                hovertemplate=(
                    f"Fecha: %{{x}}<br>" +
                    f"Valor: %{{y:.2f}}<br>" +
                    f"Jornada: {jornada}<br>"
                )
            ),
            row=row, col=col
        )
   
    fig.update_xaxes(title_text="Fecha", row=row, col=col)
    fig.update_yaxes(title_text="Cantidad", row=row, col=col)

def add_correlation_matrix(fig, metrics, row, col):
    """Agrega matriz de correlación"""
    corr_data = calculate_correlation_matrix(metrics)
   
    fig.add_trace(
        go.Heatmap(
            z=corr_data['values'],
            x=corr_data['labels'],
            y=corr_data['labels'],
            colorscale='RdBu',
            zmid=0,
            showscale=True,
            text=np.round(corr_data['values'], 2),
            texttemplate="%{text}",
            textfont={"size": 10},
            hoverongaps=False
        ),
        row=row, col=col
    )

def add_workload_distribution(fig, metrics, row, col):
    """Agrega distribución de carga de trabajo"""
    for jornada, data in metrics.items():
        fig.add_trace(
            go.Bar(
                x=data['dates'],
                y=data['workload'],
                name=jornada,
                text=data['workload'],
                textposition='auto',
            ),
            row=row, col=col
        )
   
    fig.update_xaxes(title_text="Fecha", row=row, col=col)
    fig.update_yaxes(title_text="Carga de Trabajo", row=row, col=col)

def add_3d_analysis(fig, metrics, row, col):
    """Agrega análisis 3D de métricas"""
    fig.add_trace(
        go.Scatter3d(
            x=metrics['x_values'],
            y=metrics['y_values'],
            z=metrics['z_values'],
            mode='markers',
            marker=dict(
                size=6,
                color=metrics['colors'],
                colorscale='Viridis',
                opacity=0.8
            ),
            text=metrics['labels']
        ),
        row=row, col=col
    )

def add_metrics_table(fig, metrics, row, col):
    """Agrega tabla de resumen de métricas"""
    fig.add_trace(
        go.Table(
            header=dict(
                values=["Métrica", "Valor", "Tendencia"],
                align="left",
                font=dict(size=12, color="white"),
                fill_color="darkblue"
            ),
            cells=dict(
                values=[
                    metrics['names'],
                    metrics['values'],
                    metrics['trends']
                ],
                align="left",
                font=dict(size=11),
                height=30
            )
        ),
        row=row, col=col
    )

def add_kpi_indicators(fig, metrics, row, col):
    """Agrega indicadores clave de rendimiento"""
    fig.add_trace(
        go.Indicator(
            mode="number+delta",
            value=metrics['current_value'],
            delta={'reference': metrics['previous_value'],
                   'relative': True,
                   'position': "top"},
            title={'text': "Eficiencia Global"},
            domain={'row': row, 'column': col}
        ),
        row=row, col=col
    )


def analyze_advanced_metrics(df):
    """Analiza métricas avanzadas con visualización mejorada"""
    metrics = {}
   
    # Análisis por jornada
    for jornada in df['Jornada'].unique():
        df_jornada = df[df['Jornada'] == jornada]
       
        # Métricas base
        total = len(df_jornada)
        daily_load = df_jornada.groupby('fecha_asignacion').size()
       
        # Métricas avanzadas
        metrics[jornada] = {
            'total': total,
            'mean': daily_load.mean(),
            'std': daily_load.std(),
            'cv': daily_load.std() / daily_load.mean() if daily_load.mean() > 0 else 0,
            'trend': calculate_trend(daily_load),
            'stability': calculate_stability_score(df_jornada),
            'balance': calculate_balance_score(df_jornada)
        }
       
        # Predicciones
        metrics[jornada]['forecast'] = generate_advanced_forecast(df_jornada)
   
    return metrics

def calculate_trend(series):
    """Calcula tendencia usando regresión robusta"""
    X = np.arange(len(series)).reshape(-1, 1)
    model = HuberRegressor()
    model.fit(X, series)
    return {
        'slope': model.coef_[0],
        'score': model.score(X, series),
        'prediction_interval': calculate_prediction_interval(model, X, series)
    }

def calculate_stability_score(df):
    """Calcula score de estabilidad basado en múltiples factores"""
    factors = {
        'rotation_consistency': analyze_rotation_patterns(df),
        'team_continuity': analyze_team_continuity(df),
        'workload_balance': analyze_workload_distribution(df)
    }
   
    weights = {
        'rotation_consistency': 0.4,
        'team_continuity': 0.3,
        'workload_balance': 0.3
    }
   
    return sum(score * weights[factor] for factor, score in factors.items())

def generate_advanced_forecast(df, horizon=14):
    """Genera pronósticos avanzados con intervalos de confianza"""
    # Preparar datos
    X = prepare_advanced_features(df)
    y = prepare_targets(df)
   
    # Dividir datos
    X_train, X_test = temporal_split(X)
    y_train, y_test = temporal_split(y)
   
    # Entrenar modelo ensamblado
    models = {
        'rf': RandomForestRegressor(n_estimators=200),
        'gb': GradientBoostingRegressor(),
        'lr': LinearRegression()
    }
   
    predictions = {}
    for name, model in models.items():
        model.fit(X_train, y_train)
        predictions[name] = {
            'forecast': model.predict(X_test),
            'score': model.score(X_test, y_test),
            'importance': get_feature_importance(model)
        }
   
    # Combinar predicciones
    ensemble_forecast = combine_predictions(predictions, weights={'rf': 0.5, 'gb': 0.3, 'lr': 0.2})
   
    return {
        'point_forecast': ensemble_forecast,
        'intervals': calculate_confidence_intervals(predictions),
        'model_diagnostics': generate_model_diagnostics(predictions)
    }

def visualize_advanced_metrics(metrics):
    """Crea visualizaciones avanzadas de las métricas"""
    fig = make_subplots(
        rows=3, cols=2,
        specs=[[{'type': 'scatter'}, {'type': 'heatmap'}],
               [{'type': 'bar'}, {'type': 'scatter3d'}],
               [{'type': 'table'}, {'type': 'indicator'}]],
        subplot_titles=('Tendencias Temporales', 'Matriz de Correlación',
                       'Distribución de Carga', 'Análisis Multidimensional',
                       'Resumen de Métricas', 'Indicadores Clave')
    )
   
    # Configurar cada subplot con datos específicos
    add_trend_plot(fig, metrics, row=1, col=1)
    add_correlation_matrix(fig, metrics, row=1, col=2)
    add_workload_distribution(fig, metrics, row=2, col=1)
    add_3d_analysis(fig, metrics, row=2, col=2)
    add_metrics_table(fig, metrics, row=3, col=1)
    add_kpi_indicators(fig, metrics, row=3, col=2)
   
    # Actualizar layout
    fig.update_layout(height=1200, showlegend=True)
   
    return fig


def prepare_targets(df):
    """Prepara las variables objetivo para el modelo"""
    # Contar técnicos por día
    daily_counts = df.groupby('fecha_asignacion').size()
    return daily_counts.values

def evaluate_model(model, X_test, y_test):
    """Evalúa el modelo y retorna la puntuación"""
    predictions = model.predict(X_test)
    score = r2_score(y_test, predictions)
    return score

def analyze_anomaly_causes(anomalias_df):
    """Analiza las causas de las anomalías detectadas"""
    causes = {}
   
    # Análisis por día de la semana
    if 'fecha_asignacion' in anomalias_df.columns:
        dia_semana_counts = anomalias_df['fecha_asignacion'].dt.dayofweek.value_counts()
        for dia, count in dia_semana_counts.items():
            causes[f'Día {DIAS_SEMANA[dia]}'] = count
   
    # Análisis por cargo
    if 'cargo' in anomalias_df.columns:
        cargo_counts = anomalias_df['cargo'].value_counts()
        for cargo, count in cargo_counts.items():
            causes[f'Cargo {cargo}'] = count
   
    # Análisis por jornada
    if 'Jornada' in anomalias_df.columns:
        jornada_counts = anomalias_df['Jornada'].value_counts()
        for jornada, count in jornada_counts.items():
            causes[f'Jornada {jornada}'] = count
           
    return causes

def generate_recommendations(anomalias_df, jornada):
    """Genera recomendaciones basadas en el análisis de anomalías"""
    ai_text.insert(tk.END, f"\nRecomendaciones para jornada {jornada}:\n")
   
    # Analizar distribución de carga
    if 'fecha_asignacion' in anomalias_df.columns:
        daily_load = anomalias_df.groupby('fecha_asignacion').size()
        cv = daily_load.std() / daily_load.mean() if daily_load.mean() > 0 else 0
       
        if cv > 0.3:
            ai_text.insert(tk.END, "• Alta variabilidad en la carga diaria\n")
            ai_text.insert(tk.END, "  Recomendación: Implementar mejor distribución del personal\n")
   
    # Analizar patrones de rotación
    if 'letra_rotacion' in anomalias_df.columns:
        rotation_patterns = anomalias_df['letra_rotacion'].value_counts()
        if len(rotation_patterns) > 0:
            most_common = rotation_patterns.index[0]
            ai_text.insert(tk.END, f"• Patrón de rotación más común: {most_common}\n")
   
    # Analizar composición de equipos
    if 'cargo' in anomalias_df.columns:
        cargo_dist = anomalias_df['cargo'].value_counts()
        if len(cargo_dist) > 0:
            ai_text.insert(tk.END, "• Distribución de cargos en anomalías:\n")
            for cargo, count in cargo_dist.items():
                ai_text.insert(tk.END, f"  - {cargo}: {count} casos\n")

def optimize_workload(current_load, total_staff, jornada):
    """Optimiza la distribución de carga de trabajo considerando restricciones específicas"""
    optimization = {}
    mean_load = total_staff / 7  # distribución ideal
   
    # Parámetros específicos por jornada
    params = {
        'Mañana': {'peak_days': [0, 1], 'min_staff': 4},  # Lunes y Martes
        'Tarde': {'peak_days': [2, 3], 'min_staff': 4},   # Miércoles y Jueves
        'Noche': {'peak_days': [4, 5], 'min_staff': 3}    # Viernes y Sábado
    }
   
    jornada_params = params.get(jornada, params['Mañana'])
   
    for dia in range(7):
        current = current_load.get(dia, 0)
        is_peak_day = dia in jornada_params['peak_days']
       
        if is_peak_day:
            target = max(mean_load * 1.2, jornada_params['min_staff'])
        else:
            target = max(mean_load, jornada_params['min_staff'])
       
        optimization[dia] = round(target)
   
    return optimization



def configure_button_styles():
    """Configura los estilos de botones con énfasis en botones de peligro"""
    style = ttk.Style()
   
    # Definir colores para botones de peligro
    DANGER_COLORS = {
        'main': '#dc3545',      # Rojo brillante
        'hover': '#c82333',     # Rojo más oscuro para hover
        'disabled': '#f5c6cb',  # Rojo claro para disabled
        'text': '#ffffff'       # Texto blanco para mejor contraste con rojo
    }

    # Botón de peligro (rojo)
    style.configure('Danger.TButton',
        background=DANGER_COLORS['main'],
        foreground=DANGER_COLORS['text'],
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
   
    # Mapeo para estados del botón de peligro
    style.map('Danger.TButton',
        background=[
            ('active', DANGER_COLORS['hover']),
            ('disabled', DANGER_COLORS['disabled'])
        ],
        foreground=[
            ('disabled', '#6c757d'),
            ('active', DANGER_COLORS['text'])
        ]
    )

    # Botón primario (azul, para contraste)
    style.configure('Primary.TButton',
        background='#0d6efd',
        foreground='#000000',
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
   
    # Mapeo para estados del botón primario
    style.map('Primary.TButton',
        background=[
            ('active', '#0b5ed7'),
            ('disabled', '#cce5ff')
        ],
        foreground=[
            ('disabled', '#666666'),
            ('active', '#000000')
        ]
    )

    return {
        'danger': DANGER_COLORS
    }

def create_delete_button(parent, command=None, text="Eliminar Seleccionado"):
    """Crea un botón de eliminar con el estilo correcto"""
    return ttk.Button(
        parent,
        text=text,
        command=command,
        style='Danger.TButton'
    )


def configure_modern_styles():
    """Configura estilos modernos con mejor legibilidad"""
    style = ttk.Style()
   
    # Colores actualizados con mejor contraste
    COLORS = {
        'primary': {
            'main': '#0d6efd',      # Azul más visible
            'hover': '#0b5ed7',     # Azul más oscuro para hover
            'disabled': '#cce5ff',   # Azul claro para disabled
            'text': '#000000'       # Texto negro para mejor legibilidad
        },
        'danger': {
            'main': '#dc3545',      # Rojo más visible
            'hover': '#bb2d3b',     # Rojo más oscuro para hover
            'disabled': '#f8d7da',   # Rojo claro para disabled
            'text': '#000000'       # Texto negro
        },
        'success': {
            'main': '#198754',      # Verde más visible
            'hover': '#157347',     # Verde más oscuro para hover
            'disabled': '#d1e7dd',   # Verde claro para disabled
            'text': '#000000'       # Texto negro
        },
        'secondary': {
            'main': '#5c636a',      # Gris más oscuro y visible
            'hover': '#4d5154',     # Gris más oscuro para hover
            'disabled': '#e9ecef',   # Gris claro para disabled
            'text': '#000000'       # Texto negro
        }
    }

    # Botón normal
    style.configure('TButton',
        background=COLORS['secondary']['main'],
        foreground=COLORS['secondary']['text'],
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
   
    # Botón primario (azul)
    style.configure('Primary.TButton',
        background=COLORS['primary']['main'],
        foreground=COLORS['primary']['text'],
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )

    # Botón de peligro (rojo)
    style.configure('Danger.TButton',
        background=COLORS['danger']['main'],
        foreground=COLORS['danger']['text'],
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )

    # Botón de éxito (verde)
    style.configure('Success.TButton',
        background=COLORS['success']['main'],
        foreground=COLORS['success']['text'],
        padding=(10, 5),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )

    # Estilos para los botones de jornada
    jornada_styles = {
        'Morning': {
            'bg': '#ffd54f',        # Amarillo más claro
            'fg': '#000000',        # Texto negro
            'hover': '#ffb300',     # Amarillo más oscuro para hover
            'selected': '#ffa000'   # Amarillo aún más oscuro para seleccionado
        },
        'Afternoon': {
            'bg': '#4fc3f7',        # Azul claro
            'fg': '#000000',        # Texto negro
            'hover': '#039be5',     # Azul más oscuro para hover
            'selected': '#0288d1'   # Azul aún más oscuro para seleccionado
        },
        'Night': {
            'bg': '#9575cd',        # Púrpura claro
            'fg': '#000000',        # Texto negro
            'hover': '#7e57c2',     # Púrpura más oscuro para hover
            'selected': '#673ab7'   # Púrpura aún más oscuro para seleccionado
        }
    }

    # Configurar estilos para cada jornada
    for shift, colors in jornada_styles.items():
        # Estilo normal
        style.configure(f'{shift}.TButton',
            background=colors['bg'],
            foreground=colors['fg'],
            padding=(15, 8),
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            borderwidth=0
        )
       
        # Estilo seleccionado
        style.configure(f'Selected.{shift}.TButton',
            background=colors['selected'],
            foreground=colors['fg'],
            padding=(15, 8),
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            borderwidth=0
        )
       
        # Mapeo para hover
        style.map(f'{shift}.TButton',
            background=[('active', colors['hover'])],
            foreground=[('active', colors['fg'])]
        )
        style.map(f'Selected.{shift}.TButton',
            background=[('active', colors['selected'])],
            foreground=[('active', colors['fg'])]
        )

    return COLORS

# Función para aplicar los estilos a un botón específico
def apply_button_style(button, style_name):
    """Aplica un estilo específico a un botón"""
    button.configure(style=style_name)

class NotificationManager:
    def __init__(self, parent):
        self.parent = parent
        self.notifications = []
        self.notification_frame = None
        self.setup_notification_frame()
   
    def setup_notification_frame(self):
        """Configura el frame para notificaciones"""
        self.notification_frame = ttk.Frame(self.parent)
        self.notification_frame.place(relx=1.0, y=10, anchor='ne', relwidth=0.3)
   
    def show_notification(self, message, type_='info', duration=3000):
        """Muestra una notificación con animación"""
        # Crear frame para la notificación
        notification = ttk.Frame(self.notification_frame, style='Card.TFrame')
        notification.pack(pady=5, padx=10, fill='x')
       
        # Iconos para diferentes tipos
        icons = {
            'success': '✓',
            'error': '✕',
            'warning': '⚠',
            'info': 'ℹ'
        }
       
        # Colores para diferentes tipos
        colors = {
            'success': '#28a745',
            'error': '#dc3545',
            'warning': '#ffc107',
            'info': '#17a2b8'
        }
       
        # Crear contenido de la notificación
        icon_label = ttk.Label(notification,
                             text=icons.get(type_, 'ℹ'),
                             foreground=colors.get(type_, '#17a2b8'),
                             font=('Segoe UI', 14))
        icon_label.pack(side='left', padx=10, pady=5)
       
        message_label = ttk.Label(notification,
                                text=message,
                                wraplength=250,
                                font=('Segoe UI', 10))
        message_label.pack(side='left', padx=5, pady=5, fill='x', expand=True)
       
        # Botón cerrar
        close_button = ttk.Label(notification,
                               text='✕',
                               cursor='hand2',
                               font=('Segoe UI', 12))
        close_button.pack(side='right', padx=10)
        close_button.bind('<Button-1>',
                         lambda e: self.close_notification(notification))
       
        # Animación de entrada
        self.animate_notification(notification, 'in')
       
        # Auto-cerrar después de la duración especificada
        if duration:
            self.parent.after(duration,
                            lambda: self.close_notification(notification))
   
    def animate_notification(self, notification, direction):
        """Anima la entrada o salida de la notificación"""
        if direction == 'in':
            notification.place(relx=1.0)
            for i in range(10, 0, -1):
                notification.place(relx=i/10)
                self.parent.update()
                self.parent.after(20)
        else:
            for i in range(0, 11):
                notification.place(relx=i/10)
                self.parent.update()
                self.parent.after(20)
            notification.destroy()
   
    def close_notification(self, notification):
        """Cierra una notificación específica"""
        self.animate_notification(notification, 'out')
       
       
class ModernTooltip:
    def __init__(self, widget, text, delay=500):
        """
        Inicializa un tooltip moderno
        widget: El widget al que se asocia el tooltip
        text: El texto a mostrar
        delay: Retraso antes de mostrar el tooltip (ms)
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.id = None

        # Vincular eventos
        self.widget.bind('<Enter>', self.schedule_tooltip)
        self.widget.bind('<Leave>', self.hide_tooltip)
        self.widget.bind('<Button-1>', self.hide_tooltip)

    def schedule_tooltip(self, event=None):
        """Programa la aparición del tooltip"""
        self.cancel_tooltip()
        self.id = self.widget.after(self.delay, self.show_tooltip)

    def cancel_tooltip(self):
        """Cancela la aparición programada del tooltip"""
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def show_tooltip(self):
        """Muestra el tooltip con animación"""
        # Ocultar tooltip existente si hay uno
        self.hide_tooltip()

        # Obtener posición del widget
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 25

        # Crear ventana del tooltip
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)

        # Frame principal con estilo
        frame = ttk.Frame(self.tooltip_window, style='Tooltip.TFrame')
        frame.pack(padx=1, pady=1)

        # Etiqueta con el texto
        label = ttk.Label(
            frame,
            text=self.text,
            justify=tk.LEFT,
            background='#333333',
            foreground='white',
            relief=tk.SOLID,
            borderwidth=0,
            wraplength=250,
            font=('Segoe UI', 9)
        )
        label.pack(padx=5, pady=3)

        # Posicionar tooltip
        self.tooltip_window.wm_geometry(f"+{x}+{y}")

        # Animación de entrada
        self.tooltip_window.attributes('-alpha', 0.0)
        for i in range(0, 11):
            if self.tooltip_window:  # Verificar que la ventana aún existe
                self.tooltip_window.attributes('-alpha', i/10)
                self.tooltip_window.update()
                time.sleep(0.02)  # Pequeña pausa para la animación

    def hide_tooltip(self, event=None):
        """Oculta el tooltip con animación"""
        # Cancelar cualquier tooltip programado
        self.cancel_tooltip()

        # Si hay un tooltip visible, ocultarlo con animación
        if self.tooltip_window:
            try:
                # Animación de salida
                for i in range(10, -1, -1):
                    if self.tooltip_window:  # Verificar que la ventana aún existe
                        self.tooltip_window.attributes('-alpha', i/10)
                        self.tooltip_window.update()
                        time.sleep(0.02)  # Pequeña pausa para la animación
               
                # Destruir la ventana
                if self.tooltip_window:
                    self.tooltip_window.destroy()
                    self.tooltip_window = None
            except:
                # Si hay algún error, simplemente intentar destruir la ventana
                if self.tooltip_window:
                    self.tooltip_window.destroy()
                    self.tooltip_window = None

def add_tooltips(notebook):
    """Agrega tooltips a los elementos del notebook"""
    tooltips = {
        'Gestión de Turnos': 'Administre y visualice los turnos del personal',
        'Ausencias': 'Gestione y monitoree las ausencias del personal',
        'Consulta y Reportes': 'Genere informes y consulte históricos',
        'Resumen': 'Visualice estadísticas y métricas clave',
        'Análisis Avanzado': 'Acceda a análisis predictivo y tendencias'
    }

    # Agregar tooltip a cada pestaña
    for tab_id in notebook.tabs():
        tab_text = notebook.tab(tab_id, 'text').split(' ')[-1]  # Eliminar emoji
        if tab_text in tooltips:
            ModernTooltip(
                notebook.select(tab_id),  # Widget de la pestaña
                tooltips[tab_text],
                delay=500
            )

def setup_personnel_manager(tab):
    """Configura la interfaz de gestión de personal"""
    # Split frame con lista de personal y detalles
    paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
    paned.pack(fill='both', expand=True, padx=5, pady=5)

    # === Panel Izquierdo: Lista de Personal ===
    left_frame = ttk.Frame(paned)
    left_frame.grid_columnconfigure(0, weight=1)
    left_frame.grid_rowconfigure(1, weight=1)

    # Barra de herramientas
    toolbar = ttk.Frame(left_frame)
    toolbar.grid(row=0, column=0, sticky='ew', padx=5, pady=5)

    ttk.Button(
        toolbar,
        text="Nuevo Técnico",
        command=lambda: show_technician_dialog(tree, None),
        style='Primary.TButton'
    ).pack(side='left', padx=2)

    ttk.Button(
        toolbar,
        text="Eliminar",
        command=lambda: delete_technician_record(tree),
        style='Danger.TButton'
    ).pack(side='left', padx=2)

    # Búsqueda
    search_frame = ttk.Frame(toolbar)
    search_frame.pack(side='right', padx=2)
   
    ttk.Label(search_frame, text="Buscar:").pack(side='left')
    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var)
    search_entry.pack(side='left', padx=2)
   
    search_var.trace('w', lambda *args: filter_technicians_list(tree, search_var.get()))

    # Lista de personal
    tree_frame = ttk.Frame(left_frame)
    tree_frame.grid(row=1, column=0, sticky='nsew', padx=5)
    tree_frame.grid_columnconfigure(0, weight=1)
    tree_frame.grid_rowconfigure(0, weight=1)

    columns = ('nombre', 'turno', 'cargo', 'letra')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
   
    # Configurar columnas
    tree.heading('nombre', text='Nombre')
    tree.heading('turno', text='Turno')
    tree.heading('cargo', text='Cargo')
    tree.heading('letra', text='Letra')
   
    for col in columns:
        tree.column(col, width=100)

    # Scrollbars
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    # Grid
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')

    # === Panel Derecho: Detalles y Edición ===
    right_frame = ttk.LabelFrame(paned, text="Detalles del Técnico")
   
    # Campos de edición
    details_frame = ttk.Frame(right_frame)
    details_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # Variables para los campos
    details_vars = {
        'nombre': tk.StringVar(),
        'turno': tk.StringVar(),
        'cargo': tk.StringVar(),
        'letra': tk.StringVar()
    }

    # Crear campos
    for i, (field, var) in enumerate(details_vars.items()):
        ttk.Label(details_frame, text=f"{field.title()}:").grid(row=i, column=0, sticky='w', pady=5)
        entry = ttk.Entry(details_frame, textvariable=var)
        entry.grid(row=i, column=1, sticky='ew', padx=5, pady=5)
       
    # Botones de acción
    button_frame = ttk.Frame(details_frame)
    button_frame.grid(row=len(details_vars), column=0, columnspan=2, pady=10)

    ttk.Button(
        button_frame,
        text="Guardar Cambios",
        command=lambda: save_technician_changes(tree, details_vars),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
        button_frame,
        text="Cancelar",
        command=lambda: clear_details(details_vars)
    ).pack(side='left', padx=5)

    # Añadir los paneles al PanedWindow
    paned.add(left_frame, weight=2)
    paned.add(right_frame, weight=1)

    # Manejar selección en el tree
    tree.bind('<<TreeviewSelect>>', lambda e: load_technician_details(tree, details_vars))
   
    # Cargar datos iniciales
    load_technicians_list(tree)

def show_technician_dialog(tree, technician_data=None):
    """Muestra diálogo para añadir o editar técnico"""
    dialog = tk.Toplevel()
    dialog.title("Nuevo Técnico" if not technician_data else "Editar Técnico")
    dialog.transient(tree.winfo_toplevel())
    dialog.grab_set()

    # Centrar ventana
    window_width = 400
    window_height = 350
    center_window(dialog, window_width, window_height)

    # Frame principal
    main_frame = ttk.Frame(dialog, padding="20")
    main_frame.pack(fill='both', expand=True)

    # Variables
    vars_dict = {
        'nombre': tk.StringVar(value=technician_data.get('nombre', '') if technician_data else ''),
        'turno': tk.StringVar(value=technician_data.get('turno', '') if technician_data else ''),
        'cargo': tk.StringVar(value=technician_data.get('cargo', '') if technician_data else ''),
        'letra': tk.StringVar(value=technician_data.get('letra', '') if technician_data else '')
    }

    # Sección de información
    info_frame = ttk.LabelFrame(main_frame, text="Información del Técnico", padding=10)
    info_frame.pack(fill='x', expand=True, pady=(0, 10))

    # Lista de cargos disponibles
    CARGOS = [
        'Técnico Mantenimiento 2',
        'Técnico Mantenimiento 1',
        'Técnico Depanaje',
        'Supervisor'
    ]

    # Campos
    for i, (field, var) in enumerate(vars_dict.items()):
        field_frame = ttk.Frame(info_frame)
        field_frame.pack(fill='x', pady=5)

        ttk.Label(
            field_frame,
            text=f"{field.title()}:",
            font=('Segoe UI', 10),
            width=15
        ).pack(side='left')
       
        if field == 'turno':
            combo = ttk.Combobox(
                field_frame,
                textvariable=var,
                values=['A', 'B', 'C', 'prev_nocturno'],
                state='readonly',
                width=30
            )
            combo.pack(side='left', fill='x', expand=True)
            # Establecer valor por defecto
            if not var.get():
                combo.set('A')
        elif field == 'cargo':
            cargo_combo = ttk.Combobox(
                field_frame,
                textvariable=var,
                values=CARGOS,
                state='readonly',
                width=30
            )
            cargo_combo.pack(side='left', fill='x', expand=True)
            if not var.get() and CARGOS:
                cargo_combo.set(CARGOS[0])  # Establecer primer cargo como default
        else:
            entry = ttk.Entry(
                field_frame,
                textvariable=var,
                width=30
            )
            entry.pack(side='left', fill='x', expand=True)

    # Frame para mensaje de ayuda
    help_frame = ttk.Frame(main_frame)
    help_frame.pack(fill='x', pady=10)
   
    ttk.Label(
        help_frame,
        text="Todos los campos son obligatorios",
        font=('Segoe UI', 9, 'italic'),
        foreground='gray'
    ).pack(side='left')

    # Frame para botones
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill='x', pady=(20, 0))

    def validate_and_save():
        """Valida los campos antes de guardar"""
        # Validar campos obligatorios
        for field, var in vars_dict.items():
            if not var.get().strip():
                messagebox.showerror(
                    "Error",
                    f"El campo {field} es obligatorio",
                    parent=dialog
                )
                return
           
        # Validar que el cargo sea válido
        if vars_dict['cargo'].get() not in CARGOS:
            messagebox.showerror(
                "Error",
                "El cargo seleccionado no es válido",
                parent=dialog
            )
            return
       
        # Si todo está validado, guardar
        save_new_technician(dialog, tree, vars_dict)

    ttk.Button(
        button_frame,
        text="Guardar",
        command=validate_and_save,
        style='Primary.TButton'
    ).pack(side='left', padx=5, expand=True)

    ttk.Button(
        button_frame,
        text="Cancelar",
        command=dialog.destroy
    ).pack(side='left', padx=5, expand=True)

    # Configurar estilo
    style = ttk.Style()
    style.configure('Primary.TButton', font=('Segoe UI', 10))
    style.configure('TLabel', font=('Segoe UI', 10))

    # Dar foco al primer campo
    first_entry = dialog.winfo_children()[0].winfo_children()[0].winfo_children()[0].winfo_children()[1]
    first_entry.focus_set()

    # Hacer la ventana modal
    dialog.focus_force()
    dialog.grab_set()
    dialog.wait_window()

def verify_data_integrity():
    """Verifica la integridad de los datos en la base de datos"""
    try:
        df = pd.read_excel(excel_file_path)
        issues = []

        # Lista de valores válidos
        VALID_TURNOS = ['A', 'B', 'C', 'prev_nocturno']
        VALID_CARGOS = [
            'Técnico Mantenimiento 2',
            'Técnico Mantenimiento 1',
            'Técnico Depanaje',
            'Supervisor'
        ]

        # Verificar campos obligatorios
        required_fields = ['nombre_bd', 'turno_bd', 'cargo_bd', 'Letra_bd']
        for field in required_fields:
            if field not in df.columns:
                issues.append(f"Falta la columna {field}")
            else:
                null_count = df[field].isna().sum()
                if null_count > 0:
                    issues.append(f"{null_count} registros con {field} vacío")

        # Verificar duplicados
        duplicates = df[df.duplicated(subset=['nombre_bd'], keep=False)]
        if not duplicates.empty:
            issues.append(f"{len(duplicates)} registros duplicados encontrados")

        # Verificar valores válidos en turnos
        invalid_turnos = df[~df['turno_bd'].isin(VALID_TURNOS)]['turno_bd'].unique()
        if len(invalid_turnos) > 0:
            issues.append(f"Turnos inválidos encontrados: {', '.join(map(str, invalid_turnos))}")

        # Verificar valores válidos en cargos
        invalid_cargos = df[~df['cargo_bd'].isin(VALID_CARGOS)]['cargo_bd'].unique()
        if len(invalid_cargos) > 0:
            issues.append(f"Cargos inválidos encontrados: {', '.join(map(str, invalid_cargos))}")

        # Mostrar resultados
        if issues:
            result = "Se encontraron los siguientes problemas:\n\n" + "\n".join(f"• {issue}" for issue in issues)
            if messagebox.askyesno("Problemas de Integridad",
                                 f"{result}\n\n¿Desea intentar corregir automáticamente estos problemas?"):
                fix_data_integrity(df, issues)
        else:
            messagebox.showinfo("Verificación Completada",
                              "No se encontraron problemas de integridad en los datos")

    except Exception as e:
        messagebox.showerror("Error", f"Error al verificar integridad: {str(e)}")

def fix_data_integrity(df, issues):
    """Intenta corregir problemas de integridad encontrados"""
    try:
        VALID_TURNOS = ['A', 'B', 'C', 'prev_nocturno']
        VALID_CARGOS = [
            'Técnico Mantenimiento 2',
            'Técnico Mantenimiento 1',
            'Técnico Depanaje',
            'Supervisor'
        ]

        # Crear backup antes de correcciones
        backup_path = f"{excel_file_path}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        df.to_excel(backup_path, index=False)

        fixed_issues = []

        # Corregir campos vacíos
        for field in ['nombre_bd', 'turno_bd', 'cargo_bd', 'Letra_bd']:
            if df[field].isna().any():
                df[field] = df[field].fillna('')
                fixed_issues.append(f"Campos vacíos en {field} corregidos")

        # Corregir duplicados
        if len(df[df.duplicated(subset=['nombre_bd'], keep=False)]) > 0:
            df = df.drop_duplicates(subset=['nombre_bd'], keep='first')
            fixed_issues.append("Registros duplicados eliminados")

        # Corregir turnos inválidos
        mask = ~df['turno_bd'].isin(VALID_TURNOS)
        if mask.any():
            df.loc[mask, 'turno_bd'] = 'A'
            fixed_issues.append("Turnos inválidos corregidos")

        # Corregir cargos inválidos
        mask = ~df['cargo_bd'].isin(VALID_CARGOS)
        if mask.any():
            df.loc[mask, 'cargo_bd'] = 'Técnico Mantenimiento 2'
            fixed_issues.append("Cargos inválidos corregidos")

        # Guardar cambios
        df.to_excel(excel_file_path, index=False)

        messagebox.showinfo(
            "Correcciones Aplicadas",
            "Se realizaron las siguientes correcciones:\n\n" +
            "\n".join(f"• {fix}" for fix in fixed_issues) +
            f"\n\nSe creó un respaldo en: {backup_path}"
        )

    except Exception as e:
        messagebox.showerror("Error", f"Error al corregir integridad: {str(e)}")

def save_new_technician(dialog, tree, vars_dict):
    """Guarda un nuevo técnico en la base de datos"""
    try:
        # Validar campos
        for field, var in vars_dict.items():
            if not var.get().strip():
                messagebox.showerror(
                    "Error",
                    f"El campo {field} es obligatorio",
                    parent=dialog
                )
                return

        # Validar que el nombre no exista ya
        df = pd.read_excel(excel_file_path)
        if vars_dict['nombre'].get() in df['nombre_bd'].values:
            messagebox.showerror(
                "Error",
                "Ya existe un técnico con ese nombre",
                parent=dialog
            )
            return

        # Crear nuevo registro
        new_data = {
            'nombre_bd': vars_dict['nombre'].get().strip(),
            'turno_bd': vars_dict['turno'].get().strip(),
            'cargo_bd': vars_dict['cargo'].get().strip(),
            'Letra_bd': vars_dict['letra'].get().strip()
        }

        # Añadir nuevo registro
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
       
        # Guardar cambios
        df.to_excel(excel_file_path, index=False)
       
        # Actualizar lista
        load_technicians_list(tree)
       
        messagebox.showinfo(
            "Éxito",
            "Técnico agregado correctamente",
            parent=dialog
        )
        dialog.destroy()

    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al guardar técnico: {str(e)}",
            parent=dialog
        )

def delete_technician_record(tree):
    """Elimina un técnico de la base de datos"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Advertencia", "Por favor seleccione un técnico para eliminar")
        return

    if not messagebox.askyesno(
        "Confirmar Eliminación",
        "¿Está seguro de que desea eliminar este técnico?\nEsta acción no se puede deshacer."
    ):
        return

    try:
        item = tree.item(selected[0])
        nombre = item['values'][0]

        # Leer Excel
        df = pd.read_excel(excel_file_path)
       
        # Eliminar registro
        df = df[df['nombre_bd'] != nombre]
       
        # Guardar cambios
        df.to_excel(excel_file_path, index=False)
       
        # Actualizar lista
        load_technicians_list(tree)
       
        messagebox.showinfo("Éxito", "Técnico eliminado correctamente")

    except Exception as e:
        messagebox.showerror("Error", f"Error al eliminar técnico: {str(e)}")

def load_technicians_list(tree):
    """Carga la lista de técnicos en el Treeview"""
    # Limpiar tabla
    for item in tree.get_children():
        tree.delete(item)

    try:
        # Leer Excel
        df = pd.read_excel(excel_file_path)
       
        # Filtrar solo registros con nombre_bd
        df = df[df['nombre_bd'].notna()]
       
        # Insertar registros en la tabla
        for _, row in df.iterrows():
            tree.insert('', 'end', values=(
                row['nombre_bd'],
                row['turno_bd'],
                row['cargo_bd'],
                row['Letra_bd']
            ))

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar lista de técnicos: {str(e)}")

def filter_technicians_list(tree, search_text):
    """Filtra la lista de técnicos según el texto de búsqueda"""
    try:
        # Limpiar tabla
        for item in tree.get_children():
            tree.delete(item)

        # Si no hay texto de búsqueda, mostrar todos
        if not search_text:
            load_technicians_list(tree)
            return

        # Leer Excel
        df = pd.read_excel(excel_file_path)
       
        # Filtrar registros que coincidan con la búsqueda
        df = df[
            df['nombre_bd'].str.contains(search_text, case=False, na=False) |
            df['turno_bd'].str.contains(search_text, case=False, na=False) |
            df['cargo_bd'].str.contains(search_text, case=False, na=False) |
            df['Letra_bd'].str.contains(search_text, case=False, na=False)
        ]

        # Insertar registros filtrados
        for _, row in df.iterrows():
            tree.insert('', 'end', values=(
                row['nombre_bd'],
                row['turno_bd'],
                row['cargo_bd'],
                row['Letra_bd']
            ))

    except Exception as e:
        print(f"Error al filtrar técnicos: {str(e)}")

def load_technician_details(tree, details_vars):
    """Carga los detalles del técnico seleccionado"""
    selected = tree.selection()
    if not selected:
        return

    item = tree.item(selected[0])
    values = item['values']

    details_vars['nombre'].set(values[0])
    details_vars['turno'].set(values[1])
    details_vars['cargo'].set(values[2])
    details_vars['letra'].set(values[3])
   

def save_technician_changes(tree, details_vars):
    """Guarda los cambios realizados a un técnico"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Advertencia", "Por favor seleccione un técnico para modificar")
        return

    try:
        old_values = tree.item(selected[0])['values']
        old_nombre = old_values[0]

        # Validar campos
        for var in details_vars.values():
            if not var.get().strip():
                messagebox.showerror("Error", "Todos los campos son obligatorios")
                return

        # Leer Excel
        df = pd.read_excel(excel_file_path)
       
        # Actualizar registro
        mask = df['nombre_bd'] == old_nombre
        df.loc[mask, 'nombre_bd'] = details_vars['nombre'].get()
        df.loc[mask, 'turno_bd'] = details_vars['turno'].get()
        df.loc[mask, 'cargo_bd'] = details_vars['cargo'].get()
        df.loc[mask, 'Letra_bd'] = details_vars['letra'].get()
       
        # Guardar cambios
        df.to_excel(excel_file_path, index=False)
       
        # Actualizar lista
        load_technicians_list(tree)
       
        messagebox.showinfo("Éxito", "Cambios guardados correctamente")

    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar cambios: {str(e)}")
       
       
       
def clear_details(details_vars):
    """Limpia los campos de detalles"""
    for var in details_vars.values():
        var.set('')

def setup_maintenance_tab(tab):
    """Configura la pestaña de mantenimiento y respaldos"""
    # Frame principal
    main_frame = ttk.Frame(tab, padding=10)
    main_frame.pack(fill='both', expand=True)

    # === Sección de Respaldos ===
    backup_frame = ttk.LabelFrame(main_frame, text="Gestión de Respaldos", padding=10)
    backup_frame.pack(fill='x', pady=(0, 10))

    ttk.Button(
        backup_frame,
        text="Crear Respaldo",
        command=backup_database,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        backup_frame,
        text="Restaurar desde Respaldo",
        command=restore_database,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # === Sección de Mantenimiento ===
    maintenance_frame = ttk.LabelFrame(main_frame, text="Mantenimiento de Base de Datos", padding=10)
    maintenance_frame.pack(fill='x', pady=(0, 10))

    ttk.Button(
        maintenance_frame,
        text="Limpiar Registros Antiguos",
        command=clean_old_records,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        maintenance_frame,
        text="Verificar Integridad de Datos",
        command=verify_data_integrity,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # === Sección de Estadísticas ===
    stats_frame = ttk.LabelFrame(main_frame, text="Estadísticas de Base de Datos", padding=10)
    stats_frame.pack(fill='x', pady=(0, 10))

    stats_text = tk.Text(stats_frame, height=10, wrap=tk.WORD)
    stats_text.pack(fill='x')
   
    update_db_stats(stats_text)

def verify_data_integrity():
    """Verifica la integridad de los datos en la base de datos"""
    try:
        df = pd.read_excel(excel_file_path)
        issues = []

        # Lista de valores válidos
        VALID_TURNOS = ['A', 'B', 'C', 'prev_nocturno']
        VALID_CARGOS = [
            'Técnico Mantenimiento 2',
            'Técnico Mantenimiento 1',
            'Técnico Depanaje',
            'Supervisor'
        ]

        # Verificar campos obligatorios
        required_fields = ['nombre_bd', 'turno_bd', 'cargo_bd', 'Letra_bd']
        for field in required_fields:
            if field not in df.columns:
                issues.append(f"Falta la columna {field}")
            else:
                null_count = df[field].isna().sum()
                if null_count > 0:
                    issues.append(f"{null_count} registros con {field} vacío")

        # Verificar duplicados
        duplicates = df[df.duplicated(subset=['nombre_bd'], keep=False)]
        if not duplicates.empty:
            issues.append(f"{len(duplicates)} registros duplicados encontrados")

        # Verificar valores válidos en turnos
        invalid_turnos = df[~df['turno_bd'].isin(VALID_TURNOS)]['turno_bd'].unique()
        if len(invalid_turnos) > 0:
            issues.append(f"Turnos inválidos encontrados: {', '.join(map(str, invalid_turnos))}")

        # Verificar valores válidos en cargos
        invalid_cargos = df[~df['cargo_bd'].isin(VALID_CARGOS)]['cargo_bd'].unique()
        if len(invalid_cargos) > 0:
            issues.append(f"Cargos inválidos encontrados: {', '.join(map(str, invalid_cargos))}")

        # Mostrar resultados
        if issues:
            result = "Se encontraron los siguientes problemas:\n\n" + "\n".join(f"• {issue}" for issue in issues)
            if messagebox.askyesno("Problemas de Integridad",
                                 f"{result}\n\n¿Desea intentar corregir automáticamente estos problemas?"):
                fix_data_integrity(df, issues)
        else:
            messagebox.showinfo("Verificación Completada",
                              "No se encontraron problemas de integridad en los datos")

    except Exception as e:
        messagebox.showerror("Error", f"Error al verificar integridad: {str(e)}")

def fix_data_integrity(df, issues):
    """Intenta corregir problemas de integridad encontrados"""
    try:
        VALID_TURNOS = ['A', 'B', 'C', 'prev_nocturno']
        VALID_CARGOS = [
            'Técnico Mantenimiento 2',
            'Técnico Mantenimiento 1',
            'Técnico Depanaje',
            'Supervisor'
        ]

        # Crear backup antes de correcciones
        backup_path = f"{excel_file_path}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        df.to_excel(backup_path, index=False)

        fixed_issues = []

        # Corregir campos vacíos
        for field in ['nombre_bd', 'turno_bd', 'cargo_bd', 'Letra_bd']:
            if df[field].isna().any():
                df[field] = df[field].fillna('')
                fixed_issues.append(f"Campos vacíos en {field} corregidos")

        # Corregir duplicados
        if len(df[df.duplicated(subset=['nombre_bd'], keep=False)]) > 0:
            df = df.drop_duplicates(subset=['nombre_bd'], keep='first')
            fixed_issues.append("Registros duplicados eliminados")

        # Corregir turnos inválidos
        mask = ~df['turno_bd'].isin(VALID_TURNOS)
        if mask.any():
            df.loc[mask, 'turno_bd'] = 'A'
            fixed_issues.append("Turnos inválidos corregidos")

        # Corregir cargos inválidos
        mask = ~df['cargo_bd'].isin(VALID_CARGOS)
        if mask.any():
            df.loc[mask, 'cargo_bd'] = 'Técnico Mantenimiento 2'
            fixed_issues.append("Cargos inválidos corregidos")

        # Guardar cambios
        df.to_excel(excel_file_path, index=False)

        messagebox.showinfo(
            "Correcciones Aplicadas",
            "Se realizaron las siguientes correcciones:\n\n" +
            "\n".join(f"• {fix}" for fix in fixed_issues) +
            f"\n\nSe creó un respaldo en: {backup_path}"
        )

    except Exception as e:
        messagebox.showerror("Error", f"Error al corregir integridad: {str(e)}")

def update_db_stats(stats_text):
    """Actualiza las estadísticas de la base de datos"""
    try:
        df = pd.read_excel(excel_file_path)
       
        stats_text.delete(1.0, tk.END)
        stats_text.insert(tk.END, "=== Estadísticas de la Base de Datos ===\n\n")
       
        # Estadísticas generales
        stats_text.insert(tk.END, f"Total de registros: {len(df)}\n")
        stats_text.insert(tk.END, f"Técnicos activos: {df['nombre_bd'].notna().sum()}\n")
       
        # Distribución por turno
        stats_text.insert(tk.END, "\nDistribución por turno:\n")
        turno_dist = df['turno_bd'].value_counts()
        for turno, count in turno_dist.items():
            stats_text.insert(tk.END, f"• {turno}: {count}\n")
       
        # Distribución por cargo
        stats_text.insert(tk.END, "\nDistribución por cargo:\n")
        cargo_dist = df['cargo_bd'].value_counts()
        for cargo, count in cargo_dist.items():
            stats_text.insert(tk.END, f"• {cargo}: {count}\n")
       
        # Última actualización
        stats_text.insert(tk.END, f"\nÚltima actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
       
        stats_text.config(state='disabled')

    except Exception as e:
        stats_text.delete(1.0, tk.END)
        stats_text.insert(tk.END, f"Error al cargar estadísticas: {str(e)}")

   


def backup_database():
    """Realiza una copia de seguridad del archivo Excel"""
    try:
        # Obtener la ubicación para guardar el respaldo
        backup_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar Respaldo de Base de Datos"
        )
       
        if backup_path:
            # Crear copia del archivo
            shutil.copy2(excel_file_path, backup_path)
           
            # Registrar metadatos del respaldo
            backup_info = {
                'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'usuario': 'admin',
                'archivo_original': excel_file_path,
                'archivo_respaldo': backup_path
            }
           
            messagebox.showinfo(
                "Éxito",
                f"Respaldo creado exitosamente en:\n{backup_path}"
            )
           
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al crear respaldo: {str(e)}"
        )

def restore_database():
    """Restaura la base de datos desde un respaldo"""
    try:
        # Solicitar archivo de respaldo
        backup_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Seleccionar Archivo de Respaldo"
        )
       
        if backup_path:
            # Confirmar restauración
            if messagebox.askyesno(
                "Confirmar Restauración",
                "Esta acción reemplazará la base de datos actual. ¿Desea continuar?"
            ):
                # Crear respaldo temporal antes de restaurar
                temp_backup = excel_file_path + '.temp'
                shutil.copy2(excel_file_path, temp_backup)
               
                try:
                    # Restaurar desde el respaldo
                    shutil.copy2(backup_path, excel_file_path)
                    messagebox.showinfo(
                        "Éxito",
                        "Base de datos restaurada exitosamente"
                    )
                except Exception as e:
                    # Si hay error, intentar recuperar desde el respaldo temporal
                    shutil.copy2(temp_backup, excel_file_path)
                    raise Exception(f"Error durante la restauración: {str(e)}")
                finally:
                    # Eliminar respaldo temporal
                    if os.path.exists(temp_backup):
                        os.remove(temp_backup)
               
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al restaurar base de datos: {str(e)}"
        )

def clean_old_records():
    """Limpia registros antiguos de la base de datos"""
    try:
        # Solicitar fecha límite
        dialog = CleanupDialog()
        if dialog.result:
            days_to_keep = dialog.days
           
            # Leer datos actuales
            df = pd.read_excel(excel_file_path)
            df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
           
            # Calcular fecha límite
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
           
            # Filtrar registros
            records_before = len(df)
            df = df[df['fecha_asignacion'] >= cutoff_date]
            records_removed = records_before - len(df)
           
            # Guardar cambios
            df.to_excel(excel_file_path, index=False)
           
            messagebox.showinfo(
                "Limpieza Completada",
                f"Se eliminaron {records_removed} registros anteriores a {cutoff_date.strftime('%d/%m/%Y')}"
            )
           
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al limpiar registros: {str(e)}"
        )

class CleanupDialog:
    """Diálogo para configurar la limpieza de registros"""
    def __init__(self):
        self.dialog = tk.Toplevel()
        self.dialog.title("Configurar Limpieza")
        self.dialog.grab_set()
       
        # Centrar ventana
        window_width = 300
        window_height = 150
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')
       
        self.result = False
        self.days = 0
       
        # Contenido
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
       
        ttk.Label(
            main_frame,
            text="Mantener registros de los últimos días:",
            font=('Segoe UI', 10)
        ).pack(pady=(0, 10))
       
        self.days_var = tk.StringVar(value="30")
        days_entry = ttk.Entry(
            main_frame,
            textvariable=self.days_var,
            width=10
        )
        days_entry.pack(pady=(0, 20))
       
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')
       
        ttk.Button(
            button_frame,
            text="Aceptar",
            command=self.on_accept,
            style='Primary.TButton'
        ).pack(side='left', expand=True, padx=5)
       
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=self.on_cancel
        ).pack(side='left', expand=True, padx=5)
       
        self.dialog.wait_window()
   
    def on_accept(self):
        try:
            self.days = int(self.days_var.get())
            if self.days <= 0:
                raise ValueError("El número de días debe ser positivo")
            self.result = True
            self.dialog.destroy()
        except ValueError:
            messagebox.showerror(
                "Error",
                "Por favor ingrese un número válido de días"
            )
   
    def on_cancel(self):
        self.dialog.destroy()




def generar_informe_ausencias(fecha_inicio, fecha_fin):
    """Genera un informe detallado de ausencias"""
    try:
        # Solicitar ubicación para guardar
        file_path = filedialog.asksaveasfilename(
            defaultextension='.pdf',
            filetypes=[("PDF files", "*.pdf")],
            title="Guardar Informe de Ausencias"
        )
       
        if not file_path:
            return

        # Leer y preparar datos
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
       
        fecha_inicio_dt = datetime.strptime(fecha_inicio.get_date(), '%d/%m/%Y')
        fecha_fin_dt = datetime.strptime(fecha_fin.get_date(), '%d/%m/%Y')

        # Filtrar datos por fecha
        mask = (
            (df['fecha_eliminado'].dt.date >= fecha_inicio_dt.date()) &
            (df['fecha_eliminado'].dt.date <= fecha_fin_dt.date()) &
            df['nombre_eliminado'].notna()
        )
        df_filtered = df[mask]

        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para generar el informe")
            return

        # Crear documento PDF
        doc = SimpleDocTemplate(
            file_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )

        # Contenido del informe
        elements = []
        styles = getSampleStyleSheet()
       
        # Título
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=24,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        elements.append(Paragraph("Informe de Ausencias", title_style))
        elements.append(Spacer(1, 12))

        # Período del informe
        period_text = f"Período: {fecha_inicio.get_date()} - {fecha_fin.get_date()}"
        elements.append(Paragraph(period_text, styles['Heading2']))
        elements.append(Spacer(1, 12))

        # Resumen estadístico
        elements.append(Paragraph("Resumen Estadístico", styles['Heading2']))
        elements.append(Spacer(1, 12))

        # Total de ausencias
        total_ausencias = len(df_filtered)
        elements.append(Paragraph(f"Total de ausencias: {total_ausencias}", styles['Normal']))
       
        # Ausencias por motivo
        motivos = df_filtered['motivo_eliminado'].value_counts()
        elements.append(Paragraph("Distribución por motivo:", styles['Normal']))
        for motivo, count in motivos.items():
            elements.append(Paragraph(f"• {motivo}: {count}", styles['Normal']))

        elements.append(Spacer(1, 12))

        # Tabla de ausencias
        elements.append(Paragraph("Detalle de Ausencias", styles['Heading2']))
        elements.append(Spacer(1, 12))

        # Crear tabla
        data = [['Nombre', 'Cargo', 'Fecha', 'Motivo', 'Jornada']]
        for _, row in df_filtered.iterrows():
            data.append([
                row['nombre_eliminado'],
                row['cargo_eliminado'],
                row['fecha_eliminado'].strftime('%d/%m/%Y'),
                row['motivo_eliminado'],
                row.get('jornada_eliminada', '')
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))

        elements.append(table)
       
        # Generar PDF
        doc.build(elements)
       
        messagebox.showinfo("Éxito", "Informe generado exitosamente")

    except ImportError:
        messagebox.showerror(
            "Error",
            "La biblioteca 'reportlab' no está instalada.\n"
            "Por favor, instale la biblioteca usando:\n"
            "pip install reportlab"
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error al generar informe: {str(e)}")

def actualizar_graficos_tendencias(tab):
    """Actualiza los gráficos en la pestaña de análisis y tendencias"""
    try:
        # Limpiar gráficos anteriores
        for widget in tab.winfo_children():
            widget.destroy()

        # Leer datos
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])

        # Crear figura principal
        fig = plt.Figure(figsize=(12, 8))
       
        # Gráfico 1: Tendencia temporal
        ax1 = fig.add_subplot(221)
        df_grouped = df.groupby(df['fecha_eliminado'].dt.strftime('%Y-%m')).size()
        df_grouped.plot(kind='line', ax=ax1, marker='o')
        ax1.set_title('Tendencia de Ausencias por Mes')
        ax1.set_xlabel('Mes')
        ax1.set_ylabel('Cantidad de Ausencias')
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)

        # Gráfico 2: Distribución por motivo
        ax2 = fig.add_subplot(222)
        df['motivo_eliminado'].value_counts().plot(kind='pie', ax=ax2, autopct='%1.1f%%')
        ax2.set_title('Distribución por Motivo')

        # Gráfico 3: Ausencias por jornada
        ax3 = fig.add_subplot(223)
        df['jornada_eliminada'].value_counts().plot(kind='bar', ax=ax3)
        ax3.set_title('Ausencias por Jornada')
        ax3.set_xlabel('Jornada')
        ax3.set_ylabel('Cantidad')
        plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45)

        # Gráfico 4: Heatmap de ausencias por día de la semana
        ax4 = fig.add_subplot(224)
        df['dia_semana'] = df['fecha_eliminado'].dt.day_name()
        df['mes'] = df['fecha_eliminado'].dt.month_name()
        pivot = pd.crosstab(df['dia_semana'], df['mes'])
        sns.heatmap(pivot, ax=ax4, cmap='YlOrRd', annot=True, fmt='d')
        ax4.set_title('Heatmap de Ausencias')
        plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45)

        # Ajustar layout
        fig.tight_layout()

        # Crear canvas y agregar a la pestaña
        canvas = FigureCanvasTkAgg(fig, master=tab)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        # Agregar barra de herramientas de navegación
        toolbar = NavigationToolbar2Tk(canvas, tab)
        toolbar.update()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    except Exception as e:
        messagebox.showerror("Error", f"Error al actualizar gráficos: {str(e)}")


def previsualizar_reporte(tipo_reporte):
    """Muestra una previsualización del reporte seleccionado"""
    preview_window = tk.Toplevel()
    preview_window.title(f"Previsualización - {tipo_reporte.title()}")
    preview_window.grab_set()
   
    # Configurar tamaño y posición
    window_width = 800
    window_height = 600
    screen_width = preview_window.winfo_screenwidth()
    screen_height = preview_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    preview_window.geometry(f'{window_width}x{window_height}+{x}+{y}')

    # Frame principal
    main_frame = ttk.Frame(preview_window, padding="10")
    main_frame.pack(fill='both', expand=True)

    # Título del reporte
    ttk.Label(
        main_frame,
        text=f"Reporte de Ausencias - {tipo_reporte.title()}",
        font=('Segoe UI', 14, 'bold')
    ).pack(pady=(0, 20))

    # Contenedor con scroll
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Contenido del reporte según el tipo
    if tipo_reporte == "mensual":
        crear_preview_mensual(scrollable_frame)
    elif tipo_reporte == "departamento":
        crear_preview_departamento(scrollable_frame)
    elif tipo_reporte == "supervisor":
        crear_preview_supervisor(scrollable_frame)
    elif tipo_reporte == "tendencias":
        crear_preview_tendencias(scrollable_frame)

    # Botones de acción
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill='x', pady=10)

    ttk.Button(
        button_frame,
        text="Generar PDF",
        command=lambda: generar_reporte(tipo_reporte, True),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
        button_frame,
        text="Cerrar",
        command=preview_window.destroy
    ).pack(side='left', padx=5)

    # Empaquetar canvas y scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

def crear_preview_mensual(parent):
    """Crea una previsualización del reporte mensual"""
    # Información general
    info_frame = ttk.LabelFrame(parent, text="Información General", padding=10)
    info_frame.pack(fill='x', pady=5, padx=10)

    current_month = datetime.now().strftime("%B %Y")
    ttk.Label(
        info_frame,
        text=f"Período: {current_month}",
        font=('Segoe UI', 10, 'bold')
    ).pack(anchor='w')

    # Resumen ejecutivo
    summary_frame = ttk.LabelFrame(parent, text="Resumen Ejecutivo", padding=10)
    summary_frame.pack(fill='x', pady=5, padx=10)

    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
       
        # Filtrar datos del mes actual
        current_month = datetime.now().month
        current_year = datetime.now().year
        df_month = df[
            (df['fecha_eliminado'].dt.month == current_month) &
            (df['fecha_eliminado'].dt.year == current_year)
        ]

        # Calcular estadísticas
        total_ausencias = len(df_month)
        ausencias_por_tipo = df_month['motivo_eliminado'].value_counts()

        ttk.Label(
            summary_frame,
            text=f"Total de ausencias: {total_ausencias}",
            font=('Segoe UI', 10)
        ).pack(anchor='w')

        # Mostrar distribución por tipo
        if not ausencias_por_tipo.empty:
            ttk.Label(
                summary_frame,
                text="\nDistribución por tipo:",
                font=('Segoe UI', 10, 'bold')
            ).pack(anchor='w', pady=(10,5))

            for tipo, cantidad in ausencias_por_tipo.items():
                ttk.Label(
                    summary_frame,
                    text=f"• {tipo}: {cantidad}",
                    font=('Segoe UI', 10)
                ).pack(anchor='w')

        # Gráfico de distribución
        fig = Figure(figsize=(8, 4))
        ax = fig.add_subplot(111)
        ausencias_por_tipo.plot(kind='bar', ax=ax)
        ax.set_title('Distribución de Ausencias')
        ax.set_ylabel('Cantidad')
        fig.autofmt_xdate()  # Rotar etiquetas del eje x

        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10, padx=10)

    except Exception as e:
        ttk.Label(
            summary_frame,
            text=f"Error al cargar datos: {str(e)}",
            font=('Segoe UI', 10)
        ).pack(anchor='w')

def crear_preview_departamento(parent):
    """Crea una previsualización del reporte por departamento"""
    ttk.Label(
        parent,
        text="Análisis por Departamento",
        font=('Segoe UI', 12, 'bold')
    ).pack(pady=10)

    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
       
        # Agrupar por cargo
        ausencias_por_cargo = df.groupby('cargo_eliminado').size()

        # Mostrar estadísticas por cargo
        if not ausencias_por_cargo.empty:
            for cargo, cantidad in ausencias_por_cargo.items():
                frame = ttk.Frame(parent)
                frame.pack(fill='x', pady=5, padx=10)
               
                ttk.Label(
                    frame,
                    text=f"{cargo}:",
                    font=('Segoe UI', 10, 'bold')
                ).pack(side='left')
               
                ttk.Label(
                    frame,
                    text=f" {cantidad} ausencias",
                    font=('Segoe UI', 10)
                ).pack(side='left')

        # Gráfico de distribución por cargo
        fig = Figure(figsize=(8, 4))
        ax = fig.add_subplot(111)
        ausencias_por_cargo.plot(kind='pie', ax=ax)
        ax.set_title('Distribución de Ausencias por Cargo')

        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10, padx=10)

    except Exception as e:
        ttk.Label(
            parent,
            text=f"Error al cargar datos: {str(e)}",
            font=('Segoe UI', 10)
        ).pack(pady=10)

def crear_preview_supervisor(parent):
    """Crea una previsualización del reporte para supervisores"""
    ttk.Label(
        parent,
        text="Resumen para Supervisores",
        font=('Segoe UI', 12, 'bold')
    ).pack(pady=10)

    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])

        # Crear secciones de información crítica
        secciones = [
            ("Ausencias Críticas", df[df['motivo_eliminado'].str.contains('Licencia', na=False)]),
            ("Cambios de Turno", df[df['motivo_eliminado'].str.contains('Cambio', na=False)]),
            ("Personal en Capacitación", df[df['motivo_eliminado'].str.contains('Capacitación', na=False)])
        ]

        for titulo, data in secciones:
            section_frame = ttk.LabelFrame(parent, text=titulo, padding=10)
            section_frame.pack(fill='x', pady=5, padx=10)

            if not data.empty:
                for _, row in data.iterrows():
                    info_text = f"{row['nombre_eliminado']} - {row['cargo_eliminado']}"
                    if pd.notna(row['fecha_eliminado']):
                        info_text += f" - {row['fecha_eliminado'].strftime('%d/%m/%Y')}"
                   
                    ttk.Label(
                        section_frame,
                        text=info_text,
                        font=('Segoe UI', 10)
                    ).pack(anchor='w')
            else:
                ttk.Label(
                    section_frame,
                    text="No hay registros",
                    font=('Segoe UI', 10, 'italic')
                ).pack(anchor='w')

    except Exception as e:
        ttk.Label(
            parent,
            text=f"Error al cargar datos: {str(e)}",
            font=('Segoe UI', 10)
        ).pack(pady=10)

def crear_preview_tendencias(parent):
    """Crea una previsualización del reporte de tendencias"""
    ttk.Label(
        parent,
        text="Análisis de Tendencias",
        font=('Segoe UI', 12, 'bold')
    ).pack(pady=10)

    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])

        # Tendencia mensual
        df['mes'] = df['fecha_eliminado'].dt.strftime('%Y-%m')
        tendencia_mensual = df.groupby('mes').size()

        fig = Figure(figsize=(8, 6))
       
        # Gráfico de tendencia
        ax1 = fig.add_subplot(211)
        tendencia_mensual.plot(kind='line', ax=ax1, marker='o')
        ax1.set_title('Tendencia Mensual de Ausencias')
        ax1.set_ylabel('Cantidad')
        fig.autofmt_xdate()

        # Gráfico de distribución por motivo
        ax2 = fig.add_subplot(212)
        df['motivo_eliminado'].value_counts().plot(kind='bar', ax=ax2)
        ax2.set_title('Distribución por Motivo')
        ax2.set_ylabel('Cantidad')
        fig.autofmt_xdate()

        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10, padx=10)

        # Añadir estadísticas descriptivas
        stats_frame = ttk.LabelFrame(parent, text="Estadísticas", padding=10)
        stats_frame.pack(fill='x', pady=5, padx=10)

        promedio = tendencia_mensual.mean()
        maximo = tendencia_mensual.max()
        tendencia = "Ascendente" if tendencia_mensual.iloc[-1] > tendencia_mensual.iloc[0] else "Descendente"

        ttk.Label(
            stats_frame,
            text=f"Promedio mensual: {promedio:.1f}",
            font=('Segoe UI', 10)
        ).pack(anchor='w')
       
        ttk.Label(
            stats_frame,
            text=f"Máximo mensual: {maximo}",
            font=('Segoe UI', 10)
        ).pack(anchor='w')
       
        ttk.Label(
            stats_frame,
            text=f"Tendencia: {tendencia}",
            font=('Segoe UI', 10)
        ).pack(anchor='w')

    except Exception as e:
        ttk.Label(
            parent,
            text=f"Error al cargar datos: {str(e)}",
            font=('Segoe UI', 10)
        ).pack(pady=10)

def generar_reporte(tipo_reporte, from_preview=False):
    """Genera y guarda el reporte seleccionado"""
    try:
        # Solicitar ubicación para guardar
        file_path = filedialog.asksaveasfilename(
            defaultextension='.pdf',
            filetypes=[("PDF files", "*.pdf")],
            title="Guardar Reporte"
        )
       
        if not file_path:
            return

        messagebox.showinfo(
            "Información",
            "La generación de PDF no está implementada actualmente.\n"
            "Se puede implementar usando la biblioteca 'reportlab' para generar PDFs."
        )

    except Exception as e:
        messagebox.showerror("Error", f"Error al generar reporte: {str(e)}")



def setup_ausencias_tab(tab):
    """Configura la pestaña de visualización de ausencias con funcionalidades mejoradas"""
    # Notebook interno para múltiples vistas
    ausencias_notebook = ttk.Notebook(tab)
    ausencias_notebook.pack(fill='both', expand=True, padx=5, pady=5)

    # === Pestaña principal de listado ===
    lista_tab = ttk.Frame(ausencias_notebook)
    ausencias_notebook.add(lista_tab, text="Listado de Ausencias")
   
    # Configuración del grid principal
    lista_tab.grid_columnconfigure(0, weight=1)
    lista_tab.grid_rowconfigure(1, weight=1)

    # === Panel Superior de Filtros Mejorado ===
    filter_frame = ttk.LabelFrame(lista_tab, text="Filtros y Acciones", padding="10")
    filter_frame.grid(row=0, column=0, sticky='ew', padx=10, pady=5)
   
    # Frame de filtros con más opciones
    filters_container = ttk.Frame(filter_frame)
    filters_container.pack(fill='x', expand=True)
   
    # Filtros de fecha
    date_frame = ttk.Frame(filters_container)
    date_frame.pack(fill='x', pady=5)
   
    ttk.Label(date_frame, text="Desde:", font=('Segoe UI', 10)).pack(side='left', padx=5)
    fecha_inicio = Calendar(date_frame,
                          selectmode='day',
                          date_pattern='dd/mm/yyyy',
                          width=12,
                          height=2)
    fecha_inicio.pack(side='left', padx=5)

    ttk.Label(date_frame, text="Hasta:", font=('Segoe UI', 10)).pack(side='left', padx=5)
    fecha_fin = Calendar(date_frame,
                        selectmode='day',
                        date_pattern='dd/mm/yyyy',
                        width=12,
                        height=2)
    fecha_fin.pack(side='left', padx=5)

    # Filtros adicionales
    filter_options = ttk.Frame(filters_container)
    filter_options.pack(fill='x', pady=5)

    # Variables para filtros
    motivo_var = tk.StringVar()
    jornada_var = tk.StringVar()

    # Comboboxes para filtros
    ttk.Label(filter_options, text="Motivo:").pack(side='left', padx=5)
    motivo_combo = ttk.Combobox(filter_options,
                               textvariable=motivo_var,
                               values=['Todos', 'Licencia Médica', 'Cambio de turno', 'Capacitación', 'Feriado Legal'],
                               width=20)
    motivo_combo.pack(side='left', padx=5)
    motivo_combo.set('Todos')

    ttk.Label(filter_options, text="Jornada:").pack(side='left', padx=5)
    jornada_combo = ttk.Combobox(filter_options,
                                textvariable=jornada_var,
                                values=['Todas', 'Mañana', 'Tarde', 'Noche'],
                                width=15)
    jornada_combo.pack(side='left', padx=5)
    jornada_combo.set('Todas')

    # Botones de acción
    button_frame = ttk.Frame(filters_container)
    button_frame.pack(fill='x', pady=10)

    ttk.Button(
        button_frame,
        text="Aplicar Filtros",
        command=lambda: actualizar_ausencias(tree, fecha_inicio, fecha_fin, counter_vars, motivo_var, jornada_var),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
        button_frame,
        text="Exportar Excel",
        command=lambda: exportar_ausencias(fecha_inicio, fecha_fin, motivo_var, jornada_var),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
    button_frame,
    text="Generar Informe",
    command=lambda: generar_informe_ausencias(fecha_inicio, fecha_fin),
    style='Primary.TButton'
    ).pack(side='left', padx=5)

    # === Panel de Estadísticas Mejorado ===
    stats_frame = ttk.LabelFrame(lista_tab, text="Dashboard", padding="10")
    stats_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
   
    # Split frame para estadísticas y tabla
    stats_paned = ttk.PanedWindow(stats_frame, orient=tk.HORIZONTAL)
    stats_paned.pack(fill='both', expand=True)

    # Panel izquierdo para estadísticas
    left_stats = ttk.Frame(stats_paned)
    stats_paned.add(left_stats, weight=1)

    # Contadores con mejor visualización
    counter_vars = create_enhanced_counters(left_stats)

    # Gráfico de tendencias
    fig = Figure(figsize=(6, 4), dpi=100)
    canvas = FigureCanvasTkAgg(fig, master=left_stats)
    canvas.get_tk_widget().pack(fill='both', expand=True, pady=10)

    # Panel derecho para la tabla
    right_panel = ttk.Frame(stats_paned)
    stats_paned.add(right_panel, weight=2)

    # Tabla mejorada
    tree = create_enhanced_treeview(right_panel)

    # === Pestaña de Análisis ===
    analisis_tab = ttk.Frame(ausencias_notebook)
    ausencias_notebook.add(analisis_tab, text="Análisis y Tendencias")
    setup_analisis_tab(analisis_tab)

    # === Pestaña de Reportes ===
    reportes_tab = ttk.Frame(ausencias_notebook)
    ausencias_notebook.add(reportes_tab, text="Reportes")
    setup_reportes_tab(reportes_tab)

    # Inicializar con datos actuales
    actualizar_ausencias(tree, fecha_inicio, fecha_fin, counter_vars, motivo_var, jornada_var)
   
    return tab

def create_enhanced_counters(parent):
    """Crea contadores mejorados con visualización más profesional"""
    counter_frame = ttk.Frame(parent)
    counter_frame.pack(fill='x', pady=5)

    counter_vars = {
        'total': {'var': tk.StringVar(value="0"), 'label': "Total Ausencias"},
        'licencia': {'var': tk.StringVar(value="0"), 'label': "Licencias Médicas"},
        'cambio': {'var': tk.StringVar(value="0"), 'label': "Cambios de Turno"},
        'capacitacion': {'var': tk.StringVar(value="0"), 'label': "Capacitaciones"},
        'feriado': {'var': tk.StringVar(value="0"), 'label': "Feriados Legales"}
    }

    for key, data in counter_vars.items():
        frame = ttk.Frame(counter_frame, style='Card.TFrame')
        frame.pack(fill='x', pady=2, padx=5)
       
        ttk.Label(
            frame,
            text=data['label'],
            font=('Segoe UI', 9),
            foreground='#666666'
        ).pack(anchor='w', padx=5, pady=(5,0))
       
        ttk.Label(
            frame,
            textvariable=data['var'],
            font=('Segoe UI', 12, 'bold')
        ).pack(anchor='w', padx=5, pady=(0,5))

    return counter_vars

def create_enhanced_treeview(parent):
    """Crea una tabla mejorada con más funcionalidades"""
    # Frame contenedor
    tree_frame = ttk.Frame(parent)
    tree_frame.pack(fill='both', expand=True)

    # Crear Treeview
    tree = ttk.Treeview(
        tree_frame,
        columns=('nombre', 'cargo', 'fecha', 'motivo', 'jornada'),
        show='headings',
        height=15
    )

    # Configurar columnas con mejor formato
    columns_config = [
        ('nombre', 'Nombre', 150),
        ('cargo', 'Cargo', 120),
        ('fecha', 'Fecha', 100),
        ('motivo', 'Motivo', 200),
        ('jornada', 'Jornada', 100)
    ]

    for col, heading, width in columns_config:
        tree.heading(col, text=heading, command=lambda c=col: sort_treeview(tree, c))
        tree.column(col, width=width, minwidth=width)

    # Scrollbars
    y_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

    # Grid
    tree.grid(row=0, column=0, sticky='nsew')
    y_scrollbar.grid(row=0, column=1, sticky='ns')
    x_scrollbar.grid(row=1, column=0, sticky='ew')

    # Configurar grid
    tree_frame.grid_columnconfigure(0, weight=1)
    tree_frame.grid_rowconfigure(0, weight=1)

    # Estilos y bindings
    tree.tag_configure('oddrow', background='#f5f5f5')
    tree.tag_configure('evenrow', background='white')
    tree.bind('<Double-1>', lambda e: show_detail_window(tree))

    return tree

def setup_analisis_tab(tab):
    """Configura la pestaña de análisis con gráficos interactivos"""
    try:
        # Frame principal usando pack
        main_frame = ttk.Frame(tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Frame para los gráficos
        charts_frame = ttk.Frame(main_frame)
        charts_frame.pack(fill='both', expand=True)

        # Leer datos
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])

        # Crear figura con subplots
        fig = Figure(figsize=(12, 8))
       
        # Gráfico 1: Tendencia temporal
        ax1 = fig.add_subplot(221)
        df_grouped = df.groupby(df['fecha_eliminado'].dt.strftime('%Y-%m')).size()
        if not df_grouped.empty:
            df_grouped.plot(kind='line', ax=ax1, marker='o')
            ax1.set_title('Tendencia de Ausencias por Mes')
            ax1.set_xlabel('Mes')
            ax1.set_ylabel('Cantidad de Ausencias')
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)

        # Gráfico 2: Distribución por motivo
        ax2 = fig.add_subplot(222)
        if 'motivo_eliminado' in df.columns:
            motivos = df['motivo_eliminado'].value_counts()
            if not motivos.empty:
                motivos.plot(kind='pie', ax=ax2, autopct='%1.1f%%')
                ax2.set_title('Distribución por Motivo')

        # Gráfico 3: Ausencias por jornada
        ax3 = fig.add_subplot(223)
        if 'jornada_eliminada' in df.columns:
            jornadas = df['jornada_eliminada'].value_counts()
            if not jornadas.empty:
                jornadas.plot(kind='bar', ax=ax3)
                ax3.set_title('Ausencias por Jornada')
                ax3.set_xlabel('Jornada')
                ax3.set_ylabel('Cantidad')
                plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45)

        # Gráfico 4: Heatmap de ausencias por día de la semana
        ax4 = fig.add_subplot(224)
        if not df.empty:
            df['dia_semana'] = df['fecha_eliminado'].dt.day_name()
            df['mes'] = df['fecha_eliminado'].dt.month_name()
            pivot = pd.crosstab(df['dia_semana'], df['mes'])
            if not pivot.empty:
                sns.heatmap(pivot, ax=ax4, cmap='YlOrRd', annot=True, fmt='d')
                ax4.set_title('Heatmap de Ausencias')
                plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45)

        # Ajustar layout
        fig.tight_layout()

        # Crear canvas y agregar a la pestaña
        canvas = FigureCanvasTkAgg(fig, master=charts_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Agregar barra de herramientas de navegación
        toolbar = NavigationToolbar2Tk(canvas, charts_frame)
        toolbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Frame para estadísticas
        stats_frame = ttk.LabelFrame(main_frame, text="Estadísticas Generales", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=10)

        # Calcular y mostrar estadísticas
        if not df.empty:
            total_ausencias = len(df)
            promedio_mensual = df_grouped.mean() if not df_grouped.empty else 0
            max_ausencias = df_grouped.max() if not df_grouped.empty else 0
           
            stats_text = (
                f"Total de ausencias: {total_ausencias}\n"
                f"Promedio mensual: {promedio_mensual:.1f}\n"
                f"Máximo mensual: {max_ausencias}"
            )
           
            ttk.Label(
                stats_frame,
                text=stats_text,
                font=('Segoe UI', 10),
                justify='left'
            ).pack(anchor='w', padx=5, pady=5)

    except Exception as e:
        print(f"Error en setup_analisis_tab: {str(e)}")
        # Mostrar mensaje de error en la interfaz
        error_label = ttk.Label(
            tab,
            text=f"Error al cargar los análisis: {str(e)}",
            font=('Segoe UI', 10),
            foreground='red'
        )
        error_label.pack(pady=20)

def setup_reportes_tab(tab):
    """Configura la pestaña de reportes"""
    # Frame principal
    main_frame = ttk.Frame(tab)
    main_frame.pack(fill='both', expand=True, padx=10, pady=10)

    # Opciones de reporte
    options_frame = ttk.LabelFrame(main_frame, text="Opciones de Reporte", padding=10)
    options_frame.pack(fill='x', pady=(0, 10))

    # Variables para opciones
    report_type = tk.StringVar(value="mensual")
    include_graphs = tk.BooleanVar(value=True)
    include_summary = tk.BooleanVar(value=True)

    # Tipos de reporte
    ttk.Label(
        options_frame,
        text="Tipo de Reporte:",
        font=('Segoe UI', 10, 'bold')
    ).pack(anchor='w', pady=(0, 5))

    report_types = [
        ("Reporte Mensual", "mensual"),
        ("Análisis por Departamento", "departamento"),
        ("Resumen para Supervisores", "supervisor"),
        ("Análisis de Tendencias", "tendencias")
    ]

    for text, value in report_types:
        ttk.Radiobutton(
            options_frame,
            text=text,
            value=value,
            variable=report_type
        ).pack(anchor='w', pady=2)

    # Opciones adicionales
    ttk.Checkbutton(
        options_frame,
        text="Incluir gráficos",
        variable=include_graphs
    ).pack(anchor='w', pady=2)

    ttk.Checkbutton(
        options_frame,
        text="Incluir resumen ejecutivo",
        variable=include_summary
    ).pack(anchor='w', pady=2)

    # Botones
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill='x', pady=10)

    ttk.Button(
        button_frame,
        text="Generar Reporte",
        command=lambda: generar_reporte(report_type.get()),
        style='Primary.TButton'
    ).pack(side='left', padx=5)

    ttk.Button(
        button_frame,
        text="Previsualizar",
        command=lambda: previsualizar_reporte(report_type.get())
    ).pack(side='left', padx=5)
def sort_treeview(tree, col):
    """Implementa ordenamiento en la tabla"""
    data = [(tree.set(item, col), item) for item in tree.get_children('')]
    data.sort()
    for index, (val, item) in enumerate(data):
        tree.move(item, '', index)
        # Actualizar colores alternados
        tree.item(item, tags=('evenrow' if index % 2 == 0 else 'oddrow',))

def show_detail_window(tree):
    """Muestra una ventana con detalles del registro seleccionado"""
    selection = tree.selection()
    if not selection:
        return
   
    item = tree.item(selection[0])
    values = item['values']
   
    detail_window = tk.Toplevel()
    detail_window.title("Detalles de Ausencia")
    detail_window.grab_set()
   
    # Centrar ventana
    window_width = 400
    window_height = 300
    screen_width = detail_window.winfo_screenwidth()
    screen_height = detail_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    detail_window


def actualizar_ausencias(tree, fecha_inicio, fecha_fin, counter_vars=None, motivo_filter=None, jornada_filter=None):
    """Actualiza la tabla de ausencias con los filtros aplicados"""
    try:
        # Limpiar tabla
        for item in tree.get_children():
            tree.delete(item)

        # Leer datos
        df = pd.read_excel(excel_file_path)
       
        # Convertir fechas
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
        fecha_inicio_dt = datetime.strptime(fecha_inicio.get_date(), '%d/%m/%Y')
        fecha_fin_dt = datetime.strptime(fecha_fin.get_date(), '%d/%m/%Y')

        # Filtro base por fecha
        mask = (
            (df['fecha_eliminado'].dt.date >= fecha_inicio_dt.date()) &
            (df['fecha_eliminado'].dt.date <= fecha_fin_dt.date()) &
            df['nombre_eliminado'].notna()
        )

        # Aplicar filtros adicionales si están especificados
        if motivo_filter and motivo_filter.get() != 'Todos':
            mask &= df['motivo_eliminado'].str.contains(motivo_filter.get(), na=False)
       
        if jornada_filter and jornada_filter.get() != 'Todas':
            mask &= df['jornada_eliminada'] == jornada_filter.get()

        df_filtered = df[mask]

        # Actualizar contadores si existen
        if counter_vars:
            total = len(df_filtered)
            licencias = len(df_filtered[df_filtered['motivo_eliminado'].str.contains('Licencia', na=False)])
            cambios = len(df_filtered[df_filtered['motivo_eliminado'].str.contains('Cambio', na=False)])
            capacitacion = len(df_filtered[df_filtered['motivo_eliminado'].str.contains('Capacitación', na=False)])
            feriado = len(df_filtered[df_filtered['motivo_eliminado'].str.contains('Feriado', na=False)])

            # Actualizar las variables del counter_vars actualizado
            counter_vars['total']['var'].set(str(total))
            counter_vars['licencia']['var'].set(str(licencias))
            counter_vars['cambio']['var'].set(str(cambios))
            counter_vars['capacitacion']['var'].set(str(capacitacion))
            counter_vars['feriado']['var'].set(str(feriado))

        # Llenar tabla
        for i, row in df_filtered.iterrows():
            values = (
                row['nombre_eliminado'],
                row['cargo_eliminado'],
                row['fecha_eliminado'].strftime('%d/%m/%Y'),
                row['motivo_eliminado'],
                row.get('jornada_eliminada', '')  # Usar get para manejar casos donde la columna no exista
            )
            tree.insert('', 'end', values=values, tags=('evenrow' if i % 2 == 0 else 'oddrow',))

        # Actualizar la visualización de la tabla
        tree.update_idletasks()

    except Exception as e:
        print(f"Error detallado: {str(e)}")  # Para debugging
        messagebox.showerror("Error", f"Error al actualizar ausencias: {str(e)}")

def exportar_ausencias(fecha_inicio, fecha_fin, motivo_filter=None, jornada_filter=None):
    """Exporta las ausencias filtradas a Excel"""
    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
       
        fecha_inicio_dt = datetime.strptime(fecha_inicio.get_date(), '%d/%m/%Y')
        fecha_fin_dt = datetime.strptime(fecha_fin.get_date(), '%d/%m/%Y')

        # Aplicar filtros
        mask = (
            (df['fecha_eliminado'].dt.date >= fecha_inicio_dt.date()) &
            (df['fecha_eliminado'].dt.date <= fecha_fin_dt.date()) &
            df['nombre_eliminado'].notna()
        )

        # Filtros adicionales
        if motivo_filter and motivo_filter.get() != 'Todos':
            mask &= df['motivo_eliminado'].str.contains(motivo_filter.get(), na=False)
       
        if jornada_filter and jornada_filter.get() != 'Todas':
            mask &= df['jornada_eliminada'] == jornada_filter.get()

        df_filtered = df[mask]

        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Exportar Ausencias"
        )

        if file_path:
            # Crear un DataFrame con solo las columnas relevantes
            df_export = df_filtered[[
                'nombre_eliminado',
                'cargo_eliminado',
                'fecha_eliminado',
                'motivo_eliminado',
                'jornada_eliminada'
            ]].copy()

            # Renombrar columnas para mejor presentación
            df_export.columns = ['Nombre', 'Cargo', 'Fecha', 'Motivo', 'Jornada']

            # Dar formato a la fecha
            df_export['Fecha'] = df_export['Fecha'].dt.strftime('%d/%m/%Y')

            # Exportar a Excel con formato
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Ausencias')
               
                # Obtener la hoja de trabajo
                worksheet = writer.sheets['Ausencias']
               
                # Ajustar anchos de columna
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            messagebox.showinfo("Éxito", "Datos exportados exitosamente")

    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar ausencias: {str(e)}")

def exportar_ausencias(fecha_inicio, fecha_fin):
    """Exporta las ausencias filtradas a Excel"""
    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_eliminado'] = pd.to_datetime(df['fecha_eliminado'])
       
        fecha_inicio_dt = datetime.strptime(fecha_inicio.get_date(), '%d/%m/%Y')
        fecha_fin_dt = datetime.strptime(fecha_fin.get_date(), '%d/%m/%Y')

        mask = (
            (df['fecha_eliminado'].dt.date >= fecha_inicio_dt.date()) &
            (df['fecha_eliminado'].dt.date <= fecha_fin_dt.date()) &
            df['nombre_eliminado'].notna()
        )
        df_filtered = df[mask]

        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Exportar Ausencias"
        )

        if file_path:
            # Crear un DataFrame con solo las columnas relevantes
            df_export = df_filtered[[
                'nombre_eliminado',
                'cargo_eliminado',
                'fecha_eliminado',
                'motivo_eliminado',
                'jornada_eliminada'
            ]].copy()

            # Renombrar columnas para mejor presentación
            df_export.columns = ['Nombre', 'Cargo', 'Fecha', 'Motivo', 'Jornada']

            # Dar formato a la fecha
            df_export['Fecha'] = df_export['Fecha'].dt.strftime('%d/%m/%Y')

            # Exportar a Excel con formato
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Ausencias')
               
                # Obtener la hoja de trabajo
                worksheet = writer.sheets['Ausencias']
               
                # Ajustar anchos de columna
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            messagebox.showinfo("Éxito", "Datos exportados exitosamente")

    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar ausencias: {str(e)}")



def ensure_temp_dir():
    """Asegura que el directorio temporal existe"""
    try:
        if not os.path.exists(TEMP_DIR):
            os.makedirs(TEMP_DIR, exist_ok=True)
            print(f"Directorio temporal creado en: {TEMP_DIR}")
        return True
    except Exception as e:
        print(f"Error al crear directorio temporal: {str(e)}")
        return False

def configure_styles():
    style = ttk.Style()
    style.theme_use('clam')
   
    style.configure('.', background='#f0f0f0')
    style.configure('TLabel', padding=5, font=('Segoe UI', 10))
    style.configure('TButton', padding=10, font=('Segoe UI', 10))
    style.configure('TEntry', padding=5)
    style.configure('Treeview', rowheight=25, font=('Segoe UI', 9))
    style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
   
    style.configure('Primary.TButton', background='#007bff', foreground='white')
    style.configure('Danger.TButton', background='#dc3545', foreground='white')
    style.configure('Tooltip.TFrame',
                   background='#333333',
                   relief='solid',
                   borderwidth=1)


     
 #ruta antigua con excel compartido en máquina local excel_file_path = r'\\MET-ARR_L0279\Base de datos turnos\DB.xlsx'
 
#excel_file_path = r'C:\Users\mayko\Downloads\DB.xlsx'


def create_dashboard(parent):
    """Crea un dashboard moderno con métricas clave"""
    dashboard = ttk.Frame(parent, padding=20)
    dashboard.pack(fill='x', pady=10)
   
    # Grid de tarjetas de métricas
    metrics_frame = ttk.Frame(dashboard)
    metrics_frame.pack(fill='x')
   
    # Configurar grid
    for i in range(4):
        metrics_frame.grid_columnconfigure(i, weight=1)
   
    # Estilo para las tarjetas
    style = ttk.Style()
    style.configure('Card.TFrame',
                   background='white',
                   relief='solid',
                   borderwidth=1)
   
    # Crear tarjetas de métricas
    metrics = [
        {"title": "Total Técnicos", "value": get_total_technicians(), "trend": "↑"},
        {"title": "Promedio Diario", "value": f"{get_daily_average():.1f}", "trend": "→"},
        {"title": "Máximo Diario", "value": get_daily_maximum(), "trend": "↑"},
        {"title": "Mínimo Diario", "value": get_daily_minimum(), "trend": "↓"}
    ]
   
    for i, metric in enumerate(metrics):
        card = ttk.Frame(metrics_frame, style='Card.TFrame')
        card.grid(row=0, column=i, padx=5, sticky='nsew')
       
        ttk.Label(card,
                 text=metric["title"],
                 font=('Segoe UI', 10)).pack(pady=(10, 5))
       
        ttk.Label(card,
                 text=metric["value"],
                 font=('Segoe UI', 16, 'bold')).pack()
       
        ttk.Label(card,
                 text=metric["trend"],
                 font=('Segoe UI', 12),
                 foreground='#28a745' if metric["trend"] == "↑" else
                          '#dc3545' if metric["trend"] == "↓" else
                          '#ffc107').pack(pady=(0, 10))

def configure_shift_button_styles():
    """Configura estilos específicos para los botones de jornada"""
    style = ttk.Style()
   
    # Colores para los botones de jornada
    SHIFT_COLORS = {
        'morning': {
            'main': '#ffc107',      # Amarillo/naranja para mañana
            'hover': '#e0a800',
            'selected': '#ff9800',
            'text': '#000000'       # Texto negro para mejor contraste
        },
        'afternoon': {
            'main': '#0dcaf0',      # Celeste para tarde
            'hover': '#0bacce',
            'selected': '#0995b5',
            'text': '#000000'
        },
        'night': {
            'main': '#6f42c1',      # Púrpura para noche
            'hover': '#5a37a1',
            'selected': '#4e2d8b',
            'text': '#ffffff'
        }
    }

    # Estilo para botón de mañana
    style.configure('Morning.TButton',
        background=SHIFT_COLORS['morning']['main'],
        foreground=SHIFT_COLORS['morning']['text'],
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
    style.map('Morning.TButton',
        background=[
            ('active', SHIFT_COLORS['morning']['hover']),
            ('selected', SHIFT_COLORS['morning']['selected'])
        ],
        foreground=[('disabled', '#666666')]
    )

    # Estilo para botón de tarde
    style.configure('Afternoon.TButton',
        background=SHIFT_COLORS['afternoon']['main'],
        foreground=SHIFT_COLORS['afternoon']['text'],
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
    style.map('Afternoon.TButton',
        background=[
            ('active', SHIFT_COLORS['afternoon']['hover']),
            ('selected', SHIFT_COLORS['afternoon']['selected'])
        ],
        foreground=[('disabled', '#666666')]
    )

    # Estilo para botón de noche
    style.configure('Night.TButton',
        background=SHIFT_COLORS['night']['main'],
        foreground=SHIFT_COLORS['night']['text'],
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold'),
        relief='flat',
        borderwidth=0
    )
    style.map('Night.TButton',
        background=[
            ('active', SHIFT_COLORS['night']['hover']),
            ('selected', SHIFT_COLORS['night']['selected'])
        ],
        foreground=[('disabled', '#666666')]
    )

    return SHIFT_COLORS

def update_shift_buttons(jornada_var, btn_mañana, btn_tarde, btn_noche):
    """Actualiza el estado visual de los botones de jornada"""
    selected = jornada_var.get()
   
    # Actualizar estilos basados en la selección
    btn_mañana.configure(
        style='Morning.TButton' + ('.selected' if selected == 'Mañana' else '')
    )
    btn_tarde.configure(
        style='Afternoon.TButton' + ('.selected' if selected == 'Tarde' else '')
    )
    btn_noche.configure(
        style='Night.TButton' + ('.selected' if selected == 'Noche' else '')
    )

# Modificar la parte del código donde se crean los botones de jornada
def setup_shift_buttons(button_frame, jornada_var):
    """Configura los botones de jornada con el nuevo estilo"""
    configure_shift_button_styles()
   
    btn_mañana = ttk.Button(
        button_frame,
        text="Mañana",
        command=lambda: select_jornada("Mañana", jornada_var),
        style='Morning.TButton',
        width=12
    )
   
    btn_tarde = ttk.Button(
        button_frame,
        text="Tarde",
        command=lambda: select_jornada("Tarde", jornada_var),
        style='Afternoon.TButton',
        width=12
    )
   
    btn_noche = ttk.Button(
        button_frame,
        text="Noche",
        command=lambda: select_jornada("Noche", jornada_var),
        style='Night.TButton',
        width=12
    )
   
    # Organizar los botones
    btn_mañana.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
    btn_tarde.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
    btn_noche.grid(row=0, column=2, padx=5, pady=5, sticky='ew')
   
    # Configurar el grid
    button_frame.grid_columnconfigure((0,1,2), weight=1)
   
    return btn_mañana, btn_tarde, btn_noche

def select_jornada(jornada, jornada_var):
    """Función modificada para seleccionar jornada"""
    jornada_var.set(jornada)
    # Actualizar estado visual de los botones
    update_shift_buttons(jornada_var, btn_mañana, btn_tarde, btn_noche)
    # Actualizar la tabla
    show_scheduled_technicians()


def cargar_valores_turno() -> None:
    """Carga los valores de turno desde el Excel"""
    global turno_combobox
    try:
        df = pd.read_excel(excel_file_path)
        if 'turno_bd' in df.columns:
            turnos = df['turno_bd'].dropna().unique().tolist()
            turnos.sort()
            if turno_combobox is not None:
                turno_combobox['values'] = turnos
        else:
            print("La columna 'turno_bd' no existe en el Excel")
            if turno_combobox is not None:
                turno_combobox['values'] = []
    except Exception as e:
        print(f"Error al cargar valores de turno: {str(e)}")
        if turno_combobox is not None:
            turno_combobox['values'] = []

def actualizar_tecnicos(event=None) -> None:
    """Actualiza la lista de técnicos basado en el turno seleccionado"""
    global turno_combobox, name_combobox, cargo_entry, rotation_letter_entry
    try:
        turno_seleccionado = turno_combobox.get() if turno_combobox else None
        if turno_seleccionado:
            df = pd.read_excel(excel_file_path)
            if 'turno_bd' in df.columns and 'nombre_bd' in df.columns:
                tecnicos = df[df['turno_bd'] == turno_seleccionado]['nombre_bd'].dropna().unique().tolist()
                tecnicos.sort()
               
                if name_combobox is not None:
                    name_combobox['values'] = tecnicos
                    name_combobox.set('')
               
                if cargo_entry is not None:
                    cargo_entry.configure(state='normal')
                    cargo_entry.delete(0, tk.END)
                    cargo_entry.configure(state='readonly')
               
                if rotation_letter_entry is not None:
                    rotation_letter_entry.configure(state='normal')
                    rotation_letter_entry.delete(0, tk.END)
                    rotation_letter_entry.configure(state='readonly')
            else:
                print("Las columnas necesarias no existen en el Excel")
    except Exception as e:
        print(f"Error al actualizar técnicos: {str(e)}")

def actualizar_cargo(event=None) -> None:
    """Actualiza el cargo y letra de rotación basado en el técnico seleccionado"""
    global name_combobox, cargo_entry, rotation_letter_entry
    try:
        tecnico_seleccionado = name_combobox.get() if name_combobox else None
        if tecnico_seleccionado:
            df = pd.read_excel(excel_file_path)
            if 'nombre_bd' in df.columns:
                tecnico_data = df[df['nombre_bd'] == tecnico_seleccionado]
                if not tecnico_data.empty:
                    tecnico_data = tecnico_data.iloc[0]
                   
                    if cargo_entry is not None:
                        cargo_entry.configure(state='normal')
                        cargo_entry.delete(0, tk.END)
                        if 'cargo_bd' in tecnico_data:
                            cargo_entry.insert(0, tecnico_data['cargo_bd'])
                        cargo_entry.configure(state='readonly')
                   
                    if rotation_letter_entry is not None:
                        rotation_letter_entry.configure(state='normal')
                        rotation_letter_entry.delete(0, tk.END)
                        if 'Letra_bd' in tecnico_data and pd.notna(tecnico_data['Letra_bd']):
                            rotation_letter_entry.insert(0, tecnico_data['Letra_bd'])
                        rotation_letter_entry.configure(state='readonly')
            else:
                print("La columna 'nombre_bd' no existe en el Excel")
    except Exception as e:
        print(f"Error al actualizar cargo: {str(e)}")
        if cargo_entry is not None:
            cargo_entry.configure(state='normal')
            cargo_entry.delete(0, tk.END)
            cargo_entry.configure(state='readonly')
        if rotation_letter_entry is not None:
            rotation_letter_entry.configure(state='normal')
            rotation_letter_entry.delete(0, tk.END)
            rotation_letter_entry.configure(state='readonly')

def setup_ai_analysis_tab(tab):
    """Configura la pestaña de análisis avanzado con barra de scroll"""
    # Crear contenedor principal con PanedWindow horizontal
    main_paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
    main_paned.pack(fill='both', expand=True, padx=5, pady=5)

    # === Panel Izquierdo con Scroll ===
    # Contenedor del panel izquierdo
    left_container = ttk.Frame(main_paned, width=350)
    left_container.pack_propagate(False)

    # Crear canvas y scrollbar para el panel izquierdo
    canvas = tk.Canvas(left_container, width=330)
    scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=canvas.yview)
   
    # Frame interior para los controles
    left_panel = ttk.Frame(canvas)
   
    # Configurar el scroll
    def configure_scroll(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
   
    left_panel.bind('<Configure>', configure_scroll)
    canvas_frame = canvas.create_window((0, 0), window=left_panel, anchor="nw", width=330)

    # Configurar el canvas para redimensionar el frame interior
    def configure_canvas(event):
        canvas.itemconfig(canvas_frame, width=event.width)
   
    canvas.bind('<Configure>', configure_canvas)

    # === Contenido del Panel Izquierdo ===
    # Frame para parámetros
    param_frame = ttk.LabelFrame(left_panel, text="Parámetros de Análisis", padding=10)
    param_frame.pack(fill='x', padx=5, pady=5)

    # Variables de configuración
    config_vars = {
        'confidence': tk.DoubleVar(value=0.95),
        'horizon': tk.IntVar(value=14),
        'sensitivity': tk.DoubleVar(value=0.1),
        'jornada': tk.StringVar(value="todas")
    }

    # Nivel de confianza
    ttk.Label(param_frame, text="Nivel de Confianza:").pack(anchor='w')
    confidence_scale = ttk.Scale(
        param_frame,
        from_=0.8,
        to=0.99,
        variable=config_vars['confidence'],
        orient='horizontal'
    )
    confidence_scale.pack(fill='x', pady=2)
    ttk.Label(param_frame, textvariable=config_vars['confidence']).pack(anchor='e')

    # Horizonte de predicción
    ttk.Label(param_frame, text="Horizonte de Predicción (días):").pack(anchor='w', pady=(10,0))
    ttk.Spinbox(
        param_frame,
        from_=7,
        to=90,
        textvariable=config_vars['horizon']
    ).pack(fill='x', pady=2)

    # Sensibilidad
    ttk.Label(param_frame, text="Sensibilidad de Detección:").pack(anchor='w', pady=(10,0))
    sensitivity_scale = ttk.Scale(
        param_frame,
        from_=0.01,
        to=0.2,
        variable=config_vars['sensitivity'],
        orient='horizontal'
    )
    sensitivity_scale.pack(fill='x', pady=2)
    ttk.Label(param_frame, textvariable=config_vars['sensitivity']).pack(anchor='e')

    # Filtro de jornada
    shift_frame = ttk.LabelFrame(left_panel, text="Filtro de Jornada", padding=10)
    shift_frame.pack(fill='x', padx=5, pady=5)

    for value in ["todas", "Mañana", "Tarde", "Noche"]:
        ttk.Radiobutton(
            shift_frame,
            text="Todas las Jornadas" if value == "todas" else value,
            variable=config_vars['jornada'],
            value=value
        ).pack(anchor='w', pady=2)

    # Agregar búsqueda avanzada
    df = pd.read_excel(excel_file_path)
    add_search_functionality(left_panel, df)

    # Agregar opciones de exportación
    add_export_options(left_panel, None)

    # Botones de acción
    button_frame = ttk.LabelFrame(left_panel, text="Acciones", padding=10)
    button_frame.pack(fill='x', padx=5, pady=5)

    def execute_analysis():
        try:
            results = run_analysis(config_vars, results_notebook)
           
            # Actualizar análisis adicionales
            text_widget = ttk.Frame(right_panel).children.get('!text', None)
            if text_widget and results:
                df = pd.read_excel(excel_file_path)
                add_comparative_analysis(df, text_widget)
                add_efficiency_metrics(df, text_widget)
                add_optimization_suggestions(df, text_widget)
               
            # Actualizar opciones de exportación
            add_export_options(left_panel, results)
           
        except Exception as e:
            messagebox.showerror("Error", f"Error en el análisis: {str(e)}")

    # Botones con mejor estilo
    ttk.Button(
        button_frame,
        text="📊 Ejecutar Análisis",
        command=execute_analysis,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        button_frame,
        text="🗑️ Limpiar Resultados",
        command=lambda: clear_results_from_notebook(results_notebook),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        button_frame,
        text="💾 Guardar Análisis",
        command=lambda: export_analysis_results(config_vars),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # Panel derecho para resultados
    right_panel = ttk.Frame(main_paned)

    # Notebook para resultados
    results_notebook = ttk.Notebook(right_panel)
    results_notebook.pack(fill='both', expand=True)

    # Pestañas de resultados
    tabs = {
        'overview': create_result_tab(results_notebook, "Vista General"),
        'patterns': create_result_tab(results_notebook, "Patrones"),
        'predictions': create_result_tab(results_notebook, "Predicciones"),
        'anomalies': create_result_tab(results_notebook, "Anomalías")
    }

    # Configurar scroll y empaquetado
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Configurar scroll con mousewheel
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
   
    canvas.bind_all("<MouseWheel>", on_mousewheel)

    # Agregar paneles al PanedWindow principal
    main_paned.add(left_container)
    main_paned.add(right_panel)

    # Mantener proporción de paneles
    def maintain_ratio(event=None):
        total_width = main_paned.winfo_width()
        left_width = 350
        main_paned.sashpos(0, left_width)
   
    main_paned.bind('<Configure>', maintain_ratio)
    tab.after(100, maintain_ratio)

    return tabs

def create_result_tab(notebook, title):
    """Crea una pestaña de resultados"""
    frame = ttk.Frame(notebook)
    notebook.add(frame, text=title)

    # Frame para gráfico
    graph_frame = ttk.LabelFrame(frame, text="Visualización")
    graph_frame.pack(fill='both', expand=True, padx=5, pady=5)

    # Frame para métricas
    metrics_frame = ttk.LabelFrame(frame, text="Métricas")
    metrics_frame.pack(fill='x', padx=5, pady=5)

    return {
        'frame': frame,
        'graph_frame': graph_frame,
        'metrics_frame': metrics_frame
    }

def clear_results_from_notebook(notebook):
    """Limpia los resultados anteriores del notebook"""
    if notebook:
        for tab_id in notebook.tabs():
            tab = notebook.nametowidget(tab_id)
            for child in tab.winfo_children():
                if isinstance(child, ttk.LabelFrame):
                    for widget in child.winfo_children():
                        widget.destroy()





def analyze_temporal_patterns(df):
    """Analiza patrones temporales en los datos"""
    daily_counts = df.groupby('fecha_asignacion').size()
   
    # Descomposición temporal si hay suficientes datos
    if len(daily_counts) >= 2:  # Verificar que hay suficientes datos
        try:
            # Rellenar valores faltantes si los hay
            daily_counts_filled = daily_counts.asfreq('D', fill_value=daily_counts.mean())
            decomposition = seasonal_decompose(daily_counts_filled, period=7, extrapolate_trend='freq')
           
            return {
                'daily_counts': daily_counts,
                'trend': decomposition.trend,
                'seasonal': decomposition.seasonal,
                'residual': decomposition.resid
            }
        except Exception as e:
            print(f"Error en descomposición temporal: {str(e)}")
            return {
                'daily_counts': daily_counts,
                'trend': None,
                'seasonal': None,
                'residual': None
            }
    else:
        return {
            'daily_counts': daily_counts,
            'trend': None,
            'seasonal': None,
            'residual': None
        }

def create_visualizations(patterns, predictions, anomalies, optimization):
    """Crea visualizaciones para todos los análisis"""
    # Crear figura con subplots
    fig = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=(
            'Análisis de Patrones',
            'Predicciones',
            'Detección de Anomalías',
            'Optimización'
        )
    )
   
    # Gráfico de patrones
    add_patterns_plot(fig, patterns, row=1, col=1)
   
    # Gráfico de predicciones
    add_predictions_plot(fig, predictions, row=1, col=2)
   
    # Gráfico de anomalías
    add_anomalies_plot(fig, anomalies, row=2, col=1)
   
    # Gráfico de optimización
    add_optimization_plot(fig, optimization, row=2, col=2)
   
    # Actualizar layout
    fig.update_layout(height=800, showlegend=True)
   
    return fig




def analyze_workload(df_jornada, jornada):
    """Analiza la carga de trabajo para una jornada específica"""
    # Análisis de distribución temporal
    daily_counts = df_jornada.groupby('fecha_asignacion').size()
   
    # Calcular métricas
    mean_load = daily_counts.mean()
    std_load = daily_counts.std()
    cv = std_load / mean_load if mean_load > 0 else 0
   
    ai_text.insert(tk.END, f"\nMétricas de carga - {jornada}:")
    ai_text.insert(tk.END, f"\n• Carga promedio: {mean_load:.2f}")
    ai_text.insert(tk.END, f"\n• Desviación estándar: {std_load:.2f}")
    ai_text.insert(tk.END, f"\n• Coeficiente de variación: {cv:.2f}")

    # Recomendaciones basadas en métricas
    if cv > 0.3:
        ai_text.insert(tk.END, "\n\nRecomendaciones:")
        ai_text.insert(tk.END, "\n• Alta variabilidad en la carga de trabajo")
        ai_text.insert(tk.END, "\n• Considerar redistribución más uniforme")
       
    # Análisis de patrones semanales
    weekly_pattern = df_jornada.groupby('dia_semana').size()
    ai_text.insert(tk.END, "\n\nPatrones semanales:")
   
    for dia in range(7):
        count = weekly_pattern.get(dia, 0)
        ai_text.insert(tk.END, f"\n• {DIAS_SEMANA[dia]}: {count} técnicos")

    # Análisis de balance de carga
    weekly_cv = weekly_pattern.std() / weekly_pattern.mean() if weekly_pattern.mean() > 0 else 0
    if weekly_cv > 0.2:
        ai_text.insert(tk.END, "\n\nAlerta: Desbalance en distribución semanal")
        ai_text.insert(tk.END, "\nRecomendación: Redistribuir carga entre días")

    # Análisis de rotaciones
    rotation_pattern = df_jornada.groupby('letra_rotacion').size()
    ai_text.insert(tk.END, "\n\nDistribución de rotaciones:")
    for rot, count in rotation_pattern.items():
        ai_text.insert(tk.END, f"\n• {rot}: {count} asignaciones")
       

def analyze_performance_metrics():
    """Analiza métricas de rendimiento avanzadas por jornada"""
    try:
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])

        metrics = {}
        for jornada in df['Jornada'].unique():
            df_jornada = df[df['Jornada'] == jornada]
           
            # Calcular métricas clave
            total_assignments = len(df_jornada)
            unique_technicians = df_jornada['nombre'].nunique()
            assignments_per_tech = total_assignments / unique_technicians if unique_technicians > 0 else 0
           
            # Calcular estabilidad de rotación
            rotation_changes = df_jornada.groupby('nombre')['letra_rotacion'].nunique().mean()
           
            # Calcular distribución por cargo
            cargo_distribution = df_jornada.groupby('cargo').size()
           
            # Almacenar métricas
            metrics[jornada] = {
                'total_assignments': total_assignments,
                'unique_technicians': unique_technicians,
                'assignments_per_tech': assignments_per_tech,
                'rotation_stability': rotation_changes,
                'cargo_distribution': cargo_distribution
            }

        return metrics

    except Exception as e:
        print(f"Error en análisis de métricas: {str(e)}")
        return None
       

def generate_optimization_recommendations(metrics):
    """Genera recomendaciones basadas en métricas de rendimiento"""
    recommendations = {}
   
    for jornada, metric in metrics.items():
        jornada_recommendations = []
       
        # Analizar asignaciones por técnico
        if metric['assignments_per_tech'] > 15:
            jornada_recommendations.append(
                "Alta carga de trabajo por técnico. Considerar aumentar dotación."
            )
        elif metric['assignments_per_tech'] < 5:
            jornada_recommendations.append(
                "Baja utilización de recursos. Evaluar redistribución."
            )
           
        # Analizar estabilidad de rotación
        if metric['rotation_stability'] > 2:
            jornada_recommendations.append(
                "Alta variabilidad en rotaciones. Considerar estabilizar asignaciones."
            )
           
        # Analizar distribución de cargos
        cargo_dist = metric['cargo_distribution']
        if cargo_dist.std() / cargo_dist.mean() > 0.3:
            jornada_recommendations.append(
                "Desbalance en distribución de cargos. Revisar estructura de equipo."
            )
           
        recommendations[jornada] = jornada_recommendations
       
    return recommendations

def predict_staffing_needs():
    """Predice necesidades de personal usando modelos avanzados de ML"""
    try:
        df = pd.read_excel(excel_file_path)
        predictions = {}
       
        for jornada in df['Jornada'].unique():
            df_jornada = df[df['Jornada'] == jornada]
           
            # Preparar características
            X = prepare_features(df_jornada)
            y = prepare_targets(df_jornada)
           
            # Dividir datos
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.2, random_state=42
            )
           
            # Entrenar modelo
            model = train_advanced_model(X_train, y_train)
           
            # Evaluar modelo
            score = evaluate_model(model, X_test, y_test)
           
            # Generar predicciones
            future_predictions = generate_future_predictions(model, jornada)
           
            predictions[jornada] = {
                'score': score,
                'predictions': future_predictions
            }
           
        return predictions
           
    except Exception as e:
        print(f"Error en predicción de necesidades: {str(e)}")
        return None

def prepare_features(self, df):
    """Prepara las características para el modelo predictivo"""
    try:
        # Agrupar datos por fecha
        daily_counts = df.groupby('fecha_asignacion').size().reset_index()
        daily_counts.columns = ['fecha_asignacion', 'count']

        # Crear características temporales
        features = pd.DataFrame()
        features['dia_semana'] = daily_counts['fecha_asignacion'].dt.dayofweek
        features['mes'] = daily_counts['fecha_asignacion'].dt.month
        features['dia_mes'] = daily_counts['fecha_asignacion'].dt.day
        features['dia_año'] = daily_counts['fecha_asignacion'].dt.dayofyear
        features['semana'] = daily_counts['fecha_asignacion'].dt.isocalendar().week

        # Rellenar valores faltantes
        features = features.fillna(0)

        # Asegurar que todas las características tienen el mismo número de muestras
        X = features[['dia_semana', 'mes', 'dia_mes', 'dia_año', 'semana']].values
        y = daily_counts['count'].values

        # Verificar que X e y tienen el mismo número de muestras
        if len(X) != len(y):
            raise ValueError(f"Inconsistencia en número de muestras: X={len(X)}, y={len(y)}")

        return X, y

    except Exception as e:
        print(f"Error en prepare_features: {str(e)}")
        raise

def train_advanced_model(X_train, y_train):
    """Entrena un modelo avanzado de Random Forest"""
    model = RandomForestRegressor(
        n_estimators=200,
        max_depth=None,
        min_samples_split=2,
        min_samples_leaf=1,
        bootstrap=True,
        random_state=42,
        n_jobs=-1
    )
   
    # Ajuste de hiperparámetros
    param_grid = {
        'n_estimators': [100, 200, 300],
        'max_depth': [None, 10, 20, 30],
        'min_samples_split': [2, 5, 10],
        'min_samples_leaf': [1, 2, 4]
    }
   
    grid_search = GridSearchCV(
        model,
        param_grid,
        cv=5,
        n_jobs=-1,
        scoring='neg_mean_squared_error'
    )
   
    grid_search.fit(X_train, y_train)
   
    return grid_search.best_estimator_

def generate_future_predictions(model, jornada):
    """Genera predicciones para las próximas semanas"""
    future_dates = pd.date_range(
        start=datetime.now(),
        periods=28,  # 4 semanas
        freq='D'
    )
   
    future_features = pd.DataFrame()
    future_features['dia_semana'] = future_dates.dayofweek
    future_features['mes'] = future_dates.month
    future_features['semana'] = future_dates.isocalendar().week
   
    # Normalizar características
    scaler = StandardScaler()
    future_features_scaled = scaler.fit_transform(future_features)
   
    predictions = model.predict(future_features_scaled)
   
    return pd.Series(predictions, index=future_dates)

def update_ai_interface():
    """Actualiza la interfaz de IA con los últimos análisis"""
    try:
        # Obtener métricas
        metrics = analyze_performance_metrics()
        recommendations = generate_optimization_recommendations(metrics)
        predictions = predict_staffing_needs()
       
        # Actualizar interfaz
        ai_text.delete(1.0, tk.END)
        ai_text.insert(tk.END, "=== ANÁLISIS AVANZADO DE IA ===\n\n")
       
        # Mostrar métricas por jornada
        for jornada in metrics.keys():
            ai_text.insert(tk.END, f"\nJornada: {jornada}\n")
            ai_text.insert(tk.END, "-------------------\n")
           
            # Métricas
            metric = metrics[jornada]
            ai_text.insert(tk.END, f"Total asignaciones: {metric['total_assignments']}\n")
            ai_text.insert(tk.END, f"Técnicos únicos: {metric['unique_technicians']}\n")
            ai_text.insert(tk.END, f"Asignaciones por técnico: {metric['assignments_per_tech']:.2f}\n")
           
            # Recomendaciones
            ai_text.insert(tk.END, "\nRecomendaciones:\n")
            for rec in recommendations[jornada]:
                ai_text.insert(tk.END, f"• {rec}\n")
           
            # Predicciones
            if predictions and jornada in predictions:
                pred = predictions[jornada]
                ai_text.insert(tk.END, f"\nPrecisión del modelo: {pred['score']:.2f}\n")
                ai_text.insert(tk.END, "Predicciones próximas 4 semanas:\n")
               
                for date, value in pred['predictions'].items():
                    ai_text.insert(tk.END, f"• {date.strftime('%d/%m/%Y')}: {int(value)} técnicos\n")
       
    except Exception as e:
        messagebox.showerror("Error", f"Error al actualizar interfaz de IA: {str(e)}")

# Función principal para ejecutar el análisis
 
       




def predict_future_needs():
    """Predice necesidades futuras usando Random Forest con análisis por jornada"""
    try:
        df = pd.read_excel(excel_file_path)
        ai_text.delete(1.0, tk.END)
        ai_text.insert(tk.END, "=== PREDICCIÓN DE NECESIDADES FUTURAS CON IA ===\n\n")

        # Preparar datos
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek
        df['mes'] = df['fecha_asignacion'].dt.month
        df['semana'] = df['fecha_asignacion'].dt.isocalendar().week

        # Crear modelos por jornada
        models = {}
        accuracies = {}

        for jornada in df['Jornada'].unique():
            df_jornada = df[df['Jornada'] == jornada]
           
            # Crear features agregados por jornada
            daily_counts = df_jornada.groupby([
                'fecha_asignacion',
                'dia_semana',
                'mes',
                'semana'
            ]).size().reset_index()
            daily_counts.columns = ['fecha', 'dia_semana', 'mes', 'semana', 'cantidad']

            # Preparar datos para el modelo
            X = daily_counts[['dia_semana', 'mes', 'semana']].values
            y = daily_counts['cantidad'].values

            # Dividir datos
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.2, random_state=42
            )

            # Entrenar modelo
            rf_model = RandomForestRegressor(
                n_estimators=200,
                max_depth=None,
                min_samples_split=2,
                min_samples_leaf=1,
                bootstrap=True,
                random_state=42
            )
           
            rf_model.fit(X_train, y_train)

            # Evaluar modelo
            y_pred = rf_model.predict(X_test)
            accuracy = r2_score(y_test, y_pred)
           
            models[jornada] = rf_model
            accuracies[jornada] = accuracy

        ai_text.insert(tk.END, "1. Precisión de los Modelos por Jornada:\n")
        for jornada, accuracy in accuracies.items():
            ai_text.insert(tk.END, f"• Jornada {jornada}: {accuracy*100:.2f}%\n")

        # Generar predicciones para próxima semana
        ai_text.insert(tk.END, "\n2. Predicciones para Próxima Semana por Jornada:\n")
        ultima_fecha = df['fecha_asignacion'].max()
       
        for jornada, model in models.items():
            ai_text.insert(tk.END, f"\nJornada {jornada}:\n")
           
            for i in range(7):
                fecha_prediccion = ultima_fecha + pd.Timedelta(days=i+1)
                dia = fecha_prediccion.dayofweek
                mes = fecha_prediccion.month
                semana = fecha_prediccion.isocalendar().week

                # Generar predicción
                prediccion = model.predict([[dia, mes, semana]])[0]
               
                ai_text.insert(tk.END, f"\n• {fecha_prediccion.strftime('%d/%m/%Y')} ({DIAS_SEMANA[dia]}):")
                ai_text.insert(tk.END, f"\n  Técnicos necesarios: {int(prediccion)}")

                # Análisis de confianza
                feature_importance = model.feature_importances_
                confianza = np.mean(feature_importance) * 100
                ai_text.insert(tk.END, f"\n  Confianza: {confianza:.1f}%\n")

        # Guardar modelos entrenados
        for jornada, model in models.items():
            joblib.dump(model, f'modelo_prediccion_{jornada}.joblib')

    except Exception as e:
        messagebox.showerror("Error", f"Error en predicción: {str(e)}")

def optimize_distribution(df):
    """Optimiza la distribución de personal"""
    current_distribution = analyze_current_distribution(df)
    constraints = define_optimization_constraints(df)
   
    model = create_optimization_model(current_distribution, constraints)
    solution = solve_optimization_model(model)
   
    return {
        'current': current_distribution,
        'optimal': solution,
        'improvements': calculate_improvements(current_distribution, solution)
    }


def analyze_anomaly_causes(anomalias_df):
    """Analiza las causas comunes de anomalías"""
    causes = {}
   
    # Analizar días de la semana
    dia_counts = anomalias_df['dia_semana'].value_counts()
    for dia, count in dia_counts.items():
        causes[f"Día problemático: {DIAS_SEMANA[dia]}"] = count
   
    # Analizar combinaciones cargo-rotación
    if 'cargo' in anomalias_df.columns and 'letra_rotacion' in anomalias_df.columns:
        combinaciones = anomalias_df.groupby(['cargo', 'letra_rotacion']).size()
        for (cargo, rotacion), count in combinaciones.items():
            causes[f"Combinación {cargo}-{rotacion}"] = count
   
    return causes

def optimize_workload(current_load, total_staff, jornada):
    """Optimiza la distribución de carga de trabajo por jornada"""
    # Parámetros específicos por jornada
    JORNADA_PARAMS = {
        'Mañana': {
            'min_staff': 4,
            'optimal_range': (5, 8),
            'peak_days': [0, 1]  # Lunes y Martes
        },
        'Tarde': {
            'min_staff': 4,
            'optimal_range': (5, 8),
            'peak_days': [2, 3]  # Miércoles y Jueves
        },
        'Noche': {
            'min_staff': 4,
            'optimal_range': (5, 8),
            'peak_days': [4, 5]  # Viernes y Sábado
        }
    }

    params = JORNADA_PARAMS.get(jornada, JORNADA_PARAMS['Mañana'])
    optimization = {}
    mean_load = total_staff / 7  # distribución ideal por día

    for dia in range(7):
        current = current_load.get(dia, 0)
       
        # Ajustar según día pico
        is_peak_day = dia in params['peak_days']
        min_staff = params['min_staff']
        optimal_min, optimal_max = params['optimal_range']

        if is_peak_day:
            target = max(mean_load * 1.2, optimal_min)
        else:
            target = mean_load

        if current < min_staff:
            optimization[dia] = max(min_staff, round(target * 1.1))
        elif current > optimal_max:
            optimization[dia] = round(target * 0.9)
        else:
            optimization[dia] = round(target)

    return optimization


def setup_rotation_tab(notebook):
    """Configuración de la pestaña de análisis de rotación"""
    tab = ttk.Frame(notebook)
   
    # Frame para controles
    control_frame = ttk.LabelFrame(tab, text="Controles de Análisis")
    control_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        control_frame,
        text="Analizar Patrones de Rotación",
        command=analyze_rotation_patterns,
        style='Primary.TButton'
    ).pack(pady=5, padx=5)

    # Frame para resultados
    global rotation_result_frame
    rotation_result_frame = ttk.LabelFrame(tab, text="Resultados del Análisis")
    rotation_result_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # ScrolledText para mostrar resultados
    global rotation_text
    rotation_text = tk.Text(rotation_result_frame, wrap=tk.WORD, height=10)
    rotation_text.pack(fill='both', expand=True, padx=5, pady=5)
   
    scrollbar = ttk.Scrollbar(rotation_result_frame, command=rotation_text.yview)
    scrollbar.pack(side='right', fill='y')
    rotation_text.configure(yscrollcommand=scrollbar.set)

    return tab

def setup_prediction_tab(notebook):
    """Configuración de la pestaña de predicciones"""
    tab = ttk.Frame(notebook)
   
    # Frame para controles
    pred_control_frame = ttk.LabelFrame(tab, text="Controles de Predicción")
    pred_control_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        pred_control_frame,
        text="Generar Predicciones",
        command=generate_predictions,
        style='Primary.TButton'
    ).pack(pady=5, padx=5)

    # Frame para resultados
    global prediction_result_frame
    prediction_result_frame = ttk.LabelFrame(tab, text="Predicciones y Recomendaciones")
    prediction_result_frame.pack(fill='both', expand=True, padx=10, pady=5)

    global prediction_text
    prediction_text = tk.Text(prediction_result_frame, wrap=tk.WORD, height=10)
    prediction_text.pack(fill='both', expand=True, padx=5, pady=5)
   
    pred_scrollbar = ttk.Scrollbar(prediction_result_frame, command=prediction_text.yview)
    pred_scrollbar.pack(side='right', fill='y')
    prediction_text.configure(yscrollcommand=pred_scrollbar.set)

    return tab

def setup_metrics_tab(notebook):
    """Configuración de la pestaña de métricas"""
    tab = ttk.Frame(notebook)
   
    # Frame para controles
    metrics_control_frame = ttk.LabelFrame(tab, text="Controles de Métricas")
    metrics_control_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        metrics_control_frame,
        text="Calcular Métricas",
        command=calculate_advanced_metrics,
        style='Primary.TButton'
    ).pack(pady=5, padx=5)

    # Frame para gráficos y resultados
    global metrics_result_frame
    metrics_result_frame = ttk.LabelFrame(tab, text="Resultados y Visualizaciones")
    metrics_result_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # Canvas para gráficos
    global metrics_figure
    metrics_figure = Figure(figsize=(8, 6))
    global metrics_canvas
    metrics_canvas = FigureCanvasTkAgg(metrics_figure, master=metrics_result_frame)
    metrics_canvas.get_tk_widget().pack(fill='both', expand=True)

    return tab

def analyze_rotation_patterns():
    try:
        df = pd.read_excel(excel_file_path)
       
        rotation_text.delete(1.0, tk.END)
        rotation_text.insert(tk.END, "=== ANÁLISIS DE PATRONES DE ROTACIÓN ===\n\n")

        # 1. Análisis de distribución de rotaciones
        rotaciones = df['letra_rotacion'].value_counts()
        rotation_text.insert(tk.END, "1. Distribución de Rotaciones:\n")
        for rot, count in rotaciones.items():
            percentage = (count / len(df)) * 100
            rotation_text.insert(tk.END, f"   • {rot}: {count} casos ({percentage:.1f}%)\n")

        # 2. Análisis temporal
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        df['mes'] = df['fecha_asignacion'].dt.month
        df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek

        # Patrones por día de la semana
        rotation_text.insert(tk.END, "\n2. Patrones por Día:\n")
        for dia in range(7):
            day_data = df[df['dia_semana'] == dia]
            if not day_data.empty:
                rotation_text.insert(tk.END, f"\n   {DIAS_SEMANA[dia]}:\n")
                day_rotations = day_data['letra_rotacion'].value_counts()
                for rot, count in day_rotations.items():
                    rotation_text.insert(tk.END, f"      - {rot}: {count}\n")

        # 3. Análisis de secuencias
        rotation_text.insert(tk.END, "\n3. Análisis de Secuencias:\n")
       
        # Detectar secuencias comunes
        df_sorted = df.sort_values('fecha_asignacion')
        tecnico_sequences = df_sorted.groupby('nombre')['letra_rotacion'].agg(list)
       
        common_sequences = []
        for sequences in tecnico_sequences:
            if len(sequences) >= 2:
                for i in range(len(sequences)-1):
                    common_sequences.append(f"{sequences[i]} → {sequences[i+1]}")
       
        if common_sequences:
            sequence_counts = Counter(common_sequences)
            rotation_text.insert(tk.END, "   Transiciones más comunes:\n")
            for seq, count in sequence_counts.most_common(5):
                rotation_text.insert(tk.END, f"      • {seq}: {count} veces\n")

        # 4. Recomendaciones
        rotation_text.insert(tk.END, "\n4. Recomendaciones:\n")
       
        # Analizar balance de rotaciones
        rotacion_cv = rotaciones.std() / rotaciones.mean() if rotaciones.mean() > 0 else 0
        if rotacion_cv > 0.3:
            rotation_text.insert(tk.END, "   • Desequilibrio detectado en la distribución de rotaciones\n")
            rotation_text.insert(tk.END, "   • Recomendación: Equilibrar la asignación de rotaciones\n")
       
        # Analizar patrones semanales
        weekly_patterns = df.groupby(['dia_semana', 'letra_rotacion']).size().unstack(fill_value=0)
        weekly_cv = weekly_patterns.std().mean() / weekly_patterns.mean().mean()
        if weekly_cv > 0.3:
            rotation_text.insert(tk.END, "\n   • Variabilidad significativa en patrones semanales\n")
            rotation_text.insert(tk.END, "   • Recomendación: Establecer patrones más consistentes por día\n")

        rotation_text.see(1.0)

    except Exception as e:
        messagebox.showerror("Error", f"Error en el análisis de rotación: {str(e)}")




def analyze_rotation_patterns():
    try:
        # Leer datos
        df = pd.read_excel(excel_file_path)
       
        if df.empty:
            messagebox.showinfo("Información", "No hay datos para analizar.")
            return

        rotation_text.delete(1.0, tk.END)
        rotation_text.insert(tk.END, "=== ANÁLISIS DE PATRONES DE ROTACIÓN ===\n\n")

        # 1. Análisis de distribución de turnos
        turnos_count = df['letra_rotacion'].value_counts()
        rotation_text.insert(tk.END, "1. Distribución de Rotaciones:\n")
        for turno, count in turnos_count.items():
            rotation_text.insert(tk.END, f"   • {turno}: {count} asignaciones\n")

        # 2. Análisis temporal
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        df['mes'] = df['fecha_asignacion'].dt.month
        monthly_rotation = df.groupby(['mes', 'letra_rotacion']).size().unstack(fill_value=0)
       
        rotation_text.insert(tk.END, "\n2. Tendencias Mensuales:\n")
        for mes in monthly_rotation.index:
            mes_nombre = calendar.month_name[mes]
            rotation_text.insert(tk.END, f"\n   {mes_nombre}:\n")
            for rotacion in monthly_rotation.columns:
                count = monthly_rotation.loc[mes, rotacion]
                if count > 0:
                    rotation_text.insert(tk.END, f"      • {rotacion}: {count}\n")

        # 3. Análisis de balance de carga
        rotation_text.insert(tk.END, "\n3. Análisis de Balance de Carga:\n")
        std_dev = turnos_count.std()
        mean = turnos_count.mean()
        cv = std_dev / mean if mean > 0 else 0
       
        if cv < 0.2:
            balance_msg = "Excelente balance en la distribución de turnos"
        elif cv < 0.4:
            balance_msg = "Buen balance en la distribución de turnos"
        else:
            balance_msg = "Se recomienda mejorar el balance en la distribución de turnos"
           
        rotation_text.insert(tk.END, f"   • {balance_msg}\n")
        rotation_text.insert(tk.END, f"   • Coeficiente de variación: {cv:.2f}\n")

        # 4. Recomendaciones
        rotation_text.insert(tk.END, "\n4. Recomendaciones:\n")
       
        # Identificar rotaciones menos utilizadas
        under_utilized = turnos_count[turnos_count < turnos_count.mean() * 0.8]
        if not under_utilized.empty:
            rotation_text.insert(tk.END, "   • Rotaciones subutilizadas:\n")
            for rot, count in under_utilized.items():
                rotation_text.insert(tk.END, f"      - {rot}: {count} asignaciones\n")
            rotation_text.insert(tk.END, "     Se recomienda aumentar el uso de estas rotaciones\n")

        # Identificar días con sobrecarga
        daily_count = df.groupby('fecha_asignacion').size()
        overloaded_days = daily_count[daily_count > daily_count.mean() * 1.2]
        if not overloaded_days.empty:
            rotation_text.insert(tk.END, "\n   • Se detectaron días con posible sobrecarga:\n")
            for date, count in overloaded_days.items():
                rotation_text.insert(tk.END, f"      - {date.strftime('%d/%m/%Y')}: {count} técnicos\n")

        rotation_text.see(1.0)

    except Exception as e:
        messagebox.showerror("Error", f"Error en el análisis de rotación: {str(e)}")
       
       
def setup_absence_tab(notebook):
    """Configuración de la pestaña de análisis de ausentismo"""
    tab = ttk.Frame(notebook)
   
    control_frame = ttk.LabelFrame(tab, text="Controles de Análisis")
    control_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        control_frame,
        text="Analizar Patrones de Ausentismo",
        command=analyze_absence_patterns,
        style='Primary.TButton'
    ).pack(pady=5, padx=5)

    # Frame para resultados
    global absence_frame
    absence_frame = ttk.LabelFrame(tab, text="Análisis de Ausentismo")
    absence_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # Text widget para resultados
    global absence_text
    absence_text = tk.Text(absence_frame, wrap=tk.WORD, height=10)
    absence_text.pack(fill='both', expand=True, padx=5, pady=5)
   
    scrollbar = ttk.Scrollbar(absence_frame, command=absence_text.yview)
    scrollbar.pack(side='right', fill='y')
    absence_text.configure(yscrollcommand=scrollbar.set)

    return tab

def setup_team_analysis_tab(notebook):
    """Configuración de la pestaña de análisis de equipos"""
    tab = ttk.Frame(notebook)
   
    control_frame = ttk.LabelFrame(tab, text="Controles de Análisis")
    control_frame.pack(fill='x', padx=10, pady=5)

    ttk.Button(
        control_frame,
        text="Analizar Composición de Equipos",
        command=analyze_team_composition,
        style='Primary.TButton'
    ).pack(pady=5, padx=5)

    # Frame para resultados
    global team_frame
    team_frame = ttk.LabelFrame(tab, text="Análisis de Equipos")
    team_frame.pack(fill='both', expand=True, padx=10, pady=5)

    # Canvas para gráficos
    global team_figure
    team_figure = Figure(figsize=(8, 6))
    global team_canvas
    team_canvas = FigureCanvasTkAgg(team_figure, master=team_frame)
    team_canvas.get_tk_widget().pack(fill='both', expand=True)

    return tab

def analyze_absence_patterns():
    try:
        df = pd.read_excel(excel_file_path)
        absence_text.delete(1.0, tk.END)
        absence_text.insert(tk.END, "=== ANÁLISIS DE AUSENTISMO Y CAMBIOS ===\n\n")

        # 1. Análisis de motivos de eliminación
        if 'motivo_eliminado' in df.columns:
            # Usar fillna para manejar valores nulos
            motivos = df['motivo_eliminado'].fillna('Sin motivo').value_counts()
            absence_text.insert(tk.END, "1. Distribución de Motivos de Ausencia:\n")
            for motivo, count in motivos.items():
                if motivo != 'Sin motivo':  # Excluir los registros sin motivo
                    absence_text.insert(tk.END, f"   • {motivo}: {count} casos\n")

        # 2. Análisis temporal de ausencias
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        df['mes'] = df['fecha_asignacion'].dt.month
        df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek

        # Ausencias por día de la semana
        absence_text.insert(tk.END, "\n2. Patrones por Día de la Semana:\n")
        for dia in range(7):
            dia_data = df[df['dia_semana'] == dia]
            # Contar solo los registros que tienen motivo_eliminado
            ausencias = dia_data['motivo_eliminado'].notna().sum()
            total = len(dia_data)
            if total > 0:
                porcentaje = (ausencias / total) * 100
                absence_text.insert(tk.END, f"   • {DIAS_SEMANA[dia]}: {porcentaje:.1f}% de ausencias\n")

        # 3. Análisis de impacto
        absence_text.insert(tk.END, "\n3. Análisis de Impacto:\n")
        # Usar nunique() para contar días únicos
        total_dias = df['fecha_asignacion'].dt.date.nunique()
        dias_con_ausencias = df[df['motivo_eliminado'].notna()]['fecha_asignacion'].dt.date.nunique()
       
        if total_dias > 0:
            porcentaje_dias = (dias_con_ausencias / total_dias) * 100
            absence_text.insert(tk.END, f"   • {porcentaje_dias:.1f}% de los días presentan ausencias\n")

        # 4. Recomendaciones
        absence_text.insert(tk.END, "\n4. Recomendaciones:\n")
       
        # Analizar días críticos (evitar SeriesGroupBy)
        dias_criticos = df[
            (df['dia_semana'].isin([0, 4])) &
            (df['motivo_eliminado'].notna())
        ].shape[0]
       
        if dias_criticos > 0 and total_dias > 0:
            porcentaje_criticos = (dias_criticos / total_dias) * 100
            if porcentaje_criticos > 30:
                absence_text.insert(tk.END, "   • Se detectó alta incidencia de ausencias en Lunes y Viernes\n")
                absence_text.insert(tk.END, "   • Recomendación: Implementar incentivos para estos días\n")

        # Analizar patrones recurrentes
        if 'nombre' in df.columns:
            # Usar value_counts() en lugar de groupby
            ausencias_por_tecnico = df[df['motivo_eliminado'].notna()]['nombre'].value_counts()
            if len(ausencias_por_tecnico) > 0:
                media = ausencias_por_tecnico.mean()
                std = ausencias_por_tecnico.std()
                if pd.notna(media) and pd.notna(std):
                    tecnicos_frecuentes = ausencias_por_tecnico[ausencias_por_tecnico > media + std]
                   
                    if not tecnicos_frecuentes.empty:
                        absence_text.insert(tk.END, "\n   • Técnicos con ausencias frecuentes detectados:\n")
                        for tecnico, ausencias in tecnicos_frecuentes.items():
                            absence_text.insert(tk.END, f"     - {tecnico}: {ausencias} ausencias\n")

        # 5. Análisis Mensual
        absence_text.insert(tk.END, "\n5. Análisis Mensual:\n")
        df['mes'] = df['fecha_asignacion'].dt.month
        ausencias_por_mes = df[df['motivo_eliminado'].notna()].groupby('mes').size()
       
        for mes, cantidad in ausencias_por_mes.items():
            if mes in MESES:
                absence_text.insert(tk.END, f"   • {MESES[mes]}: {cantidad} ausencias\n")

        absence_text.see(1.0)

    except Exception as e:
        print(f"Error detallado: {str(e)}")  # Para debug
        messagebox.showerror("Error", f"Error en el análisis de ausentismo: {str(e)}")

def analyze_team_composition():
    try:
        df = pd.read_excel(excel_file_path)
       
        # Convertir fechas a string para la visualización
        if 'fecha_asignacion' in df.columns:
            df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
            df['mes'] = df['fecha_asignacion'].dt.strftime('%Y-%m')

        # Crear la figura con especificaciones correctas para cada tipo de gráfico
        fig = make_subplots(
            rows=2,
            cols=2,
            specs=[
                [{"type": "domain"}, {"type": "xy"}],  # Primera fila: pie chart y heatmap
                [{"type": "xy"}, {"type": "xy"}]       # Segunda fila: bar charts
            ],
            subplot_titles=(
                'Distribución por Cargo',
                'Correlación entre Cargos',
                'Evolución Mensual de Composición',
                'Distribución de Turnos por Cargo'
            )
        )

        # 1. Gráfico de torta para distribución por cargo
        cargo_dist = df['cargo'].value_counts()
        fig.add_trace(
            go.Pie(
                labels=cargo_dist.index,
                values=cargo_dist.values,
                hole=0.3,
                name='Distribución por Cargo'
            ),
            row=1, col=1
        )

        # 2. Heatmap de correlación entre cargos
        pivot_cargo = pd.crosstab(df['fecha_asignacion'].dt.date, df['cargo'])
        corr_matrix = pivot_cargo.corr()
       
        fig.add_trace(
            go.Heatmap(
                z=corr_matrix.values,
                x=corr_matrix.columns,
                y=corr_matrix.columns,
                colorscale='RdBu',
                name='Correlación'
            ),
            row=1, col=2
        )

        # 3. Gráfico de barras apiladas para evolución mensual
        monthly_comp = df.groupby([df['fecha_asignacion'].dt.strftime('%Y-%m'), 'cargo']).size().unstack(fill_value=0)
        for cargo in monthly_comp.columns:
            fig.add_trace(
                go.Bar(
                    x=monthly_comp.index,
                    y=monthly_comp[cargo],
                    name=cargo,
                ),
                row=2, col=1
            )

        # 4. Distribución de turnos por cargo
        if 'turno' in df.columns:
            turno_cargo = pd.crosstab(df['cargo'], df['turno'])
            for turno in turno_cargo.columns:
                fig.add_trace(
                    go.Bar(
                        x=turno_cargo.index,
                        y=turno_cargo[turno],
                        name=f'Turno {turno}',
                    ),
                    row=2, col=2
                )

        # Actualizar el diseño
        fig.update_layout(
            height=800,
            showlegend=True,
            title_text="Análisis de Composición de Equipos",
            barmode='stack'
        )

        # Actualizar los ejes y títulos
        fig.update_xaxes(title_text="Fecha", row=2, col=1)
        fig.update_xaxes(title_text="Cargo", row=2, col=2)
        fig.update_yaxes(title_text="Cantidad", row=2, col=1)
        fig.update_yaxes(title_text="Cantidad", row=2, col=2)

        # Ajustar layout para mejor visualización
        fig.update_layout(
            margin=dict(t=50, l=50, r=50, b=50),
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )

        # Mostrar el gráfico en el navegador
        fig.show(renderer='browser')

        # Mostrar recomendaciones
        show_team_recommendations(df)

    except Exception as e:
        print(f"Error detallado: {str(e)}")  # Para debugging
        messagebox.showerror("Error", f"Error en el análisis de equipos: {str(e)}")

def show_team_recommendations(df):
    """Muestra recomendaciones basadas en el análisis de equipos"""
    recommendation_window = tk.Toplevel()
    recommendation_window.title("Recomendaciones de Equipo")
    recommendation_window.transient()
   
    # Centrar ventana
    window_width = 500
    window_height = 400
    screen_width = recommendation_window.winfo_screenwidth()
    screen_height = recommendation_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    recommendation_window.geometry(f'{window_width}x{window_height}+{x}+{y}')

    # Crear Text widget con scroll
    text_frame = ttk.Frame(recommendation_window)
    text_frame.pack(fill='both', expand=True, padx=10, pady=10)

    text_widget = tk.Text(text_frame, wrap=tk.WORD, padx=10, pady=10)
    scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
   
    scrollbar.pack(side='right', fill='y')
    text_widget.pack(side='left', fill='both', expand=True)

    # Análisis y recomendaciones
    text_widget.insert(tk.END, "=== RECOMENDACIONES DE COMPOSICIÓN DE EQUIPOS ===\n\n")

    try:
        # 1. Balance de habilidades
        cargos_count = df['cargo'].value_counts()
        text_widget.insert(tk.END, "1. Balance de Habilidades:\n")
        for cargo, count in cargos_count.items():
            percentage = (count / len(df)) * 100
            text_widget.insert(tk.END, f"   • {cargo}: {percentage:.1f}%\n")

        # 2. Análisis de continuidad
        text_widget.insert(tk.END, "\n2. Análisis de Continuidad:\n")
        if 'letra_rotacion' in df.columns and 'fecha_asignacion' in df.columns:
            rotation_patterns = pd.crosstab(
                df['fecha_asignacion'].dt.date,
                df['letra_rotacion']
            )
            continuity_score = rotation_patterns.std().mean()
            mean_value = rotation_patterns.mean().mean()
           
            if continuity_score > mean_value * 0.5:
                text_widget.insert(tk.END, "   • Alta variabilidad en la composición de equipos\n")
                text_widget.insert(tk.END, "   • Recomendación: Establecer equipos más estables\n")
            else:
                text_widget.insert(tk.END, "   • Buena estabilidad en la composición de equipos\n")

        # 3. Recomendaciones específicas
        text_widget.insert(tk.END, "\n3. Recomendaciones Específicas:\n")
       
        # Análisis de carga por cargo
        if 'fecha_asignacion' in df.columns:
            daily_cargo = pd.crosstab(df['fecha_asignacion'].dt.date, df['cargo'])
            cargo_stats = {
                'mean': daily_cargo.mean(),
                'std': daily_cargo.std()
            }
           
            for cargo in daily_cargo.columns:
                cv = cargo_stats['std'][cargo] / cargo_stats['mean'][cargo] if cargo_stats['mean'][cargo] > 0 else 0
                if cv > 0.5:
                    text_widget.insert(tk.END, f"   • Alta variabilidad en {cargo}\n")
                    text_widget.insert(tk.END, f"     Recomendación: Estabilizar asignación de {cargo}\n")

    except Exception as e:
        text_widget.insert(tk.END, f"\nError en el análisis: {str(e)}")

    text_widget.configure(state='disabled')

    # Botón de cerrar
    ttk.Button(
        recommendation_window,
        text="Cerrar",
        command=recommendation_window.destroy
    ).pack(pady=10)





def calculate_additional_metrics(df):
    """Calcula métricas adicionales y muestra un resumen"""
    try:
        # Crear ventana de resumen
        summary_window = tk.Toplevel()
        summary_window.title("Métricas Adicionales")
        summary_window.transient()
       
        # Centrar la ventana
        window_width = 500
        window_height = 600  # Definimos window_height
        screen_width = summary_window.winfo_screenwidth()
        screen_height = summary_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        summary_window.geometry(f'{window_width}x{window_height}+{x}+{y}')  # Usamos window_height

        # Crear Text widget con scroll
        text_widget = tk.Text(summary_window, wrap=tk.WORD, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(summary_window, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
       
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Convertir fecha_asignacion a datetime si no lo es
        if 'fecha_asignacion' in df.columns:
            df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])

        # Calcular métricas
        text_widget.insert(tk.END, "=== MÉTRICAS AVANZADAS ===\n\n")

        # 1. Estadísticas de carga
        text_widget.insert(tk.END, "1. Estadísticas de Carga:\n")
        daily_load = df.groupby(df['fecha_asignacion'].dt.date).size()
        text_widget.insert(tk.END, f"   • Promedio diario: {daily_load.mean():.2f} técnicos\n")
        text_widget.insert(tk.END, f"   • Máximo diario: {daily_load.max()} técnicos\n")
        text_widget.insert(tk.END, f"   • Mínimo diario: {daily_load.min()} técnicos\n")
        text_widget.insert(tk.END, f"   • Desviación estándar: {daily_load.std():.2f}\n\n")

        # 2. Análisis de rotación
        if 'letra_rotacion' in df.columns:
            text_widget.insert(tk.END, "2. Análisis de Rotación:\n")
            rotation_stats = df['letra_rotacion'].value_counts()
            total_rotations = len(df)
           
            for rot, count in rotation_stats.items():
                if pd.notna(rot):  # Solo procesar valores no nulos
                    percentage = (count / total_rotations) * 100
                    text_widget.insert(tk.END, f"   • {rot}: {count} ({percentage:.1f}%)\n")
            text_widget.insert(tk.END, "\n")

        # 3. Métricas de eficiencia
        text_widget.insert(tk.END, "3. Métricas de Eficiencia:\n")
       
        # Calcular balance de carga
        cv = daily_load.std() / daily_load.mean() if daily_load.mean() > 0 else 0
        text_widget.insert(tk.END, f"   • Coeficiente de variación: {cv:.3f}\n")
       
        # Evaluar distribución semanal
        if 'fecha_asignacion' in df.columns:
            df['dia_semana'] = df['fecha_asignacion'].dt.dayofweek
            weekly_dist = df.groupby('dia_semana').size()
            weekly_cv = weekly_dist.std() / weekly_dist.mean() if weekly_dist.mean() > 0 else 0
            text_widget.insert(tk.END, f"   • Balance semanal: {weekly_cv:.3f}\n\n")

        # 4. Recomendaciones basadas en métricas
        text_widget.insert(tk.END, "4. Recomendaciones:\n")
       
        if cv > 0.3:
            text_widget.insert(tk.END, "   • Alta variabilidad en la carga diaria. Se recomienda mejorar la distribución.\n")
       
        if weekly_cv > 0.2:
            text_widget.insert(tk.END, "   • Distribución semanal desbalanceada. Considerar redistribuir la carga.\n")
           
        if 'dia_semana' in df.columns:
            weekly_mean = weekly_dist.mean()
            low_coverage_days = weekly_dist[weekly_dist < weekly_mean * 0.8]
            if not low_coverage_days.empty:
                text_widget.insert(tk.END, "   • Días con baja cobertura detectados:\n")
                for day, count in low_coverage_days.items():
                    day_name = calendar.day_name[day]
                    text_widget.insert(tk.END, f"     - {day_name}: {count} técnicos\n")

        # Agregar botón de cerrar
        close_button = ttk.Button(
            summary_window,
            text="Cerrar",
            command=summary_window.destroy
        )
        close_button.pack(pady=10)

        # Deshabilitar edición del texto
        text_widget.configure(state='disabled')

    except Exception as e:
        print(f"Error detallado en calculate_additional_metrics: {str(e)}")  # Para debugging
        messagebox.showerror("Error", f"Error al calcular métricas adicionales: {str(e)}")

def calculate_advanced_metrics():
    try:
        df = pd.read_excel(excel_file_path)
       
        if df.empty:
            messagebox.showinfo("Información", "No hay datos para calcular métricas.")
            return

        # 1. Gráfico de distribución temporal
        daily_counts = df.groupby('fecha_asignacion').size().reset_index(name='counts')
        fig1 = px.scatter(daily_counts, x='fecha_asignacion', y='counts', title='Distribución Temporal')

        # 2. Gráfico de distribución por rotación
        fig2 = px.pie(df, names='letra_rotacion', title='Distribución por Rotación')

        # 3. Gráfico de carga por día de la semana
        df['dia_semana'] = df['fecha_asignacion'].dt.day_name()
        weekly_data = df.groupby('dia_semana').size().reset_index(name='counts')
        fig3 = px.bar(weekly_data, x='dia_semana', y='counts', title='Carga por Día de Semana')

        # 4. Gráfico de tendencias de carga
        monthly_data = df.groupby(df['fecha_asignacion'].dt.to_period('M')).size().reset_index(name='counts')
        monthly_data['mes'] = monthly_data['fecha_asignacion'].dt.strftime('%Y-%m')
        fig4 = px.line(monthly_data, x='mes', y='counts', title='Tendencia Mensual')

        # Crear subplots
        fig = make_subplots(rows=2, cols=2, specs=[[{'type': 'xy'}, {'type': 'domain'}],
                                                    [{'type': 'xy'}, {'type': 'xy'}]],
                            subplot_titles=('Distribución Temporal', 'Distribución por Rotación',
                                            'Carga por Día de Semana', 'Tendencia Mensual'))

        fig.add_trace(fig1.data[0], row=1, col=1)
        fig.add_trace(fig2.data[0], row=1, col=2)
        fig.add_trace(fig3.data[0], row=2, col=1)
        fig.add_trace(fig4.data[0], row=2, col=2)

        # Actualizar diseño
        fig.update_layout(height=600, showlegend=False)

        # Mostrar gráfico
        fig.show(renderer='browser')

        # Calcular y mostrar métricas adicionales
        calculate_additional_metrics(df)

    except Exception as e:
        messagebox.showerror("Error", f"Error al calcular métricas: {str(e)}")




def get_current_date():
    """Retorna el año, mes y día actual"""
    current_date = datetime.now()
    return {
        'year': current_date.year,
        'month': current_date.month,
        'day': current_date.day
    }


class MotivoEliminacionDialog:
    def __init__(self, parent):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Motivo de Eliminación")
        self.dialog.transient(parent)
        self.dialog.grab_set()
       
        # Centrar la ventana
        window_width = 400
        window_height = 400
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')
       
        self.result = None
        self.motivo = None
       
        # Frame principal con padding
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
       
        # Label de instrucción
        ttk.Label(main_frame,
                 text="Seleccione el motivo de eliminación:",
                 font=('Segoe UI', 10, 'bold')).pack(pady=(0, 10))
       
        # Variable para el radiobutton
        self.selected_option = tk.StringVar(value="")  # Inicializar con valor vacío
       
        # Opciones
        self.options = [
            "Licencia Médica",
            "Cambio de turno",
            "Capacitación",
            "Permisos especiales",
            "Feriado Legal"
        ]
       
        # Frame para los radiobuttons
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(fill=tk.X, pady=10)
       
        for option in self.options:
            rb = ttk.Radiobutton(radio_frame,
                               text=option,
                               value=option,
                               variable=self.selected_option)
            rb.pack(anchor=tk.W, pady=5)
            rb.configure(command=self.on_selection_change)  # Agregar comando después de crear el radiobutton
       
        # Frame para el campo de texto de permisos especiales
        self.permiso_frame = ttk.LabelFrame(main_frame, text="Detalle del permiso especial", padding=(10, 5))
       
        self.permiso_text = ttk.Entry(self.permiso_frame, width=40)
        self.permiso_text.pack(fill=tk.X, pady=5, padx=5)
       
        # Inicialmente ocultar el frame de permisos
        self.permiso_frame.pack_forget()
       
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
       
        # Botones con estilos
        ttk.Button(button_frame,
                  text="Aceptar",
                  style='Primary.TButton',
                  command=self.on_accept).pack(side=tk.LEFT, expand=True, padx=5)
       
        ttk.Button(button_frame,
                  text="Cancelar",
                  command=self.on_cancel).pack(side=tk.LEFT, expand=True, padx=5)

    def on_selection_change(self):
        selected = self.selected_option.get()
        if selected == "Permisos especiales":
            self.permiso_frame.pack(fill=tk.X, pady=(10, 0))
            self.permiso_text.focus_set()
            # Ajustar el tamaño de la ventana para asegurar que todo sea visible
            self.dialog.update()
            self.dialog.geometry("")  # Esto hace que la ventana se ajuste al contenido
        else:
            self.permiso_frame.pack_forget()
            # Ajustar el tamaño de la ventana nuevamente
            self.dialog.update()
            self.dialog.geometry("")

    def on_accept(self):
        if not self.selected_option.get():
            messagebox.showwarning("Advertencia", "Por favor seleccione un motivo")
            return
           
        if self.selected_option.get() == "Permisos especiales":
            if not self.permiso_text.get().strip():
                messagebox.showwarning("Advertencia", "Por favor especifique el motivo del permiso especial")
                return
            self.motivo = f"Permisos especiales: {self.permiso_text.get().strip()}"
        else:
            self.motivo = self.selected_option.get()
           
        self.result = True
        self.dialog.destroy()

    def on_cancel(self):
        self.result = False
        self.dialog.destroy()
       
       
       
class MotivoAgregadoDialog:
    def __init__(self, parent):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Motivo de Agregado")
        self.dialog.transient(parent)
        self.dialog.grab_set()
       
        # Configurar ventana
        window_width = 400
        window_height = 300  # Definimos window_height
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')  # Usamos window_height
        self.dialog.resizable(False, False)
       
        self.result = None
        self.motivo = None
       
        # Crear frame principal
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
       
        # Título
        ttk.Label(main_frame,
                 text="Seleccione el motivo de agregado:",
                 font=('Segoe UI', 12, 'bold')).pack(pady=(0, 20))
       
        # Variable para el radiobutton
        self.selected_option = tk.StringVar(value="")
       
        # Frame para los radiobuttons
        radio_frame = ttk.LabelFrame(main_frame, text="Opciones disponibles", padding=(20, 10))
        radio_frame.pack(fill=tk.X, padx=10, pady=10)
       
        # Opciones
        options = [
            "Hora Extra",
            "Cambio de turno"
        ]
       
        # Crear los radiobuttons
        for option in options:
            ttk.Radiobutton(radio_frame,
                          text=option,
                          value=option,
                          variable=self.selected_option).pack(anchor=tk.W,
                                                           padx=10,
                                                           pady=10)
       
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
       
        # Botones
        ttk.Button(button_frame,
                  text="Aceptar",
                  command=self.on_accept,
                  style='Primary.TButton',
                  width=15).pack(side=tk.LEFT, padx=5, expand=True)
       
        ttk.Button(button_frame,
                  text="Cancelar",
                  command=self.on_cancel,
                  width=15).pack(side=tk.LEFT, padx=5, expand=True)
       
        # Hacer la ventana modal
        self.dialog.focus_force()
        self.dialog.wait_visibility()
        self.dialog.grab_set()

    def on_accept(self):
        if not self.selected_option.get():
            messagebox.showwarning("Advertencia",
                                 "Por favor seleccione un motivo",
                                 parent=self.dialog)
            return
           
        self.motivo = self.selected_option.get()
        self.result = True
        self.dialog.destroy()

    def on_cancel(self):
        self.result = False
        self.dialog.destroy()


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f'{width}x{height}+{x}+{y}')
    window.minsize(width, height)

def ensure_excel_file():
    if not os.path.exists(excel_file_path):
        # Crear un ExcelWriter
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            # Crear la hoja principal
            df_main = pd.DataFrame(columns=['turno', 'nombre', 'cargo', 'letra_rotacion', 'fecha_asignacion'])
            df_main.to_excel(writer, sheet_name='Sheet1', index=False)
           
            # Crear Sheet2 con turno_global
            df_turnos = pd.DataFrame({'turno_global': ['Mañana', 'Tarde', 'Noche']})  # Valores por defecto
            df_turnos.to_excel(writer, sheet_name='Sheet2', index=False)
    return True




def setup_technician_display(tab, user_type):
    global data_display, data_display_corrective, data_display_preventive, table_frame, warning_label

    # Frame principal para las tablas
    table_frame = ttk.Frame(tab)
    table_frame.grid(row=0, column=1, sticky='nsew', padx=10, pady=10)
    table_frame.grid_columnconfigure(0, weight=1)

    # === Warning Label al inicio ===
    warning_label = ttk.Label(
        table_frame,
        text="",
        font=('Segoe UI', 11, 'bold'),
        justify='center',
        wraplength=400
    )
    warning_label.pack(pady=5)

    # === Tabla Correctivo ===
    corrective_frame = ttk.LabelFrame(table_frame, text="Técnicos Programados Correctivo")
    corrective_frame.pack(fill='both', expand=True, padx=5, pady=5)

    data_display_corrective = ttk.Treeview(
        corrective_frame,
        columns=('nombre', 'cargo', 'letra_rotacion', 'fecha', 'jornada', 'motivo'),
        show='headings',
        height=10
    )

    # Configurar columnas del correctivo
    columns = [
        ('nombre', 'Nombre', 150),
        ('cargo', 'Cargo', 120),
        ('letra_rotacion', 'Rotación', 80),
        ('fecha', 'Fecha', 100),
        ('jornada', 'Jornada', 80),
        ('motivo', 'Motivo', 150)
    ]

    for col, heading, width in columns:
        data_display_corrective.heading(col, text=heading)
        data_display_corrective.column(col, width=width)

    # Scrollbars para correctivos
    scrollbar_y_corrective = ttk.Scrollbar(corrective_frame, orient="vertical", command=data_display_corrective.yview)
    scrollbar_x_corrective = ttk.Scrollbar(corrective_frame, orient="horizontal", command=data_display_corrective.xview)
    data_display_corrective.configure(yscrollcommand=scrollbar_y_corrective.set, xscrollcommand=scrollbar_x_corrective.set)

    # Grid para correctivos
    data_display_corrective.grid(row=0, column=0, sticky='nsew')
    scrollbar_y_corrective.grid(row=0, column=1, sticky='ns')
    scrollbar_x_corrective.grid(row=1, column=0, sticky='ew')
    corrective_frame.grid_columnconfigure(0, weight=1)
    corrective_frame.grid_rowconfigure(0, weight=1)

    # === Tabla Preventivo ===
    preventive_frame = ttk.LabelFrame(table_frame, text="Técnicos Programados Preventivo")
    preventive_frame.pack(fill='both', expand=True, padx=5, pady=5)

    data_display_preventive = ttk.Treeview(
        preventive_frame,
        columns=('nombre', 'cargo', 'letra_rotacion', 'fecha', 'jornada', 'motivo'),
        show='headings',
        height=10
    )

    # Configurar columnas del preventivo
    for col, heading, width in columns:
        data_display_preventive.heading(col, text=heading)
        data_display_preventive.column(col, width=width)

    # Scrollbars para preventivos
    scrollbar_y_preventive = ttk.Scrollbar(preventive_frame, orient="vertical", command=data_display_preventive.yview)
    scrollbar_x_preventive = ttk.Scrollbar(preventive_frame, orient="horizontal", command=data_display_preventive.xview)
    data_display_preventive.configure(yscrollcommand=scrollbar_y_preventive.set, xscrollcommand=scrollbar_x_preventive.set)

    # Grid para preventivos
    data_display_preventive.grid(row=0, column=0, sticky='nsew')
    scrollbar_y_preventive.grid(row=0, column=1, sticky='ns')
    scrollbar_x_preventive.grid(row=1, column=0, sticky='ew')
    preventive_frame.grid_columnconfigure(0, weight=1)
    preventive_frame.grid_rowconfigure(0, weight=1)

    # Configurar estilos para filas alternadas
    style = ttk.Style()
    style.configure("Treeview", background="white", foreground="black", fieldbackground="white")
    style.map('Treeview', background=[('selected', '#0078D7')])

    return table_frame
   

def setup_date_range_display(tab):
    """Configura la tabla para mostrar técnicos en el rango de fechas seleccionado"""
    table_frame = ttk.Frame(tab, padding="10")
    table_frame.grid(row=0, column=1, sticky='nsew')
    table_frame.grid_columnconfigure(0, weight=1)
    table_frame.grid_rowconfigure(1, weight=1)
   
    ttk.Label(table_frame, text="Técnicos en Rango de Fechas",
             font=('Segoe UI', 12, 'bold')).grid(row=0, column=0, pady=(0, 10))
   
    global range_display
    range_display = ttk.Treeview(table_frame,
                                columns=('nombre', 'cargo', 'letra_rotacion', 'fecha'),
                                show='headings')
   
    range_display.heading('nombre', text='Nombre')
    range_display.heading('cargo', text='Cargo')
    range_display.heading('letra_rotacion', text='Rotación')
    range_display.heading('fecha', text='Fecha')
   
    # Configurar anchos de columna
    range_display.column('nombre', width=150, minwidth=100)
    range_display.column('cargo', width=150, minwidth=100)
    range_display.column('letra_rotacion', width=100, minwidth=80)
    range_display.column('fecha', width=150, minwidth=100)
   
    # Configurar scrollbars
    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=range_display.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=range_display.xview)
    range_display.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
   
    # Posicionar elementos en la grilla
    range_display.grid(row=1, column=0, sticky='nsew')
    vsb.grid(row=1, column=1, sticky='ns')
    hsb.grid(row=2, column=0, sticky='ew')
   
    # Configurar estilo y eventos
    range_display.tag_configure('oddrow', background='#f0f0f0')
    range_display.tag_configure('evenrow', background='white')
   
    def alternating_row_colors():
        for i, item in enumerate(range_display.get_children()):
            if i % 2 == 0:
                range_display.item(item, tags=('evenrow',))
            else:
                range_display.item(item, tags=('oddrow',))
               
    range_display.bind('<<TreeviewSelect>>', lambda e: alternating_row_colors())    
   


def delete_technician(user_type):
    if user_type == "programador":
        messagebox.showerror("Error", "Su perfil no tiene permisos para realizar esta acción.")
        return

    # Verificar selección en ambas tablas
    selected_corrective = data_display_corrective.selection()
    selected_preventive = data_display_preventive.selection()

    # Si no hay selección en ninguna tabla
    if not selected_corrective and not selected_preventive:
        messagebox.showerror("Error", "Por favor, selecciona un técnico para eliminar.")
        return

    # Determinar de qué tabla se está eliminando
    if selected_corrective:
        selected_items = selected_corrective
        current_display = data_display_corrective
        tipo_tecnico = "correctivo"
    else:
        selected_items = selected_preventive
        current_display = data_display_preventive
        tipo_tecnico = "preventivo"

    # Mostrar diálogo de motivo de eliminación
    dialog = MotivoEliminacionDialog(current_display.winfo_toplevel())
    dialog.dialog.wait_window()

    if not dialog.result:
        return

    try:
        # Leer el Excel actual
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])

        for item in selected_items:
            values = current_display.item(item)['values']
            nombre = values[0]
            cargo = values[1]
            fecha_str = values[3]
            fecha = datetime.strptime(fecha_str, '%d/%m/%Y')

            # Encontrar la fila que coincide
            mask = (df['nombre'] == nombre) & (df['fecha_asignacion'].dt.date == fecha.date())
           
            # Guardar la información de eliminación
            df.loc[mask, 'nombre_eliminado'] = nombre
            df.loc[mask, 'cargo_eliminado'] = cargo
            df.loc[mask, 'fecha_eliminado'] = fecha
            df.loc[mask, 'motivo_eliminado'] = dialog.motivo

            # Eliminar el registro original
            df.loc[mask, 'nombre'] = ''
            df.loc[mask, 'cargo'] = ''
            df.loc[mask, 'letra_rotacion'] = ''
            
            # Registrar la acción en el log
            log_entry = pd.DataFrame({
                'nombre_log': [user_type],
                'fecha_log': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                'accion_log': [f"Eliminó técnico {tipo_tecnico}: {nombre}"]
            })
            df = pd.concat([df, log_entry], ignore_index=True)

        # Guardar los cambios en el Excel
        df.to_excel(excel_file_path, index=False)
       
        # Actualizar ambas tablas
        show_scheduled_technicians()
       
        messagebox.showinfo("Éxito",
                          f"Técnico(s) {tipo_tecnico}(s) eliminado(s) exitosamente.\n"
                          f"Motivo registrado: {dialog.motivo}")

    except Exception as e:
        messagebox.showerror("Error", f"Error al eliminar el técnico: {str(e)}")
        print(f"Error detallado: {str(e)}")



def filter_technicians(calendars, jornada_var):
    """Filtra y muestra técnicos entre las fechas seleccionadas con filtro de jornada"""
    try:
        # Obtener fechas
        date_str_inicio = calendars['cal_inicio'].get_date()
        date_str_fin = calendars['cal_fin'].get_date()

        # Convertir fechas
        start_parts = date_str_inicio.split('-')
        end_parts = date_str_fin.split('-')
       
        start_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
        end_date = datetime(
            year=int(end_parts[2]),
            month=int(end_parts[1]),
            day=int(end_parts[0])
        )

        if start_date > end_date:
            messagebox.showerror("Error", "La fecha inicial debe ser anterior a la fecha final")
            return
           
        # Leer datos
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        # Aplicar filtros básicos
        mask = (
            (df['fecha_asignacion'].dt.date >= start_date.date()) &
            (df['fecha_asignacion'].dt.date <= end_date.date()) &
            # Filtrar registros con nombre y cargo válidos
            df['nombre'].notna() &
            (df['nombre'] != '') &
            df['cargo'].notna() &
            (df['cargo'] != '') &
            # Filtrar registros con letra de rotación válida
            df['letra_rotacion'].notna() &
            (df['letra_rotacion'] != '')
        )
       
        # Aplicar filtro de jornada si no es "todas"
        jornada_seleccionada = jornada_var.get()
        if jornada_seleccionada != "todas":
            mask = mask & (df['Jornada'] == jornada_seleccionada)
           
        df_filtered = df[mask]
       
        # Limpiar tabla actual
        for item in range_display.get_children():
            range_display.delete(item)
           
        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay técnicos programados en ese rango de fechas")
            return
           
        # Mostrar resultados filtrados
        for i, row in df_filtered.iterrows():
            values = (
                row['nombre'],
                row['cargo'],
                row['letra_rotacion'],
                row['fecha_asignacion'].strftime('%d-%m-%Y'),
                row.get('Jornada', '')
            )
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            range_display.insert('', tk.END, values=values, tags=(tag,))

    except Exception as e:
        print(f"Error en filter_technicians: {str(e)}")
        messagebox.showerror("Error", f"Error al filtrar técnicos: {str(e)}")

def export_to_excel(calendars, jornada_var):
    """Exporta los datos filtrados a Excel"""
    try:
        # Obtener fechas y jornada seleccionada
        date_str_inicio = calendars['cal_inicio'].get_date()
        date_str_fin = calendars['cal_fin'].get_date()
        jornada_seleccionada = jornada_var.get()

        # Convertir fechas
        start_parts = date_str_inicio.split('-')
        end_parts = date_str_fin.split('-')
       
        start_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
        end_date = datetime(
            year=int(end_parts[2]),
            month=int(end_parts[1]),
            day=int(end_parts[0])
        )
       
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        # Aplicar filtros
        mask = (df['fecha_asignacion'].dt.date >= start_date.date()) & \
               (df['fecha_asignacion'].dt.date <= end_date.date())
               
        if jornada_seleccionada != "todas":
            mask = mask & (df['Jornada'] == jornada_seleccionada)
           
        df_filtered = df[mask]
       
        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
           
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte"
        )
       
        if file_path:
            # Crear DataFrame con las columnas deseadas
            df_export = df_filtered[[
                'nombre',
                'cargo',
                'letra_rotacion',
                'fecha_asignacion',
                'Jornada'  # Incluir jornada
            ]].copy()
           
            # Dar formato a la fecha
            df_export['fecha_asignacion'] = df_export['fecha_asignacion'].dt.strftime('%d-%m-%Y')
           
            # Renombrar columnas para el Excel
            df_export.columns = ['Nombre', 'Cargo', 'Rotación', 'Fecha', 'Jornada']
           
            # Exportar a Excel
            df_export.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "Datos exportados exitosamente")
           
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar datos: {str(e)}")

# También vamos a asegurarnos de que las fechas se manejen correctamente en la conversión
def convert_date_string(date_str):
    """Convierte una cadena de fecha al objeto datetime"""
    parts = date_str.split('-')
    return datetime(
        year=int(parts[2]),
        month=int(parts[1]),
        day=int(parts[0])
    )

def export_to_excel(calendars):
    try:
        # Convertir las fechas usando el formato correcto
        date_str_inicio = calendars['cal_inicio'].get_date()
        date_str_fin = calendars['cal_fin'].get_date()

        # Convertir las fechas usando datetime
        start_parts = date_str_inicio.split('-')
        end_parts = date_str_fin.split('-')
       
        start_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
        end_date = datetime(
            year=int(end_parts[2]),
            month=int(end_parts[1]),
            day=int(end_parts[0])
        )
       
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        mask = (df['fecha_asignacion'].dt.date >= start_date.date()) & \
               (df['fecha_asignacion'].dt.date <= end_date.date())
        df_filtered = df[mask]
       
        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
           
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte"
        )
       
        if file_path:
            df_filtered.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "Datos exportados exitosamente")
           
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar datos: {str(e)}")
       


# Modificar el diccionario USERS para incluir el nuevo usuario admin
USERS = {
    "sta": {"password": "sta123", "role": "supervisor"},
    "stb": {"password": "stb123", "role": "supervisor"},
    "stc": {"password": "stc123", "role": "supervisor"},
    "stn": {"password": "stn123", "role": "supervisor"},
    "hcifuentes": {"password": "hcifuentes123", "role": "supervisor"},
    "aprado": {"password": "aprado123", "role": "supervisor"},
    "programador": {"password": "programador123", "role": "programador"},
    "programador2": {"password": "1234", "role": "programador"},
    "atecnico": {"password": "tecnico123", "role": "programador"},
    "admin": {"password": "admin", "role": "admin"}  # Nuevo usuario admin
}

def check_login():
    user = user_entry.get().strip()
    password = password_entry.get().strip()

    if user in USERS and USERS[user]["password"] == password:
        log_action(user, "Inicio de sesión")  # Registrar el inicio de sesión
        open_main_window(USERS[user]["role"])
    else:
        messagebox.showerror("Error", "Usuario o contraseña incorrectos.")
        user_entry.delete(0, tk.END)
        password_entry.delete(0, tk.END)
        user_entry.focus()
       



def open_login_window():
    global login_window
    login_window = tk.Tk()
    login_window.title("Metro de Santiago - Sistema de Control de Turnos")
    configure_styles()
   
    # Configurar el estilo para el marco principal
    style = ttk.Style()
    style.configure('Login.TFrame', background='#f8f9fa')
    style.configure('Header.TFrame', background='#ffffff')
    style.configure('LoginForm.TFrame', background='#ffffff')
   
    # Configurar estilos para los widgets
    style.configure('Header.TLabel',
                   background='#ffffff',
                   font=('Segoe UI', 12, 'bold'))
    style.configure('Subheader.TLabel',
                   background='#ffffff',
                   font=('Segoe UI', 10),
                   foreground='#666666')
   
    # Frame principal
    main_frame = ttk.Frame(login_window, style='Login.TFrame', padding=20)
    main_frame.grid(row=0, column=0, sticky='nsew')
   
    # Configurar grid
    login_window.grid_columnconfigure(0, weight=1)
    login_window.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)
   
    # Frame del encabezado
    header_frame = ttk.Frame(main_frame, style='Header.TFrame', padding=10)
    header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
   
    # Logo frame
    logo_frame = ttk.Frame(header_frame, style='Header.TFrame')
    logo_frame.pack(pady=(0, 15))
   
    # Crear canvas para el logo
    logo_canvas = tk.Canvas(logo_frame, width=240, height=180, bg='white', highlightthickness=0)
    logo_canvas.pack()
   
    # Dibujar el logo directamente en el canvas
    # Óvalo exterior
    logo_canvas.create_oval(40, 45, 200, 135, width=12, outline='#333333')
   
    # Rombos rojos
    rombo_color = '#EE1C25'
    # Primer rombo
    logo_canvas.create_polygon(60,90, 85,65, 110,90, 85,115, fill=rombo_color)
    # Segundo rombo
    logo_canvas.create_polygon(95,90, 120,65, 145,90, 120,115, fill=rombo_color)
    # Tercer rombo
    logo_canvas.create_polygon(130,90, 155,65, 180,90, 155,115, fill=rombo_color)
   
    # Títulos
    ttk.Label(header_frame,
             text="Metro de Santiago",
             style='Header.TLabel',
             font=('Segoe UI', 16, 'bold')).pack()
   
    ttk.Label(header_frame,
             text="Señalización y Pilotaje Automático",
             style='Subheader.TLabel',
             font=('Segoe UI', 12)).pack()
   
    # Frame del formulario
    form_frame = ttk.Frame(main_frame, style='LoginForm.TFrame', padding=20)
    form_frame.grid(row=1, column=0, sticky='n', padx=50)
   
    # Título del formulario
    ttk.Label(form_frame,
             text="Iniciar Sesión",
             style='Header.TLabel').grid(row=0, column=0, columnspan=2, pady=(0, 20))
   
    # Campo de usuario
    ttk.Label(form_frame,
             text="Usuario:",
             font=('Segoe UI', 10)).grid(row=1, column=0, sticky='w', pady=(0, 5))
    global user_entry
    user_entry = ttk.Entry(form_frame, width=30)
    user_entry.grid(row=2, column=0, sticky='ew', pady=(0, 15))
   
    # Campo de contraseña
    ttk.Label(form_frame,
             text="Contraseña:",
             font=('Segoe UI', 10)).grid(row=3, column=0, sticky='w', pady=(0, 5))
    global password_entry
    password_entry = ttk.Entry(form_frame, show="•", width=30)
    password_entry.grid(row=4, column=0, sticky='ew', pady=(0, 20))
   
    # Botón de inicio de sesión
    login_button = ttk.Button(form_frame,
                            text="Ingresar",
                            command=check_login,
                            style='Primary.TButton')
    login_button.grid(row=5, column=0, sticky='ew', pady=(0, 10))
   
    # Versión y copyright
    ttk.Label(main_frame,
             text="Sistema de Control de Turnos v1.3",
             font=('Segoe UI', 8),
             foreground='#666666').grid(row=2, column=0, pady=(20, 0))
    ttk.Label(main_frame,
             text="© 2024 Metro de Santiago",
             font=('Segoe UI', 8),
             foreground='#666666').grid(row=3, column=0)
   
    # Centrar y establecer tamaño mínimo
    center_window(login_window, 500, 700)
    login_window.minsize(500, 700)
   
    # Dar foco al campo de usuario
    user_entry.focus()
   
    login_window.mainloop()

def open_main_window(user_type):
    """Abre la ventana principal con todas las pestañas y funcionalidades"""
    # Cerrar ventana de login y declarar variables globales
    login_window.destroy()
    global ai_tab, resumen_tab

    # Crear ventana principal
    main_window = tk.Tk()
    main_window.title("Control de Turnos - Metro de Santiago")
   
    # Configurar estilos modernos
    COLORS = configure_modern_styles()
   
    # Crear gestor de notificaciones
    notifications = NotificationManager(main_window)
   
    # Configurar la ventana principal
    main_window.grid_columnconfigure(0, weight=1)
    main_window.grid_rowconfigure(0, weight=1)
   
    # Crear el notebook con estilo moderno
    notebook = ttk.Notebook(main_window)
    notebook.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
   
    # Configurar las pestañas principales
    tabs = {
        'supervisor_tab': ttk.Frame(notebook, style='Card.TFrame', padding=10),
        'ausencias_tab': ttk.Frame(notebook, style='Card.TFrame', padding=10),
        'programmer_tab': ttk.Frame(notebook, style='Card.TFrame', padding=10),
        'resumen_tab': ttk.Frame(notebook, style='Card.TFrame', padding=10),
        'ai_tab': ttk.Frame(notebook, style='Card.TFrame', padding=10)
    }

    # Configurar cada pestaña
    setup_supervisor_tab(tabs['supervisor_tab'], user_type)
    setup_ausencias_tab(tabs['ausencias_tab'])
    setup_programmer_tab(tabs['programmer_tab'])
    setup_resumen_tab(tabs['resumen_tab'])
    setup_ai_analysis_tab(tabs['ai_tab'])
   
    # Guardar referencias globales
    ai_tab = tabs['ai_tab']
    resumen_tab = tabs['resumen_tab']

    # Agregar pestañas al notebook con íconos
    notebook.add(tabs['supervisor_tab'], text='📊 Gestión de Turnos')
    notebook.add(tabs['ausencias_tab'], text='👥 Ausencias')
    notebook.add(tabs['programmer_tab'], text='📝 Consulta y Reportes')
    notebook.add(tabs['resumen_tab'], text='📈 Resumen')
    notebook.add(tabs['ai_tab'], text='🤖 Análisis Avanzado')

    # Agregar pestaña de administración solo para usuarios admin
    if user_type == "admin":
        admin_tab = ttk.Frame(notebook, style='Card.TFrame', padding=10)
        setup_admin_tab(admin_tab)
        notebook.add(admin_tab, text='⚙️ Centro de Administración')

    # Configurar barra de estado
    status_frame = ttk.Frame(main_window)
    status_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=5)

    # Mostrar información del usuario
    user_label = ttk.Label(
        status_frame,
        text=f"Usuario: {user_type.capitalize()}",
        font=('Segoe UI', 9)
    )
    user_label.pack(side='left', padx=5)

    # Mostrar versión
    version_label = ttk.Label(
        status_frame,
        text="v1.3",
        font=('Segoe UI', 9)
    )
    version_label.pack(side='right', padx=5)

    # Mostrar fecha y hora
    def update_datetime():
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        datetime_label.config(text=current_time)
        main_window.after(1000, update_datetime)

    datetime_label = ttk.Label(
        status_frame,
        font=('Segoe UI', 9)
    )
    datetime_label.pack(side='right', padx=20)
    update_datetime()

    # Mensaje de bienvenida con notificación
    notifications.show_notification(
        f"Bienvenido al sistema de Control de Turnos",
        type_='success',
        duration=3000
    )

    # Agregar tooltips a elementos importantes
    add_tooltips(notebook)

    # Agregar menú principal
    menubar = tk.Menu(main_window)
    main_window.config(menu=menubar)

    # Menú Archivo
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Archivo", menu=file_menu)
    file_menu.add_command(label="Exportar Datos", command=lambda: export_data())
    file_menu.add_separator()
    file_menu.add_command(label="Salir", command=main_window.quit)

    # Menú Herramientas
    tools_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Herramientas", menu=tools_menu)
    tools_menu.add_command(label="Preferencias", command=lambda: show_preferences(main_window))
    if user_type == "admin":
        tools_menu.add_command(label="Administración", command=lambda: show_admin_tools(main_window))

    # Menú Ayuda
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Ayuda", menu=help_menu)
    help_menu.add_command(label="Manual de Usuario", command=show_user_manual)
    help_menu.add_command(label="Acerca de", command=lambda: show_about(main_window))

    # Configurar protocolo de cierre
    def on_closing():
        if messagebox.askokcancel("Salir", "¿Desea salir del programa?"):
            cleanup_temp_files()
            main_window.destroy()

    main_window.protocol("WM_DELETE_WINDOW", on_closing)

    # Centrar y configurar tamaño de ventana
    center_window(main_window, 1024, 768)
    main_window.minsize(1024, 768)

    # Bindear teclas de acceso rápido
    main_window.bind('<Control-q>', lambda e: on_closing())
    main_window.bind('<F5>', lambda e: refresh_data())
    main_window.bind('<Control-h>', lambda e: show_user_manual())

    # Iniciar mainloop
    main_window.mainloop()

def show_preferences(parent):
    """Muestra ventana de preferencias"""
    preferences_window = tk.Toplevel(parent)
    preferences_window.title("Preferencias")
    preferences_window.transient(parent)
    preferences_window.grab_set()
   
    center_window(preferences_window, 400, 300)
    # Aquí agregarías las opciones de preferencias

def show_admin_tools(parent):
    """Muestra herramientas administrativas"""
    admin_window = tk.Toplevel(parent)
    admin_window.title("Herramientas de Administración")
    admin_window.transient(parent)
    admin_window.grab_set()
   
    center_window(admin_window, 500, 400)
    # Aquí agregarías las herramientas administrativas

def show_user_manual():
    """Muestra el manual de usuario"""
    messagebox.showinfo(
        "Manual de Usuario",
        "El manual de usuario se abrirá en su navegador web predeterminado."
    )
    # Aquí agregarías el código para abrir el manual

def show_about(parent):
    """Muestra información sobre la aplicación"""
    about_window = tk.Toplevel(parent)
    about_window.title("Acerca de")
    about_window.transient(parent)
    about_window.grab_set()
   
    center_window(about_window, 400, 300)
   
    ttk.Label(
        about_window,
        text="Sistema de Control de Turnos",
        font=('Segoe UI', 14, 'bold')
    ).pack(pady=20)
   
    ttk.Label(
        about_window,
        text="Versión 1.3",
        font=('Segoe UI', 10)
    ).pack(pady=10)
   
    ttk.Label(
        about_window,
        text="© 2024 Metro de Santiago",
        font=('Segoe UI', 10)
    ).pack(pady=10)

def refresh_data():
    """Actualiza los datos en todas las pestañas"""
    show_scheduled_technicians()
    if hasattr(resumen_tab, 'update_resumen_graph'):
        update_resumen_graph()

def add_tooltips(notebook):
    """Agrega tooltips a elementos importantes de la interfaz"""
    tooltips = {
        'Gestión de Turnos': 'Administre y visualice los turnos del personal',
        'Ausencias': 'Gestione y monitoree las ausencias del personal',
        'Consulta y Reportes': 'Genere informes y consulte históricos',
        'Resumen': 'Visualice estadísticas y métricas clave',
        'Análisis Avanzado': 'Acceda a análisis predictivo y tendencias'
    }
   
    for tab in notebook.tabs():
        tab_text = notebook.tab(tab, 'text').split(' ')[-1]  # Eliminar emoji
        if tab_text in tooltips:
            ModernTooltip(notebook, tooltips[tab_text])

def setup_supervisor_tab(tab: ttk.Frame, user_type: str) -> None:
    """Configura la pestaña de supervisor"""
    # Declarar variables globales
    global turno_combobox, name_combobox, cargo_entry, rotation_letter_entry
    global cal, jornada_var, btn_mañana, btn_tarde, btn_noche
   
    # Configuración inicial del grid
    tab.grid_columnconfigure(1, weight=3)
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(0, weight=1)

    # Configurar estilos
    style = ttk.Style()
   
    # Configuración de estilos para los botones de jornada
    style.configure('Morning.TButton',
        background='#ffc107',
        foreground='black',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )
    style.configure('Selected.Morning.TButton',
        background='#ff9800',
        foreground='black',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )
   
    style.configure('Afternoon.TButton',
        background='#0dcaf0',
        foreground='black',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )
    style.configure('Selected.Afternoon.TButton',
        background='#0995b5',
        foreground='black',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )
   
    style.configure('Night.TButton',
        background='#6f42c1',
        foreground='white',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )
    style.configure('Selected.Night.TButton',
        background='#4e2d8b',
        foreground='white',
        padding=(15, 8),
        font=('Segoe UI', 10, 'bold')
    )

    # Contenedor izquierdo
    left_container = ttk.Frame(tab)
    left_container.grid(row=0, column=0, sticky='nsew', padx=(10, 5), pady=10)
    left_container.grid_columnconfigure(0, weight=1)

    # Canvas para scroll
    canvas = tk.Canvas(left_container)
    scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=300)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Configurar grid del contenedor izquierdo
    left_container.grid_columnconfigure(0, weight=1)
    left_container.grid_rowconfigure(0, weight=1)
    canvas.grid(row=0, column=0, sticky='nsew')
    scrollbar.grid(row=0, column=1, sticky='ns')

    content_frame = ttk.Frame(scrollable_frame)
    content_frame.pack(fill='x', expand=True, padx=5)

    # Frame para la sección de jornada
    jornada_frame = ttk.LabelFrame(content_frame, text="Jornada", padding=(5, 0, 5, 0))
    jornada_frame.pack(fill='x', pady=(0, 10))
   
    # Inicializar jornada_var
    jornada_var = tk.StringVar()
    jornada_var.set('Mañana')  # Valor por defecto

    ttk.Label(
        jornada_frame,
        text="Seleccione la jornada:",
        font=('Segoe UI', 10),
        anchor='center'
    ).grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky='ew')

    def select_jornada(jornada):
        """Maneja la selección de jornada"""
        jornada_var.set(jornada)
        update_button_states()
        show_scheduled_technicians()

    def update_button_states():
        """Actualiza los estilos de los botones de jornada"""
        selected = jornada_var.get()
        btn_mañana.configure(style='Selected.Morning.TButton' if selected == 'Mañana' else 'Morning.TButton')
        btn_tarde.configure(style='Selected.Afternoon.TButton' if selected == 'Tarde' else 'Afternoon.TButton')
        btn_noche.configure(style='Selected.Night.TButton' if selected == 'Noche' else 'Night.TButton')

    # Crear los botones de jornada
    btn_mañana = ttk.Button(
        jornada_frame,
        text="Mañana",
        command=lambda: select_jornada("Mañana"),
        style='Selected.Morning.TButton',
        width=10
    )
    btn_mañana.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

    btn_tarde = ttk.Button(
        jornada_frame,
        text="Tarde",
        command=lambda: select_jornada("Tarde"),
        style='Afternoon.TButton',
        width=10
    )
    btn_tarde.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

    btn_noche = ttk.Button(
        jornada_frame,
        text="Noche",
        command=lambda: select_jornada("Noche"),
        style='Night.TButton',
        width=10
    )
    btn_noche.grid(row=1, column=2, padx=5, pady=5, sticky='ew')

    # Configurar el grid para los botones
    jornada_frame.grid_columnconfigure((0, 1, 2), weight=1)

    # Sección del calendario
    calendar_section = ttk.LabelFrame(content_frame, text="Fecha", padding=(5, 5, 5, 5))
    calendar_section.pack(fill='x', pady=(0, 10))

    current_date = get_current_date()
    cal = Calendar(calendar_section,
                  selectmode='day',
                  year=current_date['year'],
                  month=current_date['month'],
                  day=current_date['day'],
                  date_pattern='dd/mm/yyyy',
                  background='white',
                  foreground='black',
                  selectbackground='#007bff')
    cal.pack(fill='x')
   
    # Agregar binding para el calendario
    cal.bind("<<CalendarSelected>>", lambda e: show_scheduled_technicians())

    # Sección de formulario
    form_section = ttk.LabelFrame(content_frame, text="Datos del Técnico", padding=(5, 5, 5, 5))
    form_section.pack(fill='x', pady=(0, 10))

    ttk.Label(form_section, text="Turno:", font=('Segoe UI', 10)).pack(anchor='w')
    turno_combobox = ttk.Combobox(form_section, state='readonly', font=('Segoe UI', 10), height=25)
    turno_combobox.pack(fill='x', pady=(0, 10))

    ttk.Label(form_section, text="Nombre:", font=('Segoe UI', 10)).pack(anchor='w')
    name_combobox = ttk.Combobox(form_section, state='readonly', font=('Segoe UI', 10), height=25)
    name_combobox.pack(fill='x', pady=(0, 10))

    ttk.Label(form_section, text="Cargo:", font=('Segoe UI', 10)).pack(anchor='w')
    cargo_entry = ttk.Entry(form_section, font=('Segoe UI', 10))
    cargo_entry.pack(fill='x', pady=(0, 10))
    cargo_entry.configure(state='readonly')

    ttk.Label(form_section, text="Letra de Rotación:", font=('Segoe UI', 10)).pack(anchor='w')
    rotation_letter_entry = ttk.Entry(form_section, font=('Segoe UI', 10))
    rotation_letter_entry.pack(fill='x', pady=(0, 10))
    if user_type == "programador":
        rotation_letter_entry.configure(state='disabled')
    else:
        rotation_letter_entry.configure(state='readonly')

    # Sección de botones
    button_section = ttk.Frame(content_frame)
    button_section.pack(fill='x', pady=(0, 10))

    # Botón Agregar
    if user_type == "programador":
        boton_agregar = ttk.Button(
            button_section,
            text="Agregar Técnico",
            state="disabled",
            style='Primary.TButton'
        )
    else:
        boton_agregar = ttk.Button(
            button_section,
            text="Agregar Técnico",
            command=lambda: add_technician(user_type),
            style='Primary.TButton'
        )
    boton_agregar.pack(fill='x', pady=(0, 5))

    # Botón Eliminar
    if user_type == "programador":
        delete_button = ttk.Button(
            button_section,
            text="Eliminar Seleccionado",
            state="disabled",
            style='Danger.TButton'
        )
    else:
        delete_button = ttk.Button(
            button_section,
            text="Eliminar Seleccionado",
            command=lambda: delete_technician(user_type),
            style='Danger.TButton'
        )
    delete_button.pack(fill='x', pady=(0, 5))

    # Botón Actualizar
    boton_actualizar = ttk.Button(
        button_section,
        text="Actualizar Lista",
        command=show_scheduled_technicians,
        style='Primary.TButton'
    )
    boton_actualizar.pack(fill='x')

    # Configurar la parte derecha con las tablas
    setup_technician_display(tab, user_type)
   
    # Cargar valores y configurar bindings
    cargar_valores_turno()
    turno_combobox.bind('<<ComboboxSelected>>', actualizar_tecnicos)
    name_combobox.bind('<<ComboboxSelected>>', actualizar_cargo)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
   
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
   
    # Mostrar técnicos iniciales
    show_scheduled_technicians()

   
def show_scheduled_technicians():
    """Muestra los técnicos programados en las tablas"""
    global data_display_corrective, data_display_preventive

    try:
        # Limpiar ambas tablas
        for table in [data_display_corrective, data_display_preventive]:
            if table:
                for item in table.get_children():
                    table.delete(item)

        # Verificar existencia del archivo Excel
        if not os.path.exists(excel_file_path):
            ensure_excel_file()
            return

        # Obtener la fecha seleccionada y la jornada
        date_str = cal.get_date()
        jornada_seleccionada = jornada_var.get()

        # Convertir fecha
        try:
            parts = date_str.split('/')
            selected_date = datetime(
                year=int(parts[2]),
                month=int(parts[1]),
                day=int(parts[0])
            )
        except Exception as e:
            print(f"Error al convertir fecha: {str(e)}")
            selected_date = pd.to_datetime(date_str)

        # Leer y preparar datos
        df = pd.read_excel(excel_file_path)
        if df.empty:
            return

        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
        selected_date_only = selected_date.date()

        # Filtrar registros
        df_filtered = df[
            (df['fecha_asignacion'].dt.date == selected_date_only) &
            (df['nombre'].notna()) &
            (df['nombre'] != '') &
            (df['Jornada'] == jornada_seleccionada)
        ]

        # Definir letras de rotación para correctivo
        letras_correctivo = ['A', 'AA', 'AAA', 'B', 'BB', 'BBB', 'C', 'CC', 'CCC']

        # Procesar y mostrar los datos
        if not df_filtered.empty:
            for _, row in df_filtered.iterrows():
                letra_rotacion = str(row['letra_rotacion']).strip().upper() if pd.notna(row['letra_rotacion']) else ''
               
                # Determinar la tabla correcta
                if letra_rotacion in letras_correctivo:
                    display = data_display_corrective
                elif letra_rotacion == 'PREVENTIVO':
                    display = data_display_preventive
                else:
                    continue

                # Obtener el motivo
                motivo = row.get('Motivo', '') if 'Motivo' in row and pd.notna(row['Motivo']) else ''

                # Insertar en la tabla correspondiente
                display.insert('', tk.END, values=(
                    row['nombre'],
                    row['cargo'],
                    row['letra_rotacion'],
                    row['fecha_asignacion'].strftime('%d/%m/%Y'),
                    row['Jornada'],
                    motivo
                ))

        # Contar técnicos correctivos
        df_correctivo = df_filtered[
            df_filtered['letra_rotacion'].str.strip().str.upper().isin(letras_correctivo)
        ]
        cantidad_correctivo = len(df_correctivo)

        # Actualizar warning label y colores según la cantidad
        update_warning_label(cantidad_correctivo, jornada_seleccionada)

        # Aplicar colores alternados
        for display in [data_display_corrective, data_display_preventive]:
            if display:
                for i, item in enumerate(display.get_children()):
                    if i % 2 == 0:
                        display.item(item, tags=('evenrow',))
                    else:
                        display.item(item, tags=('oddrow',))

    except Exception as e:
        print(f"Error en show_scheduled_technicians: {str(e)}")
        messagebox.showerror("Error", f"Error al mostrar los técnicos programados: {str(e)}")

def update_warning_label(cantidad_correctivo, jornada):
    """Actualiza el warning label según la cantidad de técnicos y la jornada"""
    # Definir estados de dotación según la jornada
    dotacion_states = {
        'Mañana': {
            'critico': {
                'range': range(0, 4),
                'color': '#FFD2D2',
                'text_color': '#DC3545',
                'message': '🚨 CRÍTICO: Dotación correctiva mañana insuficiente',
                'icon': '⚠️'
            },
            'precaucion': {
                'range': range(4, 5),
                'color': '#FFF3CD',
                'text_color': '#FFA500',
                'message': '⚠️ PRECAUCIÓN: Dotación correctiva mañana mínima',
                'icon': '⚠️'
            },
            'optimo': {
                'range': range(5, 8),
                'color': '#D1E7DD',
                'text_color': '#28A745',
                'message': '✅ ÓPTIMO: Dotación correctiva mañana adecuada',
                'icon': '✅'
            },
            'exceso': {
                'range': range(8, 999),
                'color': '#CCE5FF',
                'text_color': '#0D6EFD',
                'message': 'ℹ️ NOTA: Dotación correctiva mañana por encima del óptimo',
                'icon': 'ℹ️'
            }
        },
        'Tarde': {
            'critico': {
                'range': range(0, 4),
                'color': '#FFD2D2',
                'text_color': '#DC3545',
                'message': '🚨 CRÍTICO: Dotación correctiva tarde insuficiente',
                'icon': '⚠️'
            },
            'precaucion': {
                'range': range(4, 5),
                'color': '#FFF3CD',
                'text_color': '#FFA500',
                'message': '⚠️ PRECAUCIÓN: Dotación correctiva tarde mínima',
                'icon': '⚠️'
            },
            'optimo': {
                'range': range(5, 8),
                'color': '#D1E7DD',
                'text_color': '#28A745',
                'message': '✅ ÓPTIMO: Dotación correctiva tarde adecuada',
                'icon': '✅'
            },
            'exceso': {
                'range': range(8, 999),
                'color': '#CCE5FF',
                'text_color': '#0D6EFD',
                'message': 'ℹ️ NOTA: Dotación correctiva tarde por encima del óptimo',
                'icon': 'ℹ️'
            }
        },
        'Noche': {
            'critico': {
                'range': range(0, 4),
                'color': '#FFD2D2',
                'text_color': '#DC3545',
                'message': '🚨 CRÍTICO: Dotación correctiva noche insuficiente',
                'icon': '⚠️'
            },
            'precaucion': {
                'range': range(4, 5),
                'color': '#FFF3CD',
                'text_color': '#FFA500',
                'message': '⚠️ PRECAUCIÓN: Dotación correctiva noche mínima',
                'icon': '⚠️'
            },
            'optimo': {
                'range': range(5, 8),
                'color': '#D1E7DD',
                'text_color': '#28A745',
                'message': '✅ ÓPTIMO: Dotación correctiva noche adecuada',
                'icon': '✅'
            },
            'exceso': {
                'range': range(8, 999),
                'color': '#CCE5FF',
                'text_color': '#0D6EFD',
                'message': 'ℹ️ NOTA: Dotación correctiva noche por encima del óptimo',
                'icon': 'ℹ️'
            }
        }
    }

    # Obtener configuración para la jornada actual
    jornada_config = dotacion_states.get(jornada, dotacion_states['Mañana'])
   
    # Determinar estado actual
    current_state = None
    for state, config in jornada_config.items():
        if cantidad_correctivo in config['range']:
            current_state = config
            break

    if current_state:
        # Actualizar warning label
        warning_label.config(
            text=f"{current_state['icon']} {current_state['message']}\n"
                 f"Técnicos correctivos {jornada.lower()}: {cantidad_correctivo}",
            foreground=current_state['text_color']
        )
       
        # Actualizar color de fondo si es necesario
        if table_frame:
            table_frame.configure(style='Custom.TFrame')
            style = ttk.Style()
            style.configure('Custom.TFrame', background=current_state['color'])

def add_technician(user_type):
    if user_type == "programador":
        messagebox.showerror("Error", "Su perfil no tiene permisos para realizar esta acción.")
        return

    turno = turno_combobox.get()
    nombre = name_combobox.get()
    cargo = cargo_entry.get()
    letra_rotacion = rotation_letter_entry.get().strip().upper()
    fecha_str = cal.get_date()
    jornada = jornada_var.get()

    if not all([turno, nombre, cargo, letra_rotacion]):
        messagebox.showerror("Error", "Todos los campos son obligatorios.")
        return

    try:
        ensure_excel_file()
       
        fecha_seleccionada = datetime.strptime(fecha_str, '%d/%m/%Y')

        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])

        existente = df[
            (df['fecha_asignacion'].dt.date == fecha_seleccionada.date()) &
            (df['nombre'] == nombre) &
            (df['Jornada'] == jornada)
        ]

        if not existente.empty:
            messagebox.showerror("Error",
                               "Ya existe un registro para este técnico en la fecha y jornada seleccionada.")
            return

        dialog = MotivoAgregadoDialog(turno_combobox.winfo_toplevel())
        dialog.dialog.wait_window()

        if not dialog.result:
            return

        nueva_fila = pd.DataFrame({
            'turno': [turno],
            'nombre': [nombre],
            'cargo': [cargo],
            'letra_rotacion': [letra_rotacion],
            'fecha_asignacion': [fecha_seleccionada],
            'Jornada': [jornada],
            'Motivo': [dialog.motivo]
        })

        df = pd.concat([df, nueva_fila], ignore_index=True)
        
        # Registrar la acción en el log
        log_entry = pd.DataFrame({
            'nombre_log': [user_type],
            'fecha_log': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            'accion_log': [f"Agregó técnico: {nombre}"]
        })
        df = pd.concat([df, log_entry], ignore_index=True)
        
        df.to_excel(excel_file_path, index=False)

        messagebox.showinfo("Éxito", 
                          f"Técnico agregado exitosamente en jornada {jornada}.\nMotivo: {dialog.motivo}")

        # Limpiar campos
        turno_combobox.set('')
        name_combobox.set('')
        cargo_entry.configure(state='normal')
        cargo_entry.delete(0, tk.END)
        cargo_entry.configure(state='readonly')
        rotation_letter_entry.configure(state='normal')
        rotation_letter_entry.delete(0, tk.END)
        rotation_letter_entry.configure(state='readonly')
        turno_combobox.focus()
       
        show_scheduled_technicians()

    except Exception as e:
        messagebox.showerror("Error", f"Error al agregar el técnico: {str(e)}")


def setup_programmer_tab(tab):
    tab.grid_columnconfigure(1, weight=3)
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(0, weight=1)
   
    # Panel izquierdo con scrollbar
    left_container = ttk.Frame(tab)
    left_container.grid(row=0, column=0, sticky='nsew', padx=(10, 5))
    left_container.grid_propagate(False)  # Prevenir propagación del grid
    left_container.configure(width=320)  # Ancho fijo del panel izquierdo

    # Canvas y Scrollbar
    canvas = tk.Canvas(left_container)
    scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    # Configurar el scrolling
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=300)

    # Ajustar el ancho del frame interior cuando el canvas cambie de tamaño
    def configure_scroll_frame(event):
        canvas.itemconfig(canvas_frame, width=event.width)
   
    canvas.bind('<Configure>', configure_scroll_frame)

    # Frame para filtros
    filter_frame = ttk.LabelFrame(scrollable_frame, text="Filtros", padding=10)
    filter_frame.pack(fill='x', expand=True, pady=(0, 10))

    # Calendario Inicial
    ttk.Label(filter_frame, text="Fecha Inicial:",
             font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))
   
    current_date = get_current_date()
    cal_inicio = Calendar(filter_frame,
                         selectmode='day',
                         year=current_date['year'],
                         month=current_date['month'],
                         day=current_date['day'],
                         date_pattern='dd-mm-yyyy',
                         background='white',
                         foreground='black',
                         selectbackground='#007bff')
    cal_inicio.pack(fill='x', pady=(0, 10))

    # Calendario Final
    ttk.Label(filter_frame, text="Fecha Final:",
             font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))
    cal_fin = Calendar(filter_frame,
                      selectmode='day',
                      year=current_date['year'],
                      month=current_date['month'],
                      day=current_date['day'],
                      date_pattern='dd-mm-yyyy',
                      background='white',
                      foreground='black',
                      selectbackground='#007bff')
    cal_fin.pack(fill='x', pady=(0, 10))

    # Filtro de Jornada
    jornada_frame = ttk.LabelFrame(scrollable_frame, text="Filtrar por Jornada", padding=5)
    jornada_frame.pack(fill='x', pady=(10, 0))

    jornada_var = tk.StringVar(value="todas")

    ttk.Radiobutton(
        jornada_frame,
        text="Todas las Jornadas",
        variable=jornada_var,
        value="todas"
    ).pack(anchor='w', pady=2)

    ttk.Radiobutton(
        jornada_frame,
        text="Mañana",
        variable=jornada_var,
        value="Mañana"
    ).pack(anchor='w', pady=2)

    ttk.Radiobutton(
        jornada_frame,
        text="Tarde",
        variable=jornada_var,
        value="Tarde"
    ).pack(anchor='w', pady=2)

    ttk.Radiobutton(
        jornada_frame,
        text="Noche",
        variable=jornada_var,
        value="Noche"
    ).pack(anchor='w', pady=2)

    # Botones
    button_frame = ttk.Frame(scrollable_frame)
    button_frame.pack(fill='x', pady=10)
   
    calendars = {'cal_inicio': cal_inicio, 'cal_fin': cal_fin}
   
    ttk.Button(
        button_frame,
        text="Filtrar Técnicos",
        command=lambda: filter_technicians(calendars, jornada_var),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)
   
    ttk.Button(
        button_frame,
        text="Exportar a Excel",
        command=lambda: export_to_excel(calendars, jornada_var),
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # Configurar el canvas y scrollbar
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configurar el mousewheel scrolling
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
   
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Tabla actualizada con columna de jornada
    table_frame = ttk.Frame(tab, padding="10")
    table_frame.grid(row=0, column=1, sticky='nsew')
    table_frame.grid_columnconfigure(0, weight=1)
    table_frame.grid_rowconfigure(1, weight=1)
   
    ttk.Label(
        table_frame,
        text="Técnicos en Rango de Fechas",
        font=('Segoe UI', 12, 'bold')
    ).grid(row=0, column=0, pady=(0, 10))
   
    global range_display
    range_display = ttk.Treeview(
        table_frame,
        columns=('nombre', 'cargo', 'letra_rotacion', 'fecha', 'jornada'),
        show='headings'
    )
   
    # Configurar columnas
    range_display.heading('nombre', text='Nombre')
    range_display.heading('cargo', text='Cargo')
    range_display.heading('letra_rotacion', text='Rotación')
    range_display.heading('fecha', text='Fecha')
    range_display.heading('jornada', text='Jornada')
   
    # Configurar anchos de columna
    range_display.column('nombre', width=150, minwidth=100)
    range_display.column('cargo', width=150, minwidth=100)
    range_display.column('letra_rotacion', width=100, minwidth=80)
    range_display.column('fecha', width=100, minwidth=100)
    range_display.column('jornada', width=100, minwidth=80)
   
    # Scrollbars para la tabla
    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=range_display.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=range_display.xview)
    range_display.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
   
    # Grid
    range_display.grid(row=1, column=0, sticky='nsew')
    vsb.grid(row=1, column=1, sticky='ns')
    hsb.grid(row=2, column=0, sticky='ew')

    # Estilos para filas alternadas
    range_display.tag_configure('oddrow', background='#f5f5f5')
    range_display.tag_configure('evenrow', background='white')

    return tab
   




def setup_resumen_tab(tab):
    """Configura la pestaña de resumen con filtro de jornadas"""
    global graph_frames, resumen_tab, cal_inicio_resumen, cal_fin_resumen, fig, canvas_plot
    resumen_tab = tab

    # Configurar el grid del tab principal
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(0, weight=1)

    # Crear PanedWindow horizontal
    main_paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
    main_paned.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

    # Panel izquierdo con scroll
    left_container = ttk.Frame(main_paned)
    left_container.grid_propagate(False)
    left_container.configure(width=320)

    scroll_canvas = tk.Canvas(left_container)
    scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=scroll_canvas.yview)
    scrollable_frame = ttk.Frame(scroll_canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
    )

    scroll_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=300)
    scroll_canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    scroll_canvas.pack(side="left", fill="both", expand=True)

    # Frame para fechas
    date_frame = ttk.LabelFrame(scrollable_frame, text="Rango de Fechas", padding=10)
    date_frame.pack(fill='x', padx=5, pady=5)

    # Fecha Inicial
    ttk.Label(date_frame, text="Fecha Inicial:",
             font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=2)
   
    current_date = get_current_date()
    global cal_inicio_resumen
    cal_inicio_resumen = Calendar(date_frame,
                                selectmode='day',
                                year=current_date['year'],
                                month=current_date['month'],
                                day=current_date['day'],
                                date_pattern='dd/mm/yyyy',
                                background='white',
                                foreground='black',
                                selectbackground='#007bff',
                                width=270)
    cal_inicio_resumen.pack(pady=5)

    # Fecha Final
    ttk.Label(date_frame, text="Fecha Final:",
             font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=2)
             
    global cal_fin_resumen
    cal_fin_resumen = Calendar(date_frame,
                             selectmode='day',
                             year=current_date['year'],
                             month=current_date['month'],
                             day=current_date['day'],
                             date_pattern='dd/mm/yyyy',
                             background='white',
                             foreground='black',
                             selectbackground='#007bff',
                             width=270)
    cal_fin_resumen.pack(pady=5)

    # Frame para filtro de jornada
    shift_frame = ttk.LabelFrame(scrollable_frame, text="Filtrar por Jornada", padding=10)
    shift_frame.pack(fill='x', padx=5, pady=5)

    # Variable para jornada
    global jornada_resumen_var
    jornada_resumen_var = tk.StringVar(value="todas")

    # Radiobuttons para jornadas
    jornadas = [
        ("Todas las Jornadas", "todas"),
        ("Mañana", "Mañana"),
        ("Tarde", "Tarde"),
        ("Noche", "Noche")
    ]

    for text, value in jornadas:
        ttk.Radiobutton(
            shift_frame,
            text=text,
            variable=jornada_resumen_var,
            value=value,
            command=update_resumen_graph
        ).pack(anchor='w', pady=2)

    # Frame para opciones de visualización
    options_frame = ttk.LabelFrame(scrollable_frame, text="Opciones de Visualización", padding=10)
    options_frame.pack(fill='x', padx=5, pady=5)

    # Variables para opciones
    global show_mean_var, show_values_var
    show_mean_var = tk.BooleanVar(value=True)
    show_values_var = tk.BooleanVar(value=True)

    ttk.Checkbutton(
        options_frame,
        text="Mostrar promedio",
        variable=show_mean_var,
        command=update_resumen_graph
    ).pack(anchor='w', pady=2)

    ttk.Checkbutton(
        options_frame,
        text="Mostrar valores",
        variable=show_values_var,
        command=update_resumen_graph
    ).pack(anchor='w', pady=2)

    # Frame para estadísticas
    stats_frame = ttk.LabelFrame(scrollable_frame, text="Estadísticas", padding=10)
    stats_frame.pack(fill='x', padx=5, pady=5)

    # Variables para estadísticas
    global stats_labels
    stats_labels = {
        'total': ttk.Label(stats_frame, text="Total técnicos: -"),
        'promedio': ttk.Label(stats_frame, text="Promedio diario: -"),
        'max': ttk.Label(stats_frame, text="Máximo diario: -"),
        'min': ttk.Label(stats_frame, text="Mínimo diario: -"),
        'dias': ttk.Label(stats_frame, text="Días totales: -")
    }

    for label in stats_labels.values():
        label.pack(anchor='w', pady=2)

    # Botones
    buttons_frame = ttk.Frame(scrollable_frame)
    buttons_frame.pack(fill='x', padx=5, pady=5)

    ttk.Button(
        buttons_frame,
        text="Actualizar Gráfico",
        command=update_resumen_graph,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        buttons_frame,
        text="Exportar Gráfico",
        command=export_graph,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    ttk.Button(
        buttons_frame,
        text="Exportar Datos",
        command=export_data,
        style='Primary.TButton'
    ).pack(fill='x', pady=2)

    # Panel derecho (gráfico)
    right_frame = ttk.Frame(main_paned)
   
    graph_frame = ttk.LabelFrame(right_frame, text="Visualización de Datos")
    graph_frame.pack(fill='both', expand=True, padx=5, pady=5)

    # Crear figura de matplotlib
    global fig
    fig = Figure(figsize=(8, 6), dpi=100)
    canvas_plot = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas_plot.get_tk_widget().pack(fill='both', expand=True)

    toolbar = NavigationToolbar2Tk(canvas_plot, graph_frame, pack_toolbar=False)
    toolbar.pack(side='bottom', fill='x')

    # Añadir paneles al PanedWindow
    main_paned.add(left_container)
    main_paned.add(right_frame, weight=1)

    # Mantener proporción
    def maintain_ratio(event=None):
        total_width = main_paned.winfo_width()
        sash_position = 320
        main_paned.sashpos(0, sash_position)
   
    main_paned.bind('<Configure>', maintain_ratio)
    tab.after(100, maintain_ratio)

    # Configurar scroll
    def _on_mousewheel(event):
        scroll_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
   
    scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Inicializar el gráfico
    update_resumen_graph()
   
   
   
   


def update_resumen_graph():
    """Actualiza el gráfico de resumen con filtro de jornada"""
    try:
        print("Iniciando actualización del gráfico de resumen...")
       
        # Obtener fechas
        date_str_inicio = cal_inicio_resumen.get_date()
        date_str_fin = cal_fin_resumen.get_date()

        # Convertir fechas
        start_parts = date_str_inicio.split('/')
        end_parts = date_str_fin.split('/')
       
        start_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
        end_date = datetime(
            year=int(end_parts[2]),
            month=int(end_parts[1]),
            day=int(end_parts[0])
        )
       
        if start_date > end_date:
            messagebox.showerror("Error", "La fecha inicial debe ser anterior a la fecha final")
            return
       
        # Leer datos
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        # Aplicar filtro de fechas
        mask = (df['fecha_asignacion'].dt.date >= start_date.date()) & \
               (df['fecha_asignacion'].dt.date <= end_date.date())
       
        # Aplicar filtro de jornada
        jornada_seleccionada = jornada_resumen_var.get()
        if jornada_seleccionada != "todas":
            mask = mask & (df['Jornada'] == jornada_seleccionada)
           
        df_filtered = df[mask]

        # Limpiar figura anterior
        fig.clear()
        ax = fig.add_subplot(111)

        if df_filtered.empty:
            ax.text(0.5, 0.5, 'No hay datos para mostrar en el rango seleccionado',
                   horizontalalignment='center', verticalalignment='center',
                   fontsize=12, color='#666666')
            canvas_plot.draw()
            return

        # Preparar datos para el gráfico
        if jornada_seleccionada == "todas":
            # Agrupar por fecha y jornada
            daily_counts = df_filtered.groupby([
                df_filtered['fecha_asignacion'].dt.date,
                'Jornada'
            ]).size().unstack(fill_value=0)
           
            # Crear gráfico para cada jornada
            colors = {'Mañana': '#ffd54f', 'Tarde': '#4fc3f7', 'Noche': '#9575cd'}
            for jornada in daily_counts.columns:
                ax.plot(daily_counts.index,
                       daily_counts[jornada],
                       marker='o',
                       linestyle='-',
                       linewidth=2,
                       markersize=8,
                       color=colors.get(jornada, '#007bff'),
                       label=f'Jornada {jornada}')
        else:
            # Gráfico para jornada específica
            daily_counts = df_filtered.groupby(
                df_filtered['fecha_asignacion'].dt.date
            ).size()
           
            ax.plot(daily_counts.index,
                   daily_counts.values,
                   marker='o',
                   linestyle='-',
                   linewidth=2,
                   markersize=8,
                   color='#007bff',
                   label=f'Jornada {jornada_seleccionada}')

        # Mostrar promedio si está activado
        if show_mean_var.get():
            if jornada_seleccionada == "todas":
                for jornada in daily_counts.columns:
                    mean_value = daily_counts[jornada].mean()
                    ax.axhline(y=mean_value,
                             color=colors.get(jornada, '#dc3545'),
                             linestyle='--',
                             label=f'Promedio {jornada}: {mean_value:.1f}')
            else:
                mean_value = daily_counts.mean()
                ax.axhline(y=mean_value,
                          color='#dc3545',
                          linestyle='--',
                          label=f'Promedio: {mean_value:.1f}')

        # Mostrar valores sobre los puntos
        if show_values_var.get():
            if jornada_seleccionada == "todas":
                for jornada in daily_counts.columns:
                    for x, y in zip(daily_counts.index, daily_counts[jornada]):
                        if y > 0:  # Solo mostrar valores mayores a 0
                            ax.annotate(str(int(y)),
                                      (x, y),
                                      textcoords="offset points",
                                      xytext=(0,10),
                                      ha='center',
                                      fontsize=9,
                                      bbox=dict(boxstyle='round,pad=0.5',
                                              fc='white',
                                              ec='gray',
                                              alpha=0.7))
            else:
                for x, y in zip(daily_counts.index, daily_counts.values):
                    ax.annotate(str(int(y)),
                              (x, y),
                              textcoords="offset points",
                              xytext=(0,10),
                              ha='center',
                              fontsize=9,
                              bbox=dict(boxstyle='round,pad=0.5',
                                      fc='white',
                                      ec='gray',
                                      alpha=0.7))

        # Configurar formato de fechas y etiquetas
        date_formatter = mdates.DateFormatter('%d/%m/%Y')
        ax.xaxis.set_major_formatter(date_formatter)
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        fig.autofmt_xdate()

        ax.set_xlabel('Fecha', fontsize=10, labelpad=10)
        ax.set_ylabel('Cantidad de Técnicos', fontsize=10, labelpad=10)
       
        title = f'Cantidad de Técnicos por Día\n{start_date.strftime("%d/%m/%Y")} a {end_date.strftime("%d/%m/%Y")}'
        if jornada_seleccionada != "todas":
            title += f'\nJornada: {jornada_seleccionada}'
        ax.set_title(title, pad=20, fontsize=12)

        ax.grid(True, linestyle='--', alpha=0.3)
        ax.set_ylim(bottom=0)
        # Mostrar leyenda si es necesario
        if show_mean_var.get() or jornada_seleccionada == "todas":
            ax.legend(loc='upper right')

        # Ajustar layout y mostrar
        fig.tight_layout()
        canvas_plot.draw()

        # Actualizar estadísticas
        if jornada_seleccionada == "todas":
            # Calcular estadísticas sumando todas las jornadas
            total = df_filtered['nombre'].count()
            promedio = daily_counts.sum(axis=1).mean()
            maximo = daily_counts.sum(axis=1).max()
            minimo = daily_counts.sum(axis=1).min()
            dias = len(daily_counts)
        else:
            # Estadísticas para jornada específica
            total = len(df_filtered)
            promedio = daily_counts.mean()
            maximo = daily_counts.max()
            minimo = daily_counts.min()
            dias = len(daily_counts)

        # Actualizar etiquetas de estadísticas
        stats_labels['total'].config(text=f"Total técnicos: {int(total)}")
        stats_labels['promedio'].config(text=f"Promedio diario: {promedio:.1f}")
        stats_labels['max'].config(text=f"Máximo diario: {int(maximo)}")
        stats_labels['min'].config(text=f"Mínimo diario: {int(minimo)}")
        stats_labels['dias'].config(text=f"Días totales: {dias}")

        print("Gráfico actualizado correctamente")

    except Exception as e:
        print(f"Error en update_resumen_graph: {str(e)}")
        messagebox.showerror("Error", f"Error al actualizar el gráfico: {str(e)}")

def export_graph():
    """Exporta el gráfico actual como imagen"""
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension='.png',
            filetypes=[
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ],
            title="Guardar gráfico"
        )
       
        if file_path:
            # Obtener la jornada seleccionada para el nombre del archivo
            jornada = jornada_resumen_var.get()
            fecha_inicio = cal_inicio_resumen.get_date()
            fecha_fin = cal_fin_resumen.get_date()
           
            # Añadir información al título del gráfico antes de exportar
            ax = fig.gca()
            title = f'Cantidad de Técnicos por Día\n{fecha_inicio} a {fecha_fin}'
            if jornada != "todas":
                title += f'\nJornada: {jornada}'
            ax.set_title(title, pad=20, fontsize=12)
           
            # Guardar con la mejor calidad
            fig.savefig(
                file_path,
                dpi=300,
                bbox_inches='tight',
                pad_inches=0.1,
                format=file_path.split('.')[-1]
            )
           
            messagebox.showinfo(
                "Éxito",
                "Gráfico exportado correctamente"
            )
           
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al exportar el gráfico: {str(e)}"
        )

def export_data():
    """Exporta los datos filtrados a Excel con formato mejorado"""
    try:
        # Obtener fechas y jornada
        fecha_inicio = cal_inicio_resumen.get_date()
        fecha_fin = cal_fin_resumen.get_date()
        jornada = jornada_resumen_var.get()

        # Convertir fechas
        start_date = datetime.strptime(fecha_inicio, '%d/%m/%Y')
        end_date = datetime.strptime(fecha_fin, '%d/%m/%Y')

        # Leer datos
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])

        # Aplicar filtros
        mask = (df['fecha_asignacion'].dt.date >= start_date.date()) & \
               (df['fecha_asignacion'].dt.date <= end_date.date())

        if jornada != "todas":
            mask = mask & (df['Jornada'] == jornada)

        df_filtered = df[mask]

        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return

        # Solicitar ubicación para guardar
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte de datos"
        )

        if file_path:
            # Crear Excel writer
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Preparar datos para el resumen diario
                if jornada == "todas":
                    pivot_table = pd.pivot_table(
                        df_filtered,
                        values='nombre',
                        index='fecha_asignacion',
                        columns='Jornada',
                        aggfunc='count',
                        fill_value=0
                    )
                else:
                    daily_counts = df_filtered.groupby('fecha_asignacion').size()
                    pivot_table = pd.DataFrame(daily_counts, columns=['Cantidad'])

                # Formatear fechas
                pivot_table.index = pivot_table.index.strftime('%d/%m/%Y')

                # Guardar resumen diario
                pivot_table.to_excel(writer, sheet_name='Resumen Diario')

                # Guardar datos detallados
                detailed_data = df_filtered[[
                    'nombre',
                    'cargo',
                    'letra_rotacion',
                    'fecha_asignacion',
                    'Jornada'
                ]].copy()
                detailed_data['fecha_asignacion'] = detailed_data['fecha_asignacion'].dt.strftime('%d/%m/%Y')
                detailed_data.columns = ['Nombre', 'Cargo', 'Rotación', 'Fecha', 'Jornada']
                detailed_data.to_excel(writer, sheet_name='Datos Detallados', index=False)

                # Ajustar anchos de columna
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            messagebox.showinfo(
                "Éxito",
                "Datos exportados exitosamente"
            )

    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Error al exportar los datos: {str(e)}"
        )
       
       
def update_statistics(daily_counts):
    """Actualiza las estadísticas mostradas"""
    try:
        total = daily_counts['Cantidad'].sum()
        promedio = daily_counts['Cantidad'].mean()
        maximo = daily_counts['Cantidad'].max()
        minimo = daily_counts['Cantidad'].min()
        dias = len(daily_counts)
       
        stats_labels['total'].config(text=f"Total técnicos: {int(total)}")
        stats_labels['promedio'].config(text=f"Promedio diario: {promedio:.1f}")
        stats_labels['max'].config(text=f"Máximo diario: {int(maximo)}")
        stats_labels['min'].config(text=f"Mínimo diario: {int(minimo)}")
        stats_labels['dias'].config(text=f"Días totales: {dias}")
    except Exception as e:
        print(f"Error al actualizar estadísticas: {str(e)}")

def export_graph():
    """Exporta el gráfico actual como imagen"""
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension='.png',
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            title="Guardar gráfico"
        )
        if file_path:
            fig.savefig(file_path, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Éxito", "Gráfico exportado correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar el gráfico: {str(e)}")

def export_data():
    try:
        # Obtener las fechas
        date_str_inicio = cal_inicio_resumen.get_date()
        date_str_fin = cal_fin_resumen.get_date()

        # Convertir fechas usando el formato correcto dd/mm/yyyy
        start_parts = date_str_inicio.split('/')
        end_parts = date_str_fin.split('/')
       
        start_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
        end_date = datetime(
            year=int(start_parts[2]),
            month=int(start_parts[1]),
            day=int(start_parts[0])
        )
       
        df = pd.read_excel(excel_file_path)
        df['fecha_asignacion'] = pd.to_datetime(df['fecha_asignacion'])
       
        mask = (df['fecha_asignacion'].dt.date >= start_date.date()) & \
               (df['fecha_asignacion'].dt.date <= end_date.date())
        df_filtered = df[mask]
       
        if df_filtered.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
       
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar datos"
        )
       
        if file_path:
            df_filtered.to_excel(file_path, index=False)
            messagebox.showinfo("Éxito", "Datos exportados correctamente")
           
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar los datos: {str(e)}")
   
   
   


def cleanup_temp_files():
    try:
        if os.path.exists(TEMP_DIR):
            for file in os.listdir(TEMP_DIR):
                try:
                    file_path = os.path.join(TEMP_DIR, file)
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Error al eliminar archivo temporal: {str(e)}")
            try:
                os.rmdir(TEMP_DIR)
            except Exception as e:
                print(f"Error al eliminar directorio temporal: {str(e)}")
    except Exception as e:
        print(f"Error en cleanup_temp_files: {str(e)}")

if __name__ == "__main__":
    try:
        ensure_excel_file()
        open_login_window()
    except Exception as e:
        messagebox.showerror("Error Fatal", f"Error al iniciar la aplicación: {str(e)}")
    finally:
        cleanup_temp_files()