import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io
import pandas as pd
import numpy as np
import os

class CorrectorPsicometrico:
    def __init__(self, root):
        self.root = root
        self.root.title("Corrector del SCL-90")
        self.root.configure(bg="#f8f9fa")  # Fondo gris claro moderno

        # === ASPECTO MODERNO: Tema 'clam' + Estilos personalizados ===
        style = ttk.Style()
        style.theme_use('clam')  # Tema moderno
        style.configure('Modern.TLabel', font=('Segoe UI', 10), background='#f8f9fa', foreground='#333')
        style.configure('Modern.TEntry', font=('Consolas', 9), fieldbackground='#ffffff', justify='center')  # Centrado
        style.configure('Modern.TButton', font=('Segoe UI', 11, 'bold'), background='#0078d4', foreground='white')
        style.map('Modern.TButton', background=[('active', '#0078d4')])
        style.configure('Header.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#0078d4')

        # Tamaño adaptativo
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        w = min(750, int(screen_w * 0.9))  # Más ancho para 6 cols
        h = min(950, int(screen_h * 0.8))
        root.geometry(f"{w}x{h}")
        root.minsize(750, 950)

        # === HEADER: Nombre y Fecha ===
        header_frame = ttk.Frame(root, relief='flat', padding=10)
        header_frame.pack(fill='x', padx=10, pady=(10, 5))

        ttk.Label(header_frame, text="Nombre:", style='Header.TLabel').grid(row=0, column=0, sticky='w', padx=(0, 5))
        self.entry_nombre = ttk.Entry(header_frame, width=20, style='Modern.TEntry')
        self.entry_nombre.grid(row=0, column=1, sticky='w', padx=(0, 20))

        ttk.Label(header_frame, text="Sexo:", style='Header.TLabel').grid(row=0, column=2, sticky='w', padx=(0, 5))
        self.entry_sexo = ttk.Combobox(header_frame, width=10, textvariable = tk.StringVar(), state="readonly")
        self.entry_sexo['values'] = ('Hombre', 'Mujer')
        self.entry_sexo.current(0)
        self.entry_sexo.grid(row=0, column=3, sticky='w', padx=(0, 20))

        ttk.Label(header_frame, text="Fecha:", style='Header.TLabel').grid(row=0, column=4, sticky='w', padx=(0, 5))
        self.entry_fecha = ttk.Entry(header_frame, width=15, style='Modern.TEntry')
        self.entry_fecha.insert(0, datetime.today().strftime("%d/%m/%Y") )  # Fecha actual fija
        self.entry_fecha.grid(row=0, column=5, sticky='w')

        ttk.Label(header_frame, text="Evaluador:", style='Header.TLabel').grid(row=1, column=0, sticky='e', padx=(0, 5), ipady=10)
        self.entry_terapeuta = ttk.Entry(header_frame, width=20, style='Modern.TEntry')
        self.entry_terapeuta.grid(row=1, column=1, sticky='w', padx=(0, 20))

        # === CUADRÍCULA 6 COLS x 30 FILAS (3 subescalas: Label+Entry repetido) ===
        grid_frame = ttk.Frame(root)
        grid_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Canvas + Scroll MEJORADO (responde a rueda del ratón)
        self.canvas = tk.Canvas(grid_frame, bg='#f8f9fa', highlightthickness=0)
        scrollbar_v = ttk.Scrollbar(grid_frame, orient='vertical', command=self.canvas.yview)
        scrollable_frame = ttk.Frame(self.canvas, style='Modern.TFrame')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar_v.set)

        # Bind rueda del ratón al canvas (arregla el scroll)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))  # Linux
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))  # Linux

        # Crear labels de cabecera para subescalas
        # ttk.Label(scrollable_frame, text="Subescala 1", style='Header.TLabel').grid(row=0, column=0, columnspan=2, pady=5, sticky='nsew')
        # ttk.Label(scrollable_frame, text="Subescala 2", style='Header.TLabel').grid(row=0, column=2, columnspan=2, pady=5, sticky='nsew')
        # ttk.Label(scrollable_frame, text="Subescala 3", style='Header.TLabel').grid(row=0, column=4, columnspan=2, pady=5, sticky='nsew')

        # Crear 30 filas x 6 columnas (cols 0,2,4: Labels "Pregunta X"; cols 1,3,5: Entries)
        self.entries = []  # Lista plana de 90 entries para tab order vertical
        for sub in range(3): #3 columnas
            col_label = 2 * sub  # Cols 0,2,4 para labels
            col_entry = 2 * sub + 1  # Cols 1,3,5 para entries

            for row in range(30):  # 30 preguntas
                # Label "Pregunta X" (compartido visualmente, pero por subescala)
                label_text = f"Pregunta {row+sub*30+1}"
                ttk.Label(scrollable_frame, text=label_text, style='Modern.TLabel', width=18, anchor='center').grid(
                    row=row+1, column=col_label, sticky='nsew', padx=1, pady=1
                )

                # Fila de números
                #num_label = ttk.Label(scrollable_frame, text=str(row+1), style='Modern.TLabel', width=8, anchor='center')
                #num_label.grid(row=row+1, column=0, sticky='nsew', padx=(5, 1), pady=1)
                #num_label.grid_propagate(False)  # Fija ancho

                # Entry para Likert 0-4
                entry = ttk.Entry(scrollable_frame, width=6, style='Modern.TEntry')
                entry.grid(row=row+1, column=col_entry, sticky='nsew', padx=1, pady=1)
                self.entries.append(entry)  # Añadir a lista plana para tab

                # Validación en tiempo real (solo 0-4)
                entry.bind('<KeyRelease>', self._validar_likert)

        # Configurar pesos para resize (6 columnas)
        for i in range(6):
            scrollable_frame.grid_columnconfigure(i, weight=1)
        for i in range(31):
            scrollable_frame.grid_rowconfigure(i, weight=1)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar_v.pack(side="right", fill="y")

        # === TAB ORDER VERTICAL: Sub1 (0-29), Sub2 (30-59), Sub3 (60-89) ===
        for i, entry in enumerate(self.entries):
            entry.lift()  # Asegurar visibilidad
        self._configurar_tab_order()

        # === BOTÓN CORREGIR ===
        button_frame = ttk.Frame(root)
        button_frame.pack(fill='x', padx=10, pady=(5, 10))
        self.btn_corregir = ttk.Button(button_frame, text="CORREGIR Y GENERAR PDF", command=self.corregir, style='Modern.TButton', cursor='hand2')
        self.btn_corregir.pack()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"  # Evita propagación

    def _validar_likert(self, event):
        entry = event.widget
        valor = entry.get().strip()
        if valor and not (valor.isdigit() and 0 <= int(valor) <= 4):
            entry.delete(0, tk.END)
            entry.insert(0, "")  # Limpiar inválido
            messagebox.showwarning("Validación", "Solo números 0-4 permitidos")

    def _configurar_tab_order(self):
        # Bind Tab para order vertical: Sub1 → Sub2 → Sub3 → ciclo
        def next_tab(event):
            current = event.widget
            idx = self.entries.index(current)
            next_idx = (idx + 1) % len(self.entries)
            self.entries[next_idx].focus_set()
            return "break"

        for entry in self.entries:
            entry.bind('<Tab>', next_tab)
            entry.bind('<Shift-Tab>', lambda e: "break")  # Opcional: deshabilitar shift-tab

    def corregir(self):
        nombre = self.entry_nombre.get().strip()
        fecha = self.entry_fecha.get().strip()
        terapeuta = self.entry_terapeuta.get().strip()

        if not nombre:
            messagebox.showerror("Error", "Introduce el nombre de la persona evaluada")
            return
        
        if not terapeuta:
            messagebox.showerror("Error", "Introduce el nombre del/de la terapeuta")
            return
        
        # Nombre por defecto bonito y seguro
        nombre_por_defecto = f"SCL90R_{nombre.replace(' ', '_')}_{self.entry_fecha.get().replace('/', '-')}.pdf"

        # Abrir diálogo de guardado (funciona en Windows, Linux y macOS)
        ruta_archivo = filedialog.asksaveasfilename(
            title="Guardar informe SCL-90-R",
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf"), ("Todos los archivos", "*.*")],
            initialfile=nombre_por_defecto,           # nombre sugerido
            initialdir="~/Desktop" if os.name != "nt" else None  # en Linux/mac abre en Escritorio
        )

        # Si el usuario cancela → salir
        if not ruta_archivo:
            messagebox.showinfo("Cancelado", "Guardado cancelado por el usuario")
            return

        # ==================== SCL-90-R OFICIAL ====================
        # Índices de los ítems para cada dimensión (empezando en 0, no en 1)

        scl90r_escalas = {
            "Somatización":          [1, 4, 12, 27, 40, 42, 48, 49, 52, 53, 56, 58],      # 12 ítems
            "Obsesividad-Compulsividad": [3, 9, 10, 18, 28, 38, 45, 46, 51, 55],          # 10 ítems
            "Sensibilidad Interpersonal": [6, 21, 34, 36, 37, 41, 61, 69, 73],           # 9 ítems
            "Depresión":             [5, 14, 15, 20, 22, 26, 29, 30, 31, 32, 54, 71, 79], # 13 ítems
            "Ansiedad":              [2, 17, 23, 33, 39, 57, 72, 78, 80, 86],             # 10 ítems
            "Hostilidad":            [11, 24, 63, 67, 74, 81],                            # 6 ítems
            "Ansiedad Fóbica":       [13, 25, 47, 50, 70, 75, 82],                        # 7 ítems
            "Ideación Paranoide":    [8, 18, 43, 68, 76, 83],                             # 6 ítems
            "Psicotismo":            [7, 16, 35, 62, 77, 84, 85, 87, 88, 90]              # 10 ítems
        }

        # Ítems ADICIONALES (no entran en las 9 dimensiones, pero sí en los índices globales)
        items_adicionales = [19, 44, 59, 60, 64, 66, 89]   # 7 ítems que solo cuentan en GSI, PSDI, PST

        # Recoger respuestas Likert (90 valores)
        respuestas = []
        for entry in self.entries:
            val = entry.get().strip()
            respuestas.append(int(val) if val.isdigit() else 0)  # 0 por defecto

        # Convertir todo a float (ya son ints, no need strip)
        try:
            valores = [float(x) for x in respuestas]
        except:
            raise ValueError("Alguna respuesta no es numérica")

        if len(valores) != 90:
            messagebox.showerror("Error", "Deben haber exactamente 90 respuestas")
            return

        resultados = {}

        # Calcular las 9 dimensiones
        sub_sumas = []  # Lista de sumas brutas para las gráficas
        for escala, items in scl90r_escalas.items():
            puntaje = sum(valores[i-1] for i in items)   # -1 porque el manual cuenta desde 1
            resultados[escala] = {
                "bruto": puntaje,
                "media": round(puntaje / len(items), 2),
                "n_items": len(items)
            }
            sub_sumas.append(puntaje)

        # Índices globales
        gsi_total = sum(valores)                                   # suma de los 90 ítems
        pst = sum(1 for x in valores if x > 0)                      # número de síntomas positivos
        psdi = gsi_total / pst if pst > 0 else 0                   # intensidad media de los síntomas positivos

        resultados["Índices Globales"] = {
            "GSI"  : round(gsi_total / 90, 2),    # Global Severity Index
            "PST"  : pst,                         # Positive Symptom Total
            "PSDI" : round(psdi, 2)                # Positive Symptom Distress Index
        }

        try:
            # Gráficas
            self.generar_graficas(respuestas, sub_sumas, resultados['Índices Globales']['GSI'])

            # ← AQUÍ ESTÁ EL CAMBIO CLAVE:
            self.generar_pdf(
                nombre=nombre,
                fecha=fecha,
                terapeuta=terapeuta,
                resultados=resultados,
                scl90r_escalas=scl90r_escalas,
                respuestas=respuestas,
                gsi=resultados['Índices Globales']['GSI'],
                pst=resultados['Índices Globales']['PST'],
                ruta_pdf=ruta_archivo  # ← ¡NUEVO PARÁMETRO!
            )

            messagebox.showinfo("¡Perfecto!", f"Informe guardado correctamente en:\n{ruta_archivo}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el PDF:\n{str(e)}")

    def generar_graficas(self, respuestas, sub_sumas, gsi):

        # ====================== NORMAS ESPAÑOLAS (mujeres y hombres) ======================
        # Medias por ítem para convertir a T (aproximado T=63 ≈ media + 1.5 DT en población joven española)
        normas_es = {
            "Somatización":          {"H": 1.18, "M": 1.63},
            "Obsesividad-Compulsividad": {"H": 1.61, "M": 1.99},
            "Sensibilidad Interpersonal": {"H": 1.37, "M": 1.81},
            "Depresión":             {"H": 1.43, "M": 1.87},
            "Ansiedad":              {"H": 1.16, "M": 1.58},
            "Hostilidad":            {"H": 1.26, "M": 1.60},
            "Ansiedad Fóbica":       {"H": 0.73, "M": 1.00},
            "Ideación Paranoide":    {"H": 1.46, "M": 1.56},
            "Psicotismo":            {"H": 0.97, "M": 1.03},
        }

        sexo = self.entry_sexo.get().strip()

        # Histograma: Distribución Likert (0-4) de todas las respuestas
        fig1, ax1 = plt.subplots()
        unique, counts = np.unique(respuestas, return_counts=True)
        ax1.bar(unique, counts, color='#0078d4')
        ax1.set_title('Distribución de Respuestas (Likert 0-4)')
        ax1.set_xlabel('Valor Likert')
        ax1.set_ylabel('Frecuencia')
        ax1.set_xticks([0,1,2,3,4])

        # Barras con cortes clínicos
        etiquetas = list(normas_es.keys())

        # Medias por dimensión
        n_items = [12, 10, 9, 13, 10, 6, 7, 6, 10]
        medias = [sub_sumas[i] / n_items[i] for i in range(9)]

        # Cortes clínicos
        cortes_clinicos = []
        if sexo == 'Hombre':
            cortes_clinicos = [normas_es[dim]['H'] for dim in normas_es]
        else:
            cortes_clinicos = [normas_es[dim]['M'] for dim in normas_es]

        fig2, ax2 = plt.subplots(figsize=(11, 6))
        colores = ["#fc988dff" if m >= c else "#0078d4" for m, c in zip(medias, cortes_clinicos)]
        bars = ax2.bar(etiquetas, medias, color=colores, edgecolor='black', alpha=0.9)
    
        # Línea roja de corte clínico
        ax2.plot(etiquetas, cortes_clinicos, color="#fc1900", linewidth=3, linestyle='--', marker='o', 
                 label='Corte clínico (T≥63) - España')
    
        ax2.set_ylim(0, 4)
        ax2.set_ylabel('Media por ítem (0-4)')
        ax2.set_xticks(range(len(etiquetas)))
        ax2.set_xticklabels(etiquetas, rotation=45, ha='right', fontsize=8)
        ax2.set_title('Perfil SCL-90-R - Comparación con normas españolas')
        ax2.legend(fontsize=10)

        # Añadir valores encima de las barras
        for bar, val in zip(bars, medias):
            if val >= 0.3:
                ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() - 0.15,
                         f'{val:.2f}', ha='center', va='bottom', fontsize=10)

        # Guardar como imágenes
        self.img_hist = io.BytesIO()
        fig1.savefig(self.img_hist, format='png', bbox_inches='tight', dpi=150)
        self.img_hist.seek(0)
        plt.close(fig1)

        self.img_barras = io.BytesIO()
        fig2.savefig(self.img_barras, format='png', bbox_inches='tight', dpi=200)
        self.img_barras.seek(0)
        plt.close(fig2)

    def generar_pdf(self, nombre, fecha, terapeuta, resultados, scl90r_escalas, respuestas, gsi, pst, ruta_pdf):
        # === PROTECCIÓN CONTRA ERRORES ===
        if len(respuestas) != 90:
            messagebox.showerror("Error crítico", 
                f"Se esperaban 90 respuestas, pero se recibieron {len(respuestas)}.\n"
                "Revisa que todos las entradas tengan valor.")
            return

        doc = SimpleDocTemplate(ruta_pdf, pagesize=A4)
        story = []
        styles = getSampleStyleSheet()

        # ====================== NORMAS ESPAÑOLAS (mujeres y hombres) ======================
        # Medias por ítem para convertir a T (aproximado T=63 ≈ media + 1.5 DT en población joven española)
        normas_es = {
            "Somatización":          {"H": 1.18, "M": 1.63},
            "Obsesividad-Compulsividad": {"H": 1.61, "M": 1.99},
            "Sensibilidad Interpersonal": {"H": 1.37, "M": 1.81},
            "Depresión":             {"H": 1.43, "M": 1.87},
            "Ansiedad":              {"H": 1.16, "M": 1.58},
            "Hostilidad":            {"H": 1.26, "M": 1.60},
            "Ansiedad Fóbica":       {"H": 0.73, "M": 1.00},
            "Ideación Paranoide":    {"H": 1.46, "M": 1.56},
            "Psicotismo":            {"H": 0.97, "M": 1.03},
        }

        sexo = self.entry_sexo.get().strip()

        
        # Título
        story.append(Paragraph("Informe de resultados SCL-90-R", styles['Title']))
        story.append(Spacer(1, 20))

        # Nombre y fecha
        story.append(Table([[
            Paragraph(f"<b>Nombre:</b> {nombre}", styles['Normal']),
            Paragraph(f"<b> Sexo: </b> {sexo}"),
            Paragraph(f"<b>Fecha:</b> {fecha}", styles['Normal'])
        ]], colWidths=[260, 80, 100]))
        story.append(Table([[
            Paragraph(f"<b>Terapeuta:</b> {terapeuta}", styles['Normal']),
            Paragraph(""),
            Paragraph("")
        ]], colWidths=[300, 70, 70]))
        story.append(Spacer(1, 20))

        # === RESPUESTAS DEL PACIENTE EN 3 COLUMNAS (30 por columna) ===
        story.append(Paragraph(f"<b>Respuestas de {nombre}</b>", styles['Heading3']))
        story.append(Spacer(1, 12))

        def formatear_respuesta(i, valor):
            valor = str(valor).strip()
            if valor in ["3", "4"]:
                return Paragraph(f"Ítem {i+1:2d}: <font color='red'><b>{valor}</b></font>", styles['Normal'])
            elif valor in ["0"]:
                return Paragraph(f"Ítem {i+1:2d}: <font color='grey'>{valor}</font>", styles['Normal'])
            else:
                return Paragraph(f"Ítem {i+1:2d}: <font color='black'>{valor}</font>", styles['Normal'])
            
        # Creamos los datos: 30 filas, 3 columnas
        data_respuestas = []

        for fila in range(30):
            col1 = formatear_respuesta(fila, respuestas[fila])
            col2 = formatear_respuesta(fila+30, respuestas[fila+30])
            col3 = formatear_respuesta(fila+60, respuestas[fila+60])
            data_respuestas.append([col1, col2, col3])

        # Tabla bonita y compacta
        tabla_respuestas = Table(data_respuestas, colWidths=[140, 140, 140])

        tabla_respuestas.setStyle(TableStyle([
            ('FONTNAME',   (0,0), (-1,-1), 'Courier'),        # fuente monoespaciada = más legible en pequeño
            ('FONTSIZE',   (0,0), (-1,-1), 8.5),              # letra pequeña pero clara
            ('TEXTCOLOR',  (0,0), (-1,-1), colors.HexColor('#1a1a1a')),
            ('ALIGN',      (0,0), (-1,-1), 'LEFT'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('GRID',       (0,0), (-1,-1), 0.25, colors.HexColor('#e0e0e0')),  # rejilla muy fina
            ('LEFTPADDING',   (0,0), (-1,-1), 5),
            ('RIGHTPADDING',  (0,0), (-1,-1), 5),
            ('TOPPADDING',    (0,0), (-1,-1), 2),    # ← filas muy bajas
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),    # ← filas muy bajas
            ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ]))

        story.append(tabla_respuestas)
        story.append(Spacer(1, 40))

        # Índices globales
        ig = resultados["Índices Globales"]
        story.append(Paragraph("<b>Índices Globales</b>", styles['Heading2']))
        if  ig['GSI'] >= 1:
            story.append(Paragraph(f"• Índice de severidad global (GSI): <b> {ig['GSI']:.2f} </b>  (≥1.00 = malestar)", styles['Normal']))
            if  ig['GSI'] >= 1.5:
                story.append(Paragraph(f"• Índice de severidad global (GSI): <b><font color='red'> {ig['GSI']:.2f} </font></b>  (≥1.50 = caso clínico)", styles['Normal']))
        else:
            story.append(Paragraph(f"• Índice de severidad global (GSI): {ig['GSI']:.2f}  (≥1.00 = malestar | ≥1.50 = caso clínico)", styles['Normal']))
        
        if sexo == 'Hombre':
            if ig['PST'] > 60:
                story.append(Paragraph(f"• Total de síntomas positivos (PST): <b>{ig['PST']}</b>  (Riesgo de simulación)", styles['Normal']))
            else:
                story.append(Paragraph(f"• Total de síntomas positivos (PST): {ig['PST']} (Riesgo de simulación en hombres > 60)", styles['Normal']))
        else:
            if ig['PST'] > 70:
                story.append(Paragraph(f"• Total de síntomas positivos (PST): <b>{ig['PST']}</b>  (Riesgo de simulación)", styles['Normal']))
            else:
                story.append(Paragraph(f"• Total de síntomas positivos (PST): {ig['PST']} (Riesgo de simulación en mujeres > 70)", styles['Normal']))
        
        if ig['PSDI'] > 2.8:
            story.append(Paragraph(f"• Intensidad media de los síntomas positivos (PSDI): <b>{ig['PSDI']:.2f}</b>  (Posible dramatización)", styles['Normal']))
        else:
            story.append(Paragraph(f"• Intensidad media de los síntomas positivos (PSDI): {ig['PSDI']:.2f}  (Posible dramatización > 2.80)", styles['Normal']))

        story.append(Spacer(1, 20))

        # Subescalas
        story.append(Paragraph("<b>Puntuaciones por dimensión (Baremo español)</b>", styles['Heading2']))
        
        # Cortes clínicos
        cortes = []
        if sexo == 'Hombre':
            cortes = [normas_es[dim]['H'] for dim in normas_es]
        else:
            cortes = [normas_es[dim]['M'] for dim in normas_es]
        
        nombres_dim = list(scl90r_escalas.keys())

        data_dim = [["Dimensión", "Media", "Corte clínico", "Estado"]]

        for i, dim in enumerate(nombres_dim):
            media = resultados[dim]["media"]
            corte = cortes[i]
            if media >= corte:
                estado = Paragraph("<b><font color='red'>Clínico</font></b>", styles['Normal'])
            else:
                estado = Paragraph("Normal", styles['Normal'])
            data_dim.append([dim, f"{media:.2f}", f"{corte:.2f}", estado])
    
        t_dim = Table(data_dim, colWidths=[200, 80, 80, 80])
        t_dim.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0078d4")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#fff8f0')),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ]))
        story.append(t_dim)
        story.append(Spacer(1, 40))

        # Gráficas
        if hasattr(self, 'img_hist') and hasattr(self, 'img_barras'):
            #story.append(Paragraph("Distribución de respuestas (Likert 0-4)", styles['Heading3']))
            #story.append(Image(self.img_hist, width=300, height=200))
            #story.append(Spacer(1, 20))

            story.append(Paragraph("Puntuaciones por dimensión", styles['Heading3']))
            story.append(Image(self.img_barras, width=500, height=300))

        # Construir PDF
        try:
            doc.build(story)
            messagebox.showinfo("Éxito", "Informe generado correctamente")
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e))

# EJECUTAR
if __name__ == "__main__":
    root = tk.Tk()
    app = CorrectorPsicometrico(root)
    root.mainloop()