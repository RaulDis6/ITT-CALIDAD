import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sqlalchemy import create_engine
import urllib
import sys

#CONFIGURACIÓN DE CONEXIÓN Y BACKEND
SERVER = 'DESKTOP-4K70KRA'
DATABASE = 'ITT_Calidad'
DRIVER = '{ODBC Driver 17 for SQL Server}' 

CONNECTION_STRING_PYODBC = (
    f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;'
)

#Configuración para Pandas y SQLAlchemy
params = urllib.parse.quote_plus(CONNECTION_STRING_PYODBC)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={params}')


def conectar_sql_server():
    """Establece la conexión a SQL Server."""
    try:
        conn = pyodbc.connect(CONNECTION_STRING_PYODBC)
        return conn
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        messagebox.showerror("Error de Conexión", 
        f"Error al conectar a SQL Server: {sqlstate}")
        return None

def insertar_registro_manual(datos):
    """Inserta un único registro."""
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    
    #SQL INSERT statement
    sql_insert = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    try:
        cursor.execute(sql_insert, datos)
        conn.commit()
        messagebox.showinfo("Éxito", f"Registro de {datos[3]} exitoso.")
    except pyodbc.IntegrityError:
        messagebox.showerror("Error", f"Error: El estudiante con No. Control {datos[0]} ya existe.")
        conn.rollback()
    except Exception as e:
        messagebox.showerror("Error de Inserción", f"Error al insertar datos: {e}")
        conn.rollback()
    finally:
        if conn: conn.close()

#IMPORTACIÓN DE EXCEL
def importar_excel_a_sql(archivo_excel, nombre_tabla='Estudiantes'):    
    try:
        df = pd.read_excel(archivo_excel, sheet_name=0)
        df.to_sql(nombre_tabla, con=engine, if_exists='append', index=False)
        messagebox.showinfo("Importación Exitosa", f"¡{len(df)} filas insertadas en SQL Server!")

    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo Excel '{archivo_excel}' no fue encontrado.")
    except Exception as e:
        messagebox.showerror("Error de Importación", f"Ocurrió un error: {e}. Revise formatos.")

#Diagrama de Pareto
def generar_pareto_factores(carrera=None, semestre=None):
    """Genera el Diagrama de Pareto y devuelve la figura de Matplotlib."""
    
    factores = ['Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico', 
                'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual']
    columnas_suma = [f'SUM(CAST({f} AS INT)) AS {f}' for f in factores]
    sql_query = f"SELECT {', '.join(columnas_suma)} FROM Estudiantes"
    
    condiciones = []
    if carrera:
        condiciones.append(f"Carrera = '{carrera}'")
    if semestre:
        condiciones.append(f"Semestre = {semestre}") 
        
    if condiciones:
        sql_query += " WHERE " + " AND ".join(condiciones)

    try:
        df = pd.read_sql(sql_query, engine)
        df_factores = df.T.rename(columns={0: 'Frecuencia'})
        df_factores = df_factores[df_factores['Frecuencia'] > 0]
        
        nombres_grafico = {
            'Factor_Academico': 'Académico','Factor_Psicosocial': 'Psicosocial',
            'Factor_Economico': 'Económico','Factor_Institucional': 'Institucional',
            'Factor_Tecnologico': 'Tecnológico','Factor_Contextual': 'Contextual'
        }
        df_factores.index = df_factores.index.map(nombres_grafico)
        df_factores = df_factores.sort_values(by='Frecuencia', ascending=False)
        total_frecuencia = df_factores['Frecuencia'].sum()
        
        if total_frecuencia == 0:
            return None, "No se encontraron factores de riesgo marcados para este grupo."

        df_factores['Porcentaje'] = (df_factores['Frecuencia'] / total_frecuencia) * 100
        df_factores['Acumulado'] = df_factores['Porcentaje'].cumsum()

        fig, facto1 = plt.subplots(figsize=(8, 5))
        
        facto1.bar(df_factores.index, df_factores['Frecuencia'], color='tab:blue')
        facto1.set_xlabel('Factores de Riesgo')
        facto1.set_ylabel('Frecuencia', color='tab:blue')
        facto1.tick_params(axis='y', labelcolor='tab:blue')
        
        paret = facto1.twinx()
        paret.plot(df_factores.index, df_factores['Acumulado'], color='tab:red', marker='o')
        paret.set_ylabel('Acumulado %', color='tab:red')
        paret.tick_params(axis='y', labelcolor='tab:red')
        paret.set_ylim(0, 100)
        paret.axhline(80, color='gray', linestyle='--')

        plt.title("Análisis de Pareto: Factores de Riesgo")
        fig.tight_layout()
        
        return fig, None

    except Exception as e:
        return None, f"Error al consultar o generar gráfico: {e}"

#CLASE PRINCIPAL (DISEÑO)

class CalidadApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Análisis de Calidad ITT")
        self.geometry("950x700") 

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        #Pestañas 
        self.tab_registro = ttk.Frame(self.notebook)
        self.tab_pareto = ttk.Frame(self.notebook)
        self.tab_histograma = ttk.Frame(self.notebook)
        self.tab_dispersion = ttk.Frame(self.notebook) 
        self.tab_ishikawa = ttk.Frame(self.notebook) 
        self.tab_importar = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_registro, text='Registro de Estudiante')
        self.notebook.add(self.tab_pareto, text='Análisis de Pareto')
        self.notebook.add(self.tab_histograma, text='Histogramas')
        self.notebook.add(self.tab_dispersion, text='Dispersión')
        self.notebook.add(self.tab_ishikawa, text='Ishikawa')
        self.notebook.add(self.tab_importar, text='Importar/Exportar Datos')

        #Inicialización de las pestañas
        self.crear_tab_registro()
        self.crear_tab_pareto()
        self.crear_tab_histograma() 
        self.crear_tab_dispersion() 
        self.crear_tab_ishikawa() 
        self.crear_tab_importar()

    #FUNCIÓN DE LIMPIAR
    def limpiar_registro(self):    
        if hasattr(self, 'entry_vars'):
            for var in self.entry_vars.values():
                var.set("")
        if hasattr(self, 'factor_vars'):
            for var in self.factor_vars.values():
                var.set(0) 
        messagebox.showinfo("Limpieza", "Campos limpiados.")


    #REGISTRO DE ESTUDIANTE
    def crear_tab_registro(self):
        frame = ttk.Frame(self.tab_registro, padding="10")
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="**REGISTRO DE ESTUDIANTE**", font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=4, pady=10)
        
        #Variables de entrada
        self.entry_vars = {}
        fields = [
            ("No. Control:", "Num_Control", 1), ("Apellido Paterno:", "Apellido_Paterno", 2), 
            ("Apellido Materno:", "Apellido_Materno", 3), ("Nombre(s):", "Nombre", 4), 
            ("Carrera:", "Carrera", 5), ("Semestre:", "Semestre", 6), 
            ("Materia:", "Materia", 7), ("Asistencia %:", "Asistencia_Porcentaje", 8)
        ]
        
        for i, (label_text, var_name, row) in enumerate(fields):
            ttk.Label(frame, text=label_text).grid(row=row, column=0, padx=5, pady=2, sticky='w')
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            ttk.Entry(frame, textvariable=var, width=30).grid(row=row, column=1, padx=5, pady=2, sticky='w')

        #Calificaciones
        calif_frame = ttk.LabelFrame(frame, text="Calificaciones")
        calif_frame.grid(row=1, column=2, rowspan=5, padx=10, pady=5, sticky='n')
        for i in range(1, 6):
            ttk.Label(calif_frame, text=f"Unidad {i}:").grid(row=i, column=0, padx=5, pady=2, sticky='w')
            var = tk.StringVar()
            self.entry_vars[f"Calificacion_Unidad_{i}"] = var
            ttk.Entry(calif_frame, textvariable=var, width=10).grid(row=i, column=1, padx=5, pady=2, sticky='w')

        #Factores de Riesgo
        fact_frame = ttk.LabelFrame(frame, text="Factores de Riesgo")
        fact_frame.grid(row=6, column=2, rowspan=3, padx=10, pady=5, sticky='n')
        self.factor_vars = {}
        factores = ["Factor_Academico", "Factor_Psicosocial", "Factor_Economico", "Factor_Institucional", "Factor_Tecnologico", "Factor_Contextual"]
        nombres_fact = ["Académico", "Psicosocial", "Económico", "Institucional", "Tecnológico", "Contextual"]
        
        for i, (var_name, name) in enumerate(zip(factores, nombres_fact)):
            var = tk.IntVar()
            self.factor_vars[var_name] = var
            ttk.Checkbutton(fact_frame, text=name, variable=var).grid(row=i // 2, column=i % 2, padx=5, pady=2, sticky='w')

        #Botones Guardar y Limpiar
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=9, column=1, columnspan=2, pady=20, sticky='w')
        ttk.Button(btn_frame, text="Guardar Registro", command=self.guardar_estudiante).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar_registro).pack(side=tk.LEFT, padx=10)


    def guardar_estudiante(self):        
        try:
            #Recoger datos principales para su insercion
            datos_principales = [self.entry_vars[name].get() for name in 
                                 ["Num_Control", "Apellido_Paterno", "Apellido_Materno", "Nombre", 
                                  "Carrera", "Semestre", "Materia"]]
            
            #Recoger calificaciones y asistencia
            calificaciones = [float(self.entry_vars[f"Calificacion_Unidad_{i}"].get() or 0.0) for i in range(1, 6)]
            asistencia = [float(self.entry_vars["Asistencia_Porcentaje"].get() or 0.0)]
            
            #Recoger factores de riesgo
            factores = [self.factor_vars[name].get() for name in self.factor_vars.keys()]
            
            datos_completos = datos_principales + calificaciones + asistencia + factores
            
            insertar_registro_manual(datos_completos)

        except ValueError:
            messagebox.showerror("Error de Validación", "Las calificaciones, semestre y asistencia deben ser números válidos.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")


    #Pestaña de ANÁLISIS DE PARETO
    def crear_tab_pareto(self):        
        frame = ttk.Frame(self.tab_pareto, padding="10")
        frame.pack(fill='both', expand=True)
        
        #Panel de Control
        control_frame = ttk.LabelFrame(frame, text="Filtros")
        control_frame.pack(pady=10, padx=10, fill='x')

        ttk.Label(control_frame, text="Carrera:").grid(row=0, column=0, padx=5, pady=5)
        self.pareto_carrera = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.pareto_carrera, width=15).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(control_frame, text="Semestre:").grid(row=0, column=2, padx=5, pady=5)
        self.pareto_semestre = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.pareto_semestre, width=15).grid(row=0, column=3, padx=5, pady=5)
        
        #Botón Generar Gráfico
        ttk.Button(control_frame, text="Generar Gráfico", command=self.mostrar_pareto).grid(row=0, column=4, padx=15, pady=5)

        #Área de Gráfico
        self.pareto_canvas_frame = ttk.Frame(frame)
        self.pareto_canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.pareto_fig_canvas = None 

    def mostrar_pareto(self):
        
        carrera = self.pareto_carrera.get().strip().upper() if self.pareto_carrera.get() else None
        semestre_str = self.pareto_semestre.get().strip()
        semestre = int(semestre_str) if semestre_str.isdigit() else None
        
        #Limpiar canvas anterior
        for widget in self.pareto_canvas_frame.winfo_children():
            widget.destroy()
        
        fig, error = generar_pareto_factores(carrera, semestre)
        
        if error:
            messagebox.showinfo("Error / Info", error)
            return

        #Insertar la figura de Matplotlib en el Canvas de Tkinter
        self.pareto_fig_canvas = FigureCanvasTkAgg(fig, master=self.pareto_canvas_frame)
        self.pareto_fig_canvas.draw()
        self.pareto_fig_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)


    #Pestaña de HISTOGRAMAS
    def crear_tab_histograma(self):
        frame = ttk.Frame(self.tab_histograma, padding="10")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**HISTOGRAMAS**", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(frame, text="Lógica de Histograma.", font=('Arial', 12)).pack(pady=50)

    #Pestaña de DISPERSIÓN
    def crear_tab_dispersion(self):
        frame = ttk.Frame(self.tab_dispersion, padding="10")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**DIAGRAMA DE DISPERSIÓN**", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(frame, text="Lógica de Dispersión.", font=('Arial', 12)).pack(pady=50)

    #Pestaña de ISHIKAWA
    def crear_tab_ishikawa(self):
        frame = ttk.Frame(self.tab_ishikawa, padding="10")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**DIAGRAMA DE ISHIKAWA (CAUSA-EFECTO)**", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(frame, text="Lógica de Ishikawa.", font=('Arial', 12)).pack(pady=50)

    #Pestaña de IMPORTAR/EXPORTAR DATOS
    def crear_tab_importar(self):
        frame = ttk.Frame(self.tab_importar, padding="10")
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="**IMPORTACIÓN DE DATOS**", font=('Arial', 14, 'bold')).pack(pady=10)

        #Importar Excel 
        ttk.Label(frame, text="Ruta del Archivo Excel (.xlsx):").pack(pady=5)
        self.import_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.import_path_var, width=50).pack(pady=5)
        ttk.Button(frame, text="Seleccionar Archivo", command=self.seleccionar_archivo).pack(pady=5)
        
        #Conectaa el botón a la función de importación
        ttk.Button(frame, text="Importar a SQL Server", command=self.ejecutar_importacion).pack(pady=10)
        
        #Exportar
        ttk.Label(frame, text="\n**EXPORTACIÓN DE DATOS Y GRÁFICOS**", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(frame, text="Pendientes", font=('Arial', 12)).pack(pady=50)
        ttk.Button(frame, text="Exportar a CSV/Excel", command=lambda: messagebox.showinfo("Pendiente")).pack(pady=5)
        ttk.Button(frame, text="Generar Reporte PDF", command=lambda: messagebox.showinfo("Pendiente")).pack(pady=5)

    def seleccionar_archivo(self):      
        filepath = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if filepath:
            self.import_path_var.set(filepath)
    
    def ejecutar_importacion(self):
       
        ruta = self.import_path_var.get()
        if not ruta:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo Excel.")
            return
        importar_excel_a_sql(ruta)


if __name__ == "__main__":
    app = CalidadApp()
    app.mainloop()