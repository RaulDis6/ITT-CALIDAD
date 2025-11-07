import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sqlalchemy import create_engine
import urllib
import sys
import os 




SERVER = 'DESKTOP-4K70KRA' 
DATABASE = 'ITT_Calidad'
DRIVER = '{ODBC Driver 17 for SQL Server}' 

CONNECTION_STRING_PYODBC = (f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;')


params = urllib.parse.quote_plus(CONNECTION_STRING_PYODBC)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={params}')


CARRERAS_ITT = [
    "ISC (Sistemas)", "IIA (Industrial)", "IME (Mecánica)", "IGE (Gestión)", 
    "IEQ (Electrónica)", "IEL (Eléctrica)", "ICA (Civil)", "ARQ (Arquitectura)", 
    "LA (Administración)", "LCP (Contador Público)"
]
SEMESTRES_LIST = [str(i) for i in range(1, 16)]



# FUNCIONES DE BACKEND


def conectar_sql_server():
    """Establece la conexión a SQL Server."""
    try:
        conn = pyodbc.connect(CONNECTION_STRING_PYODBC)
        return conn
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        error_msg = f"Error al conectar a SQL Server: {sqlstate}.\nRevisa tu DRIVER, ODBC y la variable SERVER."
        messagebox.showerror("Error de Conexión", error_msg)
        print(error_msg)
        return None

def log_actividad(user_id, tipo_accion, detalle=""):
    """Registra la actividad del usuario en la base de datos."""
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    
   
    user_id_safe = user_id if user_id is not None else None 
    
    sql_log = "INSERT INTO RegistroActividad (ID_Usuario_FK, Tipo_Accion, Detalle) VALUES (?, ?, ?)"
    try:
        cursor.execute(sql_log, (user_id_safe, tipo_accion, detalle[:4000])) 
        conn.commit()
    except Exception as e:
        print(f"Error al registrar actividad (ID {user_id_safe}): {e}") 
    finally:
        if conn: conn.close()

def autenticar_usuario(nombre_usuario, contrasena):
    """Verifica credenciales del usuario."""
    conn = conectar_sql_server()
    if not conn: return None
    cursor = conn.cursor()
    sql_auth = """
    SELECT ID_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK
    FROM Usuarios 
    WHERE Nombre_Usuario = ?
    """
    try:
        cursor.execute(sql_auth, (nombre_usuario,)) 
        usuario = cursor.fetchone()
        
        if usuario and usuario[1] == contrasena: 
            user_id, _, role_id, num_control = usuario
          
            log_actividad(user_id, 'LOGIN_EXITOSO', f'Usuario {nombre_usuario} (Rol: {role_id})')
            return user_id, role_id, num_control
        else:
            return None
    except Exception as e:
        print(f"Error al consultar usuario durante login: {e}")
        return None
    finally:
        if conn: conn.close()

def usuario_ya_existe(nombre_usuario):
    conn = conectar_sql_server()
    if not conn: 
        print("Error al verificar existencia: Fallo de conexión.")
        return True, "Error de conexión con la base de datos."
    
    sql_query = "SELECT 1 FROM Usuarios WHERE Nombre_Usuario = ?"
    try:
        cursor = conn.cursor()
        cursor.execute(sql_query, (nombre_usuario,))
        existe = cursor.fetchone() is not None
        return existe, None
    except Exception as e:
        print(f"Error al verificar existencia de usuario: {e}")
        return True, f"Error al consultar la base de datos: {e}" 
    finally:
        if conn: conn.close()

def registrar_estudiante_usuario(num_control, nombre, apellido_p, apellido_m, carrera, semestre, contrasena):
    """
    Registra un nuevo estudiante y su usuario asociado (Rol 2).
    """
    conn = conectar_sql_server()
    if not conn: return "Error de conexión con la base de datos."
    cursor = conn.cursor()
    
    try:
        semestre_int = int(semestre) 
    except ValueError:
        return "Error: El Semestre debe ser un número entero válido."

    existe, error_check = usuario_ya_existe(num_control)
    if error_check:
        return error_check
    if existe:
        return f"Error: El usuario o No. Control '{num_control}' ya está registrado en el sistema."

    sql_estudiante = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """  
    
    sql_usuario = """
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES (?, ?, 2, ?)
    """
    
    try:
        cursor.execute("BEGIN TRANSACTION")
        
      
        datos_estudiante_full = (
            num_control, apellido_p, apellido_m, nombre, carrera, semestre_int, 
            'Sin Materia', 
            0.0, 0.0, 0.0, 0.0, 0.0, 
            0.0, 
            0, 0, 0, 0, 0, 0 
        )
        
        cursor.execute(sql_estudiante, datos_estudiante_full)
        
        datos_usuario = (num_control, contrasena, num_control) 
        cursor.execute(sql_usuario, datos_usuario)
        
        conn.commit() 
        log_actividad(None, 'REGISTRO_NUEVO_ALUMNO', f"Registro exitoso de estudiante: {num_control}")
        return "Registro exitoso. ¡Ahora puedes iniciar sesión!"

    except pyodbc.IntegrityError as ex:
        conn.rollback()
        print(f"⚠️ ERROR DE INTEGRIDAD (SQL): {ex}") 
        return f"Error de integridad: El Número de Control {num_control} ya existe en una tabla relacionada."
        
    except Exception as e:
        conn.rollback()
        print(f"⚠️ ERROR DE SQL DETALLADO: {e}") 
        return f"Error al registrar: {e}"
    finally:
        if conn: conn.close()
        
def registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena):
    conn = conectar_sql_server()
    if not conn: return "Error de conexión con la base de datos."
    cursor = conn.cursor()
    
    existe, error_check = usuario_ya_existe(num_control)
    if error_check:
        return error_check
    if existe:
        return f"Error: El usuario o No. Control '{num_control}' ya está registrado en el sistema."
    
    sql_usuario = """
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES (?, ?, 1, ?)
    """
    
    try:
        cursor.execute("BEGIN TRANSACTION")
        
       
        datos_usuario = (num_control, contrasena, None) 
        cursor.execute(sql_usuario, datos_usuario)
        
        conn.commit() 
        log_actividad(None, 'REGISTRO_NUEVO_PROFESOR', f"Registro exitoso de profesor: {num_control}")
        return "Registro exitoso. ¡Ahora puedes iniciar sesión como Profesor!"

    except pyodbc.IntegrityError:
        conn.rollback()
        return f"Error de integridad: El Número de Control/Usuario {num_control} ya existe."
        
    except Exception as e:
        conn.rollback()
        return f"Error al registrar: {e}"
    finally:
        if conn: conn.close()

def insertar_registro_manual(datos, user_id):
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    
    sql_insert = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    try:
        if len(datos) != 19:
            raise ValueError(f"Error interno: Se esperaban 19 parámetros (columnas), pero se recibieron {len(datos)}. Verifica la función 'guardar_estudiante'.")

        cursor.execute(sql_insert, datos)
        conn.commit() 
        log_actividad(user_id, 'INSERT_ESTUDIANTE', f"Registro exitoso de Num_Control: {datos[0]}")
        messagebox.showinfo("Éxito", f"Registro de {datos[3]} exitoso.")
    except pyodbc.IntegrityError:
        messagebox.showerror("Error", f"Error: El estudiante con No. Control {datos[0]} ya existe.")
        conn.rollback()
    except ValueError as ve:
        messagebox.showerror("Error de Parámetros", str(ve))
        conn.rollback()
    except Exception as e:
        error_msg = str(e)
        if 'La conversión del valor varchar' in error_msg:
             messagebox.showerror("Error de Tipo de Dato", 
                                 "Error de conversión (varchar a int): Revise si Carrera y Semestre tienen tipos invertidos en SQL Server."
             )
        else:
            messagebox.showerror("Error de Inserción", f"Error al insertar datos: {error_msg}")
        conn.rollback()
    finally:
        if conn: conn.close()
        
def actualizar_datos_estudiante(num_control, nuevos_datos, user_id):
    conn = conectar_sql_server()
    if not conn: return False

    sql_update = """
    UPDATE Estudiantes 
    SET Apellido_Paterno = ?, Apellido_Materno = ?, Nombre = ? 
    WHERE Num_Control = ?
    """
    
    try:
        params = nuevos_datos + (num_control,) 
        cursor = conn.cursor()
        cursor.execute(sql_update, params)
        conn.commit() 
        
        if cursor.rowcount > 0:
            log_actividad(user_id, 'UPDATE_ESTUDIANTE', f"Actualizado Nombre/Apellidos de Num_Control: {num_control}")
            messagebox.showinfo("Éxito", "Datos actualizados correctamente.")
            return True
        else:
            messagebox.showerror("Error", f"No se encontró el estudiante con No. Control: {num_control}")
            return False

    except Exception as e:
        messagebox.showerror("Error de Actualización", f"Error al actualizar datos: {e}")
        conn.rollback()
        return False
    finally:
        if conn: conn.close()

def obtener_datos_estudiante(num_control):
    conn = conectar_sql_server()
    if not conn: return None
    
    sql_query = """
    SELECT 
        Nombre, Apellido_Paterno, Apellido_Materno, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, 
        Calificacion_Unidad_4, Calificacion_Unidad_5
    FROM Estudiantes
    WHERE Num_Control = ?
    """
    try:
        cursor = conn.cursor()
        cursor.execute(sql_query, num_control)
        return cursor.fetchone()
    except Exception as e:
        print(f"Error al obtener datos del estudiante: {e}")
        return None
    finally:
        if conn: conn.close()


def importar_datos_a_sql(archivo_path, nombre_tabla, user_id):
  
    COLUMNAS_BASE = ['Num_Control', 'Apellido_Paterno', 'Apellido_Materno', 'Nombre', 'Carrera', 'Semestre']
    
   
    COLUMNAS_ACADEMICAS_DEFECTO = {
        'Materia': 'Sin Asignar',
        'Calificacion_Unidad_1': 0.0, 'Calificacion_Unidad_2': 0.0, 'Calificacion_Unidad_3': 0.0,
        'Calificacion_Unidad_4': 0.0, 'Calificacion_Unidad_5': 0.0, 'Asistencia_Porcentaje': 0.0
    }
    COLUMNAS_FACTORES_DEFECTO = {
        'Factor_Academico': 0, 'Factor_Psicosocial': 0, 'Factor_Economico': 0,
        'Factor_Institucional': 0, 'Factor_Tecnologico': 0, 'Factor_Contextual': 0
    }
    
    try:
      
        if archivo_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(archivo_path, sheet_name=0)
        elif archivo_path.lower().endswith('.csv'):
            try:
                df = pd.read_csv(archivo_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(archivo_path, encoding='latin1')
        else:
            messagebox.showerror("Error", "Formato de archivo no soportado. Use .xlsx, .xls o .csv.")
            return

        
        if not all(col in df.columns for col in COLUMNAS_BASE):
            messagebox.showerror("Error de Columnas", 
                                 f"El archivo debe contener las siguientes 6 columnas base: {', '.join(COLUMNAS_BASE)}. Revise nombres exactos.")
            return

      
        df_final = df[COLUMNAS_BASE].copy()
        
     
        df_final['Semestre'] = df_final['Semestre'].astype(int)
        
        
        for col, val in COLUMNAS_ACADEMICAS_DEFECTO.items():
            df_final[col] = val
        for col, val in COLUMNAS_FACTORES_DEFECTO.items():
            df_final[col] = val
        
        
        df_final.to_sql(nombre_tabla, con=engine, if_exists='append', index=False, method='multi')
        
        log_actividad(user_id, 'IMPORT_DATA', f"Importadas {len(df_final)} filas (Solo 6 columnas principales) desde {archivo_path}")
        messagebox.showinfo("Importación Exitosa", f"¡{len(df_final)} filas insertadas en SQL Server! Las calificaciones y factores se inicializaron a 0.")

    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo no fue encontrado.")
    except Exception as e:
        error_msg = str(e)
        if 'La conversión del valor varchar' in error_msg:
             messagebox.showerror("Error Crítico de Tipo de Dato", 
                                 "Error de conversión: ¡AÚN DEBE CORREGIR LA TABLA EN SQL SERVER! "
                                 "Carrera y Semestre tienen tipos invertidos. Ejecute el script SQL proporcionado.")
        else:
            messagebox.showerror("Error de Importación", f"Ocurrió un error: {error_msg}. Revise su archivo y los tipos de datos en la BD.")



def exportar_datos_sql(formato, user_id):
    """Exporta todos los datos de Estudiantes a CSV o Excel."""
    try:
        
        df = pd.read_sql("SELECT * FROM Estudiantes", engine)
        
        
        if formato == 'excel':
            default_extension = ".xlsx"
            filetypes = [("Archivos de Excel", "*.xlsx")]
        elif formato == 'csv':
            default_extension = ".csv"
            filetypes = [("Archivos CSV", "*.csv")]
        else:
            return 

       
        filepath = filedialog.asksaveasfilename(
            defaultextension=default_extension,
            filetypes=filetypes,
            initialfile=f'Estudiantes_Export_{pd.Timestamp.now().strftime("%Y%m%d_%H%M")}'
        )
        
        if not filepath:
            messagebox.showinfo("Cancelado", "Exportación cancelada por el usuario.")
            return

       
        if formato == 'excel':
            df.to_excel(filepath, index=False)
        elif formato == 'csv':
            df.to_csv(filepath, index=False, encoding='utf-8')

        log_actividad(user_id, 'EXPORT_DATA', f"Exportados {len(df)} filas a {os.path.basename(filepath)}")
        messagebox.showinfo("Exportación Exitosa", f"Datos exportados a:\n{filepath}")

    except Exception as e:
        messagebox.showerror("Error de Exportación", f"Ocurrió un error al exportar los datos: {e}")

def generar_pareto_factores(user_id, carrera=None, semestre=None, materia=None, num_control_filtro=None):
    factores = ['Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico', 
                'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual']
    
 
    columnas_suma = [f'SUM(CAST({f} AS INT)) AS {f}' for f in factores]
    sql_query = f"SELECT {', '.join(columnas_suma)} FROM Estudiantes"
    
    condiciones = []
    if carrera:
        condiciones.append(f"Carrera = '{carrera}'")
    if semestre:
        try:
            semestre_int = int(semestre)
            condiciones.append(f"Semestre = {semestre_int}") 
        except ValueError:
            pass 
    if materia:
        condiciones.append(f"Materia = '{materia}'") 
    
    if num_control_filtro:
        condiciones = [f"Num_Control = '{num_control_filtro}'"]
        log_detalle = f"Consulta Pareto por Alumno: {num_control_filtro}"
    else:
        log_detalle = f"Consulta Pareto (Filtros: C={carrera}, S={semestre}, M={materia})"
        
    if condiciones:
        sql_query += " WHERE " + " AND ".join(condiciones)

    try:
        df = pd.read_sql(sql_query, engine)
        log_actividad(user_id, 'CONSULTA_PARETO', log_detalle)
        
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
            return None, "No se encontraron factores de riesgo marcados para este grupo/estudiante."

        df_factores['Porcentaje'] = (df_factores['Frecuencia'] / total_frecuencia) * 100
        df_factores['Acumulado'] = df_factores['Porcentaje'].cumsum()

        plt.clf() 
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
        log_actividad(user_id, 'ERROR_CONSULTA', f"Error en Pareto: {e}")
        return None, f"Error al consultar o generar gráfico: {e}"

def obtener_registro_auditoria():
    """
    Obtiene el registro de auditoría de acceso y cierre de sesión.
    Ordena por fecha descendente (más reciente primero).
    """
    conn = conectar_sql_server()
    if not conn: return pd.DataFrame()

    sql_query = """
    SELECT 
        U.Nombre_Usuario AS Matricula,
        R.Tipo_Accion AS Accion,
        R.Fecha_Hora AS Fecha_Hora
    FROM RegistroActividad R
    JOIN Usuarios U ON R.ID_Usuario_FK = U.ID_Usuario
    WHERE R.Tipo_Accion IN ('LOGIN_EXITOSO', 'LOGOUT')
    ORDER BY R.Fecha_Hora DESC
    """
    try:
        df = pd.read_sql(sql_query, engine)
        
       
        if not df.empty:
            df['Dia'] = df['Fecha_Hora'].dt.strftime('%Y-%m-%d')
            df['Hora'] = df['Fecha_Hora'].dt.strftime('%H:%M:%S')
            
           
            df['Accion'] = df['Accion'].replace({
                'LOGIN_EXITOSO': 'Inicio de Sesión',
                'LOGOUT': 'Cierre de Sesión'
            })
            
            
            df = df[['Matricula', 'Accion', 'Dia', 'Hora']]
            
        return df
    except Exception as e:
        print(f"Error al obtener el registro de auditoría: {e}")
        return pd.DataFrame()
    finally:
        if conn: conn.close()


# =========================================================================
# 🖼️ CLASES DE VENTANAS (GUI)
# =========================================================================

class StudentRegistrationWindow(tk.Toplevel):
    """Ventana de registro inicial para nuevos estudiantes."""
    def __init__(self, master):
        super().__init__(master)
        self.title("Registro de Estudiante")
        self.geometry("400x380")
        self.resizable(False, False)
        self.grab_set()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="**REGISTRO DE ESTUDIANTE**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        fields = [
            ("No. Control:", "num_control", 1, 'entry'), 
            ("Nombre(s):", "nombre", 2, 'entry'),
            ("Apellido Paterno:", "apellido_p", 3, 'entry'), 
            ("Apellido Materno:", "apellido_m", 4, 'entry'), 
            ("Carrera:", "carrera", 5, 'combobox'), 
            ("Semestre:", "semestre", 6, 'combobox'), 
            ("Contraseña:", "contrasena", 7, 'entry')
        ]

        self.reg_vars = {}
        for i, (label_text, var_name, row, widget_type) in enumerate(fields):
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky='w')
            var = tk.StringVar()
            self.reg_vars[var_name] = var
            
            if widget_type == 'entry':
                show = '*' if var_name == 'contrasena' else ''
                ttk.Entry(main_frame, textvariable=var, width=30, show=show).grid(row=row, column=1, padx=5, pady=5, sticky='w')
            elif widget_type == 'combobox': 
                if var_name == 'carrera':
                    combo = ttk.Combobox(main_frame, textvariable=var, values=CARRERAS_ITT, width=28, state='readonly')
                    combo.set(CARRERAS_ITT[0]) 
                elif var_name == 'semestre':
                    combo = ttk.Combobox(main_frame, textvariable=var, values=SEMESTRES_LIST, width=28, state='readonly')
                    combo.set(SEMESTRES_LIST[0]) 
                combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=8, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy).pack(side=tk.LEFT, padx=10)
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            carrera = self.reg_vars['carrera'].get().strip() 
            semestre_str = self.reg_vars['semestre'].get().strip()
            contrasena = self.reg_vars['contrasena'].get().strip()

            if not all([num_control, nombre, apellido_p, carrera, semestre_str, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
                return

            resultado = registrar_estudiante_usuario(num_control, nombre, apellido_p, apellido_m, carrera, semestre_str, contrasena)
            
            if "Error" in resultado:
                messagebox.showerror("Error de Registro", resultado)
            else:
                messagebox.showinfo("Registro Exitoso", resultado)
                self.destroy() 

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado durante la validación: {e}")

class TeacherRegistrationWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Registro de Profesor")
        self.geometry("400x280")
        self.resizable(False, False)
        self.grab_set()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="**REGISTRO DE PROFESOR**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        fields = [
            ("Usuario (Ej. No. Empleado):", "num_control", 1, ''), 
            ("Nombre(s):", "nombre", 2, ''),
            ("Apellido Paterno:", "apellido_p", 3, ''), 
            ("Apellido Materno:", "apellido_m", 4, ''), 
            ("Contraseña:", "contrasena", 5, '*')
        ]

        self.reg_vars = {}
        for i, (label_text, var_name, row, show_char) in enumerate(fields):
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky='w')
            var = tk.StringVar()
            self.reg_vars[var_name] = var
            show = '*' if var_name == 'contrasena' else ''
            ttk.Entry(main_frame, textvariable=var, width=30, show=show).grid(row=row, column=1, padx=5, pady=5, sticky='w')

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy).pack(side=tk.LEFT, padx=10)
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            contrasena = self.reg_vars['contrasena'].get().strip()

            if not all([num_control, nombre, apellido_p, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
                return

            resultado = registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena)
            
            if "Error" in resultado:
                messagebox.showerror("Error de Registro", resultado)
            else:
                messagebox.showinfo("Registro Exitoso", resultado)
                self.destroy() 

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado durante la validación: {e}")

class LoginWindow(tk.Toplevel):
    """Ventana de inicio de sesión."""
    def __init__(self, master):
        super().__init__(master)
        self.title("Inicio de Sesión")
        self.master = master 
        self.geometry("350x200")
        
        
        self.protocol("WM_DELETE_WINDOW", self.master.quit) 
        self.resizable(False, False)
        self.grab_set()
        
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
        self.create_widgets()

    def create_widgets(self):
        cred_frame = ttk.Frame(self, padding="10")
        cred_frame.grid(row=0, column=0, columnspan=2)

        ttk.Label(cred_frame, text="Usuario:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.user_var = tk.StringVar(value='ProfesorITT')
        ttk.Entry(cred_frame, textvariable=self.user_var, width=25).grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(cred_frame, text="Contraseña:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        self.pass_var = tk.StringVar(value='123')
        self.pass_entry = ttk.Entry(cred_frame, textvariable=self.pass_var, show="*", width=25)
        self.pass_entry.grid(row=1, column=1, padx=10, pady=5)
        self.pass_entry.bind('<Return>', lambda event: self.handle_login())

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Entrar", command=self.handle_login).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Registrar Alumno", command=self.open_student_registration).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Registrar Maestro", command=self.open_teacher_registration).pack(side=tk.LEFT, padx=5)

    def handle_login(self):
        username = self.user_var.get().strip()
        password = self.pass_var.get().strip()
        
        user_data = autenticar_usuario(username, password)
        
        if user_data:
            user_id, role_id, num_control = user_data
            self.master.show_main_window(user_id, role_id, num_control)
            self.destroy() 
        else:
            messagebox.showerror("Login Fallido", "Credenciales incorrectas o usuario no encontrado.", parent=self)
            try:
                log_actividad(None, 'LOGIN_FALLIDO', f'Intento fallido con usuario: {username}')
            except:
                pass

    def open_student_registration(self):
        StudentRegistrationWindow(self.master)

    def open_teacher_registration(self):
        TeacherRegistrationWindow(self.master)

class CalidadApp(ttk.Frame):
    """Ventana principal de la aplicación con pestañas."""
    def __init__(self, master, user_id, role_id, num_control):
        super().__init__(master)
        self.master.title(f"Sistema de Análisis de Calidad ITT - {'Profesor' if role_id == 1 else 'Estudiante'}")
        self.user_id = user_id
        self.role_id = role_id
        self.num_control = num_control
        self.pack(fill='both', expand=True)
        self.create_tabs()
        
    def create_tabs(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill='both', expand=True)

       
        self.tab_perfil = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_perfil, text='Perfil / Actualizar')
        self.crear_tab_perfil()
        
       
        if self.role_id == 1:
            self.tab_registro = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_registro, text='Registro Manual')
            self.crear_tab_registro()

            self.tab_pareto = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_pareto, text='Análisis de Riesgo (Pareto)')
            self.crear_tab_pareto()
            
            self.tab_importar = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_importar, text='Importar/Exportar')
            self.crear_tab_importar()
            
            self.tab_auditoria = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_auditoria, text='Registro de Actividad')
            self.crear_tab_auditoria()
            
        
        elif self.role_id == 2:
            self.tab_datos = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_datos, text='Mis Datos Académicos')
            self.crear_tab_datos_alumno()

    def cerrar_sesion(self):
        log_actividad(self.user_id, 'LOGOUT', 'Cierre de sesión manual.')
        
        self.master.main_window = None
        self.master.login_window = LoginWindow(self.master)
        self.master.geometry("350x200")
        self.destroy()

    #Pestaña de Registro Manual
    
    def crear_tab_registro(self):
        frame = ttk.Frame(self.tab_registro, padding="15")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**REGISTRO MANUAL DE ESTUDIANTES**", font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=4, pady=10)

        self.entry_vars = {}
        fields = [
            ("No. Control:", "Num_Control", 1, 'entry'),
            ("Apellido Paterno:", "Apellido_Paterno", 2, 'entry'),
            ("Apellido Materno:", "Apellido_Materno", 3, 'entry'),
            ("Nombre(s):", "Nombre", 4, 'entry'),
            ("Carrera:", "Carrera", 5, 'combobox'),
            ("Semestre:", "Semestre", 6, 'combobox'),
            ("Materia:", "Materia", 7, 'entry'),
            ("Asistencia %:", "Asistencia_Porcentaje", 8, 'entry')
        ]

        for i, (label_text, var_name, row, widget_type) in enumerate(fields):
            ttk.Label(frame, text=label_text).grid(row=row, column=0, padx=5, pady=2, sticky='w')
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            
            if widget_type == 'entry':
                ttk.Entry(frame, textvariable=var, width=30).grid(row=row, column=1, padx=5, pady=2, sticky='w')
            elif widget_type == 'combobox': 
                if var_name == 'Carrera':
                    combo = ttk.Combobox(frame, textvariable=var, values=CARRERAS_ITT, width=28, state='readonly')
                    combo.set(CARRERAS_ITT[0])
                elif var_name == 'Semestre':
                    combo = ttk.Combobox(frame, textvariable=var, values=SEMESTRES_LIST, width=28, state='readonly')
                    combo.set(SEMESTRES_LIST[0])
                combo.grid(row=row, column=1, padx=5, pady=2, sticky='w')

        # Calificaciones
        calif_frame = ttk.LabelFrame(frame, text="Calificaciones")
        calif_frame.grid(row=1, column=2, rowspan=5, padx=10, pady=5, sticky='n')
        for i in range(1, 6):
            ttk.Label(calif_frame, text=f"Unidad {i}:").grid(row=i, column=0, padx=5, pady=2, sticky='w')
            var = tk.StringVar()
            self.entry_vars[f"Calificacion_Unidad_{i}"] = var
            ttk.Entry(calif_frame, textvariable=var, width=10).grid(row=i, column=1, padx=5, pady=2, sticky='w')

        # Factores de Riesgo
        fact_frame = ttk.LabelFrame(frame, text="Factores de Riesgo")
        fact_frame.grid(row=6, column=2, rowspan=3, padx=10, pady=5, sticky='n')
        self.factor_vars = {}
        factores = ["Factor_Academico", "Factor_Psicosocial", "Factor_Economico", "Factor_Institucional", "Factor_Tecnologico", "Factor_Contextual"]
        nombres_fact = ["Académico", "Psicosocial", "Económico", "Institucional", "Tecnológico", "Contextual"]
        for i, (var_name, name) in enumerate(zip(factores, nombres_fact)):
            var = tk.IntVar()
            self.factor_vars[var_name] = var
            ttk.Checkbutton(fact_frame, text=name, variable=var).grid(row=i//2, column=i%2, padx=5, pady=2, sticky='w')

       
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=9, column=0, columnspan=4, pady=20)
        ttk.Button(btn_frame, text="Guardar Estudiante", command=self.guardar_estudiante).pack(padx=10)
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar_registro).pack(padx=10)
        
    def limpiar_registro(self):
        for var in self.entry_vars.values():
            var.set('')
        for var in self.factor_vars.values():
            var.set(0)
            
    def guardar_estudiante(self):
        try:
            #Recolección de datos principales
            datos_principales = [self.entry_vars[name].get().strip() for name in ['Num_Control', 'Apellido_Paterno', 'Apellido_Materno', 'Nombre', 'Carrera']]
            semestre_raw = self.entry_vars['Semestre'].get().strip()
            materia = self.entry_vars['Materia'].get().strip()
            
            if not all(datos_principales) or not semestre_raw or not materia:
                messagebox.showwarning("Advertencia", "Los campos principales (Control, Nombre, Apellidos, Carrera, Semestre, Materia) son obligatorios.")
                return
            
            #Validación y conversión de tipos
            try:
                semestre_final = int(semestre_raw)
                datos_principales.append(semestre_final)
                datos_principales.append(materia)
            except ValueError:
                messagebox.showerror("Error de Validación", "El Semestre debe ser un número entero.")
                return

            calificaciones_raw = [self.entry_vars[f"Calificacion_Unidad_{i}"].get().strip() for i in range(1, 6)]
            asistencia_raw = self.entry_vars['Asistencia_Porcentaje'].get().strip()
            
            try:
               
                calificaciones_finales = [float(c) if c else 0.0 for c in calificaciones_raw]
                asistencia_final = float(asistencia_raw) if asistencia_raw else 0.0
            except ValueError:
                messagebox.showerror("Error de Validación", "Las calificaciones y la asistencia deben ser números válidos.")
                return
            
            #Recolección de Factores
            factores = [self.factor_vars[name].get() for name in self.factor_vars.keys()]

           
            datos_completos = datos_principales + calificaciones_finales + [asistencia_final] + factores
            
            #Insertar en la base de datos
            insertar_registro_manual(datos_completos, self.user_id)
            self.limpiar_registro()

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado al guardar: {e}")

    #Pestaña de Perfil
    
    def crear_tab_perfil(self):
        
        frame = ttk.Frame(self.tab_perfil, padding="20")
        frame.pack(fill='both', expand=True)

        self.perfil_vars = {}
        datos_estudiante = obtener_datos_estudiante(self.num_control) if self.role_id == 2 else None

        ttk.Label(frame, text=f"**PERFIL DE USUARIO**", font=('Arial', 14, 'bold')).pack(pady=10)
        
        
        data_map = {
            "No. Control:": self.num_control if self.role_id == 2 else self.master.login_window.user_var.get(),
            "Nombre(s):": datos_estudiante[0] if datos_estudiante and len(datos_estudiante) > 0 else "",
            "Apellido Paterno:": datos_estudiante[1] if datos_estudiante and len(datos_estudiante) > 1 else "",
            "Apellido Materno:": datos_estudiante[2] if datos_estudiante and len(datos_estudiante) > 2 else "",
            "Carrera:": datos_estudiante[3] if datos_estudiante and len(datos_estudiante) > 3 else "N/A (Profesor)",
            "Semestre:": datos_estudiante[4] if datos_estudiante and len(datos_estudiante) > 4 else "N/A (Profesor)"
        }
        
        fields = [
            ("No. Control:", "Num_Control"), 
            ("Nombre(s):", "Nombre"), 
            ("Apellido Paterno:", "Apellido_Paterno"), 
            ("Apellido Materno:", "Apellido_Materno"), 
            ("Carrera:", "Carrera"), 
            ("Semestre:", "Semestre")
        ]
        
        for label_text, var_name in fields:
            sub_frame = ttk.Frame(frame)
            sub_frame.pack(fill='x', pady=5, padx=50)
            
            ttk.Label(sub_frame, text=label_text, width=15, anchor='w').pack(side=tk.LEFT)
            
            var = tk.StringVar()
            self.perfil_vars[var_name] = var
            var.set(data_map.get(label_text, ""))
            
            is_editable = var_name in ["Nombre", "Apellido_Paterno", "Apellido_Materno"]
            state = 'normal' if is_editable and self.role_id == 2 else 'readonly'
            
            ttk.Entry(sub_frame, textvariable=var, width=30, state=state).pack(side=tk.LEFT, fill='x', expand=True)

        if self.role_id == 1:
            ttk.Label(frame, text="Rol: Profesor", font=('Arial', 10, 'italic')).pack(pady=10, padx=50, anchor='w')
            
        if self.role_id == 2:
            ttk.Label(frame, text="Solo puedes editar tu Nombre y Apellidos.", font=('Arial', 10, 'italic')).pack(pady=5, padx=50, anchor='w')
       
             
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)
        ttk.Button(frame, text="Actualizar Datos", command=self.actualizar_perfil).pack(pady=5)
        ttk.Button(frame, text="Cerrar Sesión", command=self.cerrar_sesion).pack(pady=5)

    def actualizar_perfil(self):
        if self.role_id != 2:
            messagebox.showinfo("Información", "Solo los estudiantes pueden actualizar sus datos desde esta ventana.")
            return

        nombre = self.perfil_vars['Nombre'].get().strip()
        apellido_p = self.perfil_vars['Apellido_Paterno'].get().strip()
        apellido_m = self.perfil_vars['Apellido_Materno'].get().strip()
        
        if not all([nombre, apellido_p]):
             messagebox.showwarning("Advertencia", "El nombre y el apellido paterno son obligatorios.")
             return
        
        nuevos_datos = [apellido_p, apellido_m, nombre]
        actualizar_datos_estudiante(self.num_control, nuevos_datos, self.user_id)
        
    #Pestaña de Datos Alumno
    
    def crear_tab_datos_alumno(self):
        
        frame = ttk.Frame(self.tab_datos, padding="20")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**MIS DATOS ACADÉMICOS**", font=('Arial', 14, 'bold')).pack(pady=10)
        
        datos = obtener_datos_estudiante(self.num_control)

        if datos:
            labels = ["Nombre:", "Apellido Paterno:", "Apellido Materno:", "Carrera:", "Semestre:", "Materia:", 
                      "Calificación U1:", "Calificación U2:", "Calificación U3:", "Calificación U4:", "Calificación U5:"]
            
            
            info_frame = ttk.Frame(frame)
            info_frame.pack(pady=10, padx=50, anchor='w')
            
            for i, label_text in enumerate(labels):
                value = datos[i] if i < len(datos) else "N/D"
                ttk.Label(info_frame, text=label_text, width=20, anchor='w').grid(row=i, column=0, sticky='w')
                ttk.Label(info_frame, text=str(value), font=('Arial', 10, 'bold')).grid(row=i, column=1, sticky='w')
        else:
            ttk.Label(frame, text="No se encontraron datos académicos para su número de control.").pack(pady=20)

    #Pestaña de Pareto
    
    def crear_tab_pareto(self):
        frame = ttk.Frame(self.tab_pareto, padding="10")
        frame.pack(fill='both', expand=True)
        
        control_frame = ttk.LabelFrame(frame, text="Filtros")
        control_frame.pack(pady=10, padx=10, fill='x')
        
        estado_filtro = 'disabled' if self.role_id == 2 else 'readonly'
        
       
        ttk.Label(control_frame, text="Carrera:").grid(row=0, column=0, padx=5, pady=5)
        self.pareto_carrera = tk.StringVar()
        combo_carrera = ttk.Combobox(control_frame, textvariable=self.pareto_carrera, values=CARRERAS_ITT, width=15, state=estado_filtro)
        combo_carrera.grid(row=0, column=1, padx=5, pady=5)
        
      
        ttk.Label(control_frame, text="Semestre:").grid(row=0, column=2, padx=5, pady=5)
        self.pareto_semestre = tk.StringVar()
        combo_semestre = ttk.Combobox(control_frame, textvariable=self.pareto_semestre, values=SEMESTRES_LIST, width=15, state=estado_filtro)
        combo_semestre.grid(row=0, column=3, padx=5, pady=5)
        
       
        ttk.Label(control_frame, text="Materia:").grid(row=1, column=0, padx=5, pady=5)
        self.pareto_materia = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.pareto_materia, width=15, state='disabled' if self.role_id == 2 else 'normal').grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(control_frame, text="Generar Gráfico", command=self.mostrar_pareto).grid(row=1, column=3, columnspan=2, padx=15, pady=5)

        self.pareto_canvas_frame = ttk.Frame(frame)
        self.pareto_canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.pareto_fig_canvas = None

    def mostrar_pareto(self):
        
        for widget in self.pareto_canvas_frame.winfo_children():
            widget.destroy()
        
        carrera = self.pareto_carrera.get().strip()
        semestre = self.pareto_semestre.get().strip()
        materia = self.pareto_materia.get().strip()

     
        if self.role_id == 2:
            num_control_filtro = self.num_control
            carrera, semestre, materia = None, None, None 
        else:
            num_control_filtro = None

        fig, error_msg = generar_pareto_factores(self.user_id, carrera, semestre, materia, num_control_filtro)
        
        if error_msg:
            messagebox.showerror("Error de Gráfico", error_msg)
            return

        if fig:
            self.pareto_fig_canvas = FigureCanvasTkAgg(fig, master=self.pareto_canvas_frame)
            self.pareto_fig_canvas.draw()
            self.pareto_fig_canvas.get_tk_widget().pack(fill='both', expand=True)

    #Pestaña de Importar/Exportar
    
    def crear_tab_importar(self):
        frame = ttk.Frame(self.tab_importar, padding="10")
        frame.pack(fill='both', expand=True)
        
        if self.role_id == 1:
            #Sección Importar
            import_frame = ttk.LabelFrame(frame, text="Importar Datos de Estudiantes (CSV/Excel)")
            import_frame.pack(pady=20, padx=20, fill='x')
            
            ttk.Label(import_frame, text="Archivo (CSV/Excel):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            self.import_path_var = tk.StringVar()
            ttk.Entry(import_frame, textvariable=self.import_path_var, width=50, state='readonly').grid(row=0, column=1, padx=5, pady=5)
            
            ttk.Button(import_frame, text="Seleccionar Archivo", command=self.seleccionar_archivo_import).grid(row=0, column=2, padx=5, pady=5)
            ttk.Button(import_frame, text="Ejecutar Importación", command=self.ejecutar_importacion_handler).grid(row=1, column=1, columnspan=2, pady=10)

            #Sección Exportar
            export_frame = ttk.LabelFrame(frame, text="Exportar Datos de Estudiantes")
            export_frame.pack(pady=20, padx=20, fill='x')
            
            export_btns_frame = ttk.Frame(export_frame)
            export_btns_frame.pack(pady=10)
            
            ttk.Button(export_btns_frame, text="Exportar a Excel (.xlsx)", command=lambda: exportar_datos_sql('excel', self.user_id)).pack(side=tk.LEFT, padx=10)
            ttk.Button(export_btns_frame, text="Exportar a CSV (.csv)", command=lambda: exportar_datos_sql('csv', self.user_id)).pack(side=tk.LEFT, padx=10)
        else:
            ttk.Label(frame, text="Esta pestaña solo está disponible para el rol de Profesor.").pack(pady=20)
            
    def seleccionar_archivo_import(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Archivos de Datos", "*.csv *.xlsx *.xls")],
            title="Seleccionar archivo de importación"
        )
        if filepath:
            self.import_path_var.set(filepath)
            
    def ejecutar_importacion_handler(self):
        archivo_path = self.import_path_var.get()
        if not archivo_path:
            messagebox.showwarning("Advertencia", "Por favor, seleccione un archivo primero.")
            return
        
        
        importar_datos_a_sql(archivo_path, 'Estudiantes', self.user_id)

    #Pestaña de Auditoría
    
    def crear_tab_auditoria(self):
        frame = ttk.Frame(self.tab_auditoria, padding="10")
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="**REGISTRO DE INICIOS Y CIERRES DE SESIÓN**", font=('Arial', 14, 'bold')).pack(pady=10)
        
        
        columns = ("Matricula", "Accion", "Dia", "Hora")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center', width=150)
            
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill='both', expand=True)

       
        ttk.Button(frame, text="Actualizar Registro", command=self.cargar_auditoria).pack(pady=10)
        
       
        self.cargar_auditoria()
        
    def cargar_auditoria(self):
        """Carga los datos de auditoría en el Treeview."""
        if self.role_id != 1: return 
        
      
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        df = obtener_registro_auditoria()
        log_actividad(self.user_id, 'CONSULTA_AUDITORIA', 'Carga de registro de acceso')
        
        if df.empty:
            self.tree.insert("", "end", values=("No hay registros de auditoría.", "", "", ""))
            return

      
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=(row['Matricula'], row['Accion'], row['Dia'], row['Hora']))


# CLASE PRINCIPAL 


class MainApp(tk.Tk):
    """Clase principal que gestiona el ciclo de vida de la aplicación y las ventanas."""
    def __init__(self):
        super().__init__()
        self.title("Sistema de Análisis de Calidad ITT")
        self.geometry("1x1") 
        self.withdraw() 

        # --- APLICACIÓN DE DISEÑO MEJORADO (TEMA TTK) ---
        style = ttk.Style(self)
        # Aplicamos el tema 'clam' que es más personalizable que el default
        # Esto le da un aspecto más moderno que los temas clásicos.
        style.theme_use('clam') 
        
        # Estilos generales para un look más definido y espacioso
        style.configure('TButton', font=('Arial', 10, 'bold'), padding=6)
        style.configure('TLabel', font=('Arial', 10))
        style.configure('TNotebook.Tab', padding=[10, 5]) 
        # --- FIN DISEÑO MEJORADO ---

        self.main_window = None

        
        self.login_window = LoginWindow(self)
    
    def show_main_window(self, user_id, role_id, num_control):
        
        if self.main_window:
            self.main_window.destroy() 
            
        
        self.main_window = CalidadApp(self, user_id, role_id, num_control)
        
        # Muestra la ventana principal
        self.geometry("950x700") 
        self.deiconify() 
        self.main_window.lift() 
        self.main_window.focus_force()

  

if __name__ == "__main__":
    if conectar_sql_server():
        try:
            app = MainApp() 
            app.mainloop()
        except Exception as e:

            messagebox.showerror("Error Fatal de Aplicación", f"Ocurrió un error inesperado al iniciar la aplicación: {e}")
            sys.exit(1)