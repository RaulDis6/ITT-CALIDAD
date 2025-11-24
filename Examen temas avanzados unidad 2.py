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
import re 
import tkinter.font # Importar para usar la fuente de Tkinter directamente

# [ ... EL CÓDIGO INICIAL DE CONFIGURACIÓN Y CONSTANTES ES EL MISMO ... ]

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
DISCAPACIDADES_LIST = ["Ninguna", "Visual", "Auditiva", "Motriz", "Cognitiva", "Otra"]


# ====================================================================
# VARIABLES GLOBALES DE ACCESIBILIDAD
# ====================================================================

# Usaremos un tamaño de fuente base y una variable global para el tamaño actual
BASE_FONT_SIZE = 10
CURRENT_FONT_SIZE = BASE_FONT_SIZE
COLOR_INVERTED = False 
ZOOM_LEVEL = 100 # Nivel de zoom en porcentaje (100% es normal)
# Definiciones de colores para la inversión
COLOR_MAP = {
    'white': '#000000',
    '#ffffff': '#000000',
    'black': '#ffffff',
    '#000000': '#ffffff',
    # Colores comunes de Tk/Ttk
    'SystemWindow': 'SystemWindowText',
    'SystemWindowText': 'SystemWindow',
    'SystemButtonFace': 'SystemButtonText',
    'SystemButtonText': 'SystemButtonFace',
    'SystemHighlight': 'SystemHighlightText',
    'SystemHighlightText': 'SystemHighlight',
    # Para Matplotlib (si fuera necesario, pero la gráfica usa colores fijos)
    'tab:blue': 'tab:red',
    'tab:red': 'tab:blue',
}

# ====================================================================
# FUNCIONES DE BACKEND (SIN CAMBIOS FUNCIONALES)
# ====================================================================

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

def registrar_estudiante_usuario(num_control, nombre, apellido_p, apellido_m, carrera, semestre, contrasena, discapacidad):
    """
    Registra un nuevo estudiante y su usuario asociado (Rol 2).
    Añade el campo 'Discapacidad'.
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

    # ¡IMPORTANTE! Se agregó 'Discapacidad'
    # Debes asegurar que la tabla 'Estudiantes' en SQL Server tiene una columna llamada 'Discapacidad' (e.g., VARCHAR(100))
    sql_estudiante = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            0, 0, 0, 0, 0, 0,
            discapacidad # <--- NUEVO CAMPO
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
        
def registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena, discapacidad):
    """
    Registra un nuevo usuario Profesor (Rol 1).
    Se usa el No. Control como Nombre_Usuario.
    Se agregó el campo 'Discapacidad' al log, ya que la tabla 'Usuarios' es general.
    """
    conn = conectar_sql_server()
    if not conn: return "Error de conexión con la base de datos."
    cursor = conn.cursor()
    
    existe, error_check = usuario_ya_existe(num_control)
    if error_check:
        return error_check
    if existe:
        return f"Error: El usuario o No. Control '{num_control}' ya está registrado en el sistema."
    
    sql_usuario = """
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK, Discapacidad) 
    VALUES (?, ?, 1, ?, ?)
    """
    # ¡IMPORTANTE! Para profesor, estoy asumiendo que el campo Discapacidad se agrega a la tabla Usuarios.
    # Debes asegurar que la tabla 'Usuarios' en SQL Server tiene una columna llamada 'Discapacidad' (e.g., VARCHAR(100))

    try:
        cursor.execute("BEGIN TRANSACTION")
        
       
        datos_usuario = (num_control, contrasena, None, discapacidad) # <--- NUEVO CAMPO
        cursor.execute(sql_usuario, datos_usuario)
        
        conn.commit() 
        log_actividad(None, 'REGISTRO_NUEVO_PROFESOR', f"Registro exitoso de profesor: {num_control} (Discapacidad: {discapacidad})")
        return "Registro exitoso. ¡Ahora puedes iniciar sesión como Profesor!"

    except pyodbc.IntegrityError:
        conn.rollback()
        return f"Error de integridad: El Número de Control/Usuario {num_control} ya existe."
        
    except Exception as e:
        conn.rollback()
        return f"Error al registrar: {e}"
    finally:
        if conn: conn.close()

def obtener_discapacidad_usuario(num_control, rol_id):
    conn = conectar_sql_server()
    if not conn: return "N/D"
    
    try:
        cursor = conn.cursor()
        if rol_id == 2: # Estudiante
            sql_query = "SELECT Discapacidad FROM Estudiantes WHERE Num_Control = ?"
            cursor.execute(sql_query, num_control)
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        elif rol_id == 1: # Profesor
            sql_query = """
            SELECT U.Discapacidad
            FROM Usuarios U
            WHERE U.Nombre_Usuario = ? AND U.ID_Rol_FK = 1
            """
            cursor.execute(sql_query, num_control)
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        return "N/D"
    except pyodbc.ProgrammingError:
        # Esto ocurre si aún no has agregado la columna 'Discapacidad' a tus tablas.
        print("Advertencia: La columna 'Discapacidad' probablemente falta en la BD. Regresando 'N/D'.")
        return "N/D (Columna Faltante)"
    except Exception as e:
        print(f"Error al obtener discapacidad: {e}")
        return "N/D (Error)"
    finally:
        if conn: conn.close()


def insertar_registro_manual(datos, user_id):
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    
    # ¡IMPORTANTE! Se agregó 'Discapacidad' a los campos de inserción.
    # Ahora se esperan 20 campos.
    sql_insert = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    try:
        if len(datos) != 20: # <--- CAMBIO: Ahora son 20 campos
            raise ValueError(f"Error interno: Se esperaban 20 parámetros (columnas), pero se recibieron {len(datos)}. Verifica la función 'guardar_estudiante'.")

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

    # Se agregó 'Discapacidad' a la actualización
    sql_update = """
    UPDATE Estudiantes 
    SET Apellido_Paterno = ?, Apellido_Materno = ?, Nombre = ?, Discapacidad = ?
    WHERE Num_Control = ?
    """
    
    try:
        params = nuevos_datos + (num_control,) 
        cursor = conn.cursor()
        cursor.execute(sql_update, params)
        conn.commit() 
        
        if cursor.rowcount > 0:
            log_actividad(user_id, 'UPDATE_ESTUDIANTE', f"Actualizado Nombre/Apellidos/Discapacidad de Num_Control: {num_control}")
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
    
    # Se agregó 'Discapacidad' a la consulta. Ahora devuelve 12 campos.
    sql_query = """
    SELECT 
        Nombre, Apellido_Paterno, Apellido_Materno, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, 
        Calificacion_Unidad_4, Calificacion_Unidad_5, Discapacidad
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
  
    # Columnas que SIEMPRE deben estar
    COLUMNAS_BASE = ['Num_Control', 'Apellido_Paterno', 'Apellido_Materno', 'Nombre', 'Carrera', 'Semestre']
    
    # Columnas que deben tener un valor por defecto si faltan en el archivo
    # Se agregó 'Discapacidad'
    COLUMNAS_FACTORES = [
        'Materia', 'Calificacion_Unidad_1', 'Calificacion_Unidad_2', 'Calificacion_Unidad_3', 
        'Calificacion_Unidad_4', 'Calificacion_Unidad_5', 'Asistencia_Porcentaje',
        'Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico',
        'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual', 'Discapacidad' 
    ]
    
    # Valores por defecto
    # Se agregó valor por defecto para 'Discapacidad'
    VALORES_DEFECTO = {
        'Materia': 'Sin Asignar',
        'Calificacion_Unidad_1': 0.0, 'Calificacion_Unidad_2': 0.0, 'Calificacion_Unidad_3': 0.0,
        'Calificacion_Unidad_4': 0.0, 'Calificacion_Unidad_5': 0.0, 'Asistencia_Porcentaje': 0.0,
        'Factor_Academico': 0, 'Factor_Psicosocial': 0, 'Factor_Economico': 0,
        'Factor_Institucional': 0, 'Factor_Tecnologico': 0, 'Factor_Contextual': 0,
        'Discapacidad': 'Ninguna'
    }

    try:
        if archivo_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(archivo_path, sheet_name=0)
        elif archivo_path.lower().endswith('.csv'):
            df = pd.read_csv(archivo_path)
        else:
            messagebox.showerror("Error de Formato", "El archivo debe ser CSV o Excel.")
            return
        
        # 1. Verificar las columnas principales
        if not all(col in df.columns for col in COLUMNAS_BASE):
            messagebox.showerror("Error de Columnas", 
                                 f"El archivo debe contener las siguientes 6 columnas base: {', '.join(COLUMNAS_BASE)}. Revise nombres exactos.")
            return

        # 2. Seleccionar todas las columnas presentes que necesitamos
        columnas_existentes = COLUMNAS_BASE + [c for c in COLUMNAS_FACTORES if c in df.columns]
        df_final = df[columnas_existentes].copy()
        
        # 3. Aplicar valores por defecto a las columnas que FALTARON
        for col in COLUMNAS_FACTORES:
            if col not in df_final.columns:
                df_final[col] = VALORES_DEFECTO[col]
        
        # 4. Asegurar el tipo de dato de Semestre
        df_final['Semestre'] = df_final['Semestre'].astype(int)
        
        # 5. Asegurar tipos de las columnas numéricas para SQL (opcional, pero buena práctica)
        # Forzar las calificaciones a float
        for i in range(1, 6):
            col = f'Calificacion_Unidad_{i}'
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(VALORES_DEFECTO[col])
        # Forzar los factores a int
        for col in [f for f in COLUMNAS_FACTORES if f.startswith('Factor_')]:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(VALORES_DEFECTO[col]).astype(int)
            
        # 6. Insertar en SQL
        df_final.to_sql(nombre_tabla, con=engine, if_exists='append', index=False, method='multi')
        
        log_actividad(user_id, 'IMPORT_DATA', f"Importadas {len(df_final)} filas. Columnas leídas: {', '.join(df_final.columns.tolist())}")
        messagebox.showinfo("Importación Exitosa", f"¡{len(df_final)} filas insertadas en SQL Server!")

    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo no fue encontrado.")
    except Exception as e:
        error_msg = str(e)
        if 'La conversión del valor varchar' in error_msg:
             messagebox.showerror("Error Crítico de Tipo de Dato", 
                                 "Error de conversión: ¡AÚN DEBE CORREGIR LA TABLA EN SQL SERVER! "
                                 "Carrera y Semestre tienen tipos invertidos. Ejecute el script SQL proporcionado."
             )
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

# MODIFICACIÓN CLAVE: Ahora devuelve la figura Y los datos de los estudiantes
def generar_pareto_factores(user_id, carrera=None, semestre=None, materia=None, num_control_filtro=None):
    factores = ['Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico', 
                'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual']
    
 
    columnas_suma = [f'SUM(CAST({f} AS INT)) AS {f}' for f in factores]
    
    # Consulta para el Pareto: suma de factores
    sql_pareto_query = f"SELECT {', '.join(columnas_suma)} FROM Estudiantes"
    
    # Consulta para la lista de estudiantes (Num_Control, Nombre, Apellidos, Semestre, Carrera)
    sql_estudiantes_query = """
        SELECT Num_Control, Nombre, Apellido_Paterno, Apellido_Materno, Semestre, Carrera
        FROM Estudiantes
        WHERE ({}) > 0
    """
    columnas_a_sumar = [f'ISNULL(CAST({f} AS INT), 0)' for f in factores]
    suma_factores_check = " + ".join(columnas_a_sumar)
    sql_estudiantes_query = sql_estudiantes_query.format(suma_factores_check)
    
    
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
        # Si es un alumno, solo muestra sus datos y el pareto se basa en él
        condiciones = [f"Num_Control = '{num_control_filtro}'"]
        log_detalle = f"Consulta Pareto por Alumno: {num_control_filtro}"
    else:
        log_detalle = f"Consulta Pareto (Filtros: C={carrera}, S={semestre}, M={materia})"
        
    if condiciones:
        where_clause = " WHERE " + " AND ".join(condiciones)
        sql_pareto_query += where_clause
        sql_estudiantes_query += " AND " + " AND ".join(condiciones)


    try:
        # 1. Obtener datos para el gráfico de Pareto
        df_pareto = pd.read_sql(sql_pareto_query, engine)
        log_actividad(user_id, 'CONSULTA_PARETO', log_detalle)
        
        df_factores = df_pareto.T.rename(columns={0: 'Frecuencia'})
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
            return None, pd.DataFrame(), "No se encontraron factores de riesgo marcados para este grupo/estudiante."

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
        
        # 2. Obtener la lista de estudiantes
        df_estudiantes = pd.read_sql(sql_estudiantes_query, engine)
        # Combinar nombre y apellidos para la presentación en el Treeview
        df_estudiantes['Nombre Completo'] = df_estudiantes['Nombre'] + ' ' + df_estudiantes['Apellido_Paterno'] + ' ' + df_estudiantes['Apellido_Materno']
        df_estudiantes = df_estudiantes[['Num_Control', 'Nombre Completo', 'Semestre', 'Carrera']]
        
        return fig, df_estudiantes, None

    except Exception as e:
        log_actividad(user_id, 'ERROR_CONSULTA', f"Error en Pareto: {e}")
        return None, pd.DataFrame(), f"Error al consultar o generar gráfico: {e}"

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


# [ ... EL CÓDIGO DE LAS CLASES StudentRegistrationWindow, TeacherRegistrationWindow, LoginWindow ES EL MISMO ... ]

class StudentRegistrationWindow(tk.Toplevel):
    """Ventana de registro inicial para nuevos estudiantes."""
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title("Registro de Estudiante")
        self.geometry("400x440") # Ajuste de tamaño
        self.resizable(False, False)
        self.grab_set()
        
        # Aplicar el tema actual
        master.apply_theme_settings() 
        self.create_widgets()
        master.update_font_size(self) # Actualizar fuentes

    def create_widgets(self):
        # Usar fondo blanco o negro según el modo invertido
        bg_color = 'white' if not COLOR_INVERTED else 'black'
        fg_color = 'black' if not COLOR_INVERTED else 'white'
        
        self.config(bg=bg_color)
        
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="**REGISTRO DE ESTUDIANTE**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        fields = [
            ("No. Control:", "num_control", 1, 'entry', ''), 
            ("Nombre(s):", "nombre", 2, 'entry', ''),
            ("Apellido Paterno:", "apellido_p", 3, 'entry', ''), 
            ("Apellido Materno:", "apellido_m", 4, 'entry', ''), 
            ("Carrera:", "carrera", 5, 'combobox', CARRERAS_ITT), 
            ("Semestre:", "semestre", 6, 'combobox', SEMESTRES_LIST), 
            ("Discapacidad:", "discapacidad", 7, 'combobox', DISCAPACIDADES_LIST), # <--- NUEVO CAMPO
            ("Contraseña:", "contrasena", 8, 'entry', '*')
        ]

        self.reg_vars = {}
        for i, (label_text, var_name, row, widget_type, values) in enumerate(fields):
            label = ttk.Label(main_frame, text=label_text)
            label.grid(row=row, column=0, padx=5, pady=5, sticky='w')
            
            var = tk.StringVar()
            self.reg_vars[var_name] = var
            
            if widget_type == 'entry':
                show = '*' if var_name == 'contrasena' else ''
                entry = ttk.Entry(main_frame, textvariable=var, width=30, show=show)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky='w')
                if var_name == 'num_control': label.configure(underline=0); entry.focus() # Access Key + Focus
            elif widget_type == 'combobox': 
                combo = ttk.Combobox(main_frame, textvariable=var, values=values, width=28, state='readonly')
                combo.set(values[0] if values else '') 
                combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration, underline=0).pack(side=tk.LEFT, padx=10) # Access Key R
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy, underline=0).pack(side=tk.LEFT, padx=10) # Access Key C
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            carrera = self.reg_vars['carrera'].get().strip() 
            semestre_str = self.reg_vars['semestre'].get().strip()
            discapacidad = self.reg_vars['discapacidad'].get().strip() # <--- NUEVO VALOR
            contrasena = self.reg_vars['contrasena'].get().strip()

            if not all([num_control, nombre, apellido_p, carrera, semestre_str, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos principales son obligatorios.")
                return

            # Se pasa el campo 'discapacidad'
            resultado = registrar_estudiante_usuario(num_control, nombre, apellido_p, apellido_m, carrera, semestre_str, contrasena, discapacidad)
            
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
        self.master = master
        self.title("Registro de Profesor")
        self.geometry("400x320") # Ajuste de tamaño
        self.resizable(False, False)
        self.grab_set()
        
        # Aplicar el tema actual
        master.apply_theme_settings()
        self.create_widgets()
        master.update_font_size(self) # Actualizar fuentes

    def create_widgets(self):
        # Usar fondo blanco o negro según el modo invertido
        bg_color = 'white' if not COLOR_INVERTED else 'black'
        fg_color = 'black' if not COLOR_INVERTED else 'white'
        
        self.config(bg=bg_color)
        
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="**REGISTRO DE PROFESOR**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        fields = [
            ("Usuario (Ej. No. Empleado):", "num_control", 1, '', None), 
            ("Nombre(s):", "nombre", 2, '', None),
            ("Apellido Paterno:", "apellido_p", 3, '', None), 
            ("Apellido Materno:", "apellido_m", 4, '', None), 
            ("Discapacidad:", "discapacidad", 5, 'combobox', DISCAPACIDADES_LIST), # <--- NUEVO CAMPO
            ("Contraseña:", "contrasena", 6, '*', None)
        ]

        self.reg_vars = {}
        for i, (label_text, var_name, row, show_char, values) in enumerate(fields):
            label = ttk.Label(main_frame, text=label_text)
            label.grid(row=row, column=0, padx=5, pady=5, sticky='w')
            
            var = tk.StringVar()
            self.reg_vars[var_name] = var

            if var_name == 'discapacidad':
                combo = ttk.Combobox(main_frame, textvariable=var, values=values, width=28, state='readonly')
                combo.set(values[0] if values else '') 
                combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')
            else:
                show = '*' if var_name == 'contrasena' else ''
                entry = ttk.Entry(main_frame, textvariable=var, width=30, show=show)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky='w')
                if var_name == 'num_control': label.configure(underline=0); entry.focus() # Access Key + Focus


        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration, underline=0).pack(side=tk.LEFT, padx=10) # Access Key R
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy, underline=0).pack(side=tk.LEFT, padx=10) # Access Key C
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            discapacidad = self.reg_vars['discapacidad'].get().strip() # <--- NUEVO VALOR
            contrasena = self.reg_vars['contrasena'].get().strip()

            if not all([num_control, nombre, apellido_p, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos principales son obligatorios.")
                return

            # Se pasa el campo 'discapacidad'
            resultado = registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena, discapacidad)
            
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
        self.geometry("400x180") # Ajuste de tamaño
        self.master.withdraw()
        self.deiconify()
        
        
        self.protocol("WM_DELETE_WINDOW", self.master.quit) 
        self.resizable(False, False)
        self.grab_set()
        
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
        # Aplicar el tema actual
        master.apply_theme_settings()
        self.create_widgets()
        master.update_font_size(self) # Actualizar fuentes

    def create_widgets(self):
        # Usar fondo blanco o negro según el modo invertido
        bg_color = 'white' if not COLOR_INVERTED else 'black'
        fg_color = 'black' if not COLOR_INVERTED else 'white'
        
        self.config(bg=bg_color)
        
        cred_frame = ttk.Frame(self, padding="10")
        cred_frame.grid(row=0, column=0, columnspan=2)

        ttk.Label(cred_frame, text="Usuario:", underline=0).grid(row=0, column=0, padx=10, pady=5, sticky='w') # Access Key U
        self.user_var = tk.StringVar(value='ProfesorITT')
        user_entry = ttk.Entry(cred_frame, textvariable=self.user_var, width=25)
        user_entry.grid(row=0, column=1, padx=10, pady=5)
        user_entry.focus()

        ttk.Label(cred_frame, text="Contraseña:", underline=0).grid(row=1, column=0, padx=10, pady=5, sticky='w') # Access Key C
        self.pass_var = tk.StringVar(value='123')
        self.pass_entry = ttk.Entry(cred_frame, textvariable=self.pass_var, show="*", width=25)
        self.pass_entry.grid(row=1, column=1, padx=10, pady=5)
        self.pass_entry.bind('<Return>', lambda event: self.handle_login())

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Entrar", command=self.handle_login, underline=0).pack(side=tk.LEFT, padx=5) # Access Key E
        ttk.Button(btn_frame, text="Registrar Alumno", command=self.open_student_registration, underline=10).pack(side=tk.LEFT, padx=5) # Access Key A
        ttk.Button(btn_frame, text="Registrar Maestro", command=self.open_teacher_registration, underline=10).pack(side=tk.LEFT, padx=5) # Access Key M

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

        
        # 1. Pestaña de Perfil (Primera para todos)
        self.tab_perfil = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_perfil, text='Perfil / Actualizar', underline=0) # Access Key P
        self.crear_tab_perfil()
        
       
        if self.role_id == 1:
            # 2. Pestañas de Profesor
            self.tab_registro = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_registro, text='Registro Manual', underline=0) # Access Key R
            self.crear_tab_registro()

            self.tab_pareto = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_pareto, text='Análisis de Riesgo (Pareto)', underline=0) # Access Key A
            self.crear_tab_pareto() 
            
            self.tab_importar = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_importar, text='Importar/Exportar', underline=0) # Access Key I
            self.crear_tab_importar()
            
            self.tab_auditoria = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_auditoria, text='Registro de Actividad', underline=0) # Access Key R
            self.crear_tab_auditoria()
            
        
        elif self.role_id == 2:
            # 2. Pestaña de Estudiante
            self.tab_datos = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_datos, text='Mis Datos Académicos', underline=0) # Access Key M
            self.crear_tab_datos_alumno()
            
        # 3. Pestaña de Configuración (Última) <-- CAMBIO APLICADO AQUÍ
        self.tab_configuracion = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_configuracion, text='⚙️ Configuración', underline=0) # Access Key C
        self.crear_tab_configuracion()
        

    def cerrar_sesion(self):
        log_actividad(self.user_id, 'LOGOUT', 'Cierre de sesión manual.')        
        self.destroy() 
        self.master.login_window = LoginWindow(self.master) 
        self.master.withdraw() 
        self.master.main_window = None

    # ====================================================================
    # PESTAÑA DE CONFIGURACIÓN
    # ====================================================================
    def crear_tab_configuracion(self):
        global CURRENT_FONT_SIZE, COLOR_INVERTED, ZOOM_LEVEL
        
        frame = ttk.Frame(self.tab_configuracion, padding="20")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**CONFIGURACIÓN DE ACCESIBILIDAD**", font=('Arial', 14, 'bold')).pack(pady=15)
        
        # -------------------
        # 1. Tamaño de Fuente
        # -------------------
        font_frame = ttk.LabelFrame(frame, text="Ajuste de Tamaño de Fuente", padding=10)
        font_frame.pack(pady=10, padx=50, fill='x')
        
        ttk.Label(font_frame, text="Tamaño Actual:").pack(side=tk.LEFT, padx=10)
        
        self.current_font_label = ttk.Label(font_frame, text=f"{CURRENT_FONT_SIZE} puntos", font=('Arial', 10, 'bold'))
        self.current_font_label.pack(side=tk.LEFT)
        
        ttk.Button(font_frame, text="Aumentar (A)", command=lambda: self.master.adjust_font(1, self.current_font_label), underline=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(font_frame, text="Disminuir (D)", command=lambda: self.master.adjust_font(-1, self.current_font_label), underline=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(font_frame, text="Reiniciar (R)", command=lambda: self.master.adjust_font(0, self.current_font_label), underline=0).pack(side=tk.LEFT, padx=10)
        
        # -------------------
        # 2. Inversión de Colores
        # -------------------
        color_frame = ttk.LabelFrame(frame, text="Inversión de Colores (Alto Contraste)", padding=10)
        color_frame.pack(pady=10, padx=50, fill='x')
        
        self.color_var = tk.BooleanVar(value=COLOR_INVERTED)
        self.color_check = ttk.Checkbutton(color_frame, 
                                           text="Activar Inversión de Colores (I)", 
                                           variable=self.color_var,
                                           command=self.handle_color_inversion,
                                           underline=24)
        self.color_check.pack(padx=10)
        
        # -------------------
        # 3. Lupa de Zoom (Escalado de la Aplicación)
        # -------------------
        zoom_frame = ttk.LabelFrame(frame, text="Lupa de Zoom (Escala de Pantalla)", padding=10)
        zoom_frame.pack(pady=10, padx=50, fill='x')
        
        ttk.Label(zoom_frame, text="Nivel de Zoom:").pack(side=tk.LEFT, padx=10)
        
        self.zoom_level_label = ttk.Label(zoom_frame, text=f"{ZOOM_LEVEL}%", font=('Arial', 10, 'bold'))
        self.zoom_level_label.pack(side=tk.LEFT)
        
        ttk.Button(zoom_frame, text="Zoom + (Z)", command=lambda: self.master.adjust_zoom(10, self.zoom_level_label), underline=6).pack(side=tk.LEFT, padx=10)
        ttk.Button(zoom_frame, text="Zoom - (N)", command=lambda: self.master.adjust_zoom(-10, self.zoom_level_label), underline=6).pack(side=tk.LEFT, padx=10)
        ttk.Button(zoom_frame, text="Restablecer (T)", command=lambda: self.master.adjust_zoom(0, self.zoom_level_label), underline=0).pack(side=tk.LEFT, padx=10)
        
    def handle_color_inversion(self):
        """Maneja el cambio de la variable global de inversión de color y aplica el tema."""
        global COLOR_INVERTED
        COLOR_INVERTED = self.color_var.get()
        self.master.apply_theme_settings()
        self.master.update_font_size(self.master.main_window) # Forzar re-render de Treeviews, etc.
        
    # ====================================================================
    # PESTAÑA DE REGISTRO MANUAL
    # ====================================================================
    
    def crear_tab_registro(self):
        frame = ttk.Frame(self.tab_registro, padding="15")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**REGISTRO MANUAL DE ESTUDIANTES**", font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=4, pady=10)

        self.entry_vars = {}
        fields = [
            ("No. Control:", "Num_Control", 1, 'entry', None),
            ("Apellido Paterno:", "Apellido_Paterno", 2, 'entry', None),
            ("Apellido Materno:", "Apellido_Materno", 3, 'entry', None),
            ("Nombre(s):", "Nombre", 4, 'entry', None),
            ("Carrera:", "Carrera", 5, 'combobox', CARRERAS_ITT),
            ("Semestre:", "Semestre", 6, 'combobox', SEMESTRES_LIST),
            ("Materia:", "Materia", 7, 'entry', None),
            ("Asistencia %:", "Asistencia_Porcentaje", 8, 'entry', None),
            ("Discapacidad:", "Discapacidad", 9, 'combobox', DISCAPACIDADES_LIST) # <--- NUEVO CAMPO
        ]

        for i, (label_text, var_name, row, widget_type, values) in enumerate(fields):
            label = ttk.Label(frame, text=label_text)
            label.grid(row=row, column=0, padx=5, pady=2, sticky='w')
            
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            
            if widget_type == 'entry':
                ttk.Entry(frame, textvariable=var, width=30).grid(row=row, column=1, padx=5, pady=2, sticky='w')
            elif widget_type == 'combobox': 
                combo = ttk.Combobox(frame, textvariable=var, values=values, width=28, state='readonly')
                combo.set(values[0] if values else '')
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
        fact_frame.grid(row=6, column=2, rowspan=4, padx=10, pady=5, sticky='n')
        self.factor_vars = {}
        factores = ["Factor_Academico", "Factor_Psicosocial", "Factor_Economico", "Factor_Institucional", "Factor_Tecnologico", "Factor_Contextual"]
        nombres_fact = ["Académico", "Psicosocial", "Económico", "Institucional", "Tecnológico", "Contextual"]
        for i, (var_name, name) in enumerate(zip(factores, nombres_fact)):
            var = tk.IntVar()
            self.factor_vars[var_name] = var
            ttk.Checkbutton(fact_frame, text=name, variable=var).grid(row=i//2, column=i%2, padx=5, pady=2, sticky='w')

       
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=10, column=0, columnspan=4, pady=20)
        ttk.Button(btn_frame, text="Guardar Estudiante", command=self.guardar_estudiante, underline=0).pack(padx=10, side=tk.LEFT) # Access Key G
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar_registro, underline=0).pack(padx=10, side=tk.LEFT) # Access Key L
        
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
            discapacidad = self.entry_vars['Discapacidad'].get().strip() # <--- NUEVO CAMPO
            
            if not all(datos_principales) or not semestre_raw or not materia or not discapacidad:
                messagebox.showwarning("Advertencia", "Los campos principales (Control, Nombre, Apellidos, Carrera, Semestre, Materia, Discapacidad) son obligatorios.")
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

           
            # ¡IMPORTANTE! Se agregó [discapacidad] al final
            datos_completos = datos_principales + calificaciones_finales + [asistencia_final] + factores + [discapacidad]
            
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
        
        # Obtener Discapacidad para el perfil (usa una función específica ya que puede estar en Estudiantes o Usuarios)
        discapacidad_usuario = obtener_discapacidad_usuario(self.num_control, self.role_id)


        ttk.Label(frame, text=f"**PERFIL DE USUARIO**", font=('Arial', 14, 'bold')).pack(pady=10)
        
        
        data_map = {
            "No. Control:": self.num_control if self.role_id == 2 else self.master.login_window.user_var.get(),
            "Nombre(s):": datos_estudiante[0] if datos_estudiante and len(datos_estudiante) > 0 else "",
            "Apellido Paterno:": datos_estudiante[1] if datos_estudiante and len(datos_estudiante) > 1 else "",
            "Apellido Materno:": datos_estudiante[2] if datos_estudiante and len(datos_estudiante) > 2 else "",
            "Carrera:": datos_estudiante[3] if datos_estudiante and len(datos_estudiante) > 3 else "N/A (Profesor)",
            "Semestre:": datos_estudiante[4] if datos_estudiante and len(datos_estudiante) > 4 else "N/A (Profesor)",
            "Discapacidad:": discapacidad_usuario # <--- NUEVO CAMPO
        }
        
        fields = [
            ("No. Control:", "Num_Control"), 
            ("Nombre(s):", "Nombre"), 
            ("Apellido Paterno:", "Apellido_Paterno"), 
            ("Apellido Materno:", "Apellido_Materno"), 
            ("Carrera:", "Carrera"), 
            ("Semestre:", "Semestre"),
            ("Discapacidad:", "Discapacidad") # <--- NUEVO CAMPO
        ]
        
        for label_text, var_name in fields:
            sub_frame = ttk.Frame(frame)
            sub_frame.pack(fill='x', pady=5, padx=50)
            
            ttk.Label(sub_frame, text=label_text, width=15, anchor='w').pack(side=tk.LEFT)
            
            var = tk.StringVar()
            self.perfil_vars[var_name] = var
            var.set(data_map.get(label_text, ""))
            
            is_editable = var_name in ["Nombre", "Apellido_Paterno", "Apellido_Materno", "Discapacidad"]
            state = 'normal' if is_editable and self.role_id == 2 else 'readonly'
            
            if var_name == "Discapacidad" and self.role_id == 2:
                combo = ttk.Combobox(sub_frame, textvariable=var, values=DISCAPACIDADES_LIST, width=30, state='readonly')
                combo.pack(side=tk.LEFT, fill='x', expand=True)
            else:
                ttk.Entry(sub_frame, textvariable=var, width=30, state=state).pack(side=tk.LEFT, fill='x', expand=True)

        if self.role_id == 1:
            ttk.Label(frame, text="Rol: Profesor", font=('Arial', 10, 'italic')).pack(pady=10, padx=50, anchor='w')
            
        if self.role_id == 2:
            ttk.Label(frame, text="Solo puedes editar tu Nombre, Apellidos y Discapacidad.", font=('Arial', 10, 'italic')).pack(pady=5, padx=50, anchor='w')
       
             
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)
        ttk.Button(frame, text="Actualizar Datos", command=self.actualizar_perfil, underline=0).pack(pady=5) # Access Key A
        ttk.Button(frame, text="Cerrar Sesión", command=self.cerrar_sesion, underline=0).pack(pady=5) # Access Key C

    def actualizar_perfil(self):
        if self.role_id != 2:
            messagebox.showinfo("Información", "Solo los estudiantes pueden actualizar sus datos desde esta ventana.")
            return

        nombre = self.perfil_vars['Nombre'].get().strip()
        apellido_p = self.perfil_vars['Apellido_Paterno'].get().strip()
        apellido_m = self.perfil_vars['Apellido_Materno'].get().strip()
        discapacidad = self.perfil_vars['Discapacidad'].get().strip() # <--- NUEVO CAMPO
        
        if not all([nombre, apellido_p, discapacidad]):
             messagebox.showwarning("Advertencia", "El nombre, el apellido paterno y la discapacidad son obligatorios.")
             return
        
        nuevos_datos = [apellido_p, apellido_m, nombre, discapacidad] # <--- NUEVO CAMPO
        actualizar_datos_estudiante(self.num_control, nuevos_datos, self.user_id)
        
    #Pestaña de Datos Alumno
    
    def crear_tab_datos_alumno(self):
        
        frame = ttk.Frame(self.tab_datos, padding="20")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**MIS DATOS ACADÉMICOS**", font=('Arial', 14, 'bold')).pack(pady=10)
        
        datos = obtener_datos_estudiante(self.num_control)

        if datos:
            # Se agregó "Discapacidad:" a la lista. Ahora son 12 campos.
            labels = ["Nombre:", "Apellido Paterno:", "Apellido Materno:", "Carrera:", "Semestre:", "Materia:", 
                      "Calificación U1:", "Calificación U2:", "Calificación U3:", "Calificación U4:", "Calificación U5:",
                      "Discapacidad:"] 
            
            
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
        
        # ** SCROLLBAR PARA LA PESTAÑA PRINCIPAL DEL PARETO **
        # Este marco contenedor permite que todo el contenido se desplace
        canvas = tk.Canvas(frame)
        canvas.pack(side="left", fill="both", expand=True)
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        vsb.pack(side="right", fill="y")
        
        canvas.configure(yscrollcommand=vsb.set)
        
        self.scrollable_frame = ttk.Frame(canvas)
        # Se fija el ancho para que la barra lateral no comprima horizontalmente
        # NOTA: Este ancho debe ajustarse según el nivel de zoom si se quiere un layout perfecto
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        # 910 es el ancho que se ajusta a 950x700. Si el zoom cambia, este valor necesita ajustarse.
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=910 * (ZOOM_LEVEL / 100.0))

        # ** FIN DEL SCROLLBAR DE LA PESTAÑA **

        
        control_frame = ttk.LabelFrame(self.scrollable_frame, text="Filtros y Controles")
        control_frame.pack(pady=10, padx=10, fill='x')
        
        estado_filtro = 'disabled' if self.role_id == 2 else 'readonly'
        
        # Fila 0: Carrera y Semestre
        ttk.Label(control_frame, text="Carrera:", underline=0).grid(row=0, column=0, padx=5, pady=5) # Access Key C
        self.pareto_carrera = tk.StringVar()
        combo_carrera = ttk.Combobox(control_frame, textvariable=self.pareto_carrera, values=CARRERAS_ITT, width=15, state=estado_filtro)
        combo_carrera.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(control_frame, text="Semestre:", underline=0).grid(row=0, column=2, padx=5, pady=5) # Access Key S
        self.pareto_semestre = tk.StringVar()
        combo_semestre = ttk.Combobox(control_frame, textvariable=self.pareto_semestre, values=SEMESTRES_LIST, width=15, state=estado_filtro)
        combo_semestre.grid(row=0, column=3, padx=5, pady=5)
        
        # Fila 1: Materia y Botones
        ttk.Label(control_frame, text="Materia:", underline=0).grid(row=1, column=0, padx=5, pady=5) # Access Key M
        self.pareto_materia = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.pareto_materia, width=15, state='disabled' if self.role_id == 2 else 'normal').grid(row=1, column=1, padx=5, pady=5)

        # Botones
        ttk.Button(control_frame, text="Generar Gráfico", command=self.mostrar_pareto, underline=0).grid(row=1, column=2, padx=15, pady=5) # Access Key G
        ttk.Button(control_frame, text="Limpiar Filtros", command=self.limpiar_filtros_pareto, underline=0).grid(row=1, column=3, padx=15, pady=5) # Access Key L

        # Marco para el Canvas de Pareto
        self.pareto_canvas_frame = ttk.Frame(self.scrollable_frame, height=300)
        self.pareto_canvas_frame.pack(fill='x', expand=False, padx=10, pady=5)
        self.pareto_fig_canvas = None
        
        # Marco y Treeview para la lista de alumnos
        alumnos_frame = ttk.LabelFrame(self.scrollable_frame, text="Alumnos Incluidos en el Análisis")
        alumnos_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        columns = ("Num_Control", "Nombre Completo", "Semestre", "Carrera")
        self.tree_alumnos = ttk.Treeview(alumnos_frame, columns=columns, show="headings")
        
        self.tree_alumnos.heading("Num_Control", text="No. Control")
        self.tree_alumnos.heading("Nombre Completo", text="Nombre y Apellidos")
        self.tree_alumnos.heading("Semestre", text="Semestre")
        self.tree_alumnos.heading("Carrera", text="Carrera")
        
        self.tree_alumnos.column("Num_Control", width=100, anchor='center')
        self.tree_alumnos.column("Nombre Completo", width=300, anchor='w')
        self.tree_alumnos.column("Semestre", width=80, anchor='center')
        self.tree_alumnos.column("Carrera", width=150, anchor='w')
        
        # Scrollbar vertical para el Treeview (lista de alumnos)
        vsb_alumnos = ttk.Scrollbar(alumnos_frame, orient="vertical", command=self.tree_alumnos.yview)
        self.tree_alumnos.configure(yscrollcommand=vsb_alumnos.set)
        
        vsb_alumnos.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_alumnos.pack(fill='both', expand=True)
        
    def limpiar_filtros_pareto(self):
        """Limpia los campos de filtro de Carrera, Semestre y Materia."""
        self.pareto_carrera.set('')
        self.pareto_semestre.set('')
        self.pareto_materia.set('')
        
        # Opcional: limpiar la gráfica y la tabla
        for widget in self.pareto_canvas_frame.winfo_children():
            widget.destroy()
        for i in self.tree_alumnos.get_children():
            self.tree_alumnos.delete(i)


    def mostrar_pareto(self):
        
        # Limpiar el canvas y la tabla antes de la nueva consulta
        for widget in self.pareto_canvas_frame.winfo_children():
            widget.destroy()
        for i in self.tree_alumnos.get_children():
            self.tree_alumnos.delete(i)
        
        carrera = self.pareto_carrera.get().strip()
        semestre = self.pareto_semestre.get().strip()
        materia = self.pareto_materia.get().strip()

     
        if self.role_id == 2:
            num_control_filtro = self.num_control
            # Si es alumno, ignora los filtros de la interfaz
            carrera, semestre, materia = None, None, None 
        else:
            num_control_filtro = None

        # Llamada a la función de backend, que ahora devuelve fig, df_estudiantes, error_msg
        fig, df_estudiantes, error_msg = generar_pareto_factores(self.user_id, carrera, semestre, materia, num_control_filtro)
        
        if error_msg:
            messagebox.showerror("Error de Gráfico", error_msg)
            # Muestra un mensaje en la tabla
            self.tree_alumnos.insert("", "end", values=(error_msg, "", "", ""))
            return

        if fig:
            # 1. Mostrar Gráfico de Pareto
            self.pareto_fig_canvas = FigureCanvasTkAgg(fig, master=self.pareto_canvas_frame)
            self.pareto_fig_canvas.draw()
            # Usar pack, expand=False para que la gráfica no ocupe todo el espacio vertical del marco
            self.pareto_fig_canvas.get_tk_widget().pack(fill='both', expand=False) 
            
            # 2. Mostrar la lista de alumnos
            if df_estudiantes.empty:
                 self.tree_alumnos.insert("", "end", values=("No hay alumnos con factores de riesgo en estos filtros.", "", "", ""))
            else:
                for index, row in df_estudiantes.iterrows():
                    self.tree_alumnos.insert("", "end", values=(
                        row['Num_Control'], 
                        row['Nombre Completo'], 
                        row['Semestre'], 
                        row['Carrera']
                    ))


    #Pestaña de Importar/Exportar
    
    def crear_tab_importar(self):
        frame = ttk.Frame(self.tab_importar, padding="10")
        frame.pack(fill='both', expand=True)
        
        if self.role_id == 1:
            #Sección Importar
            import_frame = ttk.LabelFrame(frame, text="Importar Datos de Estudiantes (CSV/Excel)")
            import_frame.pack(pady=20, padx=20, fill='x')
            
            ttk.Label(import_frame, text="Archivo (CSV/Excel):", underline=0).grid(row=0, column=0, padx=5, pady=5, sticky='w') # Access Key A
            self.import_path_var = tk.StringVar()
            ttk.Entry(import_frame, textvariable=self.import_path_var, width=50, state='readonly').grid(row=0, column=1, padx=5, pady=5)
            
            ttk.Button(import_frame, text="Seleccionar Archivo", command=self.seleccionar_archivo_import, underline=0).grid(row=0, column=2, padx=5, pady=5) # Access Key S
            ttk.Button(import_frame, text="Ejecutar Importación", command=self.ejecutar_importacion_handler, underline=0).grid(row=1, column=1, columnspan=2, pady=10) # Access Key E

            #Sección Exportar
            export_frame = ttk.LabelFrame(frame, text="Exportar Datos de Estudiantes")
            export_frame.pack(pady=20, padx=20, fill='x')
            
            export_btns_frame = ttk.Frame(export_frame)
            export_btns_frame.pack(pady=10)
            
            ttk.Button(export_btns_frame, text="Exportar a Excel (.xlsx)", command=lambda: exportar_datos_sql('excel', self.user_id), underline=13).pack(side=tk.LEFT, padx=10) # Access Key X
            ttk.Button(export_btns_frame, text="Exportar a CSV (.csv)", command=lambda: exportar_datos_sql('csv', self.user_id), underline=13).pack(side=tk.LEFT, padx=10) # Access Key V
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

       
        ttk.Button(frame, text="Actualizar Registro", command=self.cargar_auditoria, underline=0).pack(pady=10) # Access Key A
        
       
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
    
    # Dimensiones Base de la Ventana Principal
    BASE_WIDTH = 950
    BASE_HEIGHT = 700
    
    def __init__(self):
        super().__init__()
        self.title("Sistema de Análisis de Calidad ITT")
        self.geometry("1x1") 
        self.withdraw() 
        self.style = ttk.Style(self)
        self.main_window = None

        # Configuración inicial de estilos y temas
        self.apply_theme_settings()
        
        self.login_window = LoginWindow(self)
        
    def apply_theme_settings(self):
        """Aplica la configuración de tema, tamaño de fuente e inversión de color."""
        global CURRENT_FONT_SIZE, COLOR_INVERTED
        
        # 1. Aplicar Tema Base y Fuentes
        self.style.theme_use('clam') 

        # Crear o reconfigurar fuentes con el tamaño actual
        font_size = CURRENT_FONT_SIZE
        # Usamos tkinter.font para crear las fuentes con nombre, lo que permite que ttk las use
        self.tk_normal_font = tkinter.font.Font(family="Arial", size=font_size)
        self.tk_bold_font = tkinter.font.Font(family="Arial", size=font_size, weight="bold")
        self.tk_title_font = tkinter.font.Font(family="Arial", size=font_size + 4, weight="bold")
        self.tk_italic_font = tkinter.font.Font(family="Arial", size=font_size, slant="italic")
        
        self.style.configure('.', font=self.tk_normal_font)
        self.style.configure('TButton', font=self.tk_bold_font, padding=6)
        self.style.configure('TLabel', font=self.tk_normal_font)
        self.style.configure('TNotebook.Tab', padding=[10, 5])
        self.style.configure('Treeview.Heading', font=self.tk_bold_font)
        
        # 2. Inversión de Colores
        if COLOR_INVERTED:
            # Colores para inversión
            bg_color = COLOR_MAP.get('white', 'black') # Negro
            fg_color = COLOR_MAP.get('black', 'white') # Blanco
            
            # Reconfigurar estilos comunes
            self.style.configure('TFrame', background=bg_color)
            self.style.configure('TLabel', background=bg_color, foreground=fg_color)
            self.style.configure('TNotebook', background=bg_color)
            self.style.configure('TNotebook.Tab', background=bg_color, foreground=fg_color)
            self.style.configure('TButton', background=fg_color, foreground=bg_color)
            self.style.configure('TEntry', fieldbackground=fg_color, foreground=bg_color)
            self.style.configure('Treeview', background=bg_color, foreground=fg_color, fieldbackground=bg_color)
            self.style.map('TButton', background=[('active', 'SystemHighlight')], foreground=[('active', 'SystemHighlightText')])

            # Fondo de la ventana principal
            self.configure(bg=bg_color)
            if self.main_window:
                self.main_window.configure(style='TFrame')
            
        else:
            # Colores normales
            self.style.theme_use('clam') # Recarga el tema para resetear colores si el tema subyacente lo permite
            self.style.configure('TButton', font=self.tk_bold_font, padding=6)
            self.style.configure('TLabel', font=self.tk_normal_font)
            # Asegurar que el fondo del frame de la app principal sea blanco/claro
            self.style.configure('TFrame', background='SystemWindow')
            self.style.configure('TLabel', background='SystemWindow', foreground='SystemWindowText')
            self.style.configure('Treeview', background='SystemWindow', foreground='SystemWindowText', fieldbackground='SystemWindow')
            self.configure(bg='SystemWindow')
            

    def update_font_size(self, container):
        """
        Recorre recursivamente todos los widgets en un contenedor (ventana o frame) 
        y actualiza su fuente o fondo si es necesario.
        """
        global CURRENT_FONT_SIZE, COLOR_INVERTED
        
        # Actualiza la fuente predeterminada para el contenedor y sus ttk widgets
        try:
            # Esto maneja los widgets ttk (TLabel, TButton, etc.)
            self.style.configure('.', font=self.tk_normal_font)
            self.style.configure('TButton', font=self.tk_bold_font)
            self.style.configure('Treeview.Heading', font=self.tk_bold_font)
        except:
            pass
            
        # Actualizar fuentes de widgets tk 'puros' (como los que se usan para títulos con el font=('Arial', 14, 'bold')
        for widget in container.winfo_children():
            try:
                # Actualiza fuentes de texto directo (como los titles)
                current_font = widget.cget('font')
                if isinstance(current_font, str):
                    match = re.search(r'(\w+)\s*(\d+)\s*(.*)', current_font)
                    if match:
                        family, old_size_str, style = match.groups()
                        old_size = int(old_size_str)
                        
                        # Cálculo del nuevo tamaño basado en la relación con el BASE_FONT_SIZE
                        size_ratio = old_size / BASE_FONT_SIZE
                        new_size = round(CURRENT_FONT_SIZE * size_ratio)
                        
                        new_font_str = f'{family} {new_size} {style}'.strip()
                        widget.configure(font=new_font_str)

            except tk.TclError:
                # Ocurre si el widget no tiene la opción 'font' (ej. Frames)
                pass
            except Exception as e:
                # print(f"Error al actualizar fuente de {widget}: {e}")
                pass
            
            # Recorrer widgets anidados
            self.update_font_size(widget) 
            
    def adjust_font(self, delta, label_widget):
        """Ajusta el tamaño global de la fuente y lo aplica a toda la aplicación."""
        global CURRENT_FONT_SIZE, BASE_FONT_SIZE
        
        if delta == 0:
            CURRENT_FONT_SIZE = BASE_FONT_SIZE
        else:
            new_size = CURRENT_FONT_SIZE + delta
            if new_size < 6 or new_size > 20: # Limitar el tamaño
                return
            CURRENT_FONT_SIZE = new_size
            
        # 1. Aplicar la nueva configuración de tema (que incluye el nuevo tamaño)
        self.apply_theme_settings()
        
        # 2. Actualizar las fuentes en toda la ventana principal y popups abiertos
        if self.main_window:
            self.update_font_size(self.main_window)
        
        # 3. Actualizar la etiqueta en la pestaña de configuración
        label_widget.config(text=f"{CURRENT_FONT_SIZE} puntos")
        
        # 4. Forzar el redraw de Matplotlib (si la pestaña Pareto está abierta)
        if hasattr(self.main_window, 'tab_pareto') and self.main_window.notebook.tab(self.main_window.notebook.select(), "text") == 'Análisis de Riesgo (Pareto)':
             # Esto es un workaround. Lo ideal sería volver a llamar a mostrar_pareto().
             # En este caso, simplemente se avisa.
             pass 

    def adjust_zoom(self, delta, label_widget):
        """Ajusta el nivel de zoom de la aplicación modificando el tamaño de la ventana."""
        global ZOOM_LEVEL
        
        if delta == 0:
            ZOOM_LEVEL = 100
        else:
            new_zoom = ZOOM_LEVEL + delta
            if new_zoom < 50 or new_zoom > 200: # Limitar el zoom
                return
            ZOOM_LEVEL = new_zoom
            
        scale_factor = ZOOM_LEVEL / 100.0
        
        new_width = int(self.BASE_WIDTH * scale_factor)
        new_height = int(self.BASE_HEIGHT * scale_factor)
        
        self.geometry(f"{new_width}x{new_height}")
        
        # Actualizar la etiqueta
        label_widget.config(text=f"{ZOOM_LEVEL}%")
        
        # NOTA: En aplicaciones Tkinter/Ttk puras, el zoom es difícil de aplicar globalmente 
        # sin redibujar o re-dimensionar manualmente muchos widgets (como los anchos fijos de Treeview 
        # o los tamaños fijos en grid). Este método es una simulación que solo redimensiona la ventana
        # principal, dando la sensación de zoom.

    def show_main_window(self, user_id, role_id, num_control):
        
        if self.main_window:
            self.main_window.destroy() 
            
        
        self.main_window = CalidadApp(self, user_id, role_id, num_control)
        
        # Muestra la ventana principal con el tamaño ajustado al zoom
        scale_factor = ZOOM_LEVEL / 100.0
        new_width = int(self.BASE_WIDTH * scale_factor)
        new_height = int(self.BASE_HEIGHT * scale_factor)
        
        self.geometry(f"{new_width}x{new_height}") 
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