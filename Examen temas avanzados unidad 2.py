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
import tkinter.font 
import threading
import time

try:
    from PIL import Image, ImageTk, ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Advertencia: PIL (Pillow) no está instalado. La lupa no funcionará.")

try:
    import pyttsx3
    TTS_AVAILABLE = True
except ImportError:
    TTS_AVAILABLE = False
    print("Advertencia: pyttsx3 no instalado. Lectura en voz alta desactivada.")

try:
    import speech_recognition as sr
    VOICE_AVAILABLE = True
except ImportError:
    VOICE_AVAILABLE = False
    print("Advertencia: SpeechRecognition no instalado. Control por voz desactivado.")

# ----------------------------------------------------------------

SERVER = 'DESKTOP-4K70KRA' 
DATABASE = 'ITT_Calidad'
DRIVER = '{ODBC Driver 17 for SQL Server}' 

CONNECTION_STRING_PYODBC = (f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;')

# Manejo de conexión a base de datos
try:
    params = urllib.parse.quote_plus(CONNECTION_STRING_PYODBC)
    engine = create_engine(f'mssql+pyodbc:///?odbc_connect={params}')
except Exception as e:
    print(f"Error inicializando SQLAlchemy: {e}")
    engine = None

CARRERAS_ITT = [
    "ISC (Sistemas)", "IIA (Industrial)", "IME (Mecánica)", "IGE (Gestión)", 
    "IEQ (Electrónica)", "IEL (Eléctrica)", "ICA (Civil)", "ARQ (Arquitectura)", 
    "LA (Administración)", "LCP (Contador Público)"
]
SEMESTRES_LIST = [str(i) for i in range(1, 16)]
DISCAPACIDADES_LIST = ["Ninguna", "Visual", "Auditiva", "Motriz", "Cognitiva", "Otra"]

BASE_FONT_SIZE = 10
CURRENT_FONT_SIZE = BASE_FONT_SIZE
COLOR_INVERTED = False 

# --- VARIABLES GLOBALES PARA FUENTES Y ACCESIBILIDAD ---
CURRENT_FONT_FAMILY = "Arial"
AVAILABLE_FONTS = ["Arial", "Times New Roman", "Courier New", "Verdana", "Tahoma", "Segoe UI", "Georgia", "Helvetica", "Comic Sans MS"]
DYSLEXIC_MODE = False 
# ----------------------------------------------

COLORBLIND_MODES = ["Modern Dark", "Modern Light", "Normal", "Deuteranopia (Rojo-Verde)", "Protanopia (Rojo-Verde)", "Tritanopia (Azul-Amarillo)"] 

CUSTOM_COLORS = { 
    "Modern Dark": {
        'bg_window': '#2b2b2b', 'fg_text': '#ffffff', 
        'bg_sidebar': '#1f1f1f', 'fg_sidebar': '#ffffff',
        'bg_button': '#3a3a3a', 'fg_button': '#ffffff',
        'bg_highlight': '#00adb5', 'fg_highlight': '#ffffff', # Cyan Accent
        'bg_entry': '#404040', 'fg_entry': '#ffffff', 'insert': 'white',
        'plot_bar': '#00adb5', 'plot_line': '#ff2e63'
    },
    "Modern Light": {
        'bg_window': '#f4f6f9', 'fg_text': '#333333',
        'bg_sidebar': '#ffffff', 'fg_sidebar': '#333333',
        'bg_button': '#e2e6ea', 'fg_button': '#333333',
        'bg_highlight': '#007bff', 'fg_highlight': '#ffffff',
        'bg_entry': '#ffffff', 'fg_entry': '#333333', 'insert': 'black',
        'plot_bar': '#007bff', 'plot_line': '#dc3545'
    },
    "Normal": {
        'bg_window': '#f0f0f0', 'fg_text': 'black', 
        'bg_sidebar': '#e0e0e0', 'fg_sidebar': 'black',
        'bg_button': '#e1e1e1', 'fg_button': 'black',
        'bg_highlight': '#0078d7', 'fg_highlight': 'white',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': 'tab:blue', 'plot_line': 'tab:red'
    },
    "Inversion": { 
        'bg_window': '#000000', 'fg_text': '#ffffff', 
        'bg_sidebar': '#333333', 'fg_sidebar': '#ffffff',
        'bg_button': '#333333', 'fg_button': '#ffffff',
        'bg_highlight': '#ff0000', 'fg_highlight': '#ffffff', 
        'bg_entry': '#000000', 'fg_entry': '#ffffff', 'insert': '#ffffff',
        'plot_bar': 'tab:red', 'plot_line': 'tab:blue'
    },
    "Deuteranopia (Rojo-Verde)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
        'bg_sidebar': '#e0e0e0', 'fg_sidebar': '#1c1c1c',
        'bg_button': '#e0e0e0', 'fg_button': '#1c1c1c', 
        'bg_highlight': '#008080', 'fg_highlight': '#ffffff',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': '#008080', 'plot_line': '#ffa500' 
    },
    "Protanopia (Rojo-Verde)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
        'bg_sidebar': '#e0e0e0', 'fg_sidebar': '#1c1c1c',
        'bg_button': '#e0e0e0', 'fg_button': '#1c1c1c', 
        'bg_highlight': '#ffa500', 'fg_highlight': '#1c1c1c',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': '#008080', 'plot_line': '#ffa500' 
    },
    "Tritanopia (Azul-Amarillo)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
        'bg_sidebar': '#e0e0e0', 'fg_sidebar': '#1c1c1c',
        'bg_button': '#e0e0e0', 'fg_button': '#1c1c1c', 
        'bg_highlight': '#800080', 'fg_highlight': '#ffffff',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': '#800080', 'plot_line': '#ffdb58' 
    }
}

# --- CLASE LUPA (VISUAL) ---
class MagnifierWindow(tk.Toplevel):
    def __init__(self, master, zoom=2, size=150):
        super().__init__(master)
        self.zoom = zoom
        self.size = size
        self.overrideredirect(True) 
        self.attributes("-topmost", True) 
        self.geometry(f"{size}x{size}")
        
        self.canvas = tk.Canvas(self, width=size, height=size, highlightthickness=2, highlightbackground="black")
        self.canvas.pack(fill="both", expand=True)
        self.update_magnifier()

    def update_magnifier(self):
        if not PIL_AVAILABLE:
            self.destroy()
            return
        try:
            x, y = self.winfo_pointerxy()
            self.geometry(f"+{x + 20}+{y + 20}")
            grab_radius = self.size // (2 * self.zoom)
            bbox = (x - grab_radius, y - grab_radius, x + grab_radius, y + grab_radius)
            image = ImageGrab.grab(bbox=bbox)
            image = image.resize((self.size, self.size), Image.NEAREST)
            self.photo = ImageTk.PhotoImage(image)
            self.canvas.create_image(0, 0, image=self.photo, anchor="nw")
            self.after(50, self.update_magnifier)
        except Exception as e:
            print(f"Error lupa: {e}")
            self.destroy()

# --- CLASE TECLADO EN PANTALLA (MOTORA) ---
class VirtualKeyboard(tk.Toplevel):
    def __init__(self, master, target_entry=None):
        super().__init__(master)
        self.title("Teclado Virtual Accesible")
        self.geometry("700x250")
        self.attributes("-topmost", True)
        self.target_entry = target_entry
        self.create_keys()

    def create_keys(self):
        keys = [
            ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', 'Borrar'],
            ['Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P'],
            ['A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'Ñ'],
            ['Z', 'X', 'C', 'V', 'B', 'N', 'M', ',', '.', 'Espacio']
        ]
        container = ttk.Frame(self, padding=5)
        container.pack(fill='both', expand=True)

        for r, row in enumerate(keys):
            frame_row = ttk.Frame(container)
            frame_row.pack(fill='x', expand=True)
            for key in row:
                width = 10 if key == 'Espacio' else 4
                btn = ttk.Button(frame_row, text=key, width=width,
                                 command=lambda k=key: self.press_key(k))
                btn.pack(side='left', fill='both', expand=True, padx=2, pady=2)

    def press_key(self, key):
        widget = self.master.focus_get() if not self.target_entry else self.target_entry
        if not widget or not isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox)):
             return

        try:
            if key == 'Borrar':
                if isinstance(widget, ttk.Combobox) and widget['state'] == 'readonly':
                    return
                current_pos = widget.index(tk.INSERT)
                if current_pos > 0:
                    widget.delete(current_pos - 1, tk.INSERT)
            elif key == 'Espacio':
                widget.insert(tk.INSERT, ' ')
            else:
                widget.insert(tk.INSERT, key)
        except Exception as e:
            print(f"Error teclado virtual: {e}")

# --- CLASE ASISTENTE DE VOZ ---
class VoiceAssistant:
    def __init__(self, app_reference):
        self.app = app_reference
        self.engine = pyttsx3.init() if TTS_AVAILABLE else None
        self.recognizer = sr.Recognizer() if VOICE_AVAILABLE else None
        self.is_listening = False  
        self.listen_thread = None

    def speak(self, text):
        if self.engine:
            def _speak():
                try:
                    self.engine.say(text)
                    self.engine.runAndWait()
                except: pass
            threading.Thread(target=_speak).start()

    def toggle_listening(self):
        if not self.recognizer: 
            return False
        
        if self.is_listening:
            self.is_listening = False
            self.app.update_status_voice("Voz desactivada", "#f0f0f0", "black")
            self.speak("Control por voz desactivado")
            return False
        else:
            self.is_listening = True
            self.listen_thread = threading.Thread(target=self._listen_loop)
            self.listen_thread.daemon = True 
            self.listen_thread.start()
            self.speak("Escuchando comandos")
            return True

    def _listen_loop(self):
        with sr.Microphone() as source:
            try:
                self.recognizer.adjust_for_ambient_noise(source)
            except: pass
            
            while self.is_listening:
                try:
                    self.app.update_status_voice("Escuchando...", "#ffcccc", "black")
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                    self.app.update_status_voice("Procesando...", "orange", "black")
                    text = self.recognizer.recognize_google(audio, language="es-ES")
                    print(f"Comando detectado: {text}")
                    self.app.process_voice_command(text.lower())
                except sr.WaitTimeoutError:
                    continue 
                except sr.UnknownValueError:
                    self.app.update_status_voice("...", "#f0f0f0", "black")
                except Exception as e:
                    print(f"Error en bucle de voz: {e}")
                    self.app.update_status_voice("Error mic", "red", "white")
                time.sleep(0.5)
        self.app.update_status_voice("Voz inactiva", "#f0f0f0", "black")

    def read_screen_content(self, widget):
        text_content = []
        def extract(w):
            try:
                if isinstance(w, (ttk.Label, tk.Label, ttk.Button, tk.Button)):
                    txt = w.cget("text")
                    clean_txt = str(txt).replace("**", "").strip()
                    if clean_txt:
                        text_content.append(clean_txt)
                elif isinstance(w, (ttk.Entry, tk.Entry)):
                    txt = w.get()
                    if txt: text_content.append(f"Campo de texto: {txt}")
                elif isinstance(w, ttk.Combobox):
                    txt = w.get()
                    if txt: text_content.append(f"Selección: {txt}")
                elif isinstance(w, ttk.Treeview):
                    text_content.append("Tabla de datos presente.")
            except: pass
            for child in w.winfo_children():
                extract(child)
        extract(widget)
        full_text = ". ".join(text_content)
        if not full_text: full_text = "La pantalla parece estar vacía."
        self.speak(f"Leyendo contenido: {full_text}")

# -----------------------------------------------

def conectar_sql_server():
    try:
        conn = pyodbc.connect(CONNECTION_STRING_PYODBC)
        return conn
    except pyodbc.Error as ex:
        error_msg = f"Error al conectar a SQL Server.\nDetalle: {ex}"
        messagebox.showerror("Error de Conexión", error_msg)
        return None

def log_actividad(user_id, tipo_accion, detalle=""):
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    user_id_safe = user_id if user_id is not None else None 
    sql_log = "INSERT INTO RegistroActividad (ID_Usuario_FK, Tipo_Accion, Detalle) VALUES (?, ?, ?)"
    try:
        cursor.execute(sql_log, (user_id_safe, tipo_accion, detalle[:4000])) 
        conn.commit()
    except Exception as e:
        print(f"Error al registrar actividad: {e}") 
    finally:
        if conn: conn.close()

def autenticar_usuario(nombre_usuario, contrasena):
    conn = conectar_sql_server()
    if not conn: return None
    cursor = conn.cursor()
    sql_auth = """SELECT ID_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK
    FROM Usuarios WHERE Nombre_Usuario = ?"""
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
        print(f"Error login: {e}")
        return None
    finally:
        if conn: conn.close()

def usuario_ya_existe(nombre_usuario):
    conn = conectar_sql_server()
    if not conn: return True, "Error de conexión"
    sql_query = "SELECT 1 FROM Usuarios WHERE Nombre_Usuario = ?"
    try:
        cursor = conn.cursor()
        cursor.execute(sql_query, (nombre_usuario,))
        existe = cursor.fetchone() is not None
        return existe, None
    except Exception as e:
        return True, f"Error BD: {e}" 
    finally:
        if conn: conn.close()

def registrar_estudiante_usuario(num_control, nombre, apellido_p, apellido_m, carrera, semestre, contrasena, discapacidad):
    conn = conectar_sql_server()
    if not conn: return "Error de conexión."
    cursor = conn.cursor()
    try:
        semestre_int = int(semestre) 
    except ValueError:
        return "El Semestre debe ser un número entero."

    existe, error_check = usuario_ya_existe(num_control)
    if error_check: return error_check
    if existe: return f"Usuario/Control '{num_control}' ya registrado."

    sql_estudiante = """INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""  
    sql_usuario = "INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) VALUES (?, ?, 2, ?)"
    try:
        cursor.execute("BEGIN TRANSACTION")
        datos_estudiante_full = (
            num_control, apellido_p, apellido_m, nombre, carrera, semestre_int, 
            'Sin Materia', 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
            0, 0, 0, 0, 0, 0, discapacidad 
        )
        cursor.execute(sql_estudiante, datos_estudiante_full)
        datos_usuario = (num_control, contrasena, num_control) 
        cursor.execute(sql_usuario, datos_usuario)
        conn.commit() 
        log_actividad(None, 'REGISTRO_NUEVO_ALUMNO', f"Registro estudiante: {num_control}")
        return "Registro exitoso."
    except pyodbc.IntegrityError:
        conn.rollback()
        return f"Error: El Número de Control ya existe."
    except Exception as e:
        conn.rollback()
        return f"Error al registrar: {e}"
    finally:
        if conn: conn.close()
        
def registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena, discapacidad):
    conn = conectar_sql_server()
    if not conn: return "Error de conexión."
    cursor = conn.cursor()
    existe, error_check = usuario_ya_existe(num_control)
    if error_check: return error_check
    if existe: return f"Usuario '{num_control}' ya registrado."
    
    sql_usuario = "INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK, Discapacidad) VALUES (?, ?, 1, ?, ?)"
    try:
        cursor.execute("BEGIN TRANSACTION")
        datos_usuario = (num_control, contrasena, None, discapacidad) 
        cursor.execute(sql_usuario, datos_usuario)
        conn.commit() 
        log_actividad(None, 'REGISTRO_NUEVO_PROFESOR', f"Registro profesor: {num_control}")
        return "Registro exitoso."
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
        if rol_id == 2:
            sql_query = "SELECT Discapacidad FROM Estudiantes WHERE Num_Control = ?"
            cursor.execute(sql_query, (num_control,))
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        elif rol_id == 1:
            sql_query = "SELECT U.Discapacidad FROM Usuarios U WHERE U.Nombre_Usuario = ? AND U.ID_Rol_FK = 1"
            cursor.execute(sql_query, (num_control,))
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        return "N/D"
    except Exception:
        return "N/D"
    finally:
        if conn: conn.close()

def insertar_registro_manual(datos, user_id):
    conn = conectar_sql_server()
    if not conn: return
    cursor = conn.cursor()
    sql_insert = """INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
    try:
        cursor.execute(sql_insert, datos)
        conn.commit() 
        log_actividad(user_id, 'INSERT_ESTUDIANTE', f"Registro exitoso: {datos[0]}")
        messagebox.showinfo("Éxito", f"Registro de {datos[3]} exitoso.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al insertar: {e}")
        conn.rollback()
    finally:
        if conn: conn.close()
        
def actualizar_datos_estudiante(num_control, nuevos_datos, user_id):
    conn = conectar_sql_server()
    if not conn: return False
    sql_update = """UPDATE Estudiantes SET Apellido_Paterno = ?, Apellido_Materno = ?, Nombre = ?, Discapacidad = ? WHERE Num_Control = ?"""
    try:
        params = nuevos_datos + [num_control] 
        cursor = conn.cursor()
        cursor.execute(sql_update, params)
        conn.commit() 
        if cursor.rowcount > 0:
            log_actividad(user_id, 'UPDATE_ESTUDIANTE', f"Actualizado: {num_control}")
            messagebox.showinfo("Éxito", "Datos actualizados.")
            return True
        else:
            messagebox.showerror("Error", "No se encontró el estudiante.")
            return False
    except Exception as e:
        messagebox.showerror("Error", f"Error al actualizar: {e}")
        conn.rollback()
        return False
    finally:
        if conn: conn.close()

def obtener_datos_estudiante(num_control):
    conn = conectar_sql_server()
    if not conn: return None
    sql_query = """SELECT Nombre, Apellido_Paterno, Apellido_Materno, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, 
        Calificacion_Unidad_4, Calificacion_Unidad_5, Discapacidad FROM Estudiantes WHERE Num_Control = ?"""
    try:
        cursor = conn.cursor()
        cursor.execute(sql_query, (num_control,))
        return cursor.fetchone()
    except Exception: return None
    finally:
        if conn: conn.close()

def importar_datos_a_sql(archivo_path, nombre_tabla, user_id):
    COLUMNAS_BASE = ['Num_Control', 'Apellido_Paterno', 'Apellido_Materno', 'Nombre', 'Carrera', 'Semestre']
    COLUMNAS_FACTORES = [
        'Materia', 'Calificacion_Unidad_1', 'Calificacion_Unidad_2', 'Calificacion_Unidad_3', 
        'Calificacion_Unidad_4', 'Calificacion_Unidad_5', 'Asistencia_Porcentaje',
        'Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico',
        'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual', 'Discapacidad' 
    ]
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
            messagebox.showerror("Error", "El archivo debe ser CSV o Excel.")
            return
        
        columnas_existentes = COLUMNAS_BASE + [c for c in COLUMNAS_FACTORES if c in df.columns]
        df_final = df[columnas_existentes].copy()
        
        for col in COLUMNAS_FACTORES:
            if col not in df_final.columns:
                df_final[col] = VALORES_DEFECTO[col]
        
        df_final['Semestre'] = df_final['Semestre'].astype(int)
        
        for i in range(1, 6):
            col = f'Calificacion_Unidad_{i}'
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(VALORES_DEFECTO[col])
            
        df_final.to_sql(nombre_tabla, con=engine, if_exists='append', index=False, method='multi')
        
        log_actividad(user_id, 'IMPORT_DATA', f"Importadas {len(df_final)} filas.")
        messagebox.showinfo("Importación", f"¡{len(df_final)} filas insertadas!")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def exportar_datos_sql(formato, user_id):
    try:
        df = pd.read_sql("SELECT * FROM Estudiantes", engine)
        if formato == 'excel':
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if filepath: df.to_excel(filepath, index=False)
        elif formato == 'csv':
            filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
            if filepath: df.to_csv(filepath, index=False, encoding='utf-8')
        if filepath:
            log_actividad(user_id, 'EXPORT_DATA', f"Exportados {len(df)} filas")
            messagebox.showinfo("Éxito", "Datos exportados.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar: {e}")

def generar_pareto_factores(user_id, carrera=None, semestre=None, materia=None, num_control_filtro=None, bar_color='tab:blue', line_color='tab:red'):
    factores = ['Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico', 
                'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual']
    columnas_suma = [f'SUM(CAST({f} AS INT)) AS {f}' for f in factores]
    sql_pareto_query = f"SELECT {', '.join(columnas_suma)} FROM Estudiantes"
    sql_estudiantes_query = "SELECT Num_Control, Nombre, Apellido_Paterno, Apellido_Materno, Semestre, Carrera FROM Estudiantes WHERE ({}) > 0"
    columnas_a_sumar = [f'ISNULL(CAST({f} AS INT), 0)' for f in factores]
    suma_factores_check = " + ".join(columnas_a_sumar)
    sql_estudiantes_query = sql_estudiantes_query.format(suma_factores_check)
    
    condiciones = []
    if carrera: condiciones.append(f"Carrera = '{carrera}'")
    if semestre: condiciones.append(f"Semestre = {semestre}") 
    if materia: condiciones.append(f"Materia = '{materia}'") 
    
    if num_control_filtro:
        condiciones = [f"Num_Control = '{num_control_filtro}'"]
        
    if condiciones:
        where_clause = " WHERE " + " AND ".join(condiciones)
        sql_pareto_query += where_clause
        sql_estudiantes_query += " AND " + " AND ".join(condiciones)

    try:
        df_pareto = pd.read_sql(sql_pareto_query, engine)
        log_actividad(user_id, 'CONSULTA_PARETO', "Generado gráfico Pareto")
        df_factores = df_pareto.T.rename(columns={0: 'Frecuencia'})
        df_factores = df_factores[df_factores['Frecuencia'] > 0]
        nombres_grafico = {
            'Factor_Academico': 'Académico','Factor_Psicosocial': 'Psicosocial', 'Factor_Economico': 'Económico',
            'Factor_Institucional': 'Institucional', 'Factor_Tecnologico': 'Tecnológico','Factor_Contextual': 'Contextual'
        }
        df_factores.index = df_factores.index.map(nombres_grafico)
        df_factores = df_factores.sort_values(by='Frecuencia', ascending=False)
        total_frecuencia = df_factores['Frecuencia'].sum()
        if total_frecuencia == 0: return None, pd.DataFrame(), "No hay datos."
        
        df_factores['Porcentaje'] = (df_factores['Frecuencia'] / total_frecuencia) * 100
        df_factores['Acumulado'] = df_factores['Porcentaje'].cumsum()
        
        plt.clf()
        fig, facto1 = plt.subplots(figsize=(8, 5))
        facto1.bar(df_factores.index, df_factores['Frecuencia'], color=bar_color)
        facto1.set_ylabel('Frecuencia', color=bar_color)
        facto2 = facto1.twinx()
        facto2.plot(df_factores.index, df_factores['Acumulado'], color=line_color, marker='D', ms=5)
        facto2.set_ylabel('Porcentaje Acumulado (%)', color=line_color)
        facto2.set_ylim(0, 110)
        plt.title('Diagrama de Pareto: Factores de Riesgo')
        
        df_estudiantes = pd.read_sql(sql_estudiantes_query, engine)
        return fig, df_estudiantes, None
    except Exception as e:
        return None, None, str(e)

# --- VENTANA DE REGISTRO ---
class RegisterWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Registro de Usuario")
        self.geometry("500x600")
        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.frame_alumno = ttk.Frame(notebook)
        self.frame_profesor = ttk.Frame(notebook)
        notebook.add(self.frame_alumno, text='Estudiante')
        notebook.add(self.frame_profesor, text='Profesor')
        
        self.crear_form_alumno(self.frame_alumno)
        self.crear_form_profesor(self.frame_profesor)

    def crear_form_alumno(self, parent):
        main_frame = ttk.Frame(parent, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        fields = [
            ("No. Control (*):", "num_control", 0, 'entry', None),
            ("Nombre (*):", "nombre", 1, 'entry', None),
            ("Apellido Paterno (*):", "apellido_p", 2, 'entry', None),
            ("Apellido Materno:", "apellido_m", 3, 'entry', ''),
            ("Carrera (*):", "carrera", 5, 'combobox', CARRERAS_ITT),
            ("Semestre (*):", "semestre", 6, 'combobox', SEMESTRES_LIST),
            ("Discapacidad:", "discapacidad", 7, 'combobox', DISCAPACIDADES_LIST),
            ("Contraseña (*):", "contrasena", 8, 'entry', '*')
        ]
        self.reg_vars = {}
        for label_text, var_name, row, widget_type, values in fields:
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, sticky='w', pady=5)
            var = tk.StringVar()
            self.reg_vars[var_name] = var
            if widget_type == 'entry':
                show = '*' if var_name == 'contrasena' else ''
                ttk.Entry(main_frame, textvariable=var, show=show).grid(row=row, column=1, pady=5)
            elif widget_type == 'combobox':
                cb = ttk.Combobox(main_frame, textvariable=var, values=values, state='readonly')
                cb.grid(row=row, column=1, pady=5)
                if values: cb.current(0)
        
        ttk.Button(main_frame, text="Registrar", command=self.handle_registration).grid(row=9, columnspan=2, pady=20)

    def crear_form_profesor(self, parent):
        # Simplificado para brevedad, lógica similar al de alumno
        pass 

    def handle_registration(self):
        # Lógica de registro simplificada
        data = {k: v.get().strip() for k, v in self.reg_vars.items()}
        if not all([data['num_control'], data['nombre'], data['contrasena']]):
            messagebox.showwarning("Faltan datos", "Campos obligatorios vacíos.")
            return
        res = registrar_estudiante_usuario(data['num_control'], data['nombre'], data['apellido_p'], data['apellido_m'], data['carrera'], data['semestre'], data['contrasena'], data['discapacidad'])
        if "exitoso" in res:
            messagebox.showinfo("Registro", res)
            self.destroy()
        else:
            messagebox.showerror("Error", res)

# --- VENTANA DE LOGIN ---
class LoginWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Inicio de Sesión")
        self.master = master
        self.master.login_window = self 
        self.geometry("400x250")
        self.master.withdraw() 
        self.protocol("WM_DELETE_WINDOW", self.master.quit)
        self.create_widgets()
        master.apply_theme_settings() 

    def create_widgets(self):
        frame = ttk.Frame(self, padding=20)
        frame.pack(expand=True, fill='both')
        
        ttk.Label(frame, text="SISTEMA DE CALIDAD", font=("Arial", 14, "bold")).pack(pady=10)
        
        ttk.Label(frame, text="Usuario / No. Control:").pack(anchor='w')
        self.user_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.user_var).pack(fill='x', pady=5)
        
        ttk.Label(frame, text="Contraseña:").pack(anchor='w')
        self.pass_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.pass_var, show="*").pack(fill='x', pady=5)
        
        ttk.Button(frame, text="Iniciar Sesión", command=self.login).pack(pady=10, fill='x')
        ttk.Button(frame, text="Registrarse", command=self.open_register).pack(pady=5, fill='x')

    def login(self):
        user = self.user_var.get()
        pwd = self.pass_var.get()
        res = autenticar_usuario(user, pwd)
        if res:
            self.master.show_main_window(*res)
        else:
            messagebox.showerror("Error", "Credenciales incorrectas")

    def open_register(self):
        RegisterWindow(self)

# --- VENTANA PRINCIPAL (MODERNIZADA) ---
class MainWindow(tk.Toplevel):
    def __init__(self, master, user_id, role_id, num_control):
        super().__init__(master)
        self.master = master
        self.user_id = user_id
        self.role_id = role_id
        self.num_control = num_control
        
        # Referencia inversa
        self.master.main_window = self
        
        # Inicializar helpers
        self.magnifier = None
        self.keyboard_window = None
        self.voice_assistant = VoiceAssistant(self)
        self.focus_mode_active = False

        self.title("Sistema de Gestión de Calidad Académica ITT")
        self.geometry("1100x700")
        self.protocol("WM_DELETE_WINDOW", self.cerrar_sesion)
        
        self.style = ttk.Style(self)
        self.frames = {} # Para almacenar las páginas
        
        # Aplicar tema moderno inicial
        if hasattr(self.master, 'colorblind_mode') and self.master.colorblind_mode.get() == "Normal":
             self.master.colorblind_mode.set("Modern Dark")
        
        self.master.apply_theme_settings()
        self.master.update_font_size(self)

        self.crear_barra_accesibilidad()
        self.create_widgets()
        
        # Status Bar
        self.status_label = ttk.Label(self, text="Listo", relief=tk.SUNKEN, anchor='w')
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def create_widgets(self):
        # GRID PRINCIPAL: Columna 0 = Sidebar, Columna 1 = Contenido
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1) # Row 0 is accessibility bar

        # 1. SIDEBAR (Menú Lateral)
        self.sidebar_frame = tk.Frame(self, width=220, bg='#1f1f1f')
        self.sidebar_frame.grid(row=1, column=0, sticky="ns")
        self.sidebar_frame.grid_propagate(False)

        lbl_logo = tk.Label(self.sidebar_frame, text="ITT Calidad", font=("Arial", 16, "bold"), bg='#1f1f1f', fg='white')
        lbl_logo.pack(pady=25)

        # 2. ÁREA DE CONTENIDO
        self.content_area = tk.Frame(self)
        self.content_area.grid(row=1, column=1, sticky="nsew")
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Crear Frames (Páginas)
        for F in ("Perfil", "Registro", "Pareto", "Importar", "Auditoria", "Configuracion", "DatosAlumno"):
            frame = ttk.Frame(self.content_area)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Asignar alias para compatibilidad con código antiguo
        self.tab_perfil = self.frames["Perfil"]
        self.tab_registro_manual = self.frames["Registro"]
        self.tab_pareto = self.frames["Pareto"]
        self.tab_importar = self.frames["Importar"]
        self.tab_auditoria = self.frames["Auditoria"]
        self.tab_configuracion = self.frames["Configuracion"]
        self.tab_datos = self.frames["DatosAlumno"]

        # Rellenar contenido
        self.crear_tab_perfil()
        
        if self.role_id == 1: # Profesor
            self.crear_tab_registro()
            self.crear_tab_pareto()
            self.crear_tab_importar()
            self.crear_tab_auditoria()
            # Botones Sidebar
            self.crear_boton_menu("👤 Perfil", "Perfil")
            self.crear_boton_menu("📝 Registro", "Registro")
            self.crear_boton_menu("📊 Análisis", "Pareto")
            self.crear_boton_menu("📂 Datos", "Importar")
            self.crear_boton_menu("📋 Auditoría", "Auditoria")
            
        elif self.role_id == 2: # Estudiante
            self.crear_tab_datos_alumno()
            self.crear_boton_menu("👤 Perfil", "Perfil")
            self.crear_boton_menu("🎓 Mis Datos", "DatosAlumno")
        
        self.crear_boton_menu("⚙️ Configuración", "Configuracion")
        self.crear_tab_configuracion()

        # Botón Salir
        btn_salir = tk.Button(self.sidebar_frame, text="🚪 Cerrar Sesión", command=self.cerrar_sesion, 
                              bg="#d9534f", fg="white", relief="flat", padx=20, pady=10, cursor="hand2")
        btn_salir.pack(side="bottom", fill="x", pady=20, padx=10)

        self.show_frame("Perfil")

    def crear_boton_menu(self, text, frame_name):
        btn = tk.Button(self.sidebar_frame, text=text, font=("Arial", 11), anchor="w",
                        command=lambda: self.show_frame(frame_name),
                        bg='#1f1f1f', fg='white', relief="flat", padx=20, pady=10, cursor="hand2")
        btn.pack(fill="x", pady=2)
        # Hover Effect
        btn.bind("<Enter>", lambda e: btn.config(bg="#00adb5")) 
        btn.bind("<Leave>", lambda e: btn.config(bg="#1f1f1f"))

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def crear_barra_accesibilidad(self):
        toolbar = ttk.Frame(self, padding="5", relief="raised")
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        
        ttk.Label(toolbar, text="Accesibilidad:", font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        self.btn_voice = ttk.Button(toolbar, text="🎙️ Voz", command=self.toggle_voice_handler)
        self.btn_voice.pack(side=tk.LEFT, padx=2)
        self.lbl_mic_status = ttk.Label(toolbar, text="Off", foreground="gray")
        self.lbl_mic_status.pack(side=tk.LEFT, padx=2)
        
        if TTS_AVAILABLE:
            ttk.Button(toolbar, text="🔊 Leer", command=self.read_active_tab).pack(side=tk.LEFT, padx=2)
            
        ttk.Button(toolbar, text="🔍 Lupa", command=self.toggle_magnifier).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⌨️ Teclado", command=self.toggle_virtual_keyboard).pack(side=tk.LEFT, padx=2)
        
        self.btn_focus = ttk.Button(toolbar, text="🧠 Enfoque", command=self.toggle_focus_mode)
        self.btn_focus.pack(side=tk.LEFT, padx=2)

    def toggle_voice_handler(self):
        active = self.voice_assistant.toggle_listening()
        if active:
            self.btn_voice.config(text="🛑 Detener Voz")
            self.lbl_mic_status.config(text="On", foreground="green")
        else:
            self.btn_voice.config(text="🎙️ Activar Voz")
            self.lbl_mic_status.config(text="Off", foreground="gray")

    def read_active_tab(self):
        # Identificar qué frame está visible
        visible_frame = None
        for name, frame in self.frames.items():
            # Un frame está visible si está en el tope del stack, pero en grid es difícil saber cual está arriba.
            # Sin embargo, como usamos tkraise, es el que está visualmente. 
            # Leemos el contenido de self.content_area que es el padre
            pass
        self.voice_assistant.read_screen_content(self.content_area)

    def process_voice_command(self, text):
        if "perfil" in text: self.show_frame("Perfil")
        elif "registro" in text: self.show_frame("Registro")
        elif "salir" in text: self.cerrar_sesion()
        elif "lupa" in text: self.toggle_magnifier()
        elif "leer" in text: self.read_active_tab()
        else: self.voice_assistant.speak("Comando no reconocido")

    def toggle_magnifier(self):
        if self.magnifier: 
            self.magnifier.destroy()
            self.magnifier = None
        else:
            self.magnifier = MagnifierWindow(self.master)

    def toggle_virtual_keyboard(self):
        if self.keyboard_window:
            self.keyboard_window.destroy()
            self.keyboard_window = None
        else:
            self.keyboard_window = VirtualKeyboard(self.master)

    def toggle_focus_mode(self):
        self.focus_mode_active = not self.focus_mode_active
        if self.focus_mode_active:
            self.sidebar_frame.grid_remove() # Ocultar Sidebar
            self.btn_focus.config(text="Mostrar Menú")
        else:
            self.sidebar_frame.grid() # Mostrar Sidebar
            self.btn_focus.config(text="🧠 Enfoque")

    def update_status_voice(self, text, bg, fg):
        self.status_label.config(text=f"Voz: {text}", background=bg, foreground=fg)

    # --- FUNCIONES DE CREACIÓN DE CONTENIDO (Mantenidas del original) ---
    def crear_tab_perfil(self):
        frame = ttk.Frame(self.tab_perfil, padding="30")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="PERFIL DE USUARIO", font=('Arial', 20, 'bold')).pack(pady=20)
        
        datos = obtener_datos_estudiante(self.num_control) if self.role_id == 2 else None
        
        info_frame = ttk.Frame(frame, relief="solid", borderwidth=1, padding=20)
        info_frame.pack()
        
        rol_str = "Profesor" if self.role_id == 1 else "Estudiante"
        ttk.Label(info_frame, text=f"Rol: {rol_str}", font=("Arial", 12)).pack(pady=5)
        ttk.Label(info_frame, text=f"Usuario: {self.num_control}", font=("Arial", 12, "bold")).pack(pady=5)
        
        if datos:
            ttk.Label(info_frame, text=f"Nombre: {datos[0]} {datos[1]}", font=("Arial", 12)).pack(pady=5)
            ttk.Label(info_frame, text=f"Carrera: {datos[3]}", font=("Arial", 12)).pack(pady=5)

    def crear_tab_registro(self):
        frame = ttk.Frame(self.tab_registro_manual, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Registro Manual", font=('Arial', 18)).pack(pady=10)
        # (Aquí iría el resto de tu código de registro original, simplificado para el ejemplo)
        # Puedes copiar el contenido de tu función original crear_tab_registro aquí

    def crear_tab_pareto(self):
        frame = ttk.Frame(self.tab_pareto, padding="10")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Análisis de Pareto", font=('Arial', 18)).pack()
        
        btn = ttk.Button(frame, text="Generar Gráfico", command=self.actualizar_pareto)
        btn.pack(pady=10)
        
        self.pareto_canvas_frame = ttk.Frame(frame)
        self.pareto_canvas_frame.pack(fill='both', expand=True)

    def actualizar_pareto(self):
        # Wrapper para llamar a la función global
        fig, _, err = generar_pareto_factores(self.user_id)
        if fig:
            for widget in self.pareto_canvas_frame.winfo_children(): widget.destroy()
            canvas = FigureCanvasTkAgg(fig, master=self.pareto_canvas_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
        elif err:
            messagebox.showinfo("Info", err)

    def crear_tab_importar(self):
        frame = ttk.Frame(self.tab_importar, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Importar / Exportar", font=('Arial', 18)).pack()
        
        ttk.Button(frame, text="Importar Excel/CSV", command=self.seleccionar_archivo_import).pack(pady=10)
        ttk.Button(frame, text="Exportar Excel", command=lambda: exportar_datos_sql('excel', self.user_id)).pack(pady=5)

    def seleccionar_archivo_import(self):
        path = filedialog.askopenfilename(filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls")])
        if path: importar_datos_a_sql(path, 'Estudiantes', self.user_id)

    def crear_tab_auditoria(self):
        # Simplificado
        ttk.Label(self.tab_auditoria, text="Auditoría (Logs)", font=("Arial", 16)).pack(pady=20)

    def crear_tab_datos_alumno(self):
        # Simplificado
        ttk.Label(self.tab_datos, text="Mis Datos", font=("Arial", 16)).pack(pady=20)

    def crear_tab_configuracion(self):
        frame = ttk.Frame(self.tab_configuracion, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Configuración Visual", font=("Arial", 16)).pack(pady=10)
        
        # Selector de Tema
        ttk.Label(frame, text="Tema de Color:").pack()
        self.combo_tema = ttk.Combobox(frame, values=COLORBLIND_MODES, state='readonly')
        current_mode = self.master.colorblind_mode.get() if hasattr(self.master, 'colorblind_mode') else "Modern Dark"
        self.combo_tema.set(current_mode)
        self.combo_tema.pack(pady=5)
        self.combo_tema.bind("<<ComboboxSelected>>", self.cambiar_tema)

    def cambiar_tema(self, event):
        nuevo_tema = self.combo_tema.get()
        self.master.colorblind_mode.set(nuevo_tema)
        self.master.apply_theme_settings()

    def cerrar_sesion(self):
        self.voice_assistant.is_listening = False
        if self.magnifier: self.magnifier.destroy()
        if self.keyboard_window: self.keyboard_window.destroy()
        self.destroy()
        self.master.login_window.deiconify()

# --- CLASE APP PRINCIPAL (ROOT) ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw() 
        self.title("Sistema ITT")
        self.colorblind_mode = tk.StringVar(value="Modern Dark") # Default moderno
        self.main_window = None
        self.login_window = None
        
        self.login_window = LoginWindow(self)

    def show_main_window(self, user_id, role_id, num_control):
        if self.main_window: self.main_window.destroy()
        self.main_window = MainWindow(self, user_id, role_id, num_control)

    def apply_theme_settings(self):
        mode = self.colorblind_mode.get()
        colors = CUSTOM_COLORS.get(mode, CUSTOM_COLORS["Modern Dark"])
        
        style = ttk.Style(self)
        style.theme_use('clam') 
        
        # Configurar colores TTK
        style.configure('.', background=colors['bg_window'], foreground=colors['fg_text'])
        style.configure('TFrame', background=colors['bg_window'])
        style.configure('TLabel', background=colors['bg_window'], foreground=colors['fg_text'])
        style.configure('TButton', background=colors['bg_button'], foreground=colors['fg_button'], 
                        borderwidth=1, focuscolor=colors['bg_highlight'])
        style.map('TButton', background=[('active', colors['bg_highlight'])], foreground=[('active', colors['fg_highlight'])])
        
        # Configurar Entry
        style.configure('TEntry', fieldbackground=colors['bg_entry'], foreground=colors['fg_entry'])
        
        # Actualizar ventanas hijas manualmente (Tk puro)
        if self.main_window:
            self.main_window.configure(bg=colors['bg_window'])
            if hasattr(self.main_window, 'sidebar_frame'):
                sb_bg = colors.get('bg_sidebar', '#333333')
                self.main_window.sidebar_frame.configure(bg=sb_bg)
                # Actualizar botones del sidebar
                for child in self.main_window.sidebar_frame.winfo_children():
                    if isinstance(child, tk.Button):
                        if "Cerrar" not in child['text']:
                            child.configure(bg=sb_bg, fg=colors.get('fg_sidebar', 'white'))
            
            if hasattr(self.main_window, 'content_area'):
                self.main_window.content_area.configure(bg=colors['bg_window'])

    def update_font_size(self, window):
        pass # Implementar si se requiere cambio de fuente global

if __name__ == "__main__":
    app = App()
    app.mainloop()