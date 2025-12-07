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

# Manejo de conexión a base de datos si falla la creación del engine al inicio
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

COLORBLIND_MODES = ["Normal", "Deuteranopia (Rojo-Verde)", "Protanopia (Rojo-Verde)", "Tritanopia (Azul-Amarillo)"] 

CUSTOM_COLORS = { 
    "Normal": {
        'bg_window': '#f0f0f0', 'fg_text': 'black', 
        'bg_button': '#e1e1e1', 'fg_button': 'black',
        'bg_highlight': '#0078d7', 'fg_highlight': 'white',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': 'tab:blue', 'plot_line': 'tab:red'
    },
    "Inversion": { 
        'bg_window': '#000000', 'fg_text': '#ffffff', 
        'bg_button': '#333333', 'fg_button': '#ffffff',
        'bg_highlight': '#ff0000', 'fg_highlight': '#ffffff', 
        'bg_entry': '#000000', 'fg_entry': '#ffffff', 'insert': '#ffffff',
        'plot_bar': 'tab:red', 'plot_line': 'tab:blue'
    },
    "Deuteranopia (Rojo-Verde)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
        'bg_button': '#e0e0e0', 'fg_button': '#1c1c1c', 
        'bg_highlight': '#008080', 'fg_highlight': '#ffffff',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': '#008080', 'plot_line': '#ffa500' 
    },
    "Protanopia (Rojo-Verde)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
        'bg_button': '#e0e0e0', 'fg_button': '#1c1c1c', 
        'bg_highlight': '#ffa500', 'fg_highlight': '#1c1c1c',
        'bg_entry': 'white', 'fg_entry': 'black', 'insert': 'black',
        'plot_bar': '#008080', 'plot_line': '#ffa500' 
    },
    "Tritanopia (Azul-Amarillo)": {
        'bg_window': '#f0f0f0', 'fg_text': '#1c1c1c', 
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
        # Intenta escribir en el widget que tiene foco
        widget = self.master.focus_get() if not self.target_entry else self.target_entry
        
        # Si el foco no está en un entry, intentamos buscar el último activo
        if not widget or not isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox)):
             # Fallback simple: no hace nada si no hay entry seleccionado
             return

        try:
            if key == 'Borrar':
                # Combobox readonly no permite borrar así, solo Entries
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

# --- CLASE ASISTENTE DE VOZ ACTUALIZADA (CONTINUA Y ROBUSTA) ---
class VoiceAssistant:
    def __init__(self, app_reference):
        self.app = app_reference
        self.engine = pyttsx3.init() if TTS_AVAILABLE else None
        self.recognizer = sr.Recognizer() if VOICE_AVAILABLE else None
        self.is_listening = False  # Bandera para control del bucle
        self.listen_thread = None

    def speak(self, text):
        """Lectura en voz alta (Cognitiva/Visual)"""
        if self.engine:
            def _speak():
                try:
                    self.engine.say(text)
                    self.engine.runAndWait()
                except: pass
            threading.Thread(target=_speak).start()

    def toggle_listening(self):
        """Activa o desactiva la escucha continua"""
        if not self.recognizer: 
            return False
        
        if self.is_listening:
            # Desactivar
            self.is_listening = False
            self.app.update_status_voice("Voz desactivada", "#f0f0f0", "black")
            self.speak("Control por voz desactivado")
            return False
        else:
            # Activar
            self.is_listening = True
            self.listen_thread = threading.Thread(target=self._listen_loop)
            self.listen_thread.daemon = True # Se cierra si la app se cierra
            self.listen_thread.start()
            self.speak("Escuchando comandos")
            return True

    def _listen_loop(self):
        """Bucle infinito (mientras activo) para escuchar comandos"""
        with sr.Microphone() as source:
            try:
                self.recognizer.adjust_for_ambient_noise(source)
            except: pass
            
            while self.is_listening:
                try:
                    # Feedback visual: Escuchando
                    self.app.update_status_voice("Escuchando...", "#ffcccc", "black")
                    
                    # Escucha con un límite de tiempo para no bloquear el hilo eternamente
                    # phrase_time_limit hace que corte si hablas mucho tiempo
                    # timeout hace que salte excepción si hay silencio total
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                    
                    self.app.update_status_voice("Procesando...", "orange", "black")
                    text = self.recognizer.recognize_google(audio, language="es-ES")
                    
                    print(f"Comando detectado: {text}")
                    self.app.process_voice_command(text.lower())
                    
                except sr.WaitTimeoutError:
                    # Nadie habló en el intervalo, seguimos escuchando
                    continue 
                except sr.UnknownValueError:
                    # No se entendió, seguimos escuchando sin molestar
                    self.app.update_status_voice("...", "#f0f0f0", "black")
                except Exception as e:
                    print(f"Error en bucle de voz: {e}")
                    self.app.update_status_voice("Error mic", "red", "white")
                
                # Pequeña pausa para no saturar CPU
                time.sleep(0.5)

        # Al salir del while
        self.app.update_status_voice("Voz inactiva", "#f0f0f0", "black")

    def read_screen_content(self, widget):
        """Extrae texto recursivamente de la pestaña activa"""
        text_content = []
        
        def extract(w):
            try:
                # Extraer texto de Labels y Botones
                if isinstance(w, (ttk.Label, tk.Label, ttk.Button, tk.Button)):
                    txt = w.cget("text")
                    # Evitar leer asteriscos de formato o textos vacíos
                    clean_txt = str(txt).replace("**", "").strip()
                    if clean_txt:
                        text_content.append(clean_txt)
                # Extraer valores de Entradas
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
        if not full_text: full_text = "La pantalla parece estar vacía o no tiene texto legible."
        self.speak(f"Leyendo contenido: {full_text}")

# -----------------------------------------------

def conectar_sql_server():
    """Establece la conexión a SQL Server."""
    try:
        conn = pyodbc.connect(CONNECTION_STRING_PYODBC)
        return conn
    except pyodbc.Error as ex:
        error_msg = f"Error al conectar a SQL Server.\nRevisa tu DRIVER, ODBC y la variable SERVER.\nDetalle: {ex}"
        messagebox.showerror("Error de Conexión", error_msg)
        print(error_msg)
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
        print(f"Error al registrar actividad (ID {user_id_safe}): {e}") 
    finally:
        if conn: conn.close()

def autenticar_usuario(nombre_usuario, contrasena):
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
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """  
    sql_usuario = "INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) VALUES (?, ?, 2, ?)"
    try:
        cursor.execute("BEGIN TRANSACTION")
        datos_estudiante_full = (
            num_control, apellido_p, apellido_m, nombre, carrera, semestre_int, 
            'Sin Materia', 
            0.0, 0.0, 0.0, 0.0, 0.0, 
            0.0, 
            0, 0, 0, 0, 0, 0,
            discapacidad 
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
    conn = conectar_sql_server()
    if not conn: return "Error de conexión con la base de datos."
    cursor = conn.cursor()
    existe, error_check = usuario_ya_existe(num_control)
    if error_check:
        return error_check
    if existe:
        return f"Error: El usuario o No. Control '{num_control}' ya está registrado en el sistema."
    
    sql_usuario = "INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK, Discapacidad) VALUES (?, ?, 1, ?, ?)"
    try:
        cursor.execute("BEGIN TRANSACTION")
        datos_usuario = (num_control, contrasena, None, discapacidad) 
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
            cursor.execute(sql_query, (num_control,))
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        elif rol_id == 1: # Profesor
            sql_query = "SELECT U.Discapacidad FROM Usuarios U WHERE U.Nombre_Usuario = ? AND U.ID_Rol_FK = 1"
            cursor.execute(sql_query, (num_control,))
            resultado = cursor.fetchone()
            return resultado[0] if resultado else "N/D"
        return "N/D"
    except pyodbc.ProgrammingError:
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
    sql_insert = """
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual, Discapacidad
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    try:
        if len(datos) != 20: 
            raise ValueError(f"Error interno: Se esperaban 20 parámetros (columnas), pero se recibieron {len(datos)}.")
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
             messagebox.showerror("Error de Tipo de Dato", "Error de conversión (varchar a int): Revise si Carrera y Semestre tienen tipos invertidos en SQL Server.")
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
    SET Apellido_Paterno = ?, Apellido_Materno = ?, Nombre = ?, Discapacidad = ?
    WHERE Num_Control = ?
    """
    try:
        params = nuevos_datos + [num_control] 
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
        cursor.execute(sql_query, (num_control,))
        return cursor.fetchone()
    except Exception as e:
        print(f"Error al obtener datos del estudiante: {e}")
        return None
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
            messagebox.showerror("Error de Formato", "El archivo debe ser CSV o Excel.")
            return
        
        if not all(col in df.columns for col in COLUMNAS_BASE):
            messagebox.showerror("Error de Columnas", 
                                 f"El archivo debe contener las siguientes 6 columnas base: {', '.join(COLUMNAS_BASE)}. Revise nombres exactos.")
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
        for col in [f for f in COLUMNAS_FACTORES if f.startswith('Factor_')]:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(VALORES_DEFECTO[col]).astype(int)
            
        df_final.to_sql(nombre_tabla, con=engine, if_exists='append', index=False, method='multi')
        
        log_actividad(user_id, 'IMPORT_DATA', f"Importadas {len(df_final)} filas.")
        messagebox.showinfo("Importación Exitosa", f"¡{len(df_final)} filas insertadas en SQL Server!")

    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo no fue encontrado.")
    except Exception as e:
        error_msg = str(e)
        if 'La conversión del valor varchar' in error_msg:
             messagebox.showerror("Error Crítico de Tipo de Dato", "Error de conversión: Carrera y Semestre tienen tipos invertidos.")
        else:
            messagebox.showerror("Error de Importación", f"Ocurrió un error: {error_msg}.")

def exportar_datos_sql(formato, user_id):
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
        where_clause = " WHERE " + " AND ".join(condiciones)
        sql_pareto_query += where_clause
        sql_estudiantes_query += " AND " + " AND ".join(condiciones)

    try:
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
        facto1.bar(df_factores.index, df_factores['Frecuencia'], color=bar_color)
        facto1.set_xlabel('Factores de Riesgo')
        facto1.set_ylabel('Frecuencia', color=bar_color)
        facto1.tick_params(axis='y', labelcolor=bar_color)
        
        paret = facto1.twinx()
        paret.plot(df_factores.index, df_factores['Acumulado'], color=line_color, marker='o')
        paret.set_ylabel('Acumulado %', color=line_color)
        paret.tick_params(axis='y', labelcolor=line_color)
        paret.set_ylim(0, 100)
        paret.axhline(80, color='gray', linestyle='--')

        plt.title("Análisis de Pareto: Factores de Riesgo")
        fig.tight_layout()
        
        df_estudiantes = pd.read_sql(sql_estudiantes_query, engine)
        df_estudiantes['Nombre Completo'] = df_estudiantes['Nombre'] + ' ' + df_estudiantes['Apellido_Paterno'] + ' ' + df_estudiantes['Apellido_Materno']
        df_estudiantes = df_estudiantes[['Num_Control', 'Nombre Completo', 'Semestre', 'Carrera']]
        
        return fig, df_estudiantes, None

    except Exception as e:
        log_actividad(user_id, 'ERROR_CONSULTA', f"Error en Pareto: {e}")
        return None, pd.DataFrame(), f"Error al consultar o generar gráfico: {e}"

def obtener_registro_auditoria():
    conn = conectar_sql_server()
    if not conn: return pd.DataFrame()
    sql_query = """
    SELECT U.Nombre_Usuario AS Matricula, R.Tipo_Accion AS Accion, R.Fecha_Hora AS Fecha_Hora
    FROM RegistroActividad R JOIN Usuarios U ON R.ID_Usuario_FK = U.ID_Usuario
    WHERE R.Tipo_Accion IN ('LOGIN_EXITOSO', 'LOGOUT') ORDER BY R.Fecha_Hora DESC
    """
    try:
        df = pd.read_sql(sql_query, engine)
        if not df.empty:
            df['Dia'] = df['Fecha_Hora'].dt.strftime('%Y-%m-%d')
            df['Hora'] = df['Fecha_Hora'].dt.strftime('%H:%M:%S')
            df['Accion'] = df['Accion'].replace({'LOGIN_EXITOSO': 'Inicio de Sesión', 'LOGOUT': 'Cierre de Sesión'})
            df = df[['Matricula', 'Accion', 'Dia', 'Hora']]
        return df
    except Exception as e:
        print(f"Error al obtener el registro de auditoría: {e}")
        return pd.DataFrame()
    finally:
        if conn: conn.close()

class StudentRegistrationWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title("Registro de Estudiante")
        self.geometry("400x440") 
        self.resizable(False, False)
        self.grab_set()
        master.apply_theme_settings() 
        self.create_widgets()
        master.update_font_size(self) 

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)
        ttk.Label(main_frame, text="**REGISTRO DE ESTUDIANTE**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        fields = [
            ("No. Control (*):", "num_control", 1, 'entry', ''), 
            ("Nombre(s) (*):", "nombre", 2, 'entry', ''),
            ("Apellido Paterno (*):", "apellido_p", 3, 'entry', ''), 
            ("Apellido Materno:", "apellido_m", 4, 'entry', ''), 
            ("Carrera (*):", "carrera", 5, 'combobox', CARRERAS_ITT), 
            ("Semestre (*):", "semestre", 6, 'combobox', SEMESTRES_LIST), 
            ("Discapacidad:", "discapacidad", 7, 'combobox', DISCAPACIDADES_LIST),
            ("Contraseña (*):", "contrasena", 8, 'entry', '*')
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
                if var_name == 'num_control': label.configure(underline=0); entry.focus() 
            elif widget_type == 'combobox': 
                combo = ttk.Combobox(main_frame, textvariable=var, values=values, width=28, state='readonly')
                combo.set(values[0] if values else '') 
                combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration, underline=0).pack(side=tk.LEFT, padx=10) 
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy, underline=0).pack(side=tk.LEFT, padx=10) 
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            carrera = self.reg_vars['carrera'].get().strip() 
            semestre_str = self.reg_vars['semestre'].get().strip()
            discapacidad = self.reg_vars['discapacidad'].get().strip() 
            contrasena = self.reg_vars['contrasena'].get().strip()
            if not all([num_control, nombre, apellido_p, carrera, semestre_str, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos principales son obligatorios.")
                return
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
        self.geometry("400x320") 
        self.resizable(False, False)
        self.grab_set()
        master.apply_theme_settings()
        self.create_widgets()
        master.update_font_size(self) 

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill='both', expand=True)
        ttk.Label(main_frame, text="**REGISTRO DE PROFESOR**", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        fields = [
            ("Usuario (Ej. No. Empleado) (*):", "num_control", 1, '', None), 
            ("Nombre(s) (*):", "nombre", 2, '', None),
            ("Apellido Paterno (*):", "apellido_p", 3, '', None), 
            ("Apellido Materno:", "apellido_m", 4, '', None), 
            ("Discapacidad:", "discapacidad", 5, 'combobox', DISCAPACIDADES_LIST), 
            ("Contraseña (*):", "contrasena", 6, 'entry', '*')
        ]
        self.reg_vars = {}
        for i, (label_text, var_name, row, widget_type, values) in enumerate(fields):
            label = ttk.Label(main_frame, text=label_text)
            label.grid(row=row, column=0, padx=5, pady=5, sticky='w')
            var = tk.StringVar()
            self.reg_vars[var_name] = var
            if widget_type == 'entry' or widget_type == '':
                show = '*' if var_name == 'contrasena' else ''
                entry = ttk.Entry(main_frame, textvariable=var, width=30, show=show)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky='w')
                if var_name == 'num_control': label.configure(underline=0); entry.focus()
            elif widget_type == 'combobox':
                combo = ttk.Combobox(main_frame, textvariable=var, values=values, width=28, state='readonly')
                combo.set(values[0] if values else '') 
                combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Registrarse", command=self.handle_registration, underline=0).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy, underline=0).pack(side=tk.LEFT, padx=10)
        
    def handle_registration(self):
        try:
            num_control = self.reg_vars['num_control'].get().strip()
            nombre = self.reg_vars['nombre'].get().strip()
            apellido_p = self.reg_vars['apellido_p'].get().strip()
            apellido_m = self.reg_vars['apellido_m'].get().strip()
            discapacidad = self.reg_vars['discapacidad'].get().strip() 
            contrasena = self.reg_vars['contrasena'].get().strip()
            if not all([num_control, nombre, apellido_p, contrasena]):
                messagebox.showwarning("Advertencia", "Todos los campos principales son obligatorios.")
                return
            resultado = registrar_profesor_usuario(num_control, nombre, apellido_p, apellido_m, contrasena, discapacidad)
            if "Error" in resultado:
                messagebox.showerror("Error de Registro", resultado)
            else:
                messagebox.showinfo("Registro Exitoso", resultado)
                self.destroy() 
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado durante la validación: {e}")

class LoginWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Inicio de Sesión")
        self.master = master
        
        # Corrección: Referencia explícita antes de crear widgets
        self.master.login_window = self 
        
        self.geometry("400x180") 
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
        
        # Corrección: Crear widgets PRIMERO, y luego aplicar tema
        self.create_widgets()
        master.apply_theme_settings()
        master.update_font_size(self) 

    def create_widgets(self):
        cred_frame = ttk.Frame(self, padding="10")
        cred_frame.grid(row=0, column=0, columnspan=2)
        ttk.Label(cred_frame, text="Usuario:", underline=0).grid(row=0, column=0, padx=10, pady=5, sticky='w') 
        self.user_var = tk.StringVar(value='ProfesorITT')
        user_entry = ttk.Entry(cred_frame, textvariable=self.user_var, width=25)
        user_entry.grid(row=0, column=1, padx=10, pady=5)
        user_entry.focus()
        ttk.Label(cred_frame, text="Contraseña:", underline=0).grid(row=1, column=0, padx=10, pady=5, sticky='w') 
        self.pass_var = tk.StringVar(value='123')
        self.pass_entry = ttk.Entry(cred_frame, textvariable=self.pass_var, show="*", width=25)
        self.pass_entry.grid(row=1, column=1, padx=10, pady=5)
        self.pass_entry.bind('<Return>', lambda event: self.handle_login())
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Iniciar Sesión", command=self.handle_login, underline=0).pack(side=tk.LEFT, padx=10) 
        ttk.Button(btn_frame, text="Registrar", command=self.show_registration_options, underline=0).pack(side=tk.LEFT, padx=10) 
        ttk.Button(btn_frame, text="Salir", command=self.master.quit, underline=0).pack(side=tk.LEFT, padx=10) 

    def handle_login(self):
        nombre_usuario = self.user_var.get().strip()
        contrasena = self.pass_var.get().strip()
        if not nombre_usuario or not contrasena:
            messagebox.showwarning("Advertencia", "Por favor, ingrese usuario y contraseña.")
            return
        resultado_auth = autenticar_usuario(nombre_usuario, contrasena)
        if resultado_auth:
            user_id, role_id, num_control = resultado_auth
            self.master.show_main_window(user_id, role_id, num_control)
            self.destroy() 
        else:
            messagebox.showerror("Error", "Usuario o contraseña incorrectos.")

    def show_registration_options(self):
        popup = tk.Toplevel(self)
        popup.title("Opciones de Registro")
        popup.geometry("300x120")
        popup.resizable(False, False)
        popup.grab_set() 
        popup.config(bg=self['bg']) 
        frame = ttk.Frame(popup, padding="15")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="¿Qué tipo de usuario desea registrar?").pack(pady=10)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Estudiante", command=lambda: self.open_registration_window(StudentRegistrationWindow, popup)).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Profesor", command=lambda: self.open_registration_window(TeacherRegistrationWindow, popup)).pack(side=tk.LEFT, padx=10)

    def open_registration_window(self, WindowClass, popup):
        popup.destroy()
        WindowClass(self.master)

class CalidadApp(ttk.Frame):
    def __init__(self, master, user_id, role_id, num_control):
        super().__init__(master)
        self.master = master
        self.pack(fill='both', expand=True) 
        self.user_id = user_id
        self.role_id = role_id
        self.num_control = num_control 
        self.magnifier = None 
        
        # INICIALIZACIÓN DE ASISTENTES
        self.voice_assistant = VoiceAssistant(self)
        self.keyboard_window = None
        self.focus_mode_active = False
        
        self.master.title("Sistema de Gestión de Calidad Académica ITT")
        self.master.geometry("1000x650") # Aumentado para la barra
        self.master.resizable(True, True)
        self.master.protocol("WM_DELETE_WINDOW", self.cerrar_sesion)
        
        self.style = ttk.Style(self)
        self.master.apply_theme_settings() 
        self.master.update_font_size(self)
        
        # --- BARRA GLOBAL DE ACCESIBILIDAD ---
        self.crear_barra_accesibilidad()
        
        self.create_widgets()

        # Barra de estado para feedback (Auditiva)
        self.status_label = ttk.Label(self, text="Listo", relief=tk.SUNKEN, anchor='w')
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def crear_barra_accesibilidad(self):
        """Crea una barra de herramientas persistente en la parte superior"""
        toolbar = ttk.Frame(self, padding="5", relief="raised")
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Label(toolbar, text="Accesibilidad Global:", font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=5)

        # Botón de Voz con Toggle
        self.btn_voice = ttk.Button(toolbar, text="🎙️ Activar Voz", command=self.toggle_voice_handler)
        self.btn_voice.pack(side=tk.LEFT, padx=5)
        
        # Etiqueta de estado del micrófono
        self.lbl_mic_status = ttk.Label(toolbar, text="Inactivo", foreground="gray")
        self.lbl_mic_status.pack(side=tk.LEFT, padx=2)

        # Botón de Leer Pantalla
        if TTS_AVAILABLE:
            ttk.Button(toolbar, text="🔊 Leer Pantalla Actual", command=self.read_active_tab).pack(side=tk.LEFT, padx=5)

        # Otros controles globales
        ttk.Button(toolbar, text="🔍 Lupa", command=self.toggle_magnifier).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="⌨️ Teclado", command=self.toggle_virtual_keyboard).pack(side=tk.LEFT, padx=5)
        self.btn_focus = ttk.Button(toolbar, text="🧠 Modo Enfoque", command=self.toggle_focus_mode)
        self.btn_focus.pack(side=tk.LEFT, padx=5)

    def toggle_voice_handler(self):
        """Manejador para el botón de voz que actualiza el texto del botón"""
        active = self.voice_assistant.toggle_listening()
        if active:
            self.btn_voice.config(text="🛑 Detener Voz")
            self.lbl_mic_status.config(text="Iniciando...", foreground="orange")
        else:
            self.btn_voice.config(text="🎙️ Activar Voz")
            self.lbl_mic_status.config(text="Inactivo", foreground="gray")

    def update_status_voice(self, text, bg, fg):
        """Actualiza específicamente el indicador de micrófono en la barra superior"""
        try:
            self.lbl_mic_status.config(text=text, background=bg, foreground=fg)
        except: pass
        
    def apply_colorblind_mode(self):
        self.master.apply_theme_settings()
        mode = self.master.colorblind_mode.get() 
        if mode == "Normal" and not COLOR_INVERTED:
            messagebox.showinfo("Tema Aplicado", "Modo de color normal aplicado.")
        elif mode != "Normal":
            messagebox.showinfo("Tema Aplicado", f"Modo '{mode}' aplicado.")
        else:
            messagebox.showinfo("Tema Aplicado", "Configuración de tema aplicada.")

    def update_font_size(self, font_size, label_widget):
        global CURRENT_FONT_SIZE
        if font_size < 6:
            font_size = 6
        CURRENT_FONT_SIZE = font_size
        self.master.apply_theme_settings() 
        if self.master.main_window:
            self.master.update_font_size(self.master.main_window) 
        label_widget.config(text=f"{CURRENT_FONT_SIZE} puntos")
        
    def apply_font_change(self, event=None):
        global CURRENT_FONT_FAMILY
        seleccion = self.font_family_var.get()
        if seleccion:
            CURRENT_FONT_FAMILY = seleccion
            self.master.apply_theme_settings()
            self.master.update_font_size(self.master.main_window)

    def toggle_dyslexic_mode(self):
        global DYSLEXIC_MODE, CURRENT_FONT_FAMILY, CURRENT_FONT_SIZE
        DYSLEXIC_MODE = not DYSLEXIC_MODE
        if DYSLEXIC_MODE:
            CURRENT_FONT_FAMILY = "Comic Sans MS"
            if CURRENT_FONT_SIZE < 12:
                CURRENT_FONT_SIZE = 12
        else:
            CURRENT_FONT_FAMILY = "Arial"
            if CURRENT_FONT_SIZE == 12: 
                CURRENT_FONT_SIZE = 10
        self.recursive_letter_spacing(self.master, DYSLEXIC_MODE)
        if hasattr(self, 'font_family_var'):
            self.font_family_var.set(CURRENT_FONT_FAMILY)
        if hasattr(self, 'current_font_label'):
            self.current_font_label.config(text=f"{CURRENT_FONT_SIZE} puntos")
        self.master.apply_theme_settings()
        self.master.update_font_size(self.master.main_window)
    
    def recursive_letter_spacing(self, widget, enable):
        try:
            if isinstance(widget, (tk.Label, ttk.Label, tk.Button, ttk.Button, ttk.Checkbutton, ttk.Radiobutton)):
                has_var = False
                try:
                    if widget.cget("textvariable") != "":
                        has_var = True
                except: pass
                if not has_var:
                    current_text = str(widget.cget("text"))
                    if enable:
                        if not hasattr(widget, "origin_text"):
                            widget.origin_text = current_text
                        if current_text and len(current_text) > 1 and not (current_text[1] == " " and current_text[0] != " "):
                            new_text = " ".join(list(widget.origin_text))
                            widget.configure(text=new_text)
                    else:
                        if hasattr(widget, "origin_text"):
                            widget.configure(text=widget.origin_text)
        except Exception: pass
        for child in widget.winfo_children():
            self.recursive_letter_spacing(child, enable)

    def toggle_color_inversion(self):
        global COLOR_INVERTED
        COLOR_INVERTED = not COLOR_INVERTED
        if COLOR_INVERTED and self.master.colorblind_mode.get() != "Normal":
            self.master.colorblind_mode.set("Normal")
            if hasattr(self, 'colorblind_mode_var'):
                 self.colorblind_mode_var.set("Normal")
        self.master.apply_theme_settings()
        
    def toggle_magnifier(self):
        if self.magnifier and self.magnifier.winfo_exists():
            self.magnifier.destroy()
            self.magnifier = None
        else:
            if not PIL_AVAILABLE:
                messagebox.showerror("Error", "La librería Pillow no está instalada.")
                return
            self.magnifier = MagnifierWindow(self.master)

    # --- MÉTODOS NUEVOS DE ACCESIBILIDAD ---
    def toggle_virtual_keyboard(self):
        if self.keyboard_window and self.keyboard_window.winfo_exists():
            self.keyboard_window.destroy()
        else:
            self.keyboard_window = VirtualKeyboard(self.master)

    def toggle_focus_mode(self):
        self.focus_mode_active = not self.focus_mode_active
        if self.focus_mode_active:
            # Ocultar pestañas (Simplificación Visual)
            self.style.layout("TNotebook.Tab", []) 
            self.btn_focus.config(text="Desactivar Modo Enfoque")
            messagebox.showinfo("Modo Enfoque", "Pestañas ocultas para reducir distracciones.")
        else:
            # Restaurar pestañas
            self.style.layout("TNotebook.Tab", [('Notebook.tab', {'sticky': 'nswe', 'children': [('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children': [('Notebook.focus', {'side': 'top', 'sticky': 'nswe', 'children': [('Notebook.label', {'side': 'top', 'sticky': ''})]})]})]})])
            self.btn_focus.config(text="🧠 Activar Modo Enfoque")
            self.master.apply_theme_settings() 

    def update_status(self, text, bg, fg):
        """Actualiza la barra de estado con colores (Auditiva/Visual)"""
        self.status_label.config(text=text, background=bg, foreground=fg)
        self.update()

    def process_voice_command(self, text):
        """Procesa comandos de voz de manera continua"""
        # Elimina acentos para facilitar la coincidencia
        text = text.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
        
        print(f"DEBUG: Procesando comando '{text}'")
        
        if "perfil" in text:
            self.notebook.select(self.tab_perfil)
            self.voice_assistant.speak("Abriendo perfil")
        elif "configuracion" in text or "ajustes" in text:
            self.notebook.select(self.tab_configuracion)
            self.voice_assistant.speak("Abriendo configuración")
        elif "importar" in text and self.role_id == 1:
            self.notebook.select(self.tab_importar)
            self.voice_assistant.speak("Abriendo importar y exportar")
        elif "registro" in text and self.role_id == 1:
            self.notebook.select(self.tab_registro_manual)
            self.voice_assistant.speak("Abriendo registro manual")
        elif "pareto" in text or "riesgo" in text:
            if self.role_id == 1:
                self.notebook.select(self.tab_pareto)
                self.voice_assistant.speak("Abriendo análisis de Pareto")
        elif "salir" in text or "cerrar sesion" in text:
            self.voice_assistant.speak("Cerrando sesión")
            # Necesitamos usar after para ejecutar esto en el hilo principal de tkinter
            self.after(100, self.cerrar_sesion)
        elif "leer" in text or "pantalla" in text:
            self.after(100, self.read_active_tab)
        elif "hola" in text:
            self.voice_assistant.speak("Hola, estoy escuchando tus comandos")
        else:
            print("Comando no coincide con ninguna acción.")

    def read_active_tab(self):
        """Lee el contenido de la pestaña actual"""
        current_tab_id = self.notebook.select()
        if current_tab_id:
            widget = self.notebook.nametowidget(current_tab_id)
            self.voice_assistant.read_screen_content(widget)

    def programar_recordatorio(self, mensaje, delay_ms):
        def mostrar():
            if TTS_AVAILABLE:
                self.voice_assistant.speak(mensaje)
            # Flash visual para auditiva
            original_bg = self.master.cget("bg")
            def flash():
                self.master.config(bg="blue")
                self.master.update()
                time.sleep(0.3)
                self.master.config(bg=original_bg)
            threading.Thread(target=flash).start()
            messagebox.showinfo("Recordatorio Inteligente", mensaje)
        self.after(delay_ms, mostrar)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, expand=True, fill='both')
        
        self.tab_perfil = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_perfil, text='👤 Perfil', underline=0) 
        self.crear_tab_perfil()
        
        if self.role_id == 1: # Profesor
            self.tab_registro_manual = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_registro_manual, text='Registro Manual', underline=0) 
            self.crear_tab_registro()
            self.tab_pareto = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_pareto, text='Análisis de Riesgo (Pareto)', underline=0) 
            self.crear_tab_pareto()
            self.tab_importar = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_importar, text='Importar/Exportar', underline=0) 
            self.crear_tab_importar()
            self.tab_auditoria = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_auditoria, text='Registro de Actividad', underline=0) 
            self.crear_tab_auditoria()
            self.tab_configuracion = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_configuracion, text='⚙️ Configuración', underline=0) 
            self.crear_tab_configuracion()
        elif self.role_id == 2: # Estudiante
            self.tab_datos = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_datos, text='Mis Datos Académicos', underline=0) 
            self.crear_tab_datos_alumno()
            self.tab_configuracion = ttk.Frame(self.notebook)
            self.notebook.add(self.tab_configuracion, text='⚙️ Configuración', underline=0) 
            self.crear_tab_configuracion()

    def cerrar_sesion(self):
        log_actividad(self.user_id, 'LOGOUT', 'Cierre de sesión manual.')
        
        # Detener hilo de voz si está activo
        if hasattr(self, 'voice_assistant'):
             self.voice_assistant.is_listening = False

        # 1. Limpieza de ventanas auxiliares
        if self.magnifier and self.magnifier.winfo_exists():
            self.magnifier.destroy()
        if self.keyboard_window and self.keyboard_window.winfo_exists():
            self.keyboard_window.destroy()
        
        # --- CORRECCIÓN DEL BUG DE RE-LOGIN ---
        # Guardamos la referencia a la raíz antes de destruir el frame
        root = self.master
        
        # IMPORTANTE: Desvinculamos main_window ANTES de crear el nuevo Login.
        root.main_window = None 

        # 2. Destruimos el frame actual
        self.destroy()
        
        # 3. Creamos la nueva ventana de Login y la vinculamos a la raíz
        root.login_window = LoginWindow(root)
        
        # 4. Ocultamos la raíz (por si acaso no lo hace el LoginWindow)
        root.withdraw()

    def crear_tab_configuracion(self):
        global CURRENT_FONT_SIZE, COLOR_INVERTED, CURRENT_FONT_FAMILY
        global COLORBLIND_MODES, AVAILABLE_FONTS
        
        frame = ttk.Frame(self.tab_configuracion, padding="20")
        frame.pack(fill='both', expand=True)
        
        ttk.Label(frame, text="**CONFIGURACIÓN DE ACCESIBILIDAD**", font=('Arial', 14, 'bold')).pack(pady=15)

        font_frame = ttk.LabelFrame(frame, text="Ajuste de Tamaño de Fuente", padding=10)
        font_frame.pack(pady=10, padx=50, fill='x')
        ttk.Label(font_frame, text="Tamaño Actual:").pack(side=tk.LEFT, padx=10)
        self.current_font_label = ttk.Label(font_frame, text=f"{CURRENT_FONT_SIZE} puntos", font=('Arial', 10, 'bold'))
        self.current_font_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(font_frame, text="Disminuir", command=lambda: self.update_font_size(CURRENT_FONT_SIZE - 2, self.current_font_label)).pack(side=tk.LEFT, padx=10)
        ttk.Button(font_frame, text="Aumentar", command=lambda: self.update_font_size(CURRENT_FONT_SIZE + 2, self.current_font_label), underline=0).pack(side=tk.LEFT, padx=10)
        ttk.Button(font_frame, text="Restablecer", command=lambda: self.update_font_size(BASE_FONT_SIZE, self.current_font_label), underline=0).pack(side=tk.LEFT, padx=10)

        family_frame = ttk.LabelFrame(frame, text="Tipo de Letra (Fuente)", padding=10)
        family_frame.pack(pady=10, padx=50, fill='x')
        ttk.Label(family_frame, text="Seleccionar Fuente:").pack(side=tk.LEFT, padx=10)
        self.font_family_var = tk.StringVar(value=CURRENT_FONT_FAMILY)
        font_combo = ttk.Combobox(family_frame, textvariable=self.font_family_var, values=AVAILABLE_FONTS, state='readonly', width=20)
        font_combo.pack(side=tk.LEFT, padx=10)
        font_combo.bind("<<ComboboxSelected>>", self.apply_font_change) 
        ttk.Button(family_frame, text="Aplicar Fuente", command=self.apply_font_change).pack(side=tk.LEFT, padx=10)

        daltonismo_frame = ttk.LabelFrame(frame, text="Ajustes de Daltonismo (Color)", padding=10)
        daltonismo_frame.pack(pady=10, padx=50, fill='x')
        ttk.Label(daltonismo_frame, text="Modo de Color:").pack(side=tk.LEFT, padx=10)
        self.colorblind_mode_var = self.master.colorblind_mode 
        combo_daltonismo = ttk.Combobox(daltonismo_frame, textvariable=self.colorblind_mode_var, values=COLORBLIND_MODES, width=28, state='readonly')
        combo_daltonismo.set(self.master.colorblind_mode.get()) 
        combo_daltonismo.pack(side=tk.LEFT, padx=10, fill='x', expand=True)
        ttk.Button(daltonismo_frame, text="Aplicar Color", command=self.apply_colorblind_mode, underline=7).pack(side=tk.LEFT, padx=10)
        
        color_frame = ttk.LabelFrame(frame, text="Otras Opciones de Accesibilidad", padding=10)
        color_frame.pack(pady=10, padx=50, fill='x')
        
        self.color_inverted_var = tk.BooleanVar(value=COLOR_INVERTED)
        chk_inversion = ttk.Checkbutton(color_frame, text="Invertir Colores (Alto Contraste)", variable=self.color_inverted_var, command=self.toggle_color_inversion, underline=0)
        chk_inversion.pack(side=tk.LEFT, padx=10)
        
        self.dyslexic_var = tk.BooleanVar(value=DYSLEXIC_MODE)
        chk_dyslexia = ttk.Checkbutton(color_frame, text="Modo Dislexia", variable=self.dyslexic_var, command=self.toggle_dyslexic_mode)
        chk_dyslexia.pack(side=tk.LEFT, padx=10)

        ttk.Button(color_frame, text="⏰ Recordatorio (5s)", command=lambda: self.programar_recordatorio("Es momento de tomar un descanso visual y mental.", 5000)).pack(side=tk.LEFT, padx=5)

    def crear_tab_registro(self):
        frame = ttk.Frame(self.tab_registro_manual, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**REGISTRO MANUAL DE ESTUDIANTES**", font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        self.entry_vars = {}
        fields = [
            ("No. Control (*):", "Num_Control", 1, 'entry', None),
            ("Apellido Paterno (*):", "Apellido_Paterno", 2, 'entry', None),
            ("Apellido Materno:", "Apellido_Materno", 3, 'entry', None),
            ("Nombre(s) (*):", "Nombre", 4, 'entry', None),
            ("Carrera (*):", "Carrera", 5, 'combobox', CARRERAS_ITT),
            ("Semestre (*):", "Semestre", 6, 'combobox', SEMESTRES_LIST),
            ("Materia:", "Materia", 7, 'entry', None),
            ("Asistencia %:", "Asistencia_Porcentaje", 8, 'entry', None),
            ("Discapacidad:", "Discapacidad", 9, 'combobox', DISCAPACIDADES_LIST) 
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

        calif_frame = ttk.LabelFrame(frame, text="Calificaciones")
        calif_frame.grid(row=1, column=2, rowspan=5, padx=10, pady=5, sticky='n')
        for i in range(1, 6):
            ttk.Label(calif_frame, text=f"Unidad {i}:").grid(row=i, column=0, padx=5, pady=2, sticky='w')
            var = tk.StringVar()
            self.entry_vars[f"Calificacion_Unidad_{i}"] = var
            ttk.Entry(calif_frame, textvariable=var, width=10).grid(row=i, column=1, padx=5, pady=2, sticky='w')

        fact_frame = ttk.LabelFrame(frame, text="Factores de Riesgo")
        fact_frame.grid(row=6, column=2, rowspan=4, padx=10, pady=5, sticky='n')
        self.factor_vars = {}
        factores = ['Factor_Academico', 'Factor_Psicosocial', 'Factor_Economico', 
                    'Factor_Institucional', 'Factor_Tecnologico', 'Factor_Contextual']
        for i, factor in enumerate(factores):
            var = tk.StringVar(value='0')
            self.factor_vars[factor] = var
            chk = ttk.Checkbutton(fact_frame, text=factor.replace('_', ' ').replace('Factor ', ''), variable=var, onvalue='1', offvalue='0')
            chk.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=10, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="Guardar Estudiante", command=self.guardar_estudiante, underline=0).pack(side=tk.LEFT, padx=10) 
        ttk.Button(btn_frame, text="Limpiar Campos", command=self.limpiar_registro, underline=0).pack(side=tk.LEFT, padx=10) 

    def guardar_estudiante(self):
        try:
            datos_principales = [
                self.entry_vars['Num_Control'].get().strip(),
                self.entry_vars['Apellido_Paterno'].get().strip(),
                self.entry_vars['Apellido_Materno'].get().strip(),
                self.entry_vars['Nombre'].get().strip(),
                self.entry_vars['Carrera'].get().strip(),
                self.entry_vars['Semestre'].get().strip(),
                self.entry_vars['Materia'].get().strip()
            ]
            discapacidad = self.entry_vars['Discapacidad'].get().strip() 

            if not all(datos_principales[:5]):
                messagebox.showwarning("Advertencia", "Los campos principales son obligatorios.")
                return
            calificaciones = []
            for i in range(1, 6):
                val = self.entry_vars[f"Calificacion_Unidad_{i}"].get().strip()
                try:
                    calif = float(val) if val else 0.0
                    if not 0.0 <= calif <= 100.0: raise ValueError("Fuera de rango")
                    calificaciones.append(calif)
                except ValueError:
                    messagebox.showerror("Error", f"La Calificación de la Unidad {i} debe ser un número entre 0 y 100.")
                    return
            
            asistencia_str = self.entry_vars['Asistencia_Porcentaje'].get().strip()
            try:
                asistencia_final = float(asistencia_str) if asistencia_str else 0.0
                if not 0.0 <= asistencia_final <= 100.0: raise ValueError("Fuera de rango")
            except ValueError:
                messagebox.showerror("Error", "El Porcentaje de Asistencia debe ser un número entre 0 y 100.")
                return
            
            calificaciones_finales = calificaciones 
            factores = [self.factor_vars[name].get() for name in self.factor_vars.keys()]
            datos_completos = datos_principales + calificaciones_finales + [asistencia_final] + factores + [discapacidad]
            insertar_registro_manual(datos_completos, self.user_id)
            self.limpiar_registro()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado al guardar: {e}")
            
    def limpiar_registro(self):
        for var in self.entry_vars.values():
            var.set('')
        for var in self.factor_vars.values():
            var.set('0')
        if 'Carrera' in self.entry_vars and CARRERAS_ITT:
            self.entry_vars['Carrera'].set(CARRERAS_ITT[0])
        if 'Semestre' in self.entry_vars and SEMESTRES_LIST:
            self.entry_vars['Semestre'].set(SEMESTRES_LIST[0])
        if 'Discapacidad' in self.entry_vars and DISCAPACIDADES_LIST:
            self.entry_vars['Discapacidad'].set(DISCAPACIDADES_LIST[0])

    def crear_tab_perfil(self):
        frame = ttk.Frame(self.tab_perfil, padding="20")
        frame.pack(fill='both', expand=True)
        self.perfil_vars = {}
        datos_estudiante = obtener_datos_estudiante(self.num_control) if self.role_id == 2 else None
        discapacidad_usuario = obtener_discapacidad_usuario(self.num_control, self.role_id)

        ttk.Label(frame, text=f"**PERFIL DE USUARIO**", font=('Arial', 14, 'bold')).pack(pady=10)
        data_frame = ttk.Frame(frame)
        data_frame.pack(side=tk.LEFT, padx=50, anchor='n')

        data_map = {
            "No. Control:": self.num_control if self.role_id == 2 else self.master.login_window.user_var.get(),
            "Nombre(s):": datos_estudiante[0] if datos_estudiante and len(datos_estudiante) > 0 else "",
            "Apellido Paterno:": datos_estudiante[1] if datos_estudiante and len(datos_estudiante) > 1 else "",
            "Apellido Materno:": datos_estudiante[2] if datos_estudiante and len(datos_estudiante) > 2 else "",
            "Carrera:": datos_estudiante[3] if datos_estudiante and len(datos_estudiante) > 3 else "N/A (Profesor)",
            "Semestre:": datos_estudiante[4] if datos_estudiante and len(datos_estudiante) > 4 else "N/A (Profesor)",
            "Discapacidad:": discapacidad_usuario 
        }
        
        row_num = 0
        for label_text, value in data_map.items():
            ttk.Label(data_frame, text=label_text, width=20, anchor='w').grid(row=row_num, column=0, sticky='w', pady=2)
            ttk.Label(data_frame, text=str(value), font=('Arial', 10, 'bold')).grid(row=row_num, column=1, sticky='w', pady=2)
            row_num += 1

        if self.role_id == 2 and datos_estudiante:
            edit_frame = ttk.LabelFrame(frame, text="Editar Datos Personales y Discapacidad")
            edit_frame.pack(side=tk.RIGHT, padx=50, anchor='n', fill='x', expand=True)
            fields = [
                ("Nombre(s):", "Nombre", datos_estudiante[0]), 
                ("Apellido Paterno:", "Apellido_Paterno", datos_estudiante[1]),
                ("Apellido Materno:", "Apellido_Materno", datos_estudiante[2]),
                ("Discapacidad:", "Discapacidad", discapacidad_usuario) 
            ]
            for i, (label_text, var_name, initial_value) in enumerate(fields):
                ttk.Label(edit_frame, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky='w')
                var = tk.StringVar(value=initial_value)
                self.perfil_vars[var_name] = var
                if var_name != 'Discapacidad':
                    ttk.Entry(edit_frame, textvariable=var, width=30).grid(row=i, column=1, padx=5, pady=5, sticky='w')
                else:
                    combo = ttk.Combobox(edit_frame, textvariable=var, values=DISCAPACIDADES_LIST, width=28, state='readonly')
                    combo.set(initial_value)
                    combo.grid(row=i, column=1, padx=5, pady=5, sticky='w')
            ttk.Button(edit_frame, text="Actualizar Datos", command=self.actualizar_perfil, underline=0).grid(row=len(fields), column=0, columnspan=2, pady=15)
        
        btn_frame_logout = ttk.Frame(frame)
        btn_frame_logout.pack(side=tk.BOTTOM, pady=30)
        ttk.Button(btn_frame_logout, text="Cerrar Sesión", command=self.cerrar_sesion, width=20).pack()
            
    def actualizar_perfil(self):
        nombre = self.perfil_vars['Nombre'].get().strip()
        apellido_p = self.perfil_vars['Apellido_Paterno'].get().strip()
        apellido_m = self.perfil_vars['Apellido_Materno'].get().strip()
        discapacidad = self.perfil_vars['Discapacidad'].get().strip() 

        if not nombre or not apellido_p:
            messagebox.showwarning("Advertencia", "El nombre y el apellido paterno son obligatorios.")
            return
        nuevos_datos = [apellido_p, apellido_m, nombre, discapacidad] 
        actualizar_datos_estudiante(self.num_control, nuevos_datos, self.user_id)

    def crear_tab_datos_alumno(self):
        frame = ttk.Frame(self.tab_datos, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**MIS DATOS ACADÉMICOS**", font=('Arial', 14, 'bold')).pack(pady=10)
        datos = obtener_datos_estudiante(self.num_control)
        if datos:
            labels = ["Nombre:", "Apellido Paterno:", "Apellido Materno:", "Carrera:", "Semestre:", "Materia:", "Calificación U1:", "Calificación U2:", "Calificación U3:", "Calificación U4:", "Calificación U5:", "Discapacidad:"]
            info_frame = ttk.Frame(frame)
            info_frame.pack(pady=10, padx=50, anchor='w')
            for i, label_text in enumerate(labels):
                value = datos[i] if i < len(datos) else "N/D"
                ttk.Label(info_frame, text=label_text, width=20, anchor='w').grid(row=i, column=0, sticky='w')
                ttk.Label(info_frame, text=str(value), font=('Arial', 10, 'bold')).grid(row=i, column=1, sticky='w')
        else:
            ttk.Label(frame, text="No se encontraron datos académicos para su número de control.").pack(pady=20)

    def crear_tab_pareto(self):
        frame = ttk.Frame(self.tab_pareto, padding="10")
        frame.pack(fill='both', expand=True)
        canvas = tk.Canvas(frame)
        canvas.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        vsb.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vsb.set)
        
        self.scrollable_frame = ttk.Frame(canvas, width=950) 
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=950)
        
        ttk.Label(self.scrollable_frame, text="**FILTROS DE ANÁLISIS**", font=('Arial', 14, 'bold')).pack(pady=10)
        control_frame = ttk.Frame(self.scrollable_frame, padding=10)
        control_frame.pack(padx=10, pady=5, fill='x')
        ttk.Label(control_frame, text="Carrera:").grid(row=0, column=0, padx=5, pady=5)
        self.pareto_carrera = tk.StringVar()
        ttk.Combobox(control_frame, textvariable=self.pareto_carrera, values=[''] + CARRERAS_ITT, width=20, state='readonly').grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(control_frame, text="Semestre:").grid(row=0, column=2, padx=5, pady=5)
        self.pareto_semestre = tk.StringVar()
        ttk.Combobox(control_frame, textvariable=self.pareto_semestre, values=[''] + SEMESTRES_LIST, width=10, state='readonly').grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(control_frame, text="Materia:").grid(row=1, column=0, padx=5, pady=5)
        self.pareto_materia = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.pareto_materia, width=20).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(control_frame, text="Generar Gráfica", command=self.generar_grafica_pareto, underline=0).grid(row=1, column=2, padx=15, pady=5) 
        ttk.Button(control_frame, text="Limpiar Filtros", command=self.limpiar_filtros_pareto, underline=0).grid(row=1, column=3, padx=15, pady=5) 

        self.pareto_canvas_frame = ttk.Frame(self.scrollable_frame, height=300)
        self.pareto_canvas_frame.pack(fill='x', expand=False, padx=10, pady=5)
        self.pareto_fig_canvas = None

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
        vsb_alumnos = ttk.Scrollbar(alumnos_frame, orient="vertical", command=self.tree_alumnos.yview)
        self.tree_alumnos.configure(yscrollcommand=vsb_alumnos.set)
        vsb_alumnos.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_alumnos.pack(fill='both', expand=True)

    def limpiar_filtros_pareto(self):
        self.pareto_carrera.set('')
        self.pareto_semestre.set('')
        self.pareto_materia.set('')
        for widget in self.pareto_canvas_frame.winfo_children():
            widget.destroy()
        self.pareto_fig_canvas = None 
        for i in self.tree_alumnos.get_children():
            self.tree_alumnos.delete(i)
        
    def generar_grafica_pareto(self):
        global COLOR_INVERTED, CUSTOM_COLORS
        mode = self.master.colorblind_mode.get()
        if mode != "Normal":
            scheme = CUSTOM_COLORS.get(mode)
        elif COLOR_INVERTED:
            scheme = CUSTOM_COLORS.get("Inversion")
        else:
            scheme = CUSTOM_COLORS.get("Normal")
        bar_color = scheme['plot_bar']
        line_color = scheme['plot_line']
        carrera = self.pareto_carrera.get()
        semestre = self.pareto_semestre.get()
        materia = self.pareto_materia.get()
        
        fig, df_estudiantes, error_msg = generar_pareto_factores(
            self.user_id, carrera, semestre, materia, 
            bar_color=bar_color, line_color=line_color 
        )
        if self.pareto_fig_canvas:
            self.pareto_fig_canvas.get_tk_widget().destroy()
            self.pareto_fig_canvas = None
            
        for i in self.tree_alumnos.get_children():
            self.tree_alumnos.delete(i)
        if error_msg:
            messagebox.showwarning("Advertencia", error_msg)
            return
        if fig:
            self.pareto_fig_canvas = FigureCanvasTkAgg(fig, master=self.pareto_canvas_frame)
            self.pareto_fig_canvas.draw()
            self.pareto_fig_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        if not df_estudiantes.empty:
            for _, row in df_estudiantes.iterrows():
                self.tree_alumnos.insert("", tk.END, values=row.tolist())

    def crear_tab_importar(self):
        if self.role_id == 1:
            frame = ttk.Frame(self.tab_importar, padding="20")
            frame.pack(fill='both', expand=True)
            ttk.Label(frame, text="**IMPORTAR Y EXPORTAR DATOS**", font=('Arial', 14, 'bold')).pack(pady=10)
            import_frame = ttk.LabelFrame(frame, text="Importar Datos de Estudiantes (CSV/Excel)")
            import_frame.pack(pady=20, padx=20, fill='x')
            ttk.Label(import_frame, text="Ruta del Archivo:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            self.import_path_var = tk.StringVar()
            ttk.Entry(import_frame, textvariable=self.import_path_var, width=50, state='readonly').grid(row=0, column=1, padx=5, pady=5)
            ttk.Button(import_frame, text="Seleccionar Archivo", command=self.seleccionar_archivo_import, underline=0).grid(row=0, column=2, padx=5, pady=5) 
            ttk.Button(import_frame, text="Ejecutar Importación", command=self.ejecutar_importacion_handler, underline=0).grid(row=1, column=1, columnspan=2, pady=10) 
            export_frame = ttk.LabelFrame(frame, text="Exportar Datos de Estudiantes")
            export_frame.pack(pady=20, padx=20, fill='x')
            export_btns_frame = ttk.Frame(export_frame)
            export_btns_frame.pack(pady=10)
            ttk.Button(export_btns_frame, text="Exportar a Excel (.xlsx)", command=lambda: exportar_datos_sql('excel', self.user_id), underline=13).pack(side=tk.LEFT, padx=10) 
            ttk.Button(export_btns_frame, text="Exportar a CSV (.csv)", command=lambda: exportar_datos_sql('csv', self.user_id), underline=13).pack(side=tk.LEFT, padx=10) 
        else:
            ttk.Label(self.tab_importar, text="Esta pestaña solo está disponible para el rol de Profesor.").pack(pady=20)

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

    def crear_tab_auditoria(self):
        frame = ttk.Frame(self.tab_auditoria, padding="10")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="**REGISTRO DE INICIOS Y CIERRES DE SESIÓN**", font=('Arial', 14, 'bold')).pack(pady=10)
        columns = ("Matricula", "Accion", "Dia", "Hora")
        self.tree_auditoria = ttk.Treeview(frame, columns=columns, show="headings")
        self.tree_auditoria.heading("Matricula", text="Matrícula")
        self.tree_auditoria.heading("Accion", text="Acción")
        self.tree_auditoria.heading("Dia", text="Día")
        self.tree_auditoria.heading("Hora", text="Hora")
        self.tree_auditoria.column("Matricula", width=150, anchor='center')
        self.tree_auditoria.column("Accion", width=150, anchor='center')
        self.tree_auditoria.column("Dia", width=100, anchor='center')
        self.tree_auditoria.column("Hora", width=100, anchor='center')
        vsb_auditoria = ttk.Scrollbar(frame, orient="vertical", command=self.tree_auditoria.yview)
        self.tree_auditoria.configure(yscrollcommand=vsb_auditoria.set)
        vsb_auditoria.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_auditoria.pack(fill='both', expand=True, padx=10, pady=10)
        self.cargar_registro_auditoria()
        
    def cargar_registro_auditoria(self):
        for i in self.tree_auditoria.get_children():
            self.tree_auditoria.delete(i)
        df_auditoria = obtener_registro_auditoria()
        if not df_auditoria.empty:
            for _, row in df_auditoria.iterrows():
                self.tree_auditoria.insert("", tk.END, values=row.tolist())
        else:
            self.tree_auditoria.insert("", tk.END, values=("N/D", "Sin Registros de Auditoría", "N/D", "N/D"))

class App(tk.Tk):
    BASE_WIDTH = 1400 
    BASE_HEIGHT = 800
    
    def __init__(self):
        super().__init__()
        self.withdraw() 
        self.title("Sistema de Calidad Académica")
        self.style = ttk.Style(self)
        self.style.theme_use('clam') 
        self.colorblind_mode = tk.StringVar(self, value=COLORBLIND_MODES[0]) 
        self.main_window = None
        self.login_window = None
        self.login_window = LoginWindow(self)
        
    def update_font_size(self, container_widget):
        global CURRENT_FONT_SIZE, BASE_FONT_SIZE, CURRENT_FONT_FAMILY
        font_size = max(BASE_FONT_SIZE, CURRENT_FONT_SIZE) 
        font_family = CURRENT_FONT_FAMILY 

        self.tk_normal_font = tkinter.font.Font(family=font_family, size=font_size)
        self.tk_bold_font = tkinter.font.Font(family=font_family, size=font_size, weight="bold")
        self.style.configure('.', font=self.tk_normal_font)
        pad_val = 6 
        self.style.configure('TButton', font=self.tk_bold_font, padding=pad_val)
        self.style.configure('TLabel', font=self.tk_normal_font, padding=(2, 2))
        self.style.configure('TNotebook.Tab', padding=[pad_val+4, 5])
        self.style.configure('Treeview.Heading', font=self.tk_bold_font)
        
        for widget in container_widget.winfo_children():
            try:
                if isinstance(widget, (tk.Label, tk.Button, tk.Entry, tk.Checkbutton, tk.Radiobutton)):
                    widget.configure(font=self.tk_normal_font)
            except tk.TclError:
                pass 
            if widget.winfo_children():
                self.update_font_size(widget)
                
    def apply_theme_settings(self):
        global CURRENT_FONT_SIZE, BASE_FONT_SIZE, CURRENT_FONT_FAMILY
        global COLOR_INVERTED
        global CUSTOM_COLORS

        font_size = max(BASE_FONT_SIZE, CURRENT_FONT_SIZE)
        font_family = CURRENT_FONT_FAMILY 
        pad_val = 6

        self.tk_normal_font = tkinter.font.Font(family=font_family, size=font_size)
        self.tk_bold_font = tkinter.font.Font(family=font_family, size=font_size, weight="bold")
        self.style.configure('.', font=self.tk_normal_font)
        self.style.configure('TButton', font=self.tk_bold_font, padding=pad_val)
        self.style.configure('TLabel', font=self.tk_normal_font)
        
        mode = self.colorblind_mode.get() 
        if mode != "Normal":
            scheme = CUSTOM_COLORS.get(mode)
            if self.main_window and hasattr(self.main_window, 'color_inverted_var'):
                 COLOR_INVERTED = False 
                 self.main_window.color_inverted_var.set(False)
        elif COLOR_INVERTED:
            scheme = CUSTOM_COLORS.get("Inversion")
        else:
            scheme = CUSTOM_COLORS.get("Normal")
            
        if not scheme: scheme = CUSTOM_COLORS.get("Normal") 
            
        bg = scheme['bg_window']
        fg = scheme['fg_text']
        btn_bg = scheme['bg_button']
        btn_fg = scheme['fg_button']
        hl_bg = scheme['bg_highlight']
        hl_fg = scheme['fg_highlight']
        entry_bg = scheme.get('bg_entry', 'white')
        entry_fg = scheme.get('fg_entry', 'black')
        insert_color = scheme.get('insert', 'black')

        self.style.configure('TFrame', background=bg)
        self.style.configure('TLabelframe', background=bg, foreground=fg)
        self.style.configure('TLabelframe.Label', background=bg, foreground=fg)
        self.style.configure('TLabel', background=bg, foreground=fg)
        self.style.configure('TNotebook', background=bg)
        self.style.configure('TNotebook.Tab', background=bg, foreground=fg)
        self.style.configure('TButton', background=btn_bg, foreground=btn_fg)
        
        self.style.configure('TEntry', fieldbackground=entry_bg, foreground=entry_fg, insertcolor=insert_color)
        self.style.configure('TCombobox', fieldbackground=entry_bg, background=btn_bg, foreground=entry_fg, arrowcolor=fg)
        self.style.map('TCombobox', fieldbackground=[('readonly', entry_bg)], selectbackground=[('readonly', hl_bg)])
        self.style.map('TEntry', fieldbackground=[('disabled', bg)])
        self.style.configure('Treeview', background=entry_bg, foreground=entry_fg, fieldbackground=entry_bg)
        self.style.configure('Treeview.Heading', background=btn_bg, foreground=btn_fg)
        self.style.map('Treeview', background=[('selected', hl_bg)], foreground=[('selected', hl_fg)])
        self.style.map('TButton', background=[('active', hl_bg)], foreground=[('active', hl_fg)])
        self.style.map('TCheckbutton', background=[('active', bg), ('!disabled', bg)], foreground=[('!disabled', fg)])
        self.style.configure('TCheckbutton', background=bg, foreground=fg)

        self.configure(bg=bg)
        if self.main_window:
            self.main_window.configure(bg=bg) 
            self.recursive_widget_update(self.main_window, bg, fg, entry_bg, entry_fg, insert_color)
            
        if self.login_window:
             self.login_window.configure(bg=bg)
             self.recursive_widget_update(self.login_window, bg, fg, entry_bg, entry_fg, insert_color)

    def recursive_widget_update(self, widget, bg, fg, entry_bg, entry_fg, insert_color):
        try:
            if isinstance(widget, (tk.Entry, tk.Text)):
                widget.configure(bg=entry_bg, fg=entry_fg, insertbackground=insert_color)
            elif isinstance(widget, (tk.Listbox, tk.Canvas)):
                widget.configure(bg=bg)
            elif isinstance(widget, tk.Label): 
                 widget.configure(bg=bg, fg=fg)
        except Exception:
            pass
        for child in widget.winfo_children():
            self.recursive_widget_update(child, bg, fg, entry_bg, entry_fg, insert_color)

    def show_main_window(self, user_id, role_id, num_control):
        if self.main_window:
            self.main_window.destroy() 
        self.main_window = CalidadApp(self, user_id, role_id, num_control)
        self.geometry(f"{self.BASE_WIDTH}x{self.BASE_HEIGHT}") 
        self.deiconify() 
        self.main_window.focus_force()

if __name__ == "__main__":
    app = App()
    app.mainloop()