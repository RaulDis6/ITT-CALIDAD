
    CREATE DATABASE ITT_Calidad;



USE ITT_Calidad;
GO

-- ***********************************************
-- PASO 2: CREACIÓN DE TABLAS
-- ***********************************************

-- 1. Tabla Roles: Define los tipos de usuarios (Profesor, Estudiante).
-- NOTA: Esta tabla NO tiene un 'Identity' para que los IDs sean fijos (1=Profesor, 2=Estudiante).
IF OBJECT_ID('Roles', 'U') IS NOT NULL DROP TABLE Roles;
CREATE TABLE Roles (
    ID_Rol INT PRIMARY KEY,
    Nombre_Rol VARCHAR(50) NOT NULL
);
GO

-- 2. Tabla Estudiantes: Almacena la información académica y los factores de riesgo de los alumnos.
IF OBJECT_ID('Estudiantes', 'U') IS NOT NULL DROP TABLE Estudiantes;
CREATE TABLE Estudiantes (
    -- Clave Primaria
    Num_Control VARCHAR(20) PRIMARY KEY, 
    
    -- Datos Personales y Académicos
    Apellido_Paterno VARCHAR(50) NOT NULL,
    Apellido_Materno VARCHAR(50) NULL,
    Nombre VARCHAR(50) NOT NULL,
    Carrera VARCHAR(50) NOT NULL,
    Semestre INT NOT NULL,
    Materia VARCHAR(100) NOT NULL DEFAULT 'Sin Materia',
    
    -- Calificaciones y Asistencia (Valores Flotantes)
    Calificacion_Unidad_1 DECIMAL(5, 2) DEFAULT 0.0,
    Calificacion_Unidad_2 DECIMAL(5, 2) DEFAULT 0.0,
    Calificacion_Unidad_3 DECIMAL(5, 2) DEFAULT 0.0,
    Calificacion_Unidad_4 DECIMAL(5, 2) DEFAULT 0.0,
    Calificacion_Unidad_5 DECIMAL(5, 2) DEFAULT 0.0,
    Asistencia_Porcentaje DECIMAL(5, 2) DEFAULT 0.0,

    -- Factores de Riesgo (Valores Binarios 0/1)
    -- Se usan INT porque en Python se mapean con Checkbox (True/False o 1/0)
    Factor_Academico INT DEFAULT 0, 
    Factor_Psicosocial INT DEFAULT 0,
    Factor_Economico INT DEFAULT 0,
    Factor_Institucional INT DEFAULT 0,
    Factor_Tecnologico INT DEFAULT 0,
    Factor_Contextual INT DEFAULT 0
);
GO

-- 3. Tabla Usuarios: Almacena credenciales y asocia a un Rol y, opcionalmente, a un Estudiante.
IF OBJECT_ID('Usuarios', 'U') IS NOT NULL DROP TABLE Usuarios;
CREATE TABLE Usuarios (
    ID_Usuario INT IDENTITY(1,1) PRIMARY KEY,
    Nombre_Usuario VARCHAR(50) UNIQUE NOT NULL, -- Matrícula o Usuario de Profesor
    Contrasena_Hash VARCHAR(255) NOT NULL,      -- Contraseña (el código asume que no está hasheada)
    
    -- Claves Foráneas
    ID_Rol_FK INT NOT NULL,                     -- 1 (Profesor) o 2 (Estudiante)
    Num_Control_FK VARCHAR(20) NULL,            -- Solo para Estudiantes (Matrícula)

    -- Definición de Foráneas
    CONSTRAINT FK_Usuario_Rol FOREIGN KEY (ID_Rol_FK) REFERENCES Roles(ID_Rol),
    -- La FK a Estudiantes permite NULL porque los profesores no tienen un registro en la tabla Estudiantes
    CONSTRAINT FK_Usuario_Estudiante FOREIGN KEY (Num_Control_FK) REFERENCES Estudiantes(Num_Control)
);
GO

-- 4. Tabla RegistroActividad: Para auditoría de inicios y cierres de sesión.
IF OBJECT_ID('RegistroActividad', 'U') IS NOT NULL DROP TABLE RegistroActividad;
CREATE TABLE RegistroActividad (
    ID_Registro INT IDENTITY(1,1) PRIMARY KEY,
    ID_Usuario_FK INT NULL,                      -- Puede ser NULL si es un login fallido o registro nuevo
    Tipo_Accion VARCHAR(50) NOT NULL,           -- 'LOGIN_EXITOSO', 'LOGOUT', 'REGISTRO_NUEVO_ALUMNO', etc.
    Detalle VARCHAR(4000) NULL,
    Fecha_Hora DATETIME DEFAULT GETDATE(),

    -- Definición de Foránea (Puede ser NULL)
    CONSTRAINT FK_Actividad_Usuario FOREIGN KEY (ID_Usuario_FK) REFERENCES Usuarios(ID_Usuario)
);
GO

-- INSERCIÓN DE ROLES
IF NOT EXISTS (SELECT 1 FROM Roles WHERE ID_Rol = 1) BEGIN
    INSERT INTO Roles (ID_Rol, Nombre_Rol) VALUES (1, 'Profesor');
END
IF NOT EXISTS (SELECT 1 FROM Roles WHERE ID_Rol = 2) BEGIN
    INSERT INTO Roles (ID_Rol, Nombre_Rol) VALUES (2, 'Estudiante');
END
GO

-- INSERCIÓN DE ESTUDIANTE DE PRUEBA
IF NOT EXISTS (SELECT 1 FROM Estudiantes WHERE Num_Control = 'A2024001')
BEGIN
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual 
    )
    VALUES (
        'A2024001', 'García', 'López', 'Andrea', 'ISC (Sistemas)', 1, 'Programación Orientada a Objetos',
        90.5, 95.0, 88.0, 0.0, 0.0, 100.0,
        0, 0, 0, 0, 0, 0 
    );
END
GO

-- INSERCIÓN DE USUARIOS DE PRUEBA
-- 1. Profesor
IF NOT EXISTS (SELECT 1 FROM Usuarios WHERE Nombre_Usuario = 'ProfesorITT')
BEGIN
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES ('ProfesorITT', '123', 1, NULL); 
END

-- 2. Estudiante
IF NOT EXISTS (SELECT 1 FROM Usuarios WHERE Nombre_Usuario = 'A2024001')
BEGIN
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES ('A2024001', 'pass', 2, 'A2024001');
END
GO


-- ***********************************************
-- PASO 5: VERIFICACIÓN FINAL
-- ***********************************************

SELECT 'Roles' AS Tabla, * FROM Roles;
SELECT 'Estudiantes' AS Tabla, * FROM Estudiantes WHERE Num_Control IN ('A2024001');
SELECT 'Usuarios' AS Tabla, * FROM Usuarios WHERE Nombre_Usuario IN ('ProfesorITT', 'A2024001');