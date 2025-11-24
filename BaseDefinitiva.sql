
    CREATE DATABASE ITT_Calidad;



USE ITT_Calidad;
GO



-- 1. Tabla Roles: Define los tipos de usuarios (Profesor, Estudiante).

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
 
    Factor_Academico INT DEFAULT 0, 
    Factor_Psicosocial INT DEFAULT 0,
    Factor_Economico INT DEFAULT 0,
    Factor_Institucional INT DEFAULT 0,
    Factor_Tecnologico INT DEFAULT 0,
    Factor_Contextual INT DEFAULT 0
);
GO

-- 3. Tabla Usuarios
IF OBJECT_ID('Usuarios', 'U') IS NOT NULL DROP TABLE Usuarios;
CREATE TABLE Usuarios (
    ID_Usuario INT IDENTITY(1,1) PRIMARY KEY,
    Nombre_Usuario VARCHAR(50) UNIQUE NOT NULL, 
    Contrasena_Hash VARCHAR(255) NOT NULL,      
    
    -- Claves Foráneas
    ID_Rol_FK INT NOT NULL,                     
    Num_Control_FK VARCHAR(20) NULL,            

    -- Definición de Foráneas
    CONSTRAINT FK_Usuario_Rol FOREIGN KEY (ID_Rol_FK) REFERENCES Roles(ID_Rol),
    
    CONSTRAINT FK_Usuario_Estudiante FOREIGN KEY (Num_Control_FK) REFERENCES Estudiantes(Num_Control)
);
GO

-- 4. Tabla RegistroActividad
IF OBJECT_ID('RegistroActividad', 'U') IS NOT NULL DROP TABLE RegistroActividad;
CREATE TABLE RegistroActividad (
    ID_Registro INT IDENTITY(1,1) PRIMARY KEY,
    ID_Usuario_FK INT NULL,                     
    Tipo_Accion VARCHAR(50) NOT NULL,          
    Detalle VARCHAR(4000) NULL,
    Fecha_Hora DATETIME DEFAULT GETDATE(),

    -- Definición de Foránea 
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

IF NOT EXISTS (SELECT 1 FROM Estudiantes WHERE Num_Control = '20211813')
BEGIN
    INSERT INTO Estudiantes (
        Num_Control, Apellido_Paterno, Apellido_Materno, Nombre, Carrera, Semestre, Materia,
        Calificacion_Unidad_1, Calificacion_Unidad_2, Calificacion_Unidad_3, Calificacion_Unidad_4, Calificacion_Unidad_5, Asistencia_Porcentaje,
        Factor_Academico, Factor_Psicosocial, Factor_Economico, Factor_Institucional, Factor_Tecnologico, Factor_Contextual 
    )
    VALUES (
        '20211813', 'Molina', 'Mendez', 'Raul', 'ISC (Sistemas)', 11, 'Temas avanzados',
        90.5, 95.0, 88.0, 90.0, 80.0, 100.0,
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

IF NOT EXISTS (SELECT 1 FROM Usuarios WHERE Nombre_Usuario = '2025')
BEGIN
    INSERT INTO Usuarios (Nombre_Usuario, Apellido_Paterno, Apellido_Materno, Nombre, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES ('2025', 'Maribel', 'Guerrero', 'Luis', '123', 2, 2025); 
END
Go


-- 2. Estudiante
IF NOT EXISTS (SELECT 1 FROM Usuarios WHERE Nombre_Usuario = 'A2024001')
BEGIN
    INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES ('A2024001', 'pass', 2, 'A2024001');
	
END
GO

IF NOT EXISTS (SELECT 1 FROM Usuarios WHERE Nombre_Usuario = '20211813')
BEGIN
INSERT INTO Usuarios (Nombre_Usuario, Contrasena_Hash, ID_Rol_FK, Num_Control_FK) 
    VALUES ('20211813', '123', 2, '20211813');
	END
GO


SELECT 'Roles' AS Tabla, * FROM Roles;
SELECT 'Estudiantes' AS Tabla, * FROM Estudiantes WHERE Num_Control IN ('A2024001');
SELECT 'Estudiantes' AS Tabla, * FROM Estudiantes WHERE Num_Control IN ('20211813');
SELECT 'Usuarios' AS Tabla, * FROM Usuarios WHERE Nombre_Usuario IN ('ProfesorITT', 'A2024001');
SELECT 'Usuarios' AS Tabla, * FROM Usuarios WHERE Nombre_Usuario IN ('2025');

USE ITT_Calidad;
GO

DELETE FROM Estudiantes;
PRINT 'Tabla Estudiantes (Datos Académicos) vaciada.';
