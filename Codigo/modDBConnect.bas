Attribute VB_Name = "modDBConnect"
'---------------------------------------------------------------------------------------
' Module    : modDBConnect
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public CN                       As ADODB.Connection
Public rs                       As New ADODB.Recordset

' Requiere una referencia a Microsoft ActiveX Data Objects 2.x Library
Private Const CONNECTION_STRING As String = "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;UID=root;PWD=;PORT=3306;OPTION=3;ConnectionLifetime=0;ConnectionTimeout=0"
Private Const CONNECTION_STRING_DB As String = "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;DATABASE=eiv;UID=root;PWD=;PORT=3306"

'---------------------------------------------------------------------------------------
' Procedure : DBExists
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DBExists(ByVal dbName As String) As Boolean

Dim CN                          As ADODB.Connection
Dim rs                          As ADODB.Recordset
Dim query                       As String

    On Error GoTo DBExists_Error

    Set CN = New ADODB.Connection
    CN.ConnectionString = CONNECTION_STRING
    CN.Open

    query = "SHOW DATABASES LIKE '" & dbName & "';"
    Set rs = CN.Execute(query)

    DBExists = Not rs.EOF

    Call SaveSetting(App.Path, "EIV_SOFTWARE", "IsDBAlreadyExists", IIf(DBExists, "1", "0"))

    rs.Close
    CN.Close
    Set rs = Nothing
    Set CN = Nothing

    On Error GoTo 0
    Exit Function

DBExists_Error:
    DBExists = False
    Call Logs.LogError("Error verificando la base de datos: " & Err.Number & " (" & Err.Description & ") en procedimiento DBExists de Módulo modDBConnect línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : DBCreate
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DBCreate(ByVal dbName As String) As Boolean

Dim CN                          As ADODB.Connection
Dim query                       As String

    On Error GoTo DBCreate_Error

    Set CN = New ADODB.Connection
    CN.ConnectionString = CONNECTION_STRING
    CN.Open

    query = "CREATE DATABASE " & dbName & ";"
    CN.Execute query

    MsgBox "Base de datos '" & dbName & "' creada exitosamente.", vbInformation
    CN.Close
    Set CN = Nothing
    DBCreate = True

    Exit Function

ErrorHandler:
    DBCreate = False
    MsgBox "Error creando la base de datos: " & Err.Description, vbCritical

    On Error GoTo 0
    Exit Function

DBCreate_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento DBCreate de Módulo modDBConnect línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : DbConnect
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DbConnect()

   On Error GoTo DbConnect_Error

10        Set CN = New ADODB.Connection
20        CN.ConnectionString = CONNECTION_STRING
30        CN.Open

   On Error GoTo 0
   Exit Sub

DbConnect_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento DbConnect de Módulo modDBConnect línea: " & Erl())

End Sub

Public Sub CreateTables()

    On Error GoTo ErrorHandler
    
    Dim query As String
    Dim CN As ADODB.Connection
    Set CN = New ADODB.Connection
    query = CONNECTION_STRING_DB
    CN.Open query
    
    ' Crear tabla tipos_documentos
    query = "CREATE TABLE IF NOT EXISTS tipos_documentos (" & _
            "id_tipodocumento INT NOT NULL AUTO_INCREMENT, " & _
            "nombre VARCHAR(200) NOT NULL, " & _
            "abreviatura VARCHAR(5), " & _
            "validar_como_cuit BIT DEFAULT 0, " & _
            "PRIMARY KEY (id_tipodocumento), " & _
            "UNIQUE KEY uk_abreviatura (abreviatura), " & _
            "UNIQUE KEY uk_nombre (nombre)" & _
            ");"
    CN.Execute query

    ' Crear tabla provincias
    query = "CREATE TABLE IF NOT EXISTS provincias (" & _
            "id_provincia INT NOT NULL AUTO_INCREMENT, " & _
            "nombre VARCHAR(400) NOT NULL, " & _
            "region CHAR(3), " & _
            "PRIMARY KEY (id_provincia), " & _
            "UNIQUE KEY uk_nombre (nombre)" & _
            ");"
    CN.Execute query

    ' Crear tabla localidades
    query = "CREATE TABLE IF NOT EXISTS localidades (" & _
            "id_localidad INT NOT NULL AUTO_INCREMENT, " & _
            "nombre VARCHAR(300) NOT NULL, " & _
            "id_provincia INT NOT NULL, " & _
            "codigo_postal VARCHAR(10), " & _
            "PRIMARY KEY (id_localidad), " & _
            "UNIQUE KEY uk_localidades_nombre_id_provincia (nombre, id_provincia), " & _
            "INDEX fk_provincias_localidades_idx (id_provincia), " & _
            "FOREIGN KEY (id_provincia) REFERENCES provincias (id_provincia)" & _
            ");"
    CN.Execute query

    ' Crear tabla personas
    query = "CREATE TABLE IF NOT EXISTS personas (" & _
            "id_tipodocumento INT NOT NULL, " & _
            "num_documento BIGINT NOT NULL, " & _
            "nombre_apellido VARCHAR(400) NOT NULL, " & _
            "fecha_nacimiento DATE, " & _
            "genero CHAR(1), " & _
            "es_argentino BIT DEFAULT 1, " & _
            "correo_electronico VARCHAR(300), " & _
            "foto_cara BLOB, " & _
            "id_localidad INT, " & _
            "codigo_postal VARCHAR(10), " & _
            "PRIMARY KEY (id_tipodocumento, num_documento), " & _
            "UNIQUE KEY uk_nombre_apellido (nombre_apellido), " & _
            "INDEX fk_localidades_id_localidad_idx (id_localidad), " & _
            "FOREIGN KEY (id_tipodocumento) REFERENCES tipos_documentos (id_tipodocumento), " & _
            "FOREIGN KEY (id_localidad) REFERENCES localidades (id_localidad)" & _
            ");"
    CN.Execute query

    ' Crear tabla usuarios
    query = "CREATE TABLE IF NOT EXISTS usuarios (" & _
            "id_tipodocumento INT NOT NULL, " & _
            "num_documento BIGINT NOT NULL, " & _
            "nombre_usuario VARCHAR(50) NOT NULL, " & _
            "hashed_pwd VARCHAR(200) NOT NULL, " & _
            "PRIMARY KEY (id_tipodocumento, num_documento), " & _
            "UNIQUE KEY uk_nombre_usuario (nombre_usuario), " & _
            "FOREIGN KEY (id_tipodocumento, num_documento) REFERENCES personas (id_tipodocumento, num_documento)" & _
            ");"
    CN.Execute query

    MsgBox "Todas las tablas fueron creadas exitosamente.", vbInformation

    Exit Sub

ErrorHandler:
    Call Logs.LogError("Error al crear las tablas: " & Err.Number & " " & Err.Description & " línea: " & Erl())

End Sub


Public Sub SeedDatabase()
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Dim query As String
    
    ' Conexión a MySQL
    Set conn = New ADODB.Connection
    conn.ConnectionString = CONNECTION_STRING_DB
    conn.Open

    ' Insertar datos en la tabla provincias
    query = "INSERT INTO provincias (nombre, region) VALUES " & _
            "('Buenos Aires', 'CBA'), " & _
            "('Santa Fe', 'LIT'), " & _
            "('Córdoba', 'CEN'), " & _
            "('Mendoza', 'CUY'), " & _
            "('Tucumán', 'NOA');"
    conn.Execute query

    ' Insertar datos en la tabla localidades
    query = "INSERT INTO localidades (nombre, id_provincia, codigo_postal) VALUES " & _
            "('Rosario', 2, '2000'), " & _
            "('Santa Fe', 2, '3000'), " & _
            "('Córdoba', 3, '5000'), " & _
            "('Mendoza', 4, '5500'), " & _
            "('San Miguel de Tucumán', 5, '4000');"
    conn.Execute query

    ' Insertar datos en la tabla tipos_documentos
    query = "INSERT INTO tipos_documentos (nombre, abreviatura, validar_como_cuit) VALUES " & _
            "('DNI', 'DNI', 1), " & _
            "('Pasaporte', 'PASS', 0), " & _
            "('Libreta Cívica', 'LC', 0), " & _
            "('Cédula de Identidad', 'CI', 0), " & _
            "('Carnet de Extranjería', 'CE', 0);"
    conn.Execute query

    ' Insertar datos en la tabla personas
    query = "INSERT INTO personas (id_tipodocumento, num_documento, nombre_apellido, fecha_nacimiento, genero, es_argentino, correo_electronico, id_localidad, codigo_postal) VALUES " & _
            "(1, 12345678, 'Juan Pérez', '1985-05-12', 'M', 1, 'juan.perez@mail.com', 1, '2000'), " & _
            "(1, 87654321, 'María López', '1990-08-23', 'F', 1, 'maria.lopez@mail.com', 3, '5000'), " & _
            "(2, 11223344, 'John Doe', '1982-11-05', 'M', 0, 'john.doe@mail.com', 4, '5500'), " & _
            "(1, 44332211, 'Ana González', '1975-03-14', 'F', 1, 'ana.gonzalez@mail.com', 2, '3000'), " & _
            "(2, 99887766, 'Jane Smith', '1995-07-19', 'F', 0, 'jane.smith@mail.com', 5, '4000');"
    conn.Execute query

    ' Insertar datos en la tabla usuarios
    query = "INSERT INTO usuarios (id_tipodocumento, num_documento, nombre_usuario, hashed_pwd) VALUES " & _
            "(1, 12345678, 'juanp', 'hashed_pwd1'), " & _
            "(1, 87654321, 'marial', 'hashed_pwd2'), " & _
            "(2, 11223344, 'johnd', 'hashed_pwd3'), " & _
            "(1, 44332211, 'anag', 'hashed_pwd4'), " & _
            "(2, 99887766, 'janes', 'hashed_pwd5');"
    conn.Execute query

    MsgBox "Semilla de datos insertada exitosamente.", vbInformation

    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al insertar datos de semilla: " & Err.Description, vbCritical
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
End Sub


Private Sub LoadPersonas()

Dim conn                        As ADODB.Connection
Dim rs                          As ADODB.Recordset
Dim query                       As String

    On Error GoTo ErrorHandler

    ' Conexión a la base de datos
    Set conn = New ADODB.Connection
    conn.ConnectionString = CONNECTION_STRING_DB
    conn.Open

    ' Query para listar personas
    query = "SELECT p.id_tipodocumento, t.nombre AS tipo_documento, " & _
            "p.num_documento, p.nombre_apellido, p.fecha_nacimiento, " & _
            "p.genero, l.nombre AS localidad, p.codigo_postal " & _
            "FROM personas p " & _
            "INNER JOIN tipos_documentos t ON p.id_tipodocumento = t.id_tipodocumento " & _
            "INNER JOIN localidades l ON p.id_localidad = l.id_localidad"

    Set rs = conn.Execute(query)

    ' Llenar la grilla con los datos
    Dim row                     As Integer
    row = 1
    While Not rs.EOF
        MSFlexGrid_Persons.AddItem ""
        MSFlexGrid_Persons.TextMatrix(row, 0) = rs("tipo_documento")
        MSFlexGrid_Persons.TextMatrix(row, 1) = rs("num_documento")
        MSFlexGrid_Persons.TextMatrix(row, 2) = rs("nombre_apellido")
        MSFlexGrid_Persons.TextMatrix(row, 3) = rs("fecha_nacimiento")
        MSFlexGrid_Persons.TextMatrix(row, 4) = rs("genero")
        MSFlexGrid_Persons.TextMatrix(row, 5) = rs("localidad")
        MSFlexGrid_Persons.TextMatrix(row, 6) = rs("codigo_postal")
        rs.MoveNext
        row = row + 1
    Wend

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar personas: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


