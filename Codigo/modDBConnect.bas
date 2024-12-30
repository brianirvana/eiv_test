Attribute VB_Name = "modDBConnect"
'---------------------------------------------------------------------------------------
' Module    : modDBConnect
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public cn                       As ADODB.Connection
Public RS                       As New ADODB.Recordset

' Requiere una referencia a Microsoft ActiveX Data Objects 2.x Library
Public CONNECTION_STRING As String  '= "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;UID=root;PWD=;PORT=3306;OPTION=3;ConnectionLifetime=0;ConnectionTimeout=0"
Public CONNECTION_STRING_DB As String  '= "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;DATABASE=eiv;UID=root;PWD=;PORT=3306"

' Declaración de la variable global para almacenar los datos de configuración
Public Type Param
    dbCustomCs                  As String
    dbDriverVer                 As String
    dbServer                    As String
    dbName                      As String
    dbUser                      As String
    dbPasswd                    As String
    dbDesc                      As String
    dbPort                      As Long
End Type

Public cDB                      As Param

' Función para cargar la configuración de la base de datos desde el archivo db.ini
'---------------------------------------------------------------------------------------
' Procedure : LoadDBConfig
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadDBConfig()

    On Error GoTo LoadDBConfig_Error

10  cDB.dbDesc = GetVar(App.Path & "\DB.ini", "DATABASE", "DESC")
20  cDB.dbCustomCs = GetVar(App.Path & "\DB.ini", "DATABASE", "CUSTOMCS")
30  cDB.dbName = GetVar(App.Path & "\DB.ini", "DATABASE", "DBNAME")
40  cDB.dbPasswd = GetVar(App.Path & "\DB.ini", "DATABASE", "DBPASSWD")
50  cDB.dbUser = GetVar(App.Path & "\DB.ini", "DATABASE", "DBUSER")
60  cDB.dbDriverVer = GetVar(App.Path & "\DB.ini", "DATABASE", "DriverVer")
70  cDB.dbServer = GetVar(App.Path & "\DB.ini", "DATABASE", "DBSERVER")
80  cDB.dbPort = Val(GetVar(App.Path & "\DB.ini", "DATABASE", "DBPORT"))

    ' Validar el puerto como número
90  If IsNumeric(cDB.dbPort) Then
100     cDB.dbPort = CLng(cDB.dbPort)
110 Else
120     cDB.dbPort = 3306    ' Valor por defecto si no se encuentra o no es válido
130 End If

    ' Crear las cadenas de conexión
140 CONNECTION_STRING = "DRIVER={" & cDB.dbDriverVer & "};SERVER=" & cDB.dbServer & ";" & _
                        "UID=" & cDB.dbUser & ";PWD=" & cDB.dbPasswd & ";" & _
                        "PORT=" & cDB.dbPort & ";OPTION=3;ConnectionLifetime=0;ConnectionTimeout=0"

150 CONNECTION_STRING_DB = "DRIVER={" & cDB.dbDriverVer & "};SERVER=" & cDB.dbServer & ";" & _
                           "DATABASE=" & cDB.dbName & ";UID=" & cDB.dbUser & ";" & _
                           "PWD=" & cDB.dbPasswd & ";PORT=" & cDB.dbPort

    On Error GoTo 0
    Exit Sub

LoadDBConfig_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadDBConfig de Módulo modDBConnect línea: " & Erl())

End Sub

' Función para leer el contenido de un archivo
Private Function ReadFile(filePath As String) As String
    Dim fileNumber As Integer
    Dim fileContents As String
    
    On Error GoTo ErrorHandler
    
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContents = Input(LOF(fileNumber), fileNumber)
    Close #fileNumber
    ReadFile = fileContents
    Exit Function

ErrorHandler:
    MsgBox "Error al leer el archivo: " & Err.Description, vbCritical, "Error"
    Close #fileNumber
    ReadFile = ""
End Function


'---------------------------------------------------------------------------------------
' Procedure : DBExists
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DBExists(ByVal dbName As String) As Boolean

Dim cn                          As ADODB.Connection
Dim RS                          As ADODB.Recordset
Dim query                       As String

    On Error GoTo DBExists_Error

    Set cn = New ADODB.Connection
    cn.ConnectionString = CONNECTION_STRING
    cn.Open

    query = "SHOW DATABASES LIKE '" & dbName & "';"
    Set RS = cn.Execute(query)

    DBExists = Not RS.EOF

    Call SaveSetting(App.Path, "EIV_SOFTWARE", "IsDBAlreadyExists", IIf(DBExists, "1", "0"))

    RS.Close
    cn.Close
    Set RS = Nothing
    Set cn = Nothing

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

Dim cn                          As ADODB.Connection
Dim query                       As String

10  On Error GoTo DBCreate_Error

20  Set cn = New ADODB.Connection
30  cn.ConnectionString = CONNECTION_STRING
40  cn.Open

80  cn.CursorLocation = adUseClient

90  query = "CREATE DATABASE " & dbName & ";"
100 cn.Execute query
'50  cn.ConnectionTimeout = 0
'60  cn.CommandTimeout = 0
110 MsgBox "Base de datos '" & dbName & "' creada exitosamente.", vbInformation
120 cn.Close
130 Set cn = Nothing
140 DBCreate = True

150 Exit Function

ErrorHandler:
160 DBCreate = False
170 MsgBox "Error creando la base de datos: " & Err.Description, vbCritical

180 On Error GoTo 0
190 Exit Function

DBCreate_Error:

200 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento DBCreate de Módulo modDBConnect línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : DbConnect
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DbConnect()

10  On Error GoTo DbConnect_Error

20  Set cn = New ADODB.Connection
30  cn.ConnectionTimeout = 0
40  cn.CommandTimeout = 0
50  cn.ConnectionString = CONNECTION_STRING
60  cn.Open
70  cn.CursorLocation = adUseClient

80  cn.Execute ("USE eiv"), , adOpenForwardOnly

90  On Error GoTo 0
100 Exit Sub

DbConnect_Error:

110 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento DbConnect de Módulo modDBConnect línea: " & Erl())

End Sub

Public Sub CreateTables()

    On Error GoTo ErrorHandler
    
    Dim query As String
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    query = CONNECTION_STRING_DB
    cn.Open query
    
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
    cn.Execute query

    ' Crear tabla provincias
    query = "CREATE TABLE IF NOT EXISTS provincias (" & _
            "id_provincia INT NOT NULL AUTO_INCREMENT, " & _
            "nombre VARCHAR(400) NOT NULL, " & _
            "region CHAR(3), " & _
            "PRIMARY KEY (id_provincia), " & _
            "UNIQUE KEY uk_nombre (nombre)" & _
            ");"
    cn.Execute query

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
    cn.Execute query

    ' Crear tabla personas
    query = "CREATE TABLE IF NOT EXISTS personas (" & _
            "id INT NOT NULL AUTO_INCREMENT, " & _
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
            "UNIQUE KEY uk_id (id), " & _
            "UNIQUE KEY uk_nombre_apellido (nombre_apellido), " & _
            "INDEX fk_localidades_id_localidad_idx (id_localidad), " & _
            "FOREIGN KEY (id_tipodocumento) REFERENCES tipos_documentos (id_tipodocumento), " & _
            "FOREIGN KEY (id_localidad) REFERENCES localidades (id_localidad)" & _
            ");"

    cn.Execute query

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
    cn.Execute query

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

'---------------------------------------------------------------------------------------
' Procedure : DropDatabase
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   : Función para borrar la base de datos
'---------------------------------------------------------------------------------------
'
Public Function DropDatabase(dbName As String) As Boolean

Dim sql                         As String

    ' Validar el nombre de la base de datos
    On Error GoTo DropDatabase_Error

10  If Trim(dbName) = "" Then
20      MsgBox "El nombre de la base de datos no puede estar vacío.", vbExclamation, "Error"
30      Exit Function
40  End If

    ' Crear la cadena SQL para eliminar la base de datos
50  sql = "DROP DATABASE IF EXISTS " & dbName

    ' Ejecutar la instrucción SQL para borrar la base de datos
90  cn.Execute sql

    ' Confirmar la eliminación
100 MsgBox "La base de datos '" & cDB.dbName & "' ha sido eliminada con éxito.", vbInformation, "Éxito"
110 DropDatabase = True

    On Error GoTo 0
    Exit Function

DropDatabase_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento DropDatabase de Módulo modDBConnect línea: " & Erl())

End Function
