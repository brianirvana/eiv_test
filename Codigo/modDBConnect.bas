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
Private Const CONNECTION_STRING As String = "Driver={MySQL ODBC 8.0 Driver};Server=localhost;User=root;Password=;"

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
