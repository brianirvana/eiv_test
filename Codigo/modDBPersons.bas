Attribute VB_Name = "modDBPersons"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : LoadPersonas
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadPersonas()

    Dim row                         As Integer
    Dim query                       As String
    Dim RS                          As New ADODB.Recordset

    ' Query para listar personas
    On Error GoTo LoadPersonas_Error

    query = "SELECT p.id_tipodocumento, t.nombre AS tipo_documento, " & _
            "p.num_documento, p.nombre_apellido, p.fecha_nacimiento, " & _
            "p.genero, l.nombre AS localidad, p.codigo_postal " & _
            "FROM personas p " & _
            "INNER JOIN tipos_documentos t ON p.id_tipodocumento = t.id_tipodocumento " & _
            "INNER JOIN localidades l ON p.id_localidad = l.id_localidad"

    Set RS = cn.Execute(query)

    ' Llenar la grilla con los datos
    row = 1
    While Not RS.EOF
        frmAbmPersons.MSFlexGrid_Persons.AddItem ""
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 0) = RS("tipo_documento")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 1) = RS("num_documento")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 2) = RS("nombre_apellido")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 3) = RS("fecha_nacimiento")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 4) = RS("genero")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 5) = RS("localidad")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 6) = RS("codigo_postal")
        RS.MoveNext
        row = row + 1
    Wend

    Set RS = Nothing

    On Error GoTo 0
    Exit Sub

LoadPersonas_Error:
    If Not RS Is Nothing Then RS.Close
    Set RS = Nothing
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadPersonas de Módulo modDBPersons línea: " & Erl())
End Sub


