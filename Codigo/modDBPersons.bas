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
            "p.genero, l.nombre AS localidad, pv.nombre as provincia, p.codigo_postal, p.correo_electronico, p.es_argentino " & _
            "FROM personas p " & _
            "INNER JOIN tipos_documentos t ON p.id_tipodocumento = t.id_tipodocumento " & _
            "INNER JOIN localidades l ON p.id_localidad = l.id_localidad " & _
            "INNER JOIN provincias pv ON pv.id_provincia = l.id_provincia "

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
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 5) = RS("localidad") & " - " & RS("provincia")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 6) = RS("codigo_postal")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 7) = RS("correo_electronico")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 8) = IIf(CBool(RS("es_argentino")), "SI", "NO")
        
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

Public Sub LoadDNITypes(ByRef tmpForm As Form)

      Dim sQuery                      As String
      Dim RS                          As ADODB.Recordset
      Dim cmbIndex                    As Long

          ' Consulta para obtener los datos de tipos_documentos
   On Error GoTo LoadDNITypes_Error

10        sQuery = "SELECT id_tipodocumento, nombre, abreviatura FROM tipos_documentos"

          ' Ejecutar la consulta y abrir un Recordset
20        Set RS = New ADODB.Recordset
30        RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

          ' Limpiar el ComboBox antes de cargar datos
40        tmpForm.cmbIdDNIType.Clear

          ' Validar si hay registros en el Recordset
50        If Not RS.EOF Then
60            Do While Not RS.EOF
                  ' Agregar el ítem al ComboBox con el formato "nombre - abreviatura"
70                tmpForm.cmbIdDNIType.AddItem RS.Fields("nombre").Value & " - " & RS.Fields("abreviatura").Value

                  ' Obtener el índice del ítem agregado
80                cmbIndex = tmpForm.cmbIdDNIType.NewIndex

                  ' Asignar el id_tipodocumento al ItemData del ítem actual
90                tmpForm.cmbIdDNIType.ItemData(cmbIndex) = RS.Fields("id_tipodocumento").Value

                  ' Avanzar al siguiente registro
100               RS.MoveNext
110           Loop
120       End If

          ' Cerrar el Recordset
130       RS.Close
140       Set RS = Nothing

   On Error GoTo 0
   Exit Sub

LoadDNITypes_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadDNITypes de Módulo modDBPersons línea: " & Erl())

End Sub

Public Sub LoadStates(ByRef tmpForm As Form)

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset
Dim cmbIndex                    As Long

    On Error GoTo LoadStates_Error

10  Debug.Print "Form: " & tmpForm.Name

    ' Consulta para obtener los datos de tipos_documentos
20  sQuery = "SELECT id_provincia, nombre, region FROM provincias "

    ' Ejecutar la consulta y abrir un Recordset
30  Set RS = New ADODB.Recordset
40  RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

    ' Limpiar el ComboBox antes de cargar datos
50  tmpForm.cmbState.Clear

    ' Validar si hay registros en el Recordset
60  If Not RS.EOF Then
70      Do While Not RS.EOF
            ' Agregar el ítem al ComboBox con el formato "nombre - abreviatura"
80          tmpForm.cmbState.AddItem RS.Fields("nombre").Value & " - " & RS.Fields("region")

            ' Obtener el índice del ítem agregado
90          cmbIndex = tmpForm.cmbState.NewIndex

            ' Asignar el id_tipodocumento al ItemData del ítem actual
100         tmpForm.cmbState.ItemData(cmbIndex) = RS.Fields("id_provincia").Value

            ' Avanzar al siguiente registro
110         RS.MoveNext
120     Loop
130 End If

    ' Cerrar el Recordset
140 RS.Close
150 Set RS = Nothing

    On Error GoTo 0
    Exit Sub

LoadStates_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadStates de Módulo modDBPersons línea: " & Erl())

End Sub

Public Sub LoadLocality(ByRef tmpForm As Form)

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset
Dim cmbIndex                    As Long

    On Error GoTo LoadLocality_Error

10  Debug.Print "Form: " & tmpForm.Name

    ' Consulta para obtener los datos de tipos_documentos
20  sQuery = "SELECT id_localidad, nombre, id_provincia, codigo_postal FROM localidades "

    ' Ejecutar la consulta y abrir un Recordset
30  Set RS = New ADODB.Recordset
40  RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

    ' Limpiar el ComboBox antes de cargar datos
50  tmpForm.cmbLocality.Clear

    ' Validar si hay registros en el Recordset
60  If Not RS.EOF Then
70      Do While Not RS.EOF
            ' Agregar el ítem al ComboBox con el formato "nombre - abreviatura"
80          tmpForm.cmbLocality.AddItem RS.Fields("nombre").Value & " - " & RS.Fields("codigo_postal")

            ' Obtener el índice del ítem agregado
90          cmbIndex = tmpForm.cmbLocality.NewIndex

            ' Asignar el id_tipodocumento al ItemData del ítem actual
100         tmpForm.cmbLocality.ItemData(cmbIndex) = RS.Fields("id_localidad").Value

            ' Avanzar al siguiente registro
110         RS.MoveNext
120     Loop
130 End If

    ' Cerrar el Recordset
140 RS.Close
150 Set RS = Nothing

    On Error GoTo 0
    Exit Sub

LoadLocality_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadLocality de Módulo modDBPersons línea: " & Erl())

End Sub
