Attribute VB_Name = "modDBPersons"
Option Explicit

Public Sub LoadPerson(ByRef tmpForm As Form)

Dim i                           As Integer
Dim row                         As Integer
Dim query                       As String
Dim RS                          As New ADODB.Recordset

    ' Query para listar personas
    On Error GoTo LoadPersonas_Error

    query = "SELECT p.id, p.id_tipodocumento as id_tipo_documento, t.nombre AS tipo_documento, " & _
            "p.num_documento, p.nombre_apellido, p.fecha_nacimiento, " & _
            "p.genero, l.id_localidad, l.nombre AS localidad, l.id_provincia, pv.nombre as provincia, p.codigo_postal, p.correo_electronico, p.es_argentino " & _
            "FROM personas p " & _
            "LEFT JOIN tipos_documentos t ON p.id_tipodocumento = t.id_tipodocumento " & _
            "LEFT JOIN localidades l ON p.id_localidad = l.id_localidad " & _
            "LEFT JOIN provincias pv ON pv.id_provincia = l.id_provincia " & _
            " WHERE p.id = " & tmpUserEdit.Person.id

    Set RS = cn.Execute(query)

    tmpUserEdit.Person.id = Val(RS.Fields("id"))
    tmpUserEdit.Person.Name = RS.Fields("nombre_apellido")
    tmpUserEdit.Person.id_dni = Val(RS.Fields("id_tipo_documento"))
    tmpUserEdit.Person.dni = Val(RS.Fields("num_documento"))
    tmpUserEdit.Person.DateBirth = FormatDateForVB6(RS.Fields("fecha_nacimiento") & vbNullString)
    tmpUserEdit.Person.Genre = RS.Fields("genero") & vbNullString
    tmpUserEdit.Person.id_locality = Val(RS.Fields("id_localidad") & vbNullString)
    tmpUserEdit.Person.id_state = Val(RS.Fields("id_provincia") & vbNullString)
    tmpUserEdit.Person.zip_code = Val(RS.Fields("codigo_postal") & vbNullString)
    tmpUserEdit.Person.Email = RS.Fields("correo_electronico")
    tmpUserEdit.Person.is_argentine = IIf(CBool(RS.Fields("es_argentino")), True, False)

    If tmpUserEdit.Person.id <= 0 Then
        MsgBox "Error al cargar la persona " & tmpUserEdit.Person.Name & ". Al parecer no tiene id en la base de datos."
        Exit Sub
    End If

    tmpForm.txtName = tmpUserEdit.Person.Name

    'Seleccionamos el tipo de dni
    For i = 0 To tmpForm.cmbIdDNIType.ListCount - 1
        If tmpForm.cmbIdDNIType.ItemData(i) = tmpUserEdit.Person.id_dni Then
            tmpForm.cmbIdDNIType.ListIndex = i    ' Seleccionar el ítem
            Exit For
        End If
    Next i

    tmpForm.txtDNI.Text = tmpUserEdit.Person.dni
    tmpForm.txtDateBirth.Text = tmpUserEdit.Person.DateBirth
    tmpForm.cmbGenre = tmpUserEdit.Person.Genre

    'Seleccionamos la provincia:
    If Len(RS.Fields("provincia")) > 0 Then
        For i = 0 To tmpForm.cmbState.ListCount - 1
            If tmpForm.cmbState.ItemData(i) = RS.Fields("id_provincia") Then
                tmpForm.cmbState.ListIndex = i    ' Seleccionar el ítem
                Exit For
            End If
        Next i
    End If
    
    'Cargamos y seleccionamos la localidad, para poder obtener el código postal.
    If Len(RS.Fields("localidad")) > 0 Then
        Call LoadLocality(frmPerson)
        For i = 0 To tmpForm.cmbLocality.ListCount - 1
            If tmpForm.cmbLocality.ItemData(i) = RS.Fields("id_localidad") Then
                tmpForm.cmbLocality.ListIndex = i    ' Seleccionar el ítem
                Exit For
            End If
        Next i
    End If

    'Seleccionamos el género de la persona:
    For i = 0 To tmpForm.cmbGenre.ListCount - 1
        If tmpForm.cmbGenre.List(i) = RS.Fields("genero") Then
            tmpForm.cmbGenre.ListIndex = i    ' Seleccionar el ítem
            Exit For
        End If
    Next i

    tmpForm.txtEmail.Text = RS("correo_electronico")
    tmpForm.chkIsArgentine.Value = IIf(CBool(RS("es_argentino")), vbChecked, vbUnchecked)

    Set RS = Nothing

    On Error GoTo 0
    Exit Sub

LoadPersonas_Error:
    If Not RS Is Nothing Then RS.Close
    Set RS = Nothing
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadPersonas de Módulo modDBPersons línea: " & Erl())

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadPersons
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadPersons()

Dim row                         As Integer
Dim query                       As String
Dim RS                          As New ADODB.Recordset

    ' Query para listar personas
    On Error GoTo LoadPersons_Error

    query = "SELECT p.id, p.id_tipodocumento, t.nombre AS tipo_documento, " & _
            "p.num_documento, p.nombre_apellido, p.fecha_nacimiento, " & _
            "p.genero, l.nombre AS localidad, pv.nombre as provincia, p.codigo_postal, p.correo_electronico, p.es_argentino " & _
            "FROM personas p " & _
            "LEFT JOIN tipos_documentos t ON p.id_tipodocumento = t.id_tipodocumento " & _
            "LEFT JOIN localidades l ON p.id_localidad = l.id_localidad " & _
            "LEFT JOIN provincias pv ON pv.id_provincia = l.id_provincia "

    Set RS = cn.Execute(query)

    Debug.Print "Count persons: " & RS.RecordCount

    ' Llenar la grilla con los datos
    row = 1
    While Not RS.EOF
        frmAbmPersons.MSFlexGrid_Persons.AddItem ""

        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 0) = RS("id")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 1) = RS("tipo_documento")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 2) = RS("num_documento")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 3) = RS("nombre_apellido")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 4) = RS("fecha_nacimiento") & vbNullString
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 5) = RS("genero") & vbNullString
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 6) = RS("localidad") & " - " & RS("provincia")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 7) = RS("codigo_postal") & vbNullString
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 8) = RS("correo_electronico")
        frmAbmPersons.MSFlexGrid_Persons.TextMatrix(row, 9) = IIf(CBool(RS("es_argentino")), "SI", "NO")

        RS.MoveNext
        row = row + 1
    Wend

    Set RS = Nothing

    On Error GoTo 0
    Exit Sub

LoadPersons_Error:
    If Not RS Is Nothing Then RS.Close
    Set RS = Nothing
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LoadPersons de Módulo modDBPersons línea: " & Erl())

End Sub

Public Sub LoadDNITypes(ByRef tmpForm As Form)

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset
Dim cmbIndex                    As Long

    ' Consulta para obtener los datos de tipos_documentos
    On Error GoTo LoadDNITypes_Error

10  sQuery = "SELECT id_tipodocumento, nombre, abreviatura FROM tipos_documentos"

    ' Ejecutar la consulta y abrir un Recordset
20  Set RS = New ADODB.Recordset
30  RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

    ' Limpiar el ComboBox antes de cargar datos
40  tmpForm.cmbIdDNIType.Clear

    ' Validar si hay registros en el Recordset
50  If Not RS.EOF Then
60      Do While Not RS.EOF
            ' Agregar el ítem al ComboBox con el formato "nombre - abreviatura"
70          tmpForm.cmbIdDNIType.AddItem RS.Fields("nombre").Value & " - " & RS.Fields("abreviatura").Value

            ' Obtener el índice del ítem agregado
80          cmbIndex = tmpForm.cmbIdDNIType.NewIndex

            ' Asignar el id_tipodocumento al ItemData del ítem actual
90          tmpForm.cmbIdDNIType.ItemData(cmbIndex) = RS.Fields("id_tipodocumento").Value

            ' Avanzar al siguiente registro
100         RS.MoveNext
110     Loop
120 End If

    ' Cerrar el Recordset
130 RS.Close
140 Set RS = Nothing

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

'---------------------------------------------------------------------------------------
' Procedure : GetZipCodeFromLocality
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 29/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function GetZipCodeFromLocality(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Long

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset

10  On Error GoTo GetZipCodeFromLocality_Error

20  sQuery = "SELECT codigo_postal FROM localidades WHERE id_localidad = " & tmpUser.Person.id_locality

    ' Ejecutar la consulta y abrir un Recordset
30  Set RS = New ADODB.Recordset
40  RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

50  If Not RS.EOF Then
60      GetZipCodeFromLocality = Val(RS.Fields("codigo_postal"))
        RS.Close
        Set RS = Nothing
        Exit Function
70  Else
80      sErrorMsg = "Error al intentar obtener el código postal de la localidad seleccionada."
90      Exit Function
100 End If

140 On Error GoTo 0
150 Exit Function

GetZipCodeFromLocality_Error:

110 GetZipCodeFromLocality = -1
120 RS.Close
130 Set RS = Nothing

160 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento GetZipCodeFromLocality de Módulo modDBPersons línea: " & Erl())

End Function

Public Sub LoadLocality(ByRef tmpForm As Form)

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset
Dim cmbIndex                    As Long

    On Error GoTo LoadLocality_Error

10  Debug.Print "Form: " & tmpForm.Name

    ' Consulta para obtener los datos de tipos_documentos
    
    If tmpForm.cmbState.ListIndex < 0 Then
        Exit Sub
    End If
    
20  sQuery = "SELECT id_localidad, nombre, id_provincia, codigo_postal FROM localidades WHERE id_provincia = " & tmpForm.cmbState.ItemData(tmpForm.cmbState.ListIndex)

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

Function PersonCreate(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim sQuery                      As String
Dim cHash                       As New CSHA256
Dim RS                          As ADODB.Recordset

    'Validamos la existencia del nombre de usuario
10  On Error GoTo PersonCreate_Error

    'Puede existir una persona con el mismo nombre...
      If modDB.Exists("personas", "nombre_apellido", tmpUser.Person.Name) Then
          sErrorMsg = "El nombre y apellido elegido ya están en uso, por favor utilice otro."
          Exit Function
      End If

    'Validamos la existencia del correo electrónico
20  If modDB.Exists("personas", "correo_electronico", tmpUser.Person.Email) Then
30      sErrorMsg = "El correo electrónico ya está en uso, por favor utilice otro."
40      Exit Function
50  End If

60  ReDim tmpFields(1 To 2) As String
70  ReDim tmpValues(1 To 2) As String

80  tmpFields(1) = "id_tipodocumento"
90  tmpFields(2) = "num_documento"

100 ReDim tmpValues(1 To 2) As String
110 ReDim tmpValues(1 To 2) As String
120 tmpValues(1) = tmpUser.Person.id_dni
130 tmpValues(2) = tmpUser.Person.dni

    'Validamos la existencia del DNI
140 If modDB.ExistsArr("personas", tmpFields, tmpValues) Then
150     sErrorMsg = "El dni ya está en uso, por favor utilice otro."
160     Exit Function
170 End If

180 tmpUser.HashedPwd = cHash.SHA256(tmpUser.Password)

    'Primero insertamos los datos del usuario en la tabla "personas" para garantizar que las claves foráneas requeridas en la tabla "usuarios" (id_tipodocumento, num_documento) existan.
    'Esto evita conflictos de integridad referencial al insertar en la tabla "usuarios".
190 sQuery = "INSERT INTO personas (id_tipodocumento, num_documento, nombre_apellido, fecha_nacimiento, genero, es_argentino, correo_electronico, id_localidad, codigo_postal)  VALUES ( " & tmpUser.Person.id_dni & "," & tmpUser.Person.dni & ",'" & tmpUser.Person.Name & "','" & FormatDateForMySQL(tmpUser.Person.DateBirth) & "','" & tmpUser.Person.Genre & "'," & IIf(CBool(tmpUser.Person.is_argentine), 1, 0) & ",'" & tmpUser.Person.Email & "'," & tmpUser.Person.id_locality & "," & tmpUser.Person.zip_code & ")"
200 Set RS = cn.Execute(sQuery, , adOpenForwardOnly)

230 PersonCreate = True

240 On Error GoTo 0
250 Exit Function

PersonCreate_Error:
260 PersonCreate = False
270 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento PersonCreate de Módulo modDBUser línea: " & Erl())

End Function

Function PersonEdit(ByRef tmpUser As tUser) As Boolean

Dim sQuery                      As String
Dim cHash                       As New CSHA256
Dim RS                          As ADODB.Recordset

    'Validamos la existencia del nombre de usuario
10  On Error GoTo PersonEdit_Error

20  sQuery = "UPDATE personas SET " & _
             "id_tipodocumento = " & tmpUser.Person.id_dni & ", " & _
             "num_documento = " & tmpUser.Person.dni & ", " & _
             "nombre_apellido = '" & tmpUser.Person.Name & "', " & _
             "fecha_nacimiento = '" & FormatDateForMySQL(tmpUser.Person.DateBirth) & "', " & _
             "genero = '" & tmpUser.Person.Genre & "', " & _
             "es_argentino = " & IIf(CBool(tmpUser.Person.is_argentine), 1, 0) & ", " & _
             "correo_electronico = '" & tmpUser.Person.Email & "', " & _
             "id_localidad = " & tmpUser.Person.id_locality & ", " & _
             "codigo_postal = '" & tmpUser.Person.zip_code & "' " & _
             "WHERE id = " & tmpUser.Person.id

30  Set RS = cn.Execute(sQuery, , adOpenForwardOnly)
40  PersonEdit = True

50  On Error GoTo 0
60  Exit Function

PersonEdit_Error:
70  PersonEdit = False
80  Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento PersonEdit de Módulo modDBUser línea: " & Erl())

End Function

Public Function FormatDateForMySQL(ByVal vb6Date As String) As String
Dim dayPart                     As String
Dim monthPart                   As String
Dim yearPart                    As String

    ' Validar que la fecha tiene el formato esperado
    If InStr(vb6Date, "/") > 0 Then
        dayPart = Split(vb6Date, "/")(0)
        monthPart = Split(vb6Date, "/")(1)
        yearPart = Split(vb6Date, "/")(2)
        FormatDateForMySQL = yearPart & "-" & monthPart & "-" & dayPart
    Else
        ' Devuelve una cadena vacía si la fecha no es válida
        FormatDateForMySQL = ""
    End If
End Function

Public Function FormatDateForVB6(ByVal mysqlDate As String) As String
Dim yearPart                    As String
Dim monthPart                   As String
Dim dayPart                     As String

    ' Validar que la fecha tiene el formato esperado
    If InStr(mysqlDate, "/") > 0 Then
        yearPart = Split(mysqlDate, "/")(2)
        monthPart = Split(mysqlDate, "/")(1)
        dayPart = Split(mysqlDate, "/")(0)
        
        If Len(monthPart) = 1 Then monthPart = "0" & monthPart
        
        FormatDateForVB6 = dayPart & "/" & monthPart & "/" & yearPart
    Else
        ' Devuelve una cadena vacía si la fecha no es válida
        FormatDateForVB6 = ""
    End If
End Function

Function PersonDelete(ByRef tmpUser As tUser) As Boolean

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset

   On Error GoTo PersonDelete_Error

11    sQuery = "DELETE FROM usuarios WHERE num_documento = " & tmpUser.Person.dni & " AND id_tipodocumento = " & tmpUser.Person.id_dni
21    Set RS = cn.Execute(sQuery, , adOpenForwardOnly)

10    sQuery = "DELETE FROM personas WHERE id = " & tmpUser.Person.id
20    Set RS = cn.Execute(sQuery, , adOpenForwardOnly)
30    PersonDelete = True

   On Error GoTo 0
   Exit Function

PersonDelete_Error:
    PersonDelete = False
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento PersonDelete de Módulo modDBPersons línea: " & Erl())

End Function
