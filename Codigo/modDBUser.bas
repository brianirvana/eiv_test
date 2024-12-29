Attribute VB_Name = "modDBUser"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ValidateDBPassword
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ValidateDBPassword(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim sQuery                      As String
Dim tmpHashedPwd                As String
Dim tmpDBHashedPwd              As String
Dim cHash                       As New CSHA256
Dim RS                          As ADODB.Recordset

10  On Error GoTo ValidateDBPassword_Error

20  tmpHashedPwd = cHash.SHA256(tmpUser.Password)

30  sQuery = "SELECT hashed_pwd FROM usuarios WHERE nombre_usuario = '" & tmpUser.UserName & "'"
40  Set RS = cn.Execute(sQuery, , adOpenForwardOnly)

50  If RS.EOF Or RS.BOF Then
60      sErrorMsg = "El usuario no existe. Por favor, cree un nuevo usuario."
70      Exit Function
80  End If

90  tmpDBHashedPwd = RS.Fields("hashed_pwd")

100 If modDB.Exists("usuarios", "nombre_usuario", tmpUser.UserName) Then
110     If StrComp(tmpHashedPwd, tmpDBHashedPwd) <> 0 Then
120         sErrorMsg = "El usuario y la contraseña no coinciden."
130         Exit Function
140     End If
150 Else
160     sErrorMsg = "El usuario (" & tmpUser.UserName & ") a validar la contraseña, no existe."
170     Exit Function
180 End If

190 ValidateDBPassword = True

200 On Error GoTo 0
210 Exit Function

ValidateDBPassword_Error:

220 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateDBPassword de Módulo modDBUser línea: " & Erl())

End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateUser
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function UserCreate(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim sQuery                      As String
Dim cHash                       As New CSHA256
Dim RS                          As ADODB.Recordset

    'Validamos la existencia del nombre de usuario
10  On Error GoTo UserCreate_Error

20  If modDB.Exists("usuarios", "nombre_usuario", tmpUser.UserName) Then
30      sErrorMsg = "El usuario elegido ya está en uso, por favor utilice otro."
40      Exit Function
50  End If

    'Validamos la existencia del correo electrónico
60  If modDB.Exists("personas", "correo_electronico", tmpUser.Person.Email) Then
70      sErrorMsg = "El correo electrónico ya está en uso, por favor utilice otro."
80      Exit Function
90  End If

    Dim tmpFields()             As String
    Dim tmpValues()             As String

100 ReDim tmpFields(1 To 2) As String
110 ReDim tmpValues(1 To 2) As String

120 tmpFields(1) = "id_tipodocumento"
130 tmpFields(2) = "num_documento"

140 ReDim tmpValues(1 To 2) As String
150 ReDim tmpValues(1 To 2) As String
160 tmpValues(1) = tmpUser.Person.id_dni
170 tmpValues(2) = tmpUser.Person.dni

    'Validamos la existencia del DNI
180 If modDB.ExistsArr("personas", tmpFields, tmpValues) Then
190     sErrorMsg = "El dni ya está en uso, por favor utilice otro."
200     Exit Function
210 End If

220 tmpUser.HashedPwd = cHash.SHA256(tmpUser.Password)

    'Primero insertamos los datos del usuario en la tabla "personas" para garantizar que las claves foráneas requeridas en la tabla "usuarios" (id_tipodocumento, num_documento) existan.
    'Esto evita conflictos de integridad referencial al insertar en la tabla "usuarios".
230 sQuery = "INSERT INTO personas (id_tipodocumento, num_documento, nombre_apellido, correo_electronico)  VALUES ( " & tmpUser.Person.id_dni & "," & tmpUser.Person.dni & ",'" & tmpUser.Person.FirstName & " " & tmpUser.Person.LastName & "','" & tmpUser.Person.Email & "')"
240 Set RS = cn.Execute(sQuery, , adOpenForwardOnly)

250 sQuery = "INSERT INTO usuarios (id_tipodocumento, num_documento, nombre_usuario, hashed_pwd)  VALUES ( " & tmpUser.Person.id_dni & "," & tmpUser.Person.dni & ",'" & tmpUser.UserName & "','" & tmpUser.HashedPwd & "')"
260 Set RS = cn.Execute(sQuery, , adOpenForwardOnly)

270 UserCreate = True

280 On Error GoTo 0
290 Exit Function

UserCreate_Error:
300 UserCreate = False
310 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento UserCreate de Módulo modDBUser línea: " & Erl())

End Function
