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
40  Set RS = CN.Execute(sQuery, , adOpenForwardOnly)

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
Function CreateUser(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim sQuery                      As String
Dim cHash                       As New CSHA256
Dim RS                          As ADODB.Recordset

    'Validamos la existencia del nombre de usuario
10  On Error GoTo CreateUser_Error

20  If modDB.Exists("usuarios", "nombre_usuario", tmpUser.UserName) Then
30      sErrorMsg = "El usuario elegido ya está en uso, por favor utilice otro."
40      Exit Function
50  End If

    'Validamos la existencia del correo electrónico
60  If modDB.Exists("personas", "correo_electronico", tmpUser.Email) Then
70      sErrorMsg = "El correo electrónico ya está en uso, por favor utilice otro."
80      Exit Function
90  End If

    'Validamos la existencia del DNI
100 If modDB.Exists("personas", "num_documento", CStr(tmpUser.dni)) Then
110     sErrorMsg = "El correo electrónico ya está en uso, por favor utilice otro."
120     Exit Function
130 End If

140 tmpUser.HashedPwd = cHash.SHA256(tmpUser.Password)

    'Primero insertamos los datos del usuario en la tabla "personas" para garantizar que las claves foráneas requeridas en la tabla "usuarios" (id_tipodocumento, num_documento) existan.
    'Esto evita conflictos de integridad referencial al insertar en la tabla "usuarios".
150 sQuery = "INSERT INTO personas (id_tipodocumento, num_documento, nombre_apellido, correo_electronico)  VALUES ( " & tmpUser.id_dni & "," & tmpUser.dni & ",'" & tmpUser.FirstName & " " & tmpUser.LastName & "','" & tmpUser.Email & "')"
160 Set RS = CN.Execute(sQuery, , adOpenForwardOnly)

170 sQuery = "INSERT INTO usuarios (id_tipodocumento, num_documento, nombre_usuario, hashed_pwd)  VALUES ( " & tmpUser.id_dni & "," & tmpUser.dni & ",'" & tmpUser.UserName & "','" & tmpUser.HashedPwd & "')"
180 Set RS = CN.Execute(sQuery, , adOpenForwardOnly)

190 CreateUser = True

200 On Error GoTo 0
210 Exit Function

CreateUser_Error:
220 CreateUser = False
230 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CreateUser de Módulo modDBUser línea: " & Erl())

End Function
