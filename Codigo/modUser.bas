Attribute VB_Name = "modUser"
Option Explicit

'Public aUser                    As tUser

'Esto debería ir en modPerson, pero debido a las dependencias circulares de vb6 no fue posible.
Public Type tPerson
    id                          As Long         'id auto incremental
    id_dni                      As Long         'Tipo documento
    dni                         As Long         'DNI Número
    Name                        As String       'Nombre
    DateBirth                   As String       'Fecha nacimiento
    Genre                       As String * 1   'Género
    is_argentine                As Boolean      'Es argentino?
    email                       As String       'Correo electrónico
    pic_face                    As Variant      'Foto cara
    Id_Locality                 As Long         'Localidad
    id_state                    As Long         'Provincia
    zip_code                    As Long         'Codigo Postal
End Type

Public Type tUser
    UserName                    As String
    Password                    As String
    HashedPwd                   As String
    Person                      As tPerson
End Type

'---------------------------------------------------------------------------------------
' Procedure : NumberToPunctuatedString
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Agregamos puntitos en los miles del DNI para mejorar su visualización.
'---------------------------------------------------------------------------------------
'
Public Function NumberToPunctuatedString(ByVal Cantidad As Double) As String

Dim i                           As Double
Dim CantidadStr                 As String
Dim FinalStr                    As String

    On Error GoTo NumberToPunctuatedString_Error

10  CantidadStr = CStr(Cantidad)

20  For i = Len(CantidadStr) To 1 Step -3
30      FinalStr = Right$(Left$(CantidadStr, i), 3) & "." & FinalStr
40  Next i

50  NumberToPunctuatedString = Left$(FinalStr, Len(FinalStr) - 1)

    On Error GoTo 0
    Exit Function

NumberToPunctuatedString_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento NumberToPunctuatedString de Módulo modUserCreate línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidatePassword
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ValidatePassword(ByRef UserPassword As String, ByRef sErrorMsg As String) As Boolean

Dim ValidNumeric                As Boolean
Dim ValidLetters                As Boolean
Dim loopc                       As Byte
Dim CharAscii                   As String
Dim tmpIsNumber                 As String

    On Error GoTo ValidatePassword_Error

    ' Verificamos la longitud de la contraseña
    If Len(UserPassword) < 6 Or Len(UserPassword) > 32 Then
        sErrorMsg = "La contraseña debe tener entre 6 y 32 caracteres."
        ValidatePassword = False
        Exit Function
    End If

    ' Recorremos cada carácter de la contraseña
    For loopc = 1 To Len(UserPassword)
        CharAscii = Mid$(UserPassword, loopc, 1)
        tmpIsNumber = Asc(CharAscii)

        ' Verificamos si el carácter es válido (por ejemplo, letras o números)
        If Not LegalCharacter(tmpIsNumber) Then
            sErrorMsg = "Password inválido. El carácter '" & CharAscii & "' no está permitido."
            ValidatePassword = False
            Exit Function
        ElseIf IsNumeric(CharAscii) Then
            ValidNumeric = True
        ElseIf tmpIsNumber >= 65 And tmpIsNumber <= 90 Or tmpIsNumber >= 97 And tmpIsNumber <= 122 Then
            ' Consideramos letras (mayúsculas y minúsculas)
            ValidLetters = True
        End If
    Next loopc

    ' Verificamos si cumple con los requisitos de contenido
    If Not ValidLetters Or Not ValidNumeric Then
        sErrorMsg = "La contraseña debe contener al menos un número y al menos una letra."
        ValidatePassword = False
        Exit Function
    End If

    ' Si pasa todas las validaciones
    ValidatePassword = True
    Exit Function

ValidatePassword_Error:
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidatePassword de Módulo modUser línea: " & Erl())
    ValidatePassword = False

End Function

'---------------------------------------------------------------------------------------
' Procedure : LegalCharacter
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
    On Error GoTo LegalCharacter_Error

10  If KeyAscii = 8 Then
20      LegalCharacter = True
30      Exit Function
40  End If

    'Only allow space, numbers, letters and special characters
50  If KeyAscii < 32 Or KeyAscii = 44 Then
60      Exit Function
70  End If

80  If KeyAscii > 126 Then
90      Exit Function
100 End If

    'Check for bad special characters in between
110 If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
120     Exit Function
130 End If

    'else everything is cool
140 LegalCharacter = True

    On Error GoTo 0
    Exit Function

LegalCharacter_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LegalCharacter de Módulo modUser línea: " & Erl())
End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidateUserLogin
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Validaciones de login de usuario.
'---------------------------------------------------------------------------------------
'
Function ValidateUserLogin(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

10  On Error GoTo ValidateUserLogin_Error

20  If tmpUser.UserName = "Usuario" Then
30      sErrorMsg = "Ingrese un nombre de usuario válido por favor."
40      Exit Function
50  ElseIf Len(tmpUser.UserName) < 3 Or Len(tmpUser.UserName) > 50 Then
60      sErrorMsg = "Ingrese un nombre de usuario válido por favor, el mismo debe tener al menos 3 caracteres y 50 como máximo."
70      Exit Function
80  ElseIf tmpUser.Password = "Contraseña" Then
90      sErrorMsg = "Ingrese una contraseña válida por favor."
100     Exit Function
110 ElseIf Len(tmpUser.Password) < 6 Or Len(tmpUser.Password) > 50 Then
120     sErrorMsg = "Ingrese una contraseña válida por favor, la misma debe tener al menos 6 caracteres y 50 como máximo."
130     Exit Function
140 End If

150 ValidateUserLogin = True

160 On Error GoTo 0
170 Exit Function

ValidateUserLogin_Error:

180 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateUserLogin de Módulo modUserCreate línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateUserCreate
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Validaciones en la creación de un usuario.
'---------------------------------------------------------------------------------------
'
Function ValidateUserCreate(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim tmpFields()                 As String
Dim tmpValues()                 As String

10  On Error GoTo ValidateUserCreate_Error

20  ReDim tmpFields(1 To 2) As String
30  ReDim tmpValues(1 To 2) As String
40  tmpFields(1) = "id_tipodocumento"
50  tmpFields(2) = "num_documento"

60  ReDim tmpValues(1 To 2) As String
70  ReDim tmpValues(1 To 2) As String
80  tmpValues(1) = tmpUser.Person.id_dni
90  tmpValues(2) = tmpUser.Person.dni

100 If tmpUser.UserName = "Usuario" Then
110     sErrorMsg = "Ingrese un nombre de usuario válido por favor."
120     Exit Function
130 ElseIf Not Len(tmpUser.UserName) > 3 Then
140     sErrorMsg = "El nombre de usuario debe contener al menos 3 letras."
150     Exit Function
160 ElseIf tmpUser.Person.Name = "Nombre y apellido" Then
170     sErrorMsg = "Ingrese un nombre y apellido válido por favor."
180     Exit Function
190 ElseIf Not Len(tmpUser.Person.Name) > 3 Then
200     sErrorMsg = "El nombre y apellido deben contener al menos 3 letras."
210     Exit Function
220 ElseIf CStr(tmpUser.Person.id_dni) = 0 Then
230     sErrorMsg = "Ingrese un tipo de documento por favor."
240 ElseIf CStr(tmpUser.Person.dni) = "DNI" Then
250     sErrorMsg = "Ingrese un DNI válido por favor."
260     Exit Function
270 ElseIf Not ValidateDNI(tmpUser.Person.dni) Then
280     sErrorMsg = "Ingrese un DNI válido por favor (que tenga entre 7 u 8 caracteres y sean sólo números)."
290     Exit Function
300 ElseIf ExistsArr("personas", tmpFields, tmpValues) Then
310     sErrorMsg = "El número y tipo de documentos seleccionados, ya están registrados. La persona ya existe en la base de datos."
320     Exit Function
330 ElseIf tmpUser.Password = "Contraseña" Then
340     sErrorMsg = "Ingrese una contraseña válida por favor."
350     Exit Function
360 ElseIf Not modPerson.ValidateEmail(tmpUser.Person.email) Then
370     sErrorMsg = "El e-mail tiene un formato incorrecto."
380     Exit Function
390 End If

400 If Not ValidatePassword(tmpUser.Password, sErrorMsg) Then
410     Exit Function
420 End If

430 ValidateUserCreate = True

440 On Error GoTo 0
450 Exit Function

ValidateUserCreate_Error:

460 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateUserCreate de Módulo modUser línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateDNI
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Utilicé CHAT GPT para crear esta función, con el propósito de validar la autenticidad del DNI.
'---------------------------------------------------------------------------------------
'
Public Function ValidateDNI(ByVal dni As String) As Boolean

' Verificar que no esté vacío
    On Error GoTo ValidateDNI_Error

    If InStrB(1, dni, ".") > 0 Then
        dni = Replace(dni, ".", "")
    End If

10  If Trim(dni) = "" Then
20      ValidateDNI = False
30      Exit Function
40  End If

    ' Verificar que todos los caracteres sean números
    Dim i                       As Integer
50  For i = 1 To Len(dni)
60      If Not IsNumeric(Mid(dni, i, 1)) Then
70          ValidateDNI = False
80          Exit Function
90      End If
100 Next i

    ' Verificar el rango de longitud (7 u 8 dígitos)
110 If Len(dni) < 7 Or Len(dni) > 8 Then
120     ValidateDNI = False
130     Exit Function
140 End If

    ' Validación exitosa
150 ValidateDNI = True
160 Exit Function

    On Error GoTo 0
    Exit Function

ValidateDNI_Error:
    ValidateDNI = False
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateDNI de Módulo modUser línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckTxtControlMouseDown
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Controlamos en los eventos MouseDown (click) el Texto que muestra el control para referenciar al usuario.
'---------------------------------------------------------------------------------------
'
Public Function CheckTxtControlMouseDown(ByRef txtControl As TextBox)

    On Error GoTo CheckTxtControlMouseDown_Error

    If txtControl.Text = "Usuario" Or txtControl.Text = "Nombre y apellido" Or txtControl.Text = "DNI" _
    Or txtControl.Text = "Contraseña" Or txtControl.Text = "E-mail" Or txtControl.Text = "Fecha nacimiento" Or txtControl.Text = "Código Postal" Then
        txtControl.Text = vbNullString
    End If

    On Error GoTo 0
    Exit Function

CheckTxtControlMouseDown_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckTxtControlMouseDown de Formulario frmUserCreate línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckTxtControlMouseUp
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Controlamos en los eventos MouseUp (soltar el click) el Texto que muestra el control para referenciar al usuario.
'---------------------------------------------------------------------------------------
'
Public Function CheckTxtControlMouseUp(ByRef txtControl As TextBox)

    On Error GoTo CheckTxtControlMouseUp_Error

    If StrComp(UCase$(txtControl.Text), vbNullString) = 0 Or StrComp(UCase$(txtControl.Text), " ") = 0 Then
        If txtControl.Name = "txtUserName" Then
            txtControl.Text = "Usuario"
        ElseIf txtControl.Name = "txtName" Then
            txtControl.Text = "Nombre y apellido"
        ElseIf txtControl.Name = "txtDNI" Then
            txtControl.Text = "DNI"
        ElseIf txtControl.Name = "txtPassword" Then
            txtControl.Text = "Contraseña"
        ElseIf txtControl.Name = "txtEmail" Then
            txtControl.Text = "E-mail"
        ElseIf txtControl.Name = "txtDateBirth" Then
            txtControl.MaxLength = 20
            txtControl.Text = "Fecha nacimiento"
            txtControl.MaxLength = 10
        ElseIf txtControl.Name = "txtZipCode" Then
            txtControl.Text = "Código Postal"
        End If
        Exit Function
    End If

    On Error GoTo 0
    Exit Function

CheckTxtControlMouseUp_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckTxtControlMouseUp de Formulario frmUserCreate línea: " & Erl())

End Function
