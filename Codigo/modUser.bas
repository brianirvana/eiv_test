Attribute VB_Name = "modUser"
Option Explicit

'Public aUser                    As tUser

Public Type tUser
    UserName                    As String
    FirstName                   As String
    LastName                    As String
    Email                       As String
    Password                    As String
    HashedPwd                   As String
    dni                         As Long
    id_dni                      As Long
End Type

'---------------------------------------------------------------------------------------
' Procedure : NumberToPunctuatedString
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Agregamos puntitos en los miles del DNI para mejorar su visualizaci�n.
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

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento NumberToPunctuatedString de M�dulo modUserCreate l�nea: " & Erl())

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

    ' Verificamos la longitud de la contrase�a
    If Len(UserPassword) < 6 Or Len(UserPassword) > 32 Then
        sErrorMsg = "La contrase�a debe tener entre 6 y 32 caracteres."
        ValidatePassword = False
        Exit Function
    End If

    ' Recorremos cada car�cter de la contrase�a
    For loopc = 1 To Len(UserPassword)
        CharAscii = Mid$(UserPassword, loopc, 1)
        tmpIsNumber = Asc(CharAscii)

        ' Verificamos si el car�cter es v�lido (por ejemplo, letras o n�meros)
        If Not LegalCharacter(tmpIsNumber) Then
            sErrorMsg = "Password inv�lido. El car�cter '" & CharAscii & "' no est� permitido."
            ValidatePassword = False
            Exit Function
        ElseIf IsNumeric(CharAscii) Then
            ValidNumeric = True
        ElseIf tmpIsNumber >= 65 And tmpIsNumber <= 90 Or tmpIsNumber >= 97 And tmpIsNumber <= 122 Then
            ' Consideramos letras (may�sculas y min�sculas)
            ValidLetters = True
        End If
    Next loopc

    ' Verificamos si cumple con los requisitos de contenido
    If Not ValidLetters Or Not ValidNumeric Then
        sErrorMsg = "La contrase�a debe contener al menos un n�mero y al menos una letra."
        ValidatePassword = False
        Exit Function
    End If

    ' Si pasa todas las validaciones
    ValidatePassword = True
    Exit Function

ValidatePassword_Error:
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidatePassword de M�dulo modUser l�nea: " & Erl())
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

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento LegalCharacter de M�dulo modUser l�nea: " & Erl())
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
30      sErrorMsg = "Ingrese un nombre de usuario v�lido por favor."
40      Exit Function
50  ElseIf Len(tmpUser.UserName) < 3 Or Len(tmpUser.UserName) > 50 Then
60      sErrorMsg = "Ingrese un nombre de usuario v�lido por favor, el mismo debe tener al menos 3 caracteres y 50 como m�ximo."
70      Exit Function
80  ElseIf tmpUser.Password = "Contrase�a" Then
90      sErrorMsg = "Ingrese una contrase�a v�lida por favor."
100     Exit Function
110 ElseIf Len(tmpUser.Password) < 6 Or Len(tmpUser.Password) > 50 Then
120     sErrorMsg = "Ingrese una contrase�a v�lida por favor, la misma debe tener al menos 6 caracteres y 50 como m�ximo."
130     Exit Function
140 End If

150 ValidateUserLogin = True

160 On Error GoTo 0
170 Exit Function

ValidateUserLogin_Error:

180 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateUserLogin de M�dulo modUserCreate l�nea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateUserCreate
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Validaciones en la creaci�n de un usuario.
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
80  tmpValues(1) = tmpUser.id_dni
90  tmpValues(2) = tmpUser.dni

100 If tmpUser.UserName = "Usuario" Then
110     sErrorMsg = "Ingrese un nombre de usuario v�lido por favor."
120     Exit Function
130 ElseIf Not Len(tmpUser.UserName) > 3 Then
140     sErrorMsg = "El nombre de usuario debe contener al menos 3 letras."
150     Exit Function
160 ElseIf tmpUser.FirstName = "Nombre" Then
170     sErrorMsg = "Ingrese un nombre personal v�lido por favor."
180     Exit Function
190 ElseIf Not Len(tmpUser.FirstName) > 3 Then
200     sErrorMsg = "El nombre personal debe contener al menos 3 letras."
210     Exit Function
220 ElseIf tmpUser.LastName = "Apellido" Then
230     sErrorMsg = "Ingrese un apellido v�lido por favor."
240     Exit Function
250 ElseIf CStr(tmpUser.id_dni) = 0 Then
260     sErrorMsg = "Ingrese un tipo de documento por favor."
270 ElseIf CStr(tmpUser.dni) = "DNI" Then
280     sErrorMsg = "Ingrese un DNI v�lido por favor."
290     Exit Function
300 ElseIf Not ValidateDNI(tmpUser.dni) Then
310     sErrorMsg = "Ingrese un DNI v�lido por favor (que tenga entre 7 u 8 caracteres y sean s�lo n�meros)."
320     Exit Function
330 ElseIf ExistsArr("personas", tmpFields, tmpValues) Then
340     sErrorMsg = "El n�mero y tipo de documentos seleccionados, ya est�n registrados. La persona ya existe en la base de datos."
350     Exit Function
360 ElseIf tmpUser.Password = "Contrase�a" Then
370     sErrorMsg = "Ingrese una contrase�a v�lida por favor."
380     Exit Function
390 End If

400 If Not ValidatePassword(tmpUser.Password, sErrorMsg) Then
410     Exit Function
420 End If

430 ValidateUserCreate = True

440 On Error GoTo 0
450 Exit Function

ValidateUserCreate_Error:

460 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateUserCreate de M�dulo modUser l�nea: " & Erl())

End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidateDNI
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Utilic� CHAT GPT para crear esta funci�n, con el prop�sito de validar la autenticidad del DNI.
'---------------------------------------------------------------------------------------
'
Public Function ValidateDNI(ByVal dni As String) As Boolean

' Verificar que no est� vac�o
    On Error GoTo ValidateDNI_Error

10  If Trim(dni) = "" Then
20      ValidateDNI = False
30      Exit Function
40  End If

    ' Verificar que todos los caracteres sean n�meros
    Dim i                       As Integer
50  For i = 1 To Len(dni)
60      If Not IsNumeric(Mid(dni, i, 1)) Then
70          ValidateDNI = False
80          Exit Function
90      End If
100 Next i

    ' Verificar el rango de longitud (7 u 8 d�gitos)
110 If Len(dni) < 7 Or Len(dni) > 8 Then
120     ValidateDNI = False
130     Exit Function
140 End If

    ' Validaci�n exitosa
150 ValidateDNI = True
160 Exit Function

    On Error GoTo 0
    Exit Function

ValidateDNI_Error:
    ValidateDNI = False
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateDNI de M�dulo modUser l�nea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckTxtControlMouseDown
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Controlamos en los eventos MouseDown (click) el texto que muestra el control para referenciar al usuario.
'---------------------------------------------------------------------------------------
'
Public Function CheckTxtControlMouseDown(ByRef txtControl As TextBox)

    On Error GoTo CheckTxtControlMouseDown_Error

    If txtControl.Text = "Usuario" Or txtControl.Text = "Nombre" Or txtControl.Text = "Apellido" Or txtControl.Text = "DNI" Or txtControl.Text = "Contrase�a" Or txtControl.Text = "E-mail" Then
        txtControl.Text = vbNullString
    End If

    On Error GoTo 0
    Exit Function

CheckTxtControlMouseDown_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckTxtControlMouseDown de Formulario frmUserCreate l�nea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckTxtControlMouseUp
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Controlamos en los eventos MouseUp (soltar el click) el texto que muestra el control para referenciar al usuario.
'---------------------------------------------------------------------------------------
'
Public Function CheckTxtControlMouseUp(ByRef txtControl As TextBox)

    On Error GoTo CheckTxtControlMouseUp_Error

    If StrComp(UCase$(txtControl.Text), vbNullString) = 0 Or StrComp(UCase$(txtControl.Text), " ") = 0 Then
        If txtControl.Name = "txtUserName" Then
            txtControl.Text = "Usuario"
        ElseIf txtControl.Name = "txtFirstName" Then
            txtControl.Text = "Nombre"
        ElseIf txtControl.Name = "txtLastName" Then
            txtControl.Text = "Apellido"
        ElseIf txtControl.Name = "txtDNI" Then
            txtControl.Text = "DNI"
        ElseIf txtControl.Name = "txtPassword" Then
            txtControl.Text = "Contrase�a"
        ElseIf txtControl.Name = "txtEmail" Then
            txtControl.Text = "E-mail"
        End If
        Exit Function
    End If

    On Error GoTo 0
    Exit Function

CheckTxtControlMouseUp_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckTxtControlMouseUp de Formulario frmUserCreate l�nea: " & Erl())

End Function
