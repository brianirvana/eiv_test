Attribute VB_Name = "modUserCreate"
Option Explicit

'Public aUser                    As tUser

Private Type tUser
    UserName                    As String
    FirstName                   As String
    LastName                    As String
    Email                       As String
    Password                    As String * 200
    DNI                         As Long
End Type

'---------------------------------------------------------------------------------------
' Procedure : NumberToPunctuatedString
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NumberToPunctuatedString(ByVal Cantidad As Double) As String

Dim i                           As Double
Dim CantidadStr                 As String
Dim FinalStr                    As String

   On Error GoTo NumberToPunctuatedString_Error

10    CantidadStr = CStr(Cantidad)

20    For i = Len(CantidadStr) To 1 Step -3
30      FinalStr = Right$(Left$(CantidadStr, i), 3) & "." & FinalStr
40    Next i

50    StringOro = Left$(FinalStr, Len(FinalStr) - 1)

   On Error GoTo 0
   Exit Function

NumberToPunctuatedString_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento NumberToPunctuatedString de Módulo modUserCreate línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateUserLogin
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Validaciones de login de usuario.
'---------------------------------------------------------------------------------------
'
Function ValidateUserLogin(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

    On Error GoTo ValidateUserLogin_Error

10  If tmpUser.UserName = "Usuario" Then
20      sErrorMsg = "Ingrese un nombre de usuario válido por favor."
30      If tmpUser.Password = "Contraseña" Then
40          sErrorMsg = "Ingrese una contraseña válida por favor."
50      End If

60      CreateUserValidate = True

        On Error GoTo 0
        Exit Function

ValidateUserLogin_Error:

        Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateUserLogin de Módulo modUserCreate línea: " & Erl())

    End Function
