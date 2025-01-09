Attribute VB_Name = "modPerson"
Option Explicit

Public tmpUserEdit              As tUser

Function ValidatePersonCreate(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim tmpFields()                 As String
Dim tmpValues()                 As String

10  On Error GoTo ValidatePersonCreate_Error

20  ReDim tmpFields(1 To 2) As String
30  ReDim tmpValues(1 To 2) As String
40  tmpFields(1) = "id_tipodocumento"
50  tmpFields(2) = "num_documento"

60  ReDim tmpValues(1 To 2) As String
70  ReDim tmpValues(1 To 2) As String
80  tmpValues(1) = tmpUser.Person.id_dni
90  tmpValues(2) = tmpUser.Person.dni

100 If tmpUser.Person.Name = "Nombre y apellido" Then
110     sErrorMsg = "Ingrese un nombre y apellido válidos por favor."
120     Exit Function
130 ElseIf Not Len(tmpUser.Person.Name) > 3 Then
140     sErrorMsg = "El nombre y apellido deben contener al menos 3 letras."
150     Exit Function
160 ElseIf CStr(tmpUser.Person.id_dni) = 0 Then
170     sErrorMsg = "Ingrese un tipo de documento por favor."
180 ElseIf CStr(tmpUser.Person.dni) = "DNI" Then
190     sErrorMsg = "Ingrese un DNI válido por favor."
200     Exit Function
210 ElseIf Not ValidateDNI(tmpUser.Person.dni) Then
220     sErrorMsg = "Ingrese un DNI válido por favor (que tenga entre 7 u 8 caracteres y sean sólo números)."
230     Exit Function
240 ElseIf ExistsArr("personas", tmpFields, tmpValues) Then
250     sErrorMsg = "El número y tipo de documentos seleccionados, ya están registrados. La persona ya existe en la base de datos."
260     Exit Function
270 ElseIf Not ValidateDateBirth(tmpUser) Then
280     sErrorMsg = "La fecha de nacimiento parece ser inválida. Utilice el formato DD/MM/YYYY (Ej: 01/05/2001)"
290     Exit Function
300 ElseIf CStr(tmpUser.Person.id_state) < 0 Then
310     sErrorMsg = "Seleccione una provincia"
320     Exit Function
330 ElseIf CStr(tmpUser.Person.Id_Locality) < 0 Then
340     sErrorMsg = "Seleccione una localidad"
350     Exit Function
360 ElseIf Not modPerson.ValidateEmail(tmpUser.Person.email) Then
370     sErrorMsg = "El e-mail tiene un formato incorrecto."
380     Exit Function
390 End If

400 ValidatePersonCreate = True

410 On Error GoTo 0
420 Exit Function

ValidatePersonCreate_Error:

430 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidatePersonCreate de Módulo modPerson línea: " & Erl())

End Function

Function ValidatePersonEdit(ByRef tmpUser As tUser, ByRef sErrorMsg As String) As Boolean

Dim tmpFields()                 As String
Dim tmpValues()                 As String

10  On Error GoTo ValidatePersonEdit_Error

20  If tmpUser.Person.Name = "Nombre y apellido" Then
30      sErrorMsg = "Ingrese un nombre y apellido válidos por favor."
40      Exit Function
50  ElseIf Not Len(tmpUser.Person.Name) > 3 Then
60      sErrorMsg = "El nombre y apellido deben contener al menos 3 letras."
70      Exit Function
80  ElseIf CStr(tmpUser.Person.id_dni) = 0 Then
90      sErrorMsg = "Ingrese un tipo de documento por favor."
100 ElseIf CStr(tmpUser.Person.dni) = "DNI" Then
110     sErrorMsg = "Ingrese un DNI válido por favor."
120     Exit Function
130 ElseIf Not ValidateDNI(tmpUser.Person.dni) Then
140     sErrorMsg = "Ingrese un DNI válido por favor (que tenga entre 7 u 8 caracteres y sean sólo números)."
150     Exit Function
160 ElseIf Not ValidateDateBirth(tmpUser) Then
170     sErrorMsg = "La fecha de nacimiento parece ser inválida. Utilice el formato DD/MM/YYYY (Ej: 01/05/2001)"
180     Exit Function
190 ElseIf Not modPerson.ValidateEmail(tmpUser.Person.email) Then
200     sErrorMsg = "El e-mail tiene un formato incorrecto."
210     Exit Function
220 End If

230 ValidatePersonEdit = True

240 On Error GoTo 0
250 Exit Function

ValidatePersonEdit_Error:

260 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidatePersonEdit de Módulo modPerson línea: " & Erl())

End Function

Public Function ValidateDateBirth(ByRef tmpUser As tUser) As Boolean

Dim inputDate                   As String
Dim day As Integer, month As Integer, year As Integer
Dim dateValue                   As Date

    ' Obtener el Texto del TextBox
    On Error GoTo ValidateDateBirth_Error

10  inputDate = tmpUser.Person.DateBirth

    ' Asegurarse de que el campo no esté vacío
20  If Len(inputDate) = 0 Then
30      ValidateDateBirth = False
40      Exit Function
50  End If

    ' Verificar que el formato sea exactamente DD/MM/YYYY
60  If Len(inputDate) <> 10 Then
70      ValidateDateBirth = False
80      Exit Function
90  End If

    ' Verificar si las posiciones 3 y 6 son las barras (/)
100 If Mid(inputDate, 3, 1) <> "/" Or Mid(inputDate, 6, 1) <> "/" Then
110     ValidateDateBirth = False
120     Exit Function
130 End If

    ' Extraer el día, mes y año del Texto
140 day = Val(Mid(inputDate, 1, 2))
150 month = Val(Mid(inputDate, 4, 2))
160 year = Val(Mid(inputDate, 7, 4))    ' El año debe ser de 4 dígitos

    ' Verificar si el año es válido (es posible que quieras ajustarlo para un rango específico)
170 If year < 1900 Or year > 2100 Then
180     ValidateDateBirth = False
190     Exit Function
200 End If

    ' Verificar si la fecha es válida
210 On Error Resume Next
220 dateValue = DateSerial(year, month, day)
230 On Error GoTo 0

    ' Si ocurrió un error al crear la fecha, la fecha no es válida
240 If dateValue = 0 Then
250     ValidateDateBirth = False
260 Else
        ' Si no ocurrió ningún error, la fecha es válida
270     ValidateDateBirth = True
280 End If

    On Error GoTo 0
    Exit Function

ValidateDateBirth_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateDateBirth de Módulo modPerson línea: " & Erl())
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateEmail
' Author    : Brian Sabatier (https://github.com/brianirvana)
' Date      : 8/1/2025
' Purpose   : Renombro la función CheckMailString a ValidateEmail para reflejar mejor su funcionalidad. Implementa validación con expresiones regulares según lo
' solicitado por Franco.
'---------------------------------------------------------------------------------------
'
Function ValidateEmail(ByVal sEmail As String) As Boolean

Dim regex                       As Object

    On Error GoTo ValidateEmail_Error

10  Set regex = CreateObject("VBScript.RegExp")
20  regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
30  regex.IgnoreCase = True
40  regex.Global = False

    ' Verificar si el email cumple con el patrón
50  ValidateEmail = regex.Test(sEmail)

    ' Liberar el objeto RegExp
60  Set regex = Nothing

    On Error GoTo 0
    Exit Function

ValidateEmail_Error:
    ValidateEmail = False
    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ValidateEmail (" & sEmail & ") de Módulo modPerson línea: " & Erl())

End Function
