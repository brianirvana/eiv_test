VERSION 5.00
Begin VB.Form frmUserCreate 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Crear usuario"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbIdDNIType 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "cmbIdTipodocumento"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtDNI 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "DNI"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtLastName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Apellido"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtFirstName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Nombre"
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdGoBack 
      Appearance      =   0  'Flat
      Caption         =   "Volver"
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton cmdCreateUser 
      Appearance      =   0  'Flat
      Caption         =   "Crear usuario"
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "E-mail"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "Contraseña"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Usuario"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   5295
   End
End
Attribute VB_Name = "frmUserCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbIdDNIType_GotFocus()
    If Len(cmbIdDNIType.Text) > 0 Then
        txtDNI.Enabled = True
    End If
End Sub

Private Sub cmbIdDNIType_LostFocus()
    If Len(cmbIdDNIType.Text) > 0 Then
        txtDNI.Enabled = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCreateUser_Click
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCreateUser_Click()

Dim sErrorMsg                   As String
Dim tmpUser                     As tUser

10  On Error GoTo cmdCreateUser_Click_Error

20  tmpUser.UserName = txtUserName.Text
30  tmpUser.FirstName = txtFirstName.Text
40  tmpUser.LastName = txtLastName.Text

50  If cmbIdDNIType.ListIndex = -1 Then
60      MsgBox "Por favor, debe seleccionar un tipo de documento."
70      Exit Sub
80  End If

90  tmpUser.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
100 tmpUser.dni = txtDNI.Text
110 tmpUser.Email = txtEmail.Text
120 tmpUser.Password = txtPassword.Text

130 If Not ValidateUserCreate(tmpUser, sErrorMsg) Then
140     Call MsgBox(sErrorMsg, vbInformation, "Error, por favor revise la información ingresada.")
150 Else
160     If modDBUser.CreateUser(tmpUser, sErrorMsg) Then
170         Call MsgBox("Nuevo usuario del sistema creado con éxito.", vbInformation, "¡Exito!")
180         frmUserLogin.Show
190         Me.Hide
200     Else
210         Call MsgBox(sErrorMsg, vbInformation, "Error al crear el usuario.")
220     End If
230 End If

240 On Error GoTo 0
250 Exit Sub

cmdCreateUser_Click_Error:

260 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdCreateUser_Click de Formulario frmUserCreate línea: " & Erl())

End Sub

Private Sub cmdGoBack_Click()
    frmUserLogin.Show
    Me.Hide
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim sQuery                      As String
Dim RS                          As ADODB.Recordset
Dim cmbIndex                    As Long

    ' Consulta para obtener los datos de tipos_documentos
    sQuery = "SELECT id_tipodocumento, nombre, abreviatura FROM tipos_documentos"

    ' Ejecutar la consulta y abrir un Recordset
    Set RS = New ADODB.Recordset
    RS.Open sQuery, cn, adOpenForwardOnly, adLockReadOnly

    ' Limpiar el ComboBox antes de cargar datos
    cmbIdDNIType.Clear

    ' Validar si hay registros en el Recordset
    If Not RS.EOF Then
        Do While Not RS.EOF
            ' Agregar el ítem al ComboBox con el formato "nombre - abreviatura"
            cmbIdDNIType.AddItem RS.Fields("nombre").Value & " - " & RS.Fields("abreviatura").Value

            ' Obtener el índice del ítem agregado
            cmbIndex = cmbIdDNIType.NewIndex

            ' Asignar el id_tipodocumento al ItemData del ítem actual
            cmbIdDNIType.ItemData(cmbIndex) = RS.Fields("id_tipodocumento").Value

            ' Avanzar al siguiente registro
            RS.MoveNext
        Loop
    End If

    ' Cerrar el Recordset
    RS.Close
    Set RS = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(cmbIdDNIType.Text) > 0 And Not txtDNI.Enabled Then
        txtDNI.Enabled = True
    End If
End Sub

Private Sub txtDNI_Change()
    If Len(txtDNI.Text) > 1000 Then
        txtDNI.Text = NumberToPunctuatedString(txtDNI.Text)
    End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEmail_Change()
    If Not CheckMailString(txtEmail.Text) Then
        lblInfo.Caption = "El e-mail parece ser inválido."
    Else
        lblInfo.Caption = ""
    End If
End Sub

Private Sub txtFirstName_Click()
    Call CheckTxtControlMouseDown(txtUserName)
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        ' Verificar si el carácter no es una letra (minúscula o mayúscula)
        If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtPassword_Change()
    If Len(txtPassword.Text) < 6 Then
        frmUserCreate.lblInfo.Caption = "La contraseña debe tener como mínimo 6 caracteres, máximo 32, debe contener al menos una letra y un número."
    Else
        frmUserCreate.lblInfo.Caption = ""
    End If
End Sub

Private Sub txtDNI_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtEmail_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtFirstName_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        ' Verificar si el carácter no es una letra (minúscula o mayúscula)
        If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtUserName_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtPassword_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtLastName_LostFocus()
    Call CheckTxtControlMouseUp(txtUserName)
End Sub

Private Sub txtLastName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtLastName)
End Sub

Private Sub txtFirstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtFirstName)
End Sub

Private Sub txtPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtPassword)
End Sub

Private Sub txtUserName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtUserName)
End Sub

Private Sub txtEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtEmail)
End Sub

Private Sub txtDNI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtDNI)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckMailString
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CheckMailString(ByRef sString As String) As Boolean

Dim lPos                        As Long
Dim lX                          As Long
Dim iAsc                        As Integer

    On Error GoTo CheckMailString_Error

10  lPos = InStr(sString, "@")    '1er test: Busca un simbolo @
20  If (lPos <> 0) Then

30      If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function    '2do test: Busca un simbolo . después de @ + 1

40      For lX = 0 To Len(sString) - 1    '3er test: Recorre todos los caracteres y los valída
50          If Not (lX = (lPos - 1)) Then    'No chequeamos la '@'
60              iAsc = Asc(Mid$(sString, (lX + 1), 1))
70              If Not CMSValidateChar(iAsc) Then Exit Function
80          End If
90      Next lX

100     CheckMailString = True    'Finale
110 End If

    On Error GoTo 0
    Exit Function

CheckMailString_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckMailString de Formulario frmUserCreate línea: " & Erl())

End Function

Private Function CMSValidateChar(ByVal iAsc As Integer) As Boolean
    CMSValidateChar = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function
