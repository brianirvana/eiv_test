VERSION 5.00
Begin VB.Form frmPersonAdd 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Persona"
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIsArgentine 
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   5400
      Width           =   195
   End
   Begin VB.ComboBox cmbGenre 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmPersonAdd.frx":0000
      Left            =   1680
      List            =   "frmPersonAdd.frx":000A
      TabIndex        =   7
      Text            =   "Genero"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ComboBox cmbLocality 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Text            =   "Localidad"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.ComboBox cmbState 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Text            =   "Provincia"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtDateBirth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "Fecha nacimiento"
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "E-mail"
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton cmdPersonCreate 
      Appearance      =   0  'Flat
      Caption         =   "Añadir persona"
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton cmdGoBack 
      Appearance      =   0  'Flat
      Caption         =   "Volver"
      Height          =   615
      Left            =   1560
      TabIndex        =   11
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox txtFirstName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Nombre"
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtLastName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "Apellido"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtDNI 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "DNI"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.ComboBox cmbIdDNIType 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Tipo DNI"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblIsArgentine 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¿Es argentino?"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   12
      Top             =   5880
      Width           =   5295
   End
End
Attribute VB_Name = "frmPersonAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    frmUserLogin.Show
    Unload Me
End Sub

Private Sub cmdPersonCreate_Click()

Dim sErrorMsg                   As String
Dim tmpUser                     As tUser

10  On Error GoTo cmdPersonCreate_Click_Error

20  tmpUser.Person.FirstName = txtFirstName.text
30  tmpUser.Person.LastName = txtLastName.text

40  If cmbIdDNIType.ListIndex = -1 Then
50      MsgBox "Por favor, debe seleccionar un tipo de documento."
60      Exit Sub
70  End If

80  tmpUser.Person.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
90  tmpUser.Person.dni = txtDNI.text
100 tmpUser.Person.DateBirth = txtDateBirth.text
110 tmpUser.Person.Email = txtEmail.text
120 tmpUser.Person.is_argentine = IIf(CBool(chkIsArgentine.Value), True, False)

130 If cmbLocality.ListIndex = -1 Then
140     MsgBox "Por favor, debe seleccionar la localidad de la persona."
150     Exit Sub
160 End If

170 tmpUser.Person.id_locality = cmbLocality.ItemData(cmbLocality.ListIndex)

180 If cmbState.ListIndex = -1 Then
190     MsgBox "Por favor, debe seleccionar la provincia de la persona."
200     Exit Sub
210 End If

220 tmpUser.Person.id_locality = cmbState.ItemData(cmbState.ListIndex)

230 If cmbGenre.ListIndex = -1 Then
240     MsgBox "Por favor, debe seleccionar el género de la persona."
250     Exit Sub
260 End If

270 tmpUser.Person.Genre = cmbGenre.text
280 tmpUser.Person.zip_code = GetZipCodeFromLocality(tmpUser, sErrorMsg)
290 If tmpUser.Person.zip_code <= 0 Then
300     MsgBox sErrorMsg
310     Exit Sub
320 End If

330 If Not ValidatePersonCreate(tmpUser, sErrorMsg) Then
340     Call MsgBox(sErrorMsg, vbInformation, "Error, por favor revise la información ingresada.")
350 Else
360     If modDBPersons.PersonCreate(tmpUser, sErrorMsg) Then
370         Call MsgBox("Nueva persona añadida con éxito.", vbInformation, "¡Exito!")
380         frmUserLogin.Show
390         Me.Hide
400     Else
410         Call MsgBox(sErrorMsg, vbInformation, "Error al añadir la persona.")
420     End If
430 End If

440 On Error GoTo 0
450 Exit Sub

cmdPersonCreate_Click_Error:

460 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdCreatePerson_Click de Formulario frmPersonAdd línea: " & Erl())

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmUserLogin.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Call modDBPersons.LoadDNITypes(Me)
    Call modDBPersons.LoadStates(Me)
    Call modDBPersons.LoadLocality(Me)
    
    Me.cmbGenre.Clear
    Me.cmbGenre.AddItem "M"
    Me.cmbGenre.AddItem "F"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(cmbIdDNIType.text) > 0 And Not txtDNI.Enabled Then
        txtDNI.Enabled = True
    End If
End Sub

Private Sub txtDateBirth_KeyPress(KeyAscii As Integer)
    ' Permitir solo números y el Backspace (código 8)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDateBirth_Change()
    Dim text As String
    Dim formattedText As String
    
    ' Obtener el texto del TextBox
    text = txtDateBirth.text
    
    ' Eliminar cualquier carácter no numérico (en caso de que el usuario pegue texto)
    text = Replace(text, "/", "")
    
    ' Si el texto tiene más de 6 caracteres, no hacer nada para evitar desbordar el formato
    If Len(text) > 6 Then Exit Sub
    
    ' Aplicar el formato DD/MM/YY mientras el usuario escribe
    If Len(text) > 4 Then
        formattedText = Mid(text, 1, 2) & "/" & Mid(text, 3, 2) & "/" & Mid(text, 5, 2)
    ElseIf Len(text) > 2 Then
        formattedText = Mid(text, 1, 2) & "/" & Mid(text, 3, 2)
    Else
        formattedText = text
    End If
    
    ' Establecer el texto formateado en el TextBox
    txtDateBirth.text = formattedText
    
    ' Posicionar el cursor al final del texto
    txtDateBirth.SelStart = Len(formattedText)
End Sub

Private Sub txtDNI_Change()
    If Not modUser.ValidateDNI(txtDNI.text) Then
        lblInfo.Caption = "El DNI es inválido al parecer."
    End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDateBirth_LostFocus()
    Call CheckTxtControlMouseUp(txtDateBirth)
End Sub

Private Sub txtDNI_LostFocus()
    Call CheckTxtControlMouseUp(txtDNI)
End Sub

Private Sub txtEmail_LostFocus()
    Call CheckTxtControlMouseUp(txtEmail)
End Sub

Private Sub txtFirstName_LostFocus()
    Call CheckTxtControlMouseUp(txtFirstName)
End Sub

Private Sub txtLastName_LostFocus()
    Call CheckTxtControlMouseUp(txtLastName)
End Sub

Private Sub txtLastName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtLastName)
End Sub

Private Sub txtFirstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtFirstName)
End Sub

Private Sub txtEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtEmail)
End Sub

Private Sub txtDNI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtDNI)
End Sub

Private Sub txtDateBirth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtDateBirth)
End Sub
