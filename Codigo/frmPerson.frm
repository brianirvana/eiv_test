VERSION 5.00
Begin VB.Form frmPerson 
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIsArgentine 
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   5400
      Width           =   195
   End
   Begin VB.ComboBox cmbGenre 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmPerson.frx":0000
      Left            =   1680
      List            =   "frmPerson.frx":000A
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   3
      Text            =   "Provincia"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
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
      TabIndex        =   5
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
      TabIndex        =   7
      Text            =   "E-mail"
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton cmdPersonAction 
      Appearance      =   0  'Flat
      Caption         =   "Añadir persona"
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmdGoBack 
      Appearance      =   0  'Flat
      Caption         =   "Volver"
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Nombre y apellido"
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   13
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
      TabIndex        =   11
      Top             =   5880
      Width           =   5295
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eTypeMode
    None = 0
    PersonCreate = 1
    PersonEdit = 2
End Enum

Public TypeMode                 As eTypeMode

Private Sub cmbState_LostFocus()
    Call modDBPersons.LoadLocality(Me)
End Sub

Private Sub cmdClose_Click()
    frmUserLogin.Show
    Unload Me
End Sub

Private Sub cmdGoBack_Click()
    frmAbmPersons.Show
    Unload Me
End Sub

Private Sub cmdPersonAction_Click()

Dim sErrorMsg                   As String

10  On Error GoTo cmdPersonAction_Click_Error

20  If cmbIdDNIType.ListIndex = -1 Then
30      MsgBox "Por favor, debe seleccionar un tipo de documento."
40      Exit Sub
50  End If

60  If cmbLocality.ListIndex = -1 Then
70      MsgBox "Por favor, debe seleccionar la localidad de la persona."
80      Exit Sub
90  End If

100 tmpUserEdit.Person.id_locality = cmbLocality.ItemData(cmbLocality.ListIndex)

110 If cmbState.ListIndex = -1 Then
120     MsgBox "Por favor, debe seleccionar la provincia de la persona."
130     Exit Sub
140 End If

150 tmpUserEdit.Person.id_locality = cmbState.ItemData(cmbState.ListIndex)
151 tmpUserEdit.Person.zip_code = GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)

160 If cmbGenre.ListIndex = -1 Then
170     MsgBox "Por favor, debe seleccionar el género de la persona."
180     Exit Sub
190 End If

200 If tmpUserEdit.Person.zip_code <= 0 Then
210     MsgBox "Por favor, debe ingresar el código postal."
220     Exit Sub
230 End If

240 Select Case TypeMode

        Case eTypeMode.None
            'Do nothing
            
250     Case eTypeMode.PersonCreate
            Dim tmpUser         As tUser
260         tmpUser.Person.Name = txtName.Text
270         tmpUser.Person.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
280         tmpUser.Person.dni = txtDNI.Text
290         tmpUser.Person.DateBirth = txtDateBirth.Text
300         tmpUser.Person.Email = txtEmail.Text
310         tmpUser.Person.is_argentine = IIf(CBool(chkIsArgentine.Value), True, False)
320         tmpUser.Person.Genre = cmbGenre.Text
330         tmpUser.Person.zip_code = GetZipCodeFromLocality(tmpUser, sErrorMsg)
            tmpUser.Person.id_state = cmbState.ItemData(cmbState.ListIndex)
            tmpUser.Person.id_locality = cmbLocality.ItemData(cmbLocality.ListIndex)
531         tmpUser.Person.zip_code = GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)

340         If Not ValidatePersonCreate(tmpUser, sErrorMsg) Then
350             Call MsgBox(sErrorMsg, vbInformation, "Error, por favor revise la información ingresada.")
360         Else
370             If modDBPersons.PersonCreate(tmpUser, sErrorMsg) Then
380                 Call MsgBox("Nueva persona: '" & tmpUser.Person.Name & "' añadida con éxito.", vbInformation, "¡Exito!")
390                 frmAbmPersons.Show
400                 Unload Me
410             Else
420                 Call MsgBox(sErrorMsg, vbInformation, "Error al añadir la persona.")
430             End If
440         End If

450     Case eTypeMode.PersonEdit

460         tmpUserEdit.Person.Name = txtName.Text
470         tmpUserEdit.Person.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
480         tmpUserEdit.Person.dni = txtDNI.Text
490         tmpUserEdit.Person.DateBirth = txtDateBirth.Text
500         tmpUserEdit.Person.Email = txtEmail.Text
510         tmpUserEdit.Person.is_argentine = IIf(CBool(chkIsArgentine.Value), True, False)
520         tmpUserEdit.Person.Genre = cmbGenre.Text

            tmpUserEdit.Person.id_state = cmbState.ItemData(cmbState.ListIndex)
            tmpUserEdit.Person.id_locality = cmbLocality.ItemData(cmbLocality.ListIndex)
530         tmpUserEdit.Person.zip_code = GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)
            
540         If Not ValidatePersonEdit(tmpUserEdit, sErrorMsg) Then
550             Call MsgBox(sErrorMsg, vbInformation, "Error, revise la información ingresada.")
560         Else
570             If modDBPersons.PersonEdit(tmpUserEdit) Then
580                 Call MsgBox("Los datos de " & tmpUserEdit.Person.Name & " fueron editados con éxito.", vbInformation, "¡Exito!")
590                 frmAbmPersons.Show
600                 Me.Hide
610             Else
620                 Call MsgBox(sErrorMsg, vbInformation, "Error al editar la persona.")
630             End If
640         End If
650 End Select

161 Call frmAbmPersons.FormatGrid
171 Call modDBPersons.LoadPersons

660 On Error GoTo 0
670 Exit Sub

cmdPersonAction_Click_Error:

680 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdCreatePerson_Click de Formulario frmPerson línea: " & Erl())

End Sub

Private Sub Form_Activate()
    Call LoadTypeModeForm
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
    'Call modDBPersons.LoadLocality(Me)
    
    Me.cmbGenre.Clear
    Me.cmbGenre.AddItem "M"
    Me.cmbGenre.AddItem "F"

    Call LoadTypeModeForm

End Sub

Private Sub LoadTypeModeForm()

    Select Case TypeMode
        Case eTypeMode.None
            'Do nothing
        Case eTypeMode.PersonCreate
            cmdPersonAction.Caption = "Añadir persona"
            
        Case eTypeMode.PersonEdit
            cmdPersonAction.Caption = "Editar persona"
    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Len(cmbIdDNIType.Text) > 0 And Not txtDNI.Enabled Then
        txtDNI.Enabled = True
    End If
    
    Static LastLocality As Long
    
'    If Len(cmbState.Text) > 0 Or LastLocality <> 0 Then
'        'Call modDBPersons.LoadLocality(Me)
'        If frmPerson.cmbLocality.ListIndex >= 0 Then
'            LastLocality = frmPerson.cmbLocality.ItemData(frmPerson.cmbLocality.ListIndex)
'        End If
'    End If
    
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

Dim Text                        As String
Dim formattedText               As String

    ' Obtener el Texto del TextBox
    On Error GoTo txtDateBirth_Change_Error

10  Text = txtDateBirth.Text

    ' Eliminar cualquier carácter no numérico (en caso de que el usuario pegue Texto)
20  Text = Replace(Text, "/", "")

    ' Si el Texto tiene más de 6 caracteres, no hacer nada para evitar desbordar el formato
30  If Len(Text) > 6 Then Exit Sub

    ' Aplicar el formato DD/MM/YY mientras el usuario escribe
40  If Len(Text) > 4 Then
50      formattedText = Mid(Text, 1, 2) & "/" & Mid(Text, 3, 2) & "/" & Mid(Text, 5, 2)
60  ElseIf Len(Text) > 2 Then
70      formattedText = Mid(Text, 1, 2) & "/" & Mid(Text, 3, 2)
80  Else
90      formattedText = Text
100 End If

    ' Establecer el Texto formateado en el TextBox
110 txtDateBirth.Text = formattedText

    ' Posicionar el cursor al final del Texto
120 txtDateBirth.SelStart = Len(formattedText)

    On Error GoTo 0
    Exit Sub

txtDateBirth_Change_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento txtDateBirth_Change de Formulario frmPerson línea: " & Erl())
End Sub

Private Sub txtDNI_Change()
    If Not modUser.ValidateDNI(txtDNI.Text) Then
        lblInfo.Caption = "El DNI es inválido al parecer."
    End If
    
    If Len(txtDNI.Text) > 3 Then
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

Private Sub txtDateBirth_LostFocus()
    Call CheckTxtControlMouseUp(txtDateBirth)
End Sub

Private Sub txtDNI_LostFocus()
    Call CheckTxtControlMouseUp(txtDNI)
End Sub

Private Sub txtEmail_LostFocus()
    Call CheckTxtControlMouseUp(txtEmail)
End Sub

Private Sub txtName_LostFocus()
    Call CheckTxtControlMouseUp(txtName)
End Sub

Private Sub txtName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtControlMouseDown(txtName)
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
