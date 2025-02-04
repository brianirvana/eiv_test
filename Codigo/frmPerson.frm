VERSION 5.00
Begin VB.Form frmPerson 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "A�adir Persona"
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
   Begin VB.TextBox txtZipCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "C�digo Postal"
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CheckBox chkIsArgentine 
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   5640
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
      TabIndex        =   7
      Text            =   "Genero"
      Top             =   4800
      Width           =   2895
   End
   Begin VB.ComboBox cmbLocality 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmPerson.frx":0014
      Left            =   1680
      List            =   "frmPerson.frx":0016
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
      ItemData        =   "frmPerson.frx":0018
      Left            =   1680
      List            =   "frmPerson.frx":001A
      TabIndex        =   3
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
      Top             =   4320
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
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdPersonAction 
      Appearance      =   0  'Flat
      Caption         =   "A�adir persona"
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   6960
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
      MaxLength       =   11
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
      Caption         =   "�Es argentino?"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5640
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
      Top             =   6000
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

Private bIsLoaded               As Boolean

Public TypeMode                 As eTypeMode

Private Sub cmbGenre_GotFocus()
    lblInfo.Caption = "Ingrese el g�nero de la persona (F/M)"
End Sub

Private Sub cmbIdDNIType_GotFocus()

    lblInfo.Caption = "Ingrese el tipo de documento"

End Sub

Private Sub cmbLocality_Change()
    'Call AutoCompleteZipCode
End Sub

Private Sub cmbLocality_Click()
    Call AutoCompleteZipCode
End Sub

Private Sub cmbLocality_GotFocus()
    If cmbLocality.Text = "Localidad" Then
        lblInfo.Caption = "Seleccione una localidad."
    Else
        lblInfo.Caption = ""
    End If
End Sub

Private Sub cmbLocality_KeyPress(KeyAscii As Integer)
    Call AutoCompleteZipCode
End Sub

Private Sub AutoCompleteZipCode()

    If cmbLocality.ListIndex > -1 Then
        Dim sErrorMsg As String
        
        txtZipCode.Text = GetZipCodeFromLocality(cmbLocality.ItemData(cmbLocality.ListIndex), sErrorMsg)
        
        If Len(sErrorMsg) > 0 Then
            MsgBox sErrorMsg
        End If
    End If

End Sub

Private Sub cmbLocality_LostFocus()
    Call AutoCompleteZipCode
End Sub

Private Sub cmbState_GotFocus()
    If cmbState.Text = "Provincia" Then
        lblInfo.Caption = "Seleccione una provincia."
    Else
        lblInfo.Caption = ""
    End If
End Sub

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

Private Sub cmdGoBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.Caption = "Volver al listado de personas"
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

100 tmpUserEdit.Person.Id_Locality = cmbLocality.ItemData(cmbLocality.ListIndex)

110 If cmbState.ListIndex = -1 Then
120     MsgBox "Por favor, debe seleccionar la provincia de la persona."
130     Exit Sub
140 End If

150 tmpUserEdit.Person.Id_Locality = cmbState.ItemData(cmbState.ListIndex)
    'tmpUserEdit.Person.zip_code = GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)

160 If cmbGenre.ListIndex = -1 Then
170     MsgBox "Por favor, debe seleccionar el g�nero de la persona."
180     Exit Sub
190 End If

    '     If tmpUserEdit.Person.zip_code <= 0 Then
    '         MsgBox "Por favor, debe ingresar el c�digo postal."
    '         Exit Sub
    '     End If

200 Select Case TypeMode

        Case eTypeMode.None
            'Do nothing

210     Case eTypeMode.PersonCreate
            Dim tmpUser         As tUser
220         tmpUser.Person.Name = txtName.Text
230         tmpUser.Person.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
240         tmpUser.Person.dni = txtDNI.Text
250         tmpUser.Person.DateBirth = txtDateBirth.Text
260         tmpUser.Person.email = txtEmail.Text
270         tmpUser.Person.is_argentine = IIf(CBool(chkIsArgentine.Value), True, False)
280         tmpUser.Person.Genre = cmbGenre.Text

290         tmpUser.Person.id_state = cmbState.ItemData(cmbState.ListIndex)
300         tmpUser.Person.Id_Locality = cmbLocality.ItemData(cmbLocality.ListIndex)

            'Dejo como opcional el c�digo postal, sin la necesidad de precargarlo en base a la localidad. Requerimiento Tarea - 001
310         tmpUser.Person.zip_code = txtZipCode.Text  'GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)

320         If Not ValidatePersonCreate(tmpUser, sErrorMsg) Then
330             Call MsgBox(sErrorMsg, vbInformation, "Error, por favor revise la informaci�n ingresada.")
340         Else
350             If modDBPersons.PersonCreate(tmpUser, sErrorMsg) Then
360                 Call MsgBox("Nueva persona: '" & tmpUser.Person.Name & "' a�adida con �xito.", vbInformation, "�Exito!")
370                 frmAbmPersons.Show
380                 Unload Me
390             Else
400                 Call MsgBox(sErrorMsg, vbInformation, "Error al a�adir la persona.")
410             End If
420         End If

430     Case eTypeMode.PersonEdit

440         tmpUserEdit.Person.Name = txtName.Text
450         tmpUserEdit.Person.id_dni = cmbIdDNIType.ItemData(cmbIdDNIType.ListIndex)
460         tmpUserEdit.Person.dni = txtDNI.Text
470         tmpUserEdit.Person.DateBirth = txtDateBirth.Text
480         tmpUserEdit.Person.email = txtEmail.Text
490         tmpUserEdit.Person.is_argentine = IIf(CBool(chkIsArgentine.Value), True, False)
500         tmpUserEdit.Person.Genre = cmbGenre.Text

510         tmpUserEdit.Person.id_state = cmbState.ItemData(cmbState.ListIndex)
520         tmpUserEdit.Person.Id_Locality = cmbLocality.ItemData(cmbLocality.ListIndex)
530         tmpUserEdit.Person.zip_code = txtZipCode.Text  'GetZipCodeFromLocality(tmpUserEdit, sErrorMsg)

540         If Not ValidatePersonEdit(tmpUserEdit, sErrorMsg) Then
550             Call MsgBox(sErrorMsg, vbInformation, "Error, revise la informaci�n ingresada.")
560         Else
570             If modDBPersons.PersonEdit(tmpUserEdit) Then
580                 Call MsgBox("Los datos de " & tmpUserEdit.Person.Name & " fueron editados con �xito.", vbInformation, "�Exito!")
590                 frmAbmPersons.Show
600                 Me.Hide
610             Else
620                 Call MsgBox(sErrorMsg, vbInformation, "Error al editar la persona.")
630             End If
640         End If
650 End Select

660 Call frmAbmPersons.FormatGrid
670 Call modDBPersons.LoadPersons

680 On Error GoTo 0
690 Exit Sub

cmdPersonAction_Click_Error:

700 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdCreatePerson_Click de Formulario frmPerson l�nea: " & Erl())

End Sub

Private Sub cmdPersonAction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case TypeMode
        Case eTypeMode.PersonCreate
            lblInfo.Caption = "A�adir nueva persona"
        Case eTypeMode.PersonEdit
            lblInfo.Caption = "Editar persona"
    End Select

End Sub

Private Sub Form_Activate()
    Call LoadTypeModeForm
    bIsLoaded = True
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
            cmdPersonAction.Caption = "A�adir persona"
            
        Case eTypeMode.PersonEdit
            cmdPersonAction.Caption = "Editar persona"
    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Len(cmbIdDNIType.Text) > 0 And Not txtDNI.Enabled Then
        txtDNI.Enabled = True
    End If
    
    Static LastLocality As Long
    
    lblInfo.Caption = ""
    
'    If Len(cmbState.Text) > 0 Or LastLocality <> 0 Then
'        'Call modDBPersons.LoadLocality(Me)
'        If frmPerson.cmbLocality.ListIndex >= 0 Then
'            LastLocality = frmPerson.cmbLocality.ItemData(frmPerson.cmbLocality.ListIndex)
'        End If
'    End If
    
End Sub

Private Sub lblIsArgentine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.Caption = "Determina si la persona es de nacionalidad argentina"
End Sub

Private Sub txtDateBirth_GotFocus()
    Call CheckTxtControlMouseDown(txtDateBirth)
    lblInfo.Caption = "Ingrese la fecha de nacimiento de la persona"
End Sub

Private Sub txtDateBirth_KeyPress(KeyAscii As Integer)
    ' Permitir solo n�meros y el Backspace (c�digo 8)
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

    ' Eliminar cualquier car�cter no num�rico (en caso de que el usuario pegue Texto)
20  Text = Replace(Text, "/", "")

    ' Si el Texto tiene m�s de 6 caracteres, no hacer nada para evitar desbordar el formato
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

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento txtDateBirth_Change de Formulario frmPerson l�nea: " & Erl())
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtDNI_Change
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 9/1/2025
' Purpose   : Deshabilit� la funci�n NumberToPunctuatedString hasta pensar una mejor soluci�n.
'---------------------------------------------------------------------------------------
'
Private Sub txtDNI_Change()

Static isUpdating               As Boolean
Dim cursorPosition              As Integer

    On Error GoTo txtDNI_Change_Error

10  If isUpdating Then Exit Sub
20  isUpdating = True

30  cursorPosition = txtDNI.SelStart

    ' Validar el DNI
40  If Not modUser.ValidateDNI(txtDNI.Text) Then
50      lblInfo.Caption = "El DNI es inv�lido al parecer."
60  Else
70      lblInfo.Caption = "El DNI parece ser correcto."
80  End If

'90  If Len(txtDNI.Text) > 3 Then
'        Dim formattedText       As String
'100     formattedText = NumberToPunctuatedString(txtDNI.Text)
'110     txtDNI.Text = formattedText
'120     txtDNI.SelStart = cursorPosition + (Len(formattedText) - Len(txtDNI.Text))
'130 End If

140 isUpdating = False

    On Error GoTo 0
    Exit Sub

txtDNI_Change_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento txtDNI_Change de Formulario frmPerson l�nea: " & Erl())

End Sub

Private Sub txtDNI_GotFocus()
    Call CheckTxtControlMouseDown(txtDNI)
    lblInfo.Caption = "Ingrese el n�mero de documento."
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

Private Sub txtEmail_Change()
    If ValidateEmail(txtEmail.Text) Then
        lblInfo.Caption = "El e-mail parecer ser v�lido."
    Else
        lblInfo.Caption = "El e-mail es inv�lido."
    End If
End Sub

Private Sub txtEmail_GotFocus()
    Call CheckTxtControlMouseDown(txtEmail)
    lblInfo.Caption = "Ingrese el e-mail de contacto"
End Sub

Private Sub txtEmail_LostFocus()
    Call CheckTxtControlMouseUp(txtEmail)
End Sub

Private Sub txtName_GotFocus()
    If Not bIsLoaded Then
        Call CheckTxtControlMouseDown(txtName)
        bIsLoaded = False
    Else
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName.Text)
        lblInfo.Caption = "Ingrese un nombre y apellido"
    End If
    
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

Private Sub txtZipCode_GotFocus()
    Call CheckTxtControlMouseDown(txtZipCode)
    lblInfo.Caption = "Ingrese el C�digo Postal"
End Sub

Private Sub txtZipCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtZipCode_LostFocus()
    Call CheckTxtControlMouseUp(txtZipCode)
End Sub
