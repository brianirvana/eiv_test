VERSION 5.00
Begin VB.Form frmUserLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Login de Usuario"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteDB 
      Caption         =   "Delete DB"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdUserCreate 
      Caption         =   "Crear usuario"
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtUserPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "#"
      TabIndex        =   1
      Text            =   "3421321a"
      Top             =   960
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
      Text            =   "admin"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   5415
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmLogin
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdDeleteDB_Click()

    If MsgBox("¿Está seguro que desea eliminar la base de datos " & cDB.dbName & "?", vbOKCancel) = vbOK Then
        Call SaveSetting(App.Path, "EIV_SOFTWARE", "IsDBAlreadyExists", "0")
        Call modDBConnect.DropDatabase(cDB.dbName)
    End If

End Sub

Private Sub cmdUserCreate_Click()

    frmUserCreate.Show
    Me.Hide

End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdLogin_Click()

Dim sErrorMsg                   As String
Dim tmpUser                     As tUser

    tmpUser.UserName = txtUserName.Text
    tmpUser.Password = txtUserPassword.Text

    If Not ValidateUserLogin(tmpUser, sErrorMsg) Then
        Call MsgBox(sErrorMsg, vbInformation, "Error, por favor revise la información ingresada.")
        lblInfo.Caption = "Error, por favor revise la información ingresada."
    Else
        If Not ValidateDBPassword(tmpUser, sErrorMsg) Then
            Call MsgBox(sErrorMsg, vbInformation, "Error, contraseña inválida.")
            lblInfo.Caption = "Error, contraseña inválida."
            Exit Sub
        End If
        
        frmAbmPersons.Show
        Me.Hide
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        End
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtUserPassword.SetFocus
    End If

    If (KeyAscii <> 8) Then
        ' Verificar si el carácter no es una letra (minúscula o mayúscula)
        If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) Then
            KeyAscii = 0
        End If
    End If

   lblInfo.Caption = vbNullString

End Sub

Private Sub txtUserName_LostFocus()
    If StrComp(UCase$(txtUserName.Text), vbNullString) = 0 Or StrComp(UCase$(txtUserName.Text), " ") = 0 Then
        txtUserName.Text = "Usuario"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtUserName_MouseUp
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtUserName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo txtUserName_MouseUp_Error

    If StrComp(UCase$(txtUserName.Text), "USUARIO") = 0 Then
        txtUserName.Text = vbNullString
    End If

    On Error GoTo 0
    Exit Sub

txtUserName_MouseUp_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento txtUserName_MouseUp de Formulario frmUserLogin línea: " & Erl())

End Sub

Private Sub txtUserPassword_KeyPress(KeyAscii As Integer)

    lblInfo.Caption = vbNullString

    If KeyAscii = 13 Then
        Call cmdLogin_Click
    End If
    
End Sub

Private Sub txtUserPassword_LostFocus()
    If StrComp(UCase$(txtUserPassword.Text), vbNullString) = 0 Or StrComp(UCase$(txtUserPassword.Text), " ") = 0 Then
        txtUserPassword.Text = "CONTRASEÑA"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtUserPassword_MouseUp
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtUserPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo txtUserPassword_MouseUp_Error

    If StrComp(UCase$(txtUserPassword.Text), "CONTRASEÑA") = 0 Then
        txtUserPassword.Text = vbNullString
    End If

    On Error GoTo 0
    Exit Sub

txtUserPassword_MouseUp_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento txtUserPassword_MouseUp de Formulario frmUserLogin línea: " & Erl())

End Sub
