VERSION 5.00
Begin VB.Form frmUserCreate 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Crear usuario"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "E-mail"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Contraseña"
      Top             =   1320
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
      Top             =   840
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
      TabIndex        =   3
      Top             =   2880
      Width           =   5295
   End
End
Attribute VB_Name = "frmUserCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtEmail_Change()
    Call CheckTxtUserNameMouseUp(txtEmail)
End Sub

Private Sub txtPassword_Change()
    If Len(txtPassword.Text) < 6 Then
        frmUserCreate.lblInfo.Caption = "La contraseña debe tener como mínimo 6 caracteres."
    End If
End Sub

Private Sub txtPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtUserNameMouseUp(txtPassword)
End Sub

Private Sub txtUserName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CheckTxtUserNameMouseUp(txtUserName)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckTxtControlMouseUp
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Controlamos en los eventos MouseUp (soltar el click) el texto que muestra el control para referenciar al usuario.
'---------------------------------------------------------------------------------------
'
Private Sub CheckTxtControlMouseUp(ByRef txtControl As TextBox)

   On Error GoTo CheckTxtControlMouseUp_Error

    If StrComp(UCase$(txtControl.Text), vbNullString) = 0 Or StrComp(UCase$(txtControl.Text), " ") = 0 Then
        If txtControl.Name = "txtUserName" Then
            txtControl.Text = "Usuario"
        ElseIf txtControl.Name = "txtPassword" Then
            txtControl.Text = "Contraseña"
        ElseIf txtControl.Name = "txtEmail" Then
            txtControl.Text = "E-mail"
        End If
        Exit Sub
    End If

    If txtControl.Text = "Usuario" Or txtControl.Text = "txtPassword" Or txtControl.Text = "txtEmail" Then
        txtControl.Text = vbNullString
    End If

   On Error GoTo 0
   Exit Sub

CheckTxtControlMouseUp_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento CheckTxtControlMouseUp de Formulario frmUserCreate línea: " & Erl())

End Sub
