VERSION 5.00
Begin VB.Form frmUserLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login de Usuario"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
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

Private Sub lblInfo_Click()

End Sub
