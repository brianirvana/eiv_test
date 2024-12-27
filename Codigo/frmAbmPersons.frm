VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAbmPersons 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alta Baja Modificación Personas"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid_Persons 
      Height          =   4815
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8493
      _Version        =   393216
      BackColor       =   12632256
      BackColorFixed  =   12640511
      Appearance      =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   555
      Left            =   2640
      TabIndex        =   2
      Top             =   5040
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   1035
   End
End
Attribute VB_Name = "frmAbmPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmAbmPersons
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

