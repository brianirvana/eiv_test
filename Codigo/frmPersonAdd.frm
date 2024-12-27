VERSION 5.00
Begin VB.Form frmPersonAdd 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Añadir Persona"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPersonAdd"
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

