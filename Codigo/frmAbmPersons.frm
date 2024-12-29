VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAbmPersons 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Alta Baja Modificación Personas"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
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

'---------------------------------------------------------------------------------------
' Procedure : FormatGrid
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 28/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub FormatGrid()

' Configura las columnas de la grilla con 7 campos
    On Error GoTo FormatGrid_Error

10  With frmAbmPersons.MSFlexGrid_Persons
20      .Cols = 7  ' Establecer el número de columnas en la grilla (debe ser 7 en este caso)

        ' Configurar encabezados de columna
30      .TextMatrix(0, 0) = "Tipo de Documento"
40      .TextMatrix(0, 1) = "Número de Documento"
50      .TextMatrix(0, 2) = "Nombre y Apellido"
60      .TextMatrix(0, 3) = "Fecha de Nacimiento"
70      .TextMatrix(0, 4) = "Género"
80      .TextMatrix(0, 5) = "Localidad"
90      .TextMatrix(0, 6) = "Código Postal"

        ' Opcional: Ajustar el ancho de las columnas
100     .ColWidth(0) = 2000   ' Ajustar ancho de la columna de tipo de documento
110     .ColWidth(1) = 2000   ' Ajustar ancho de la columna de número de documento
120     .ColWidth(2) = 3000   ' Ajustar ancho de la columna de nombre y apellido
130     .ColWidth(3) = 2000   ' Ajustar ancho de la columna de fecha de nacimiento
140     .ColWidth(4) = 1000   ' Ajustar ancho de la columna de género
150     .ColWidth(5) = 2000   ' Ajustar ancho de la columna de localidad
160     .ColWidth(6) = 1500   ' Ajustar ancho de la columna de código postal
170 End With

    On Error GoTo 0
    Exit Sub

FormatGrid_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento FormatGrid de Módulo modDBPersons línea: " & Erl())

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call modDBPersons.LoadPersonas
End Sub

