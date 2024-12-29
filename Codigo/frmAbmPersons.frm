VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAbmPersons 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Alta Baja Modificación Personas"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   255
      Left            =   11520
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid_Persons 
      Height          =   4695
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   20
      Cols            =   7
      FixedCols       =   0
      BackColor       =   12632256
      BackColorFixed  =   12640511
      AllowBigSelection=   -1  'True
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdDeletePerson 
      Caption         =   "Borrar persona"
      Height          =   555
      Left            =   6960
      TabIndex        =   2
      Top             =   5280
      Width           =   1035
   End
   Begin VB.CommandButton cmdEditPerson 
      Caption         =   "Modificar persona"
      Height          =   555
      Left            =   5520
      TabIndex        =   1
      Top             =   5280
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddPerson 
      Caption         =   "Nueva persona"
      Height          =   555
      Left            =   4200
      TabIndex        =   0
      Top             =   5280
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
10  On Error GoTo FormatGrid_Error

20  With frmAbmPersons.MSFlexGrid_Persons
30      .Cols = 9  ' Establecer el número de columnas en la grilla (debe ser 7 en este caso)

        ' Configurar encabezados de columna
40      .TextMatrix(0, 0) = "Tipo de Documento"
50      .TextMatrix(0, 1) = "Número de Documento"
60      .TextMatrix(0, 2) = "Nombre y Apellido"
70      .TextMatrix(0, 3) = "Fecha de Nacimiento"
80      .TextMatrix(0, 4) = "Género"
90      .TextMatrix(0, 5) = "Localidad"
100     .TextMatrix(0, 6) = "Código Postal"
110     .TextMatrix(0, 7) = "Correo electrónico"
120     .TextMatrix(0, 8) = "Argentino"

        ' Opcional: Ajustar el ancho de las columnas
130     .ColWidth(0) = 900   ' Ajustar ancho de la columna de tipo de documento
140     .ColWidth(1) = 900   ' Ajustar ancho de la columna de número de documento
150     .ColWidth(2) = 1500   ' Ajustar ancho de la columna de nombre y apellido
160     .ColWidth(3) = 900   ' Ajustar ancho de la columna de fecha de nacimiento
170     .ColWidth(4) = 650   ' Ajustar ancho de la columna de género
180     .ColWidth(5) = 1500  ' Ajustar ancho de la columna de localidad
190     .ColWidth(6) = 900   ' Ajustar ancho de la columna de código postal
200     .ColWidth(7) = 1900
210     .ColWidth(8) = 800

220 End With

230 On Error GoTo 0
240 Exit Sub

FormatGrid_Error:

250 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento FormatGrid de Módulo modDBPersons línea: " & Erl())

End Sub

Private Sub cmdAddPerson_Click()
    frmPersonAdd.Show
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    frmUserLogin.Show
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmUserLogin.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call modDBPersons.LoadPersonas
End Sub

 
Private Sub MSFlexGrid_Persons_Scroll()
    Debug.Print "ASD"
End Sub
