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
      BackColorFixed  =   16777215
      GridColor       =   14737632
      AllowBigSelection=   -1  'True
      FocusRect       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
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

' Configura las columnas de la grilla con 9 campos (incluyendo el ID oculto)
10  On Error GoTo FormatGrid_Error

20  With frmAbmPersons.MSFlexGrid_Persons
30      .Cols = 10  ' Establecer el número de columnas en la grilla (incluye la columna oculta para el ID)

        ' Configurar encabezados de columna
40      .TextMatrix(0, 0) = "ID"  ' Encabezado de la columna oculta
50      .TextMatrix(0, 1) = "Tipo de Documento"
60      .TextMatrix(0, 2) = "Número de Documento"
70      .TextMatrix(0, 3) = "Nombre y Apellido"
80      .TextMatrix(0, 4) = "Fecha de Nacimiento"
90      .TextMatrix(0, 5) = "Género"
100     .TextMatrix(0, 6) = "Localidad"
110     .TextMatrix(0, 7) = "Código Postal"
120     .TextMatrix(0, 8) = "Correo electrónico"
130     .TextMatrix(0, 9) = "Argentino"

        ' Ajustar el ancho de las columnas
140     .ColWidth(0) = 0      ' Ocultar la columna del ID
150     .ColWidth(1) = 900    ' Ancho de la columna de tipo de documento
160     .ColWidth(2) = 900    ' Ancho de la columna de número de documento
170     .ColWidth(3) = 1500   ' Ancho de la columna de nombre y apellido
180     .ColWidth(4) = 1200    ' Ancho de la columna de fecha de nacimiento
190     .ColWidth(5) = 650    ' Ancho de la columna de género
200     .ColWidth(6) = 1500   ' Ancho de la columna de localidad
210     .ColWidth(7) = 900    ' Ancho de la columna de código postal
220     .ColWidth(8) = 1900   ' Ancho de la columna de correo electrónico
230     .ColWidth(9) = 800    ' Ancho de la columna de argentino

240 End With

250 On Error GoTo 0
260 Exit Sub

FormatGrid_Error:

270 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento FormatGrid de Módulo modDBPersons línea: " & Erl())

End Sub

Private Sub cmdAddPerson_Click()

   On Error GoTo cmdAddPerson_Click_Error

10        If Not frmPerson Is Nothing Then
20            Call Unload(frmPerson)
30        End If
          
40        Call Load(frmPerson)
50        frmPerson.TypeMode = eTypeMode.PersonCreate
60        frmPerson.Show
70        Me.Hide

   On Error GoTo 0
   Exit Sub

cmdAddPerson_Click_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdAddPerson_Click de Formulario frmAbmPersons línea: " & Erl())
          
End Sub

Private Sub cmdClose_Click()
    frmUserLogin.Show
    Unload Me
End Sub

Private Sub cmdEditPerson_Click()

Dim selectedRow                 As Integer

    ' Verificar que hay una fila seleccionada (excepto la fila de encabezado)
    On Error GoTo cmdEditPerson_Click_Error

10  selectedRow = MSFlexGrid_Persons.row

20  If selectedRow <= 0 Then
30      MsgBox "Por favor, seleccione un registro válido.", vbExclamation, "Atención"
40      Exit Sub
50  End If

60  tmpUserEdit.Person.id = Val(MSFlexGrid_Persons.TextMatrix(selectedRow, 0))

70  If tmpUserEdit.Person.id < 0 Then
80      MsgBox "Debe seleccionar una persona para continuar."
90      Exit Sub
100 End If

110 Call LoadPerson(frmPerson)
120 frmPerson.TypeMode = eTypeMode.PersonEdit
130 Call modDBPersons.LoadPerson(frmPerson)
140 frmPerson.Show
150 Unload Me

    On Error GoTo 0
    Exit Sub

cmdEditPerson_Click_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdEditPerson_Click de Formulario frmAbmPersons línea: " & Erl())

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmUserLogin.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call modDBPersons.LoadPersons
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MSFlexGrid_Persons_DblClick
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 29/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub MSFlexGrid_Persons_DblClick()

Dim selectedRow                 As Integer
'Dim tmpUser                     As tUser

    ' Obtener la fila seleccionada
    On Error GoTo MSFlexGrid_Persons_DblClick_Error

10  selectedRow = MSFlexGrid_Persons.row

    ' Asegurarse de que la fila seleccionada no sea la cabecera (fila 0)
20  If selectedRow <= 0 Then
30      MsgBox "Seleccione una persona válida de la lista.", vbExclamation, "Advertencia"
40      Exit Sub
50  End If

    ' Obtener el ID de la persona (asumimos que está en la primera columna)
60  tmpUserEdit.Person.id = Val(MSFlexGrid_Persons.TextMatrix(selectedRow, 0))

    ' Verificar que se obtuvo un ID válido
70  If Trim(tmpUserEdit.Person.id) = "" Or tmpUserEdit.Person.id <= 0 Then
80      MsgBox "No se pudo obtener el ID de la persona seleccionada.", vbCritical, "Error"
90      Exit Sub
100 End If

    ' Abrir el formulario de edición de personas
110 Call Load(frmPerson)
120 frmPerson.TypeMode = eTypeMode.PersonEdit
    Call modDBPersons.LoadPerson(frmPerson)
    
130 frmPerson.Show
140 Unload Me

    On Error GoTo 0
    Exit Sub

MSFlexGrid_Persons_DblClick_Error:

    Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento MSFlexGrid_Persons_DblClick de Formulario frmAbmPersons línea: " & Erl())

End Sub

Private Sub MSFlexGrid_Persons_Scroll()
    Debug.Print "ASD"
End Sub
