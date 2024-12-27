Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public IsDBAlreadyExists        As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub Main()

Dim dbName                      As String

10  On Error GoTo Main_Error

20  IsDBAlreadyExists = Val(GetSetting(App.Path, "EIV_SOFTWARE", "IsDBAlreadyExists"))

30  dbName = "EIV"
40  If Not IsDBAlreadyExists Then
50      If Not DBExists(dbName) Then
60          If Not modDBConnect.DBCreate(dbName) Then
70              End
80          Else
90              Call modDBConnect.DbConnect
100             Call modDBConnect.CreateTables
110             Call modDBConnect.SeedDatabase
120         End If
130     End If
140 End If

150 Call modDBConnect.DbConnect

160 frmUserLogin.Show

170 On Error GoTo 0
180 Exit Sub

Main_Error:

190 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento Main de Módulo modMain línea: " & Erl())

End Sub
