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

20  Call modDBConnect.LoadDBConfig

30  IsDBAlreadyExists = Val(GetSetting(App.Path, "EIV_SOFTWARE", "IsDBAlreadyExists"))

40  dbName = "EIV"
50  If Not IsDBAlreadyExists Then
60      If Not DBExists(dbName) Then
70          If Not modDBConnect.DBCreate(dbName) Then
80              End
90          Else
100             Call modDBConnect.DbConnect
110             Call modDBConnect.CreateTables
120             Call modDBConnect.SeedDatabase
130         End If
140     End If
150 End If

160 Call modDBConnect.DbConnect

170 frmUserLogin.Show

180 On Error GoTo 0
190 Exit Sub

Main_Error:

200 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento Main de Módulo modMain línea: " & Erl())

End Sub
