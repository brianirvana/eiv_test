Attribute VB_Name = "modDB"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function Exists(ByVal Table As String, ByRef Field As String, ByRef Value As String)

Dim sQuery                      As String
Dim RS                          As ADODB.Recordset

10  On Error GoTo Exists_Error

110 sQuery = "SELECT " & Field & " FROM " & Table & " WHERE " & Field & "='" & Value & "'"
120 Set RS = CN.Execute(sQuery, , adOpenForwardOnly)

130 Exists = Not RS.EOF

140 On Error GoTo 0
150 Exit Function

Exists_Error:
160 Exists = False
170 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento Exists de Módulo modDB línea: " & Erl())

End Function

'---------------------------------------------------------------------------------------
' Procedure : ExistsArr
' Author    : [/About] Brian Sabatier https://github.com/brianirvana
' Date      : 27/12/2024
' Purpose   : Función que devuelve la existencia de registros de varios campos en simultáneo.
'---------------------------------------------------------------------------------------
'
Function ExistsArr(ByVal Table As String, ByRef Fields() As String, ByRef Values() As String)

Dim i                           As Integer
Dim sQuery                      As String
Dim tmpSelectFields             As String
Dim tmpWhereFields              As String
Dim RS                          As ADODB.Recordset

10  On Error GoTo Exists_Error

20  For i = 1 To UBound(Fields)
30      If i = 1 Then
40          tmpSelectFields = Fields(i)
50          tmpWhereFields = Fields(i) & "='" & Values(i) & "'"
60      Else
70          tmpSelectFields = tmpSelectFields & ", " & Fields(i)
80          tmpWhereFields = tmpWhereFields & " AND " & Fields(i) & "='" & Values(i) & "'"
90      End If
100 Next i

110 sQuery = "SELECT " & tmpSelectFields & " FROM " & Table & " WHERE " & tmpWhereFields

120 Set RS = CN.Execute(sQuery, , adOpenForwardOnly)

130 ExistsArr = Not RS.EOF

140 On Error GoTo 0
150 Exit Function

Exists_Error:
160 ExistsArr = False
170 Call Logs.LogError("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ExistsArr de Módulo modDB línea: " & Erl())

End Function
