Attribute VB_Name = "Logs"
Option Explicit
Public GetVarBuf                As String
Public Const GetVarBufTam       As Long = 3000
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub LogError(ByVal strMsg As String)
    MsgBox strMsg
End Sub

' Función para obtener un valor de una clave específica en el archivo INI
Public Function GetVar(filePath As String, section As String, key As String) As String
    Dim result As String * 255
    Dim length As Long
    
    length = GetPrivateProfileString(section, key, "", result, 255, filePath)
    GetVar = Left(result, length)
End Function
