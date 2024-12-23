Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : Brian Sabatier
' Date      : 23/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Sub main()

Dim dbName                      As String

    dbName = "EIV"

    If Not DBExists(dbName) Then
         If Not modDBConnect.DBCreate(dbName) Then
            End
         End If
    End If

End Sub
