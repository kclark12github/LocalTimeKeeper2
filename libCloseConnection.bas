Attribute VB_Name = "libCloseConnection"
Option Explicit
Public Sub CloseConnection(cn As ADODB.Connection, Destroy As Boolean)
    On Error Resume Next
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        If Destroy Then
            Set cn = Nothing
        End If
    End If
End Sub

