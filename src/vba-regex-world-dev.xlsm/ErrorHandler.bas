Attribute VB_Name = "ErrorHandler"
' Handle for Error messages
Private Const MSGBOX_LIMIT = 10

' Show messagebox error
Public Sub ShowError(errString As String)
    Static counter As Integer
    If counter < MSGBOX_LIMIT Then
        MsgBox errString
        counter = counter - 1
    End If
End Sub

