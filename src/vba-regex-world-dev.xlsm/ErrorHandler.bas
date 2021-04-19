Attribute VB_Name = "ErrorHandler"
' Handle for Error messages
Private Const MSGBOX_LIMIT = 10
Private errorLog As String

' Show messagebox error
Public Sub InitErrorLog()
    errorLog = ""
End Sub

' Show messagebox error
Public Sub WriteErrorLog()
    Dim path As String
    path = Application.ActiveWorkbook.path

    Call WriteTextFile(path, errorLog)
End Sub

' Show messagebox error
Public Sub AddErrorLog(error As String)
    errorLog = errorLog & vbCrLf & error
End Sub

' Show messagebox error
Public Sub ShowError(errString As String)
    Static counter As Integer
    If counter < MSGBOX_LIMIT Then
        MsgBox errString
        counter = counter - 1
    End If
End Sub
