Attribute VB_Name = "DataSheetHandler"
' Handle for read, write and clear data sheet
Public regexList As Collection
Public fileList As Collection

' Clear data output
Public Sub ClearOutput()

End Sub

Public Sub ParseInput()
    Dim regexItem As String
    Dim fileItem As String
    Dim lineCounter As Integer
    Set regexList = New Collection
    Set fileList = New Collection

    'Get the list regex
    lineCounter = 0
    Do
        regexItem = ThisWorkbook.Sheets(TOOL_INDEX).Cells(START_DATA + lineCounter, REGEX_COL)
        If regexItem <> "" Then
            regexList.Add regexItem
        End If
        lineCounter = lineCounter + 1
    Loop Until regexItem = ""

    'Get the list file
    lineCounter = 0
    Do
        fileItem = ThisWorkbook.Sheets(TOOL_INDEX).Cells(START_DATA + lineCounter, FILE_COL)
        If fileItem <> "" Then
            fileList.Add fileItem
        End If
        lineCounter = lineCounter + 1
    Loop Until fileItem = ""
End Sub
