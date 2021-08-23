Attribute VB_Name = "FeaureSearch"
Public Sub SearchRegex()
    Dim updateSheet As Worksheet
    Dim counter As Integer
    Dim inputList() As String
    Dim fileList() As String

    Call InitRegex
    Set updateSheet = ThisWorkbook.Sheets(SHEET_SEARCH)
    counter = START_DATA
    While updateSheet.Cells(counter, SEARCH_REGEX_COL) <> ""
        ' Search here.
        regexGo.Pattern = updateSheet.Cells(counter, SEARCH_REGEX_COL)

        ' Get the list of file to search
        Call GetListOfFile(updateSheet.Cells(counter, SEARCH_FILE_COL), inputList)

        For i = 0 To UBound(inputList)
            'MsgBox (inputList(i))
            If IsDir(inputList(i)) Then
                ' Handle for folder
                Call ReadThroughFolder(inputList(i), fileList)
                'For j = 0 To UBound(fileList)
                '    MsgBox (fileList(j))
                'Next j
                ' Apply the regex to list of file here
                For j = 0 To UBound(fileList)
                    Call ApplyRegexToFile(fileList(j), updateSheet, counter)
                Next j

            Else
                ' Handle for File
                If IsFile(inputList(i)) Then
                    ' Apply the regex to list of file here
                    Call ApplyRegexToFile(inputList(i), updateSheet, counter)
                End If
            End If
        Next i
        counter = counter + 1
    Wend
End Sub

Public Sub ApplyRegexToFile(filePath As String, updateSheet As Worksheet, row As Integer)
    Dim regexString As String
    Dim searchString As String
    Dim outputString As String
    Dim allMatches As Object
    Dim outputCounter As Long
    Dim i As Long
    Dim j As Long

    searchString = ReadTextFile(filePath)
    Set allMatches = regexGo.Execute(searchString)
    outputString = ""
    outputCounter = 0

    For i = 0 To allMatches.Count - 1
        outputString = allMatches.Item(i)
        updateSheet.Cells(row + outputCounter, SEARCH_OUTPUT_COL) = outputString
        outputCounter = outputCounter + 1
    Next i


End Sub

