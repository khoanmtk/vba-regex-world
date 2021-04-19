Attribute VB_Name = "ExcelFunc"
'Create for excel function
Public Function RegexCal(regexCell As Range, inputText As Range) As String
    Dim regexItem As New RegExp
    Dim strRegex As String
    Dim strInput As String
    Dim matches As Object

    strRegex = regexCell.Value
    strInput = inputText.Value
    RegexCal = ""

    If strRegex <> "" Then
        With regexItem
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strRegex
        End With

        'Currently only return the first match
        If regexItem.test(strInput) Then
            Set matches = regexItem.Execute(strInput)
            RegexCal = matches(0).Value
        End If
    End If
End Function
