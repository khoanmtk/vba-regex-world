Attribute VB_Name = "RegexSupport"
Public regexGo As New RegExp

Public Sub InitRegex()
    With regexGo
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = ""
    End With
End Sub
