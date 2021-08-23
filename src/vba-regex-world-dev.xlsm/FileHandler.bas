Attribute VB_Name = "FileHandler"
' Check if folder exist
Public Function IsFile(path As String)
    IsFile = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' Check if folder exist
Public Function IsDir(path As String)
    IsDir = CreateObject("Scripting.FileSystemObject").FolderExists(path)
End Function

' Read text file by get line by line method
Public Function ReadTextFileAlt(path As String) As String
    Dim fileNum As Integer
    Dim lineString As String
    Dim outString As String

    fileNum = FreeFile
    outString = ""
    Open path For Input As #fileNum

    While Not EOF(fileNum)
        Line Input #fileNum, lineString ' read in data 1 line at a time
        outString = outString & lineString
    Wend

    Close #fileNum
    ReadTextFileAlt = outString
End Function

' Read text file by parse all, it may get error in some file with japanese encoding.
' Change to use ReadTextFile2 if get error related to read file
Public Function ReadTextFile(path As String) As String
    Dim objFSO As Object
    Dim objTF As Object
    Dim strIn As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTF = objFSO.OpenTextFile(path, 1)
    strIn = objTF.readall
    objTF.Close

    ReadTextFile = strIn
End Function

' Read text file
' Return E_OK or E_NOT_OK
Public Function WriteTextFile(path As String, dataToWrite As String) As Integer
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(path, True, True)
    Fileout.Write dataToWrite
    Fileout.Close
End Function

'Go through all subfolders
Public Sub ReadThroughFolder(inputFolder As String, fileArray() As String)
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Dim fileNum As Integer
    fileNum = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder(inputFolder) 'obviously replace
    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files
            '...insert any file processing code here...
            ReDim Preserve fileArray(fileNum)
            fileArray(fileNum) = oFile.path
            fileNum = fileNum + 1
        Next oFile
    Loop
End Sub


'Get list of file
Public Sub GetListOfFile(listOfFile As String, fileArray() As String)
    If InStr(listOfFile, vbLf) Then
        fileArray = Split(listOfFile, vbLf)
    Else
        ReDim fileArray(0)
        fileArray(0) = listOfFile
    End If
End Sub

