Attribute VB_Name = "FileHandler"
' Check if folder exist
' Return E_OK if exist
' or E_NOT_OK if not exist
Public Function IsFolderExist(path As String) As Integer
    Dim folder As String
    folder = Dir(path, vbDirectory)
    If folder = "" Then
        IsFolderExist = E_NOT_OK
    Else
        IsFolderExist = E_OK
    End If
End Function

' Check if folder exist
' Return E_OK if exist
' or E_NOT_OK if not exist
Public Function IsFileExist(path As String) As Integer
    Dim fileExist As String
    fileExist = Dir(path)
    If fileExist = "" Then
        IsFileExist = E_NOT_OK
    Else
        IsFileExist = E_OK
    End If
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
    Set Fileout = fso.CreateTextFile("C:\your_path\vba.txt", True, True)
    Fileout.Write "your string goes here"
    Fileout.Close
End Function

'Go through all subfolders
Public Sub ReadThroughFolder(path As String)
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder("your folder path variable") 'obviously replace
    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files
            '...insert any file processing code here...
        Next oFile
    Loop
End Sub
