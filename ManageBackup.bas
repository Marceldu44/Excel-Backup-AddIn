Attribute VB_Name = "Module1"
Option Explicit

Dim sPath As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim sDirection As String
Dim NameFile As String

  Dim sFile As String
Sub VarVal()

    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
    sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
    sPath = sDirection & "*" & sExt
    NameFile = Dir(sPath)
    
End Sub
Sub prueba()

    Dim fList As String
    'Application.Run "Macro.xlsm!Module1.VarVal"
    VarVal
        ' The variable fName now contains the name of the files within the specified path.
        Do While NameFile <> ""
        ' Store the current file in the string fList.
            fList = fList & vbNewLine & NameFile
            ' Get the next files in the specified path.
            NameFile = Dir()
            ' The variable fName now contains the name of the next files in the specified path.
        Loop
        ' Display the list of files in a message box.
        MsgBox ("List of Files:" & fList)

End Sub

Sub FirstFile()

    Dim myFileSystemObject As FileSystemObject
    VarVal
    
        MsgBox NameFile

End Sub

Sub GetAFolder()

    With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = Application.DefaultFilePath & " \ "
    .Title = "Please select a location for the backup"
    .Show
        If .SelectedItems.count = 0 Then
            MsgBox "Canceled"
        Else
            MsgBox .SelectedItems(1)
        End If
    End With

End Sub

Sub GetImportFileName()

    Dim Finfo As String
    Dim FilterIndex As Integer
    Dim Title As String
    Dim FileName As Variant
    ' Set up list of file filters
    Finfo = "Text Files (*.txt),*.txt," & _
    "Lotus Files (*.prn),*.prn," & _
    "Comma Separated Files (*.csv),*.csv," & _
    "ASCII Files (*.asc),*.asc," & _
    "All Files (*.*),*.*"
    ' Display *.* by default
    FilterIndex = 5
    ' Set the dialog box caption
    Title = "Select a File to Import"
    ' Get the filename
    FileName = Application.GetOpenFilename(Finfo, _
    FilterIndex, Title)
    ' Handle return info from dialog box
        If FileName = False Then
            MsgBox "No file was selected."
        Else
            MsgBox "You selected " & FileName
        End If

End Sub

Sub BackUpManagement()

    Dim FSO
    Dim folder
    Dim files
    Dim file
    Dim count
    Dim sDeleteFile As String
    Dim myFileSystemObject As FileSystemObject
    VarVal
            ChDir (sDirection & "\")
            Set FSO = CreateObject("scripting.filesystemobject")
            Set folder = FSO.GetFolder(CurDir())
            Set files = folder.files
                For Each file In files
                count = count + 1
                Next
            sDeleteFile = sDirection & "\" & NameFile
'                If count > 10 Then
'                    Kill sDeleteFile
'                Else
'                    MsgBox "There's 10 or less"
'                End If
                
End Sub

Sub exeCode()

    Dim fileTest() As String
    VarVal
    fileTest = fileArray(sDirection)
        MsgBox fileTest(1)

End Sub

Function GetFolderName(openFile As String) As String


End Function

Function fileArray(sDirection As String) As String()

    Dim FSO As Object, dirFolder, files, file
    Dim arraySize As Integer
    Dim testArray() As String
    
    Set FSO = CreateObject("scripting.filesystemobject")
    Set dirFolder = FSO.GetFolder(sDirection)
    Set files = dirFolder.files
    
    arraySize = 0
    
    For Each file In files
        arraySize = arraySize + 1
        ReDim testArray(arraySize)
    Next
    
    arraySize = 0
    For Each file In files
        arraySize = arraySize + 1
        testArray(arraySize) = file
    Next

    fileArray = testArray

End Function



