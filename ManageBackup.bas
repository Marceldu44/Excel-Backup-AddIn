Attribute VB_Name = "Module1"
Option Explicit

Dim sPath As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim sDirection As String

Sub prueba()

    Dim fList As String
    Dim fName As String
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
    sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
    sPath = sDirection & "*" & sExt
        fName = Dir(sPath)
        ' The variable fName now contains the name of the files within the specified path.
        Do While fName <> ""
        ' Store the current file in the string fList.
            fList = fList & vbNewLine & fName
            ' Get the next files in the specified path.
            fName = Dir()
            ' The variable fName now contains the name of the next files in the specified path.
        Loop
        ' Display the list of files in a message box.
        MsgBox ("List of Files:" & fList)

End Sub

Sub FirstFile()

    Dim NameFile As String
    Dim myFileSystemObject As FileSystemObject
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
    sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
    sPath = sDirection & "*" & sExt
        NameFile = Dir(sPath)
    
        MsgBox NameFile

End Sub

Sub BackUpManagement()

    Dim fso
    Dim folder
    Dim files
    Dim file
    Dim count
    Dim NameFile As String
    Dim sName As String
    Dim sDeleteFile As String
    Dim myFileSystemObject As FileSystemObject
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
    sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
    sPath = sDirection & "*" & sExt
        NameFile = Dir(sPath)
        ChDir (sDirection & "\")
        Set fso = CreateObject("scripting.filesystemobject")
        Set folder = fso.GetFolder(CurDir())
        Set files = folder.files
            For Each file In files
            count = count + 1
            Next
        sDeleteFile = sDirection & "\" & NameFile
            If count > 10 Then
                Kill sDeleteFile
            Else
                MsgBox "There's 10 or less"
            End If

End Sub

Sub exeCode()

    Dim sFolder As String
    Dim fileTest() As String
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sFolder = ActiveWorkbook.Path & "\Backup " & sFile & "\"
    fileTest = fileArray(sFolder)
        MsgBox fileTest(1)

End Sub

Function GetFolderName(openFile As String) As String


End Function

Function fileArray(sFolder As String) As String()

    Dim fso As Object, dirFolder, files, file
    Dim arraySize As Integer
    Dim testArray() As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set dirFolder = fso.GetFolder(sFolder)
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

Function FsoFolder(sFolder As String) As Variant

End Function

Sub test4()
    
    Dim var1 As String
    var1 = FileName(sFil)
    MsgBox var1

End Sub


