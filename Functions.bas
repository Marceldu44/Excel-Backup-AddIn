Attribute VB_Name = "Module1"
Option Explicit

Public Sub Button()


Dim wSheet As Worksheet
On Error Resume Next
Set wSheet = Sheets("Backup")
On Error GoTo 0
    If wSheet Is Nothing Then
        Worksheets.Add.Name = "Backup"
    End If
    If wSheet.Buttons.Count = 0 Then
        ActiveSheet.Buttons.Add(95.25, 60.75, 193.5, 75.75).Select
        Selection.OnAction = "Macro.xlsm!RibbonBackUP2"
        ActiveSheet.Shapes("Button 1").IncrementLeft -95.25
        ActiveSheet.Shapes("Button 1").IncrementTop -60.75
        ActiveSheet.Shapes.Range(Array("Button 1")).Select
        Selection.Characters.Text = "Backup"
    End If
    With wSheet
        Columns("E:E").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.EntireColumn.Hidden = True
        Rows("6:6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.EntireRow.Hidden = True
    End With

End Sub


Sub ExecuteBackup()

Dim openWb As Workbook
Dim pathOpenWb As String, nameOpenWb As String
Dim testBk As Boolean

    For Each openWb In Workbooks 'Running over all opened Workbooks

        pathOpenWb = openWb.Path
        nameOpenWb = openWb.Name
        
'        MsgBox pathOpenWb & nameOpenWb

        If TestBackup(pathOpenWb, nameOpenWb) = "%" Then

            BackupFolder pathOpenWb, nameOpenWb
        End If

    Next
 
End Sub

Function TestBackup(pathName, FileName As String) As String
    
    If pathName = "" Then
        MsgBox "Tuto Backup"
    Else
        TestBackup = VBA.Mid(FileName, VBA.InStr(FileName, ".xl") - 1, 1)
    End If
End Function

Sub BackupFolder(pathName As String, FileName As String)

Dim FSO As Object, sBackupFolder As String, sFile As String

sFile = VBA.Mid(FileName, 1, VBA.InStr(FileName, ".xl") - 2)

sBackupFolder = pathName & "\Backup " & sFile

MsgBox sBackupFolder
Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.folderexists(sBackupFolder) Then
        FSO.CreateFolder (sBackupFolder)
'        BackupConfig.Show
        
    End If


End Sub


Sub BackupConfig()

BackupConfigForm.Show



End Sub

Sub testPath()
 
    MsgBox ThisWorkbook.Path
    MsgBox ThisWorkbook.Name
    MsgBox ThisWorkbook.FullName

    
End Sub

Function FileExt(sPath As String) As String


End Function

Function FileNameOnly(sPath As String) As String

'Return file name without extension
    Dim FirstString As String
    FirstString = VBA.Mid(sPath, 1, VBA.InStr(sPath, ".x") - 1)
    FileNameOnly = VBA.Mid(FirstString, InStrRev(FirstString, "\", -1, vbTextCompare) + 1)

End Function
Sub prueba()
Dim w As String
w = FileNameOnly("C:\Users\Apollo\Documents\Backup Config\Book2(04-04-18).xlsx")
MsgBox w
End Sub

Function IsBackup(sPath As String) As Boolean

'if file is filename%.ext, backup ok
    IsBackup = InStr(1, sPath, "%", vbTextCompare)
    

End Function

Function CreateFolder(sFilepath As String) As String
    
    'return FolderPath/create the folder in case it doesn't exist
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
                    If Not FSO.folderexists(sFilepath) Then
                        FSO.CreateFolder (sFilepath)
                    Else
                        MsgBox "Folder Exists!"
                    End If
    
End Function

Sub SaveFile(sFileName As String, sFolderPath As String, sExtension As String)

    Dim Fname
    ThisWorkbook.SaveAs Fname
    Fname = sFolderPath & " " & sFileName & sExtension
    
End Sub


Sub testRun()

    Dim wbPath As String
    
    wbPath = ThisWorkbook.FullName
    
    If IsBackup(wbPath) = True Then
        
        
        
    End If
    

End Sub


