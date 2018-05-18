Attribute VB_Name = "Módulo1"
Option Explicit

Private Sub cboxDate_Click()

    If cboxDate.Value = True Then
        cmbDate.Enabled = True
    Else
        cmbDate.Enabled = False
        cmbDate.Value = ""
    End If

End Sub

Private Sub cboxOtherPath_Click()

    If cboxOtherPath.Value = True Then
        textOther.Enabled = True
        cmbPath.Enabled = False
        cmbPath.Value = ""
        lblPath.Enabled = False
        lblOther.Enabled = True
    Else
        textOther.Enabled = False
        textOther.Value = ""
        cmbPath.Enabled = True
        lblPath.Enabled = True
        lblOther.Enabled = False
    End If
    
End Sub

Private Sub cbttCancel_Click()

   Unload ufrmConfig

End Sub

Public Sub cbttOk_Click()

    Dim ConfPath As String
    Dim DateConf As String
    Dim ExtConf As String
    Dim x As Integer
    On Error GoTo BlankSpace
        If Not cmbPath.Enabled = False Then
            If cmbPath.Value = "" Then
                x = 1
                GoTo BlankSpace
            Else
                ConfPath = cmbPath.Value
            End If
        End If
        If Not textOther.Enabled = False Then
            If textOther.Value = "" Then
                x = 1
                GoTo BlankSpace
            Else
                ConfPath = textOther.Value
            End If
        End If
        If Not cmbDate.Enabled = False Then
            If cmbDate.Value = "" Then
                x = 1
                GoTo BlankSpace
            Else
                DateConf = cmbDate.Value
            End If
        End If
    If cboxValues.Value = True Then
        BackupConfigurated
        SaveControls
    End If
    End

BlankSpace:
    If x <> 0 Then
        MsgBox "You have leave an information space in blank or have enter an invalid value!"
        x = 0
    End If

End Sub


Private Sub UserForm_Initialize()

    cmbPath.AddItem "Documents"
    cmbPath.AddItem "Beside this workbook"
    cmbPath.AddItem "Desktop"
    cmbDate.AddItem "Date and Time"
    cmbDate.AddItem "Only Date"
    cmbDate.AddItem "Only Time"
    textOther.Enabled = False
    lblOther.Enabled = False
    cmbDate.Enabled = False
    ReadControls
    If cmbPath.Value <> "" Or textOther.Value <> "" Then
        BackupConfigurated
        End
    End If
    
End Sub

Sub SaveControls()
    
    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    If cmbPath.Value <> "" Or textOther.Value <> "" Then
        DeleteControls
    End If
    SaveSetting var1, "ufrmConfig", "cmbPath", cmbPath.Text
    SaveSetting var1, "ufrmConfig", "cboxOtherPath", False
    If cboxOtherPath = True Then SaveSetting var1, "ufrmConfig", "cboxOtherPath", False
    SaveSetting var1, "ufrmConfig", "textOther", textOther.Text
    SaveSetting var1, "ufrmConfig", "cboxDate", False
    If cboxDate = True Then SaveSetting var1, "ufrmConfig", "cboxDate", False
    SaveSetting var1, "ufrmConfig", "cmbDate", cmbDate.Text

End Sub

Sub ReadControls(): On Error Resume Next

    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    cmbPath.Text = GetSetting(var1, "ufrmConfig", "cmbPath")
    cboxOtherPath = False
    If GetSetting(var1, "ufrmConfig", "cboxOtherPath") Then cboxOtherPath = True
    textOther = GetSetting(var1, "ufrmConfig", "textOther")
    cboxDate = False
    If GetSetting(var1, "ufrmConfig", "cboxDate") Then cboxDate = True
    cmbDate.Text = GetSetting(var1, "ufrmConfig", "cmbDate")
    
End Sub

Sub DeleteControls()

    Dim ConfPath As String
    Dim ConfPath2 As String
    Dim OtherPath As String
    Dim DateConf As String
    Dim ExtConf As String
    Dim ClicDate As String
    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    ConfPath = GetSetting(var1, "ufrmConfig", "cmbPath")
    If ConfPath <> "" Then DeleteSetting var1, "ufrmConfig", "cmbPath"
    OtherPath = GetSetting(var1, "ufrmConfig", "cboxOtherPath")
    If OtherPath = True Then DeleteSetting var1, "ufrmConfig", "cboxOtherPath"
    ConfPath2 = GetSetting(var1, "ufrmConfig", "textOther")
    If ConfPath2 <> "" Then DeleteSetting var1, "ufrmConfig", "textOther"
    ClicDate = GetSetting(var1, "ufrmConfig", "cboxDate")
    If ClicDate = True Then DeleteSetting var1, "ufrmConfig", "cboxDate"
    DateConf = GetSetting(var1, "ufrmConfig", "cmbDate")
    If DateConf <> "" Then DeleteSetting var1, "ufrmConfig", "cmbDate"

End Sub

    
Sub BackupConfigurated()
    
Dim sPath As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim sDirection As String
Dim Fname As String
Dim sDate
Dim msg1 As String
Dim VerifyPath As Boolean
Dim DateConf As String
Dim ConfPath As String
Dim FSO
Dim var1 As String
Dim DiskName As String
Dim DeskTop As String
    On Error GoTo ErrorHandler
    var1 = ActiveWorkbook.CodeName
    VerifyPath = cboxOtherPath.Value
    DateConf = cmbDate.Value
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
    msg1 = "This workbook is not saved yet!"
    DiskName = VBA.Left(Application.DefaultFilePath, 2)
    DeskTop = DiskName & "\Users\" & Application.UserName & "\Desktop" & "\Backup " & sFile & "\"
        If VerifyPath = True Then
            ConfPath = textOther.Value
        Else
            ConfPath = cmbPath.Value
        End If
    Select Case DateConf
        Case "Date and Time"
           sDate = Format(Now, "dd-mm-yy hh.mm.ss")
        Case "Only Date"
            sDate = Format(Now, "dd-mm-yy")
        Case "Only Time"
            sDate = Format(Now, "hh.mm.ss")
    End Select
    Select Case ConfPath
        Case "Documents"
            sDirection = Application.DefaultFilePath & "\Backup " & sFile & "\"
        Case "Beside this workbook"
            sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
        Case "Desktop"
            sDirection = DeskTop
        Case Else
            sDirection = GetSetting(var1, "ufrmConfig", "textOther")
    End Select
    Set FSO = CreateObject("Scripting.FileSystemObject")
        If Not FSO.folderexists(sDirection) Then
            FSO.CreateFolder (sDirection)
        End If
    If sDate = "" Then
        Fname = sDirection & " " & sFile & sExt
    Else
        Fname = sDirection & " " & sFile & "(" & sDate & ")" & sExt
    End If
    ActiveWorkbook.SaveCopyAs Fname

ErrorHandler:
    Select Case Err.Number
        Case 5
            MsgBox msg1
    End Select
End Sub
