VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmConfig 
   Caption         =   "Backup Config."
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "CodeUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    If cboxValues.Value = True Then SaveControls
    BackupConfigurated
    Unload ufrmConfig


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
    If cmbPath.Value <> "" Or textOther.Value <> "" Then    '<=========es en este c�digo que megenera el error
        BackupConfigurated                                  '<=========trato de hacer que si el UserForm ha sido
        Unload ufrmConfig                                   '<=========guardado antes, se haga el backup sin
    End If                                                  '<=========necesidad de que el usuario presione el boton.
    
End Sub

Sub SaveControls()
    
    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    If cmbPath.Value <> "" Or textOther.Value <> "" Then
        DeleteControls
    End If
    SaveSetting var1, "ufrmConfig", "cmbPath", cmbPath.Text
    SaveSetting var1, "ufrmConfig", "cboxOtherPath", False
    If cboxOtherPath = True Then SaveSetting var1, "ufrmConfig", "cboxOtherPath", True
    SaveSetting var1, "ufrmConfig", "textOther", textOther.Text
    SaveSetting var1, "ufrmConfig", "cboxDate", False
    If cboxDate = True Then SaveSetting var1, "ufrmConfig", "cboxDate", True
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

    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    If cmbPath.Text <> "" Then DeleteSetting var1, "ufrmConfig", "cmbPath"
    If cboxOtherPath = True Then DeleteSetting var1, "ufrmConfig", "cboxOtherPath"
    If textOther.Text <> "" Then DeleteSetting var1, "ufrmConfig", "textOther"
    If cboxDate = True Then DeleteSetting var1, "ufrmConfig", "cboxDate"
    If cmbDate.Text <> "" Then DeleteSetting var1, "ufrmConfig", "cmbDate"

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
