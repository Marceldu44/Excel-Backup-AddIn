VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmConfig 
   Caption         =   "Backup Config."
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "UserformConfig.frx":0000
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
    Dim DefaultVar As Boolean
    If cboxValues.Value = True Then
        DefaultVar = True
        SaveControls
    Else
        DefaultVar = False
    End If
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

End Sub

Function verifyValue(inputvalue As Object) As String

    Dim x As Integer
    If Not inputvalue.Enabled = False Then
        If inputvalue.Value = "" Then
            x = 1
        Else
            verifyValue = inputvalue.Value
        End If
    End If
    If Error Then x = 1

End Function

Sub SaveControls()
    
    Dim var1 As String
    var1 = ActiveWorkbook.CodeName
    DeleteControls
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

