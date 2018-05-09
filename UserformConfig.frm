VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufrmConfig 
   Caption         =   "Backup Config."
   ClientHeight    =   4125
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

Private Sub cbttOk_Click()
    Dim confPath As String
    Dim DateConf As String
    Dim ExtConf As String
    Dim x As Integer
        If Not cmbPath.Enabled = False Then
            If cmbPath.Value = "" Then
                x = 1
                GoTo BlankSpace
            Else
                confPath = cmbPath.Value
            End If
        End If
        If Not textOther.Enabled = False Then
            If textOther.Value = "" Then
                x = 1
                GoTo BlankSpace
            Else
                confPath = textOther.Value
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
        If cmbExt.Value = "" Then
            x = 1
            GoTo BlankSpace
        Else
            ExtConf = cmbExt.Value
        End If
    Dim z
    SaveSetting "MyApp", "Startup", "Top", 90
    z = GetSetting("MyApp", "Startup", "Top")
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
    cmbExt.AddItem "Excel Workbook" '".xlsx"
    cmbExt.AddItem "Excel Macro-Enabled Workbook" '".xlsm"
    cmbExt.AddItem "Excel Binary Workbook" '".xlsb"
    cmbExt.AddItem "Excel 97-2003 Workbook" '".xls"
    cmbExt.AddItem "CSV UFT-8(Comma Delimited)" '".csv"
    cmbExt.AddItem "XLM Data" '".xml"
    cmbExt.AddItem "Excel Template" '".xltx"
    cmbExt.AddItem "Excel Macro-Enabled Template" '".xltm"
    cmbExt.AddItem "Excel 97-2000 Template" '".xlt"
    cmbExt.AddItem "Text(Tab Delimited)" '".txt"
    cmbExt.AddItem "Unicode Text" '".txt"
    cmbExt.AddItem "XLM Spreadsheet 2003" '".xml"
    cmbExt.AddItem "Microsoft Excel 5.0/95 Workbook" '".xls"
    cmbExt.AddItem "CSV(Comma Delimited)" '".csv"
    cmbExt.AddItem "Formatted Text(Space Delimited)" '".prn"
    cmbExt.AddItem "Text(Macintosh)" '".txt"
    cmbExt.AddItem "Text(MS-DOS)" '".txt"
    cmbExt.AddItem "CSV(Macintosh)" '".csv"
    cmbExt.AddItem "CSV(MS-DOS)" '".csv"
    cmbExt.AddItem "DIF(Data Interchange Format)" '".dif"
    cmbExt.AddItem "SYLK(Symbolic Link)" '".slk"
    cmbExt.AddItem "Excel Add-in" '".xlam"
    cmbExt.AddItem "Excel 97-2003 Add-in" '".xla"
    cmbExt.AddItem "PDF" '".pdf"
    cmbExt.AddItem "XPS Document" '".xps"
    cmbExt.AddItem "Strict Open XML Spreadsheet" '".xlsx"
    cmbExt.AddItem "OpenDocument Spreadsheet" '".ods"
    textOther.Enabled = False
    lblOther.Enabled = False
    cmbDate.Enabled = False


End Sub

Private Sub UserForm_Terminate()
'Save controls
Dim x
SaveSetting "MyApp", "Startup", "Top", 90
x = GetSetting("MyApp", "Startup", "Top")

End Sub

Private Sub LeerControles()
cmbPath.Text = GetSetting("CodeForPublish", "ufrmConfig", "cmbPath")
cboxOtherPath = False
If GetSetting("CodeForPublish", "ufrmConfig", "cboxOtherPath") Then cboxOtherPath = True
textOther = GetSetting("CodeForPublish", "ufrmConfig", "textOther")
cboxDate = False
If GetSetting("CodeForPublish", "ufrmConfig", "cboxDate") Then cboxDate = True
cmbDate.Text = GetSetting("CodeForPublish", "ufrmConfig", "cmbDate")
cmbExt.Text = GetSetting("CodeForPublish", "ufrmConfig", "cmbExt")
End Sub

Private Sub GuardarControles()
SaveSetting "CodeForPublish", "ufrmConfig", "cmbPath", cmbPath.Text
SaveSetting "CodeForPublish", "ufrmConfig", "cboxOtherPath", False
If cboxOtherPath = True Then SaveSetting "CodeForPublish", "ufrmConfig", "cboxOtherPath", True
SaveSetting "CodeForPublish", "ufrmConfig", "textOther", textOther.Text
SaveSetting "CodeForPublish", "ufrmConfig", "cboxDate", False
If cboxDate = True Then SaveSetting "CodeForPublish", "ufrmConfig", "cboxDate", True
SaveSetting "CodeForPublish", "ufrmConfig", "cmbDate", cmbDate.Text
SaveSetting "CodeForPublish", "ufrmConfig", "cmbExt", cmbExt.Text
End Sub

