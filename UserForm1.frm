VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Backup Config."
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckBox1_Click()

End Sub

Private Sub cboxOtherPath_Click()

    If cboxOtherPath.Value = True Then
        textOther.Enabled = True
        cmbPath.Enabled = False
        lblPath.Enabled = False
        lblOther.Enabled = True
    Else
        textOther.Enabled = False
        cmbPath.Enabled = True
        lblPath.Enabled = True
        lblOther.Enabled = False
    End If
    
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub cmbPath_Change()

End Sub

Private Sub CommandButton1_Click()

    Dim conf As String
    Dim anPath As String
    conf = ComboBox1.Value
    anPath = TextBox1
'    MsgBox conf
'    MsgBox anPath
    

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()

    cmbPath.AddItem "Documents"
    cmbPath.AddItem "Beside this workbook"
    cmbPath.AddItem "Desktop"
    'If ComboBox1. =  Then TextBox1.Enabled = True
    
End Sub
            
Private Sub ListBox1_Click()

    

End Sub

