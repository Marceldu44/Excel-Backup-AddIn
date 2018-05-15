Attribute VB_Name = "Módulo1"
Option Explicit
Dim sPath As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim sDirection As String
Dim Fname As String
Dim sDate
Dim msg1 As String

Sub BackupConfigurated(ExtConf As String, DateConf As String, confPath As String)
    
    Dim FSO
    ufrmConfig.Show
    Select Case DateConf
        Case "Date and Time"
           sDate = Format(Now, "dd-mm-yy hh.mm.ss")
        Case "Only Date"
            sDate = Format(Now, "dd-mm-yy")
        Case "Only Time"
            sDate = Format(Now, "hh.mm.ss")
    End Select
    Select Case confPath
        Case "Documents"
            sDirection = "Libraries\Documents" & "\Backup " & sFile & "\"
        Case "Beside this workbook"
            sDirection = ActiveWorkbook.Path & "\Backup " & sFile & "\"
        Case "Desktop"
            sDirection = "Desktop" & "\Backup " & sFile & "\"
    End Select
    Select Case ExtConf
        Case "OpenDocument Spreadsheet"
            sExt = ".ods"
        Case "Strict Open XML Spreadsheet"
            sExt = ".xlsx"
        Case "XPS Document"
            sExt = ".xps"
        Case "PDF"
            sExt = ".pdf"
        Case "Excel 97-2003 Add-in"
            sExt = ".xla"
        Case "Excel Add-in"
            sExt = ".xlam"
        Case "SYLK(Symbolic Link)"
            sExt = ".slk"
        Case "DIF(Data Interchange Format)"
            sExt = ".dif"
        Case "CSV(MS-DOS)"
            sExt = ".csv"
        Case "CSV(Macintosh)"
            sExt = ".csv"
        Case "Text(MS-DOS)"
            sExt = ".txt"
        Case "Text(Macintosh)"
            sExt = ".txt"
        Case "Formatted Text(Space Delimited)"
            sExt = ".prn"
        Case "CSV(Comma Delimited)"
            sExt = ".csv"
        Case "Microsoft Excel 5.0/95 Workbook"
            sExt = ".xls"
        Case "XLM Spreadsheet 2003"
            sExt = ".xml"
        Case "Unicode Text"
            sExt = ".txt"
        Case "Text(Tab Delimited)"
            sExt = ".txt"
        Case "Excel 97-2000 Template"
            sExt = ".xlt"
        Case "Excel Macro-Enabled Template"
            sExt = ".xltm"
        Case "Excel Template"
            sExt = ".xltx"
        Case "XLM Data"
            sExt = ".xml"
        Case "CSV UFT-8(Comma Delimited)"
            sExt = ".csv"
        Case "Excel 97-2003 Workbook"
            sExt = ".xls"
        Case "Excel Binary Workbook"
            sExt = ".xlsb"
        Case "Excel Macro-Enabled Workbook"
            sExt = ".xlsm"
        Case "Excel Workbook"
            sExt = ".xlsx"
    End Select
    sFil = ActiveWorkbook.Name
    sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
    Fname = sDirection & " " & sFile & "(" & sDate & ")" & sExt
    msg1 = "This workbook is not saved yet!"
    On Error GoTo ErrorHandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
        If Not FSO.folderexists(sDirection) Then
            FSO.CreateFolder (sDirection)
        End If
    ActiveWorkbook.SaveCopyAs Fname

ErrorHandler:
    Select Case Err.Number
        Case 5
            MsgBox msg1
    End Select
End Sub
