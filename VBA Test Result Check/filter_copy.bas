Attribute VB_Name = "filter_copy"
'Sub QC_Log_Control_TEST_KONTROL()
'    filterCopy
'    DictionaryVLookupFinal_Code_TESTKONTROL
'    DictionaryVLookup2_ReasonTESTKONTROL
'
'End Sub
Sub CopyLogs()

ans = MsgBox("This process may take several minutes!" & Chr(10) & "Do you want to continue?", vbYesNo, "Copy Logs")
If ans = vbYes Then
    labTestLogCopy
    QCLogCopy
    labTestLogCopyPlate
End If
MsgBox "Process Completed"
End Sub
Sub labTestLogCopy()
Dim wbd As Workbook
Dim wbs As Workbook
Dim wsd As Worksheet
Dim wss As Worksheet
Dim FldrPicker As FileDialog
Dim lr As Long
Dim arr As Variant
Dim a, err As Integer
Dim path As String
Application.DisplayAlerts = False
Application.ScreenUpdating = False
path = sh_Settings.Cells(3, 2)
Set wbd = ThisWorkbook
Set wsd = LabTestLog

If err = 3169 Then

flpckr:
    Set FldrPicker = Application.FileDialog(msoFileDialogFilePicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        path = .SelectedItems(1)
        sh_Settings.Cells(3, 2) = path
    End With
    
    'In Case of Cancel
NextCode:

  'If myPath = "" Then GoTo ResetSettings

End If

On Error GoTo flpckr
Set wbs = Workbooks.Open(path, False, True)
Set wss = wbs.Worksheets(sh_Settings.Cells(4, 2).Value)

lr = wss.Cells(Rows.Count, 5).End(xlUp).Row

wss.Range("B6:B" & lr).Copy
wsd.Range("A1").PasteSpecial xlPasteValues
wsd.Range("A1:A" & lr).NumberFormat = "dd.mm.yyyy"

wss.Range("E6:E" & lr).Copy
wsd.Range("B1").PasteSpecial xlPasteValues

wss.Range("H6:H" & lr).Copy
wsd.Range("D1").PasteSpecial xlPasteValues

wss.Range("F6:F" & lr).Copy
wsd.Range("C1").PasteSpecial xlPasteValues

wss.Range("K6:K" & lr).Copy
wsd.Range("E1").PasteSpecial xlPasteValues

wss.Range("L6:L" & lr).Copy
wsd.Range("F1").PasteSpecial xlPasteValues

lr = wsd.Cells(Rows.Count, 2).End(xlUp).Row

arr = wsd.Range("B2:B" & lr).Value

For i = LBound(arr) To UBound(arr)
a = Len(CStr(arr(i, 1))) - InStrRev(CStr(arr(i, 1)), "-")
    arr(i, 1) = Right(CStr(arr(i, 1)), a)
    
Next i

wsd.Range("B2:B" & lr) = arr

wbs.Close

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


Sub QCLogCopy()

'source worksheet
     Dim wss, wsd As Worksheet
     Dim wbs As Workbook
     Dim wbd As Workbook
     Dim arr As Variant
     Dim arr2 As Variant
     Dim FldrPicker As FileDialog
     
     Application.ScreenUpdating = False
     Application.DisplayAlerts = False
    
     Dim sPath As String
     sPath = sh_Settings.Cells(1, 2).Value
     
     
     
     If err = 3169 Then

flpckr:
    Set FldrPicker = Application.FileDialog(msoFileDialogFilePicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sPath = .SelectedItems(1)
        sh_Settings.Cells(1, 2) = sPath
    End With
    
    'In Case of Cancel
NextCode:

  'If myPath = "" Then GoTo ResetSettings

End If

On Error GoTo flpckr
     
     Set wbs = Workbooks.Open(sPath)
     Set wss = wbs.Worksheets(sh_Settings.Cells(2, 2).Value)
     Set wsd = QCLog
     Set wbd = ThisWorkbook
     Dim lrd As Long
     Dim bul As Range
     Dim lrs, lr As Long
     
     lrd = wsd.Cells(Rows.Count, 8).End(xlUp).Row
     
     arr = wsd.Range("H1:H" & lrd).Value
     
     lrd = wsd.Cells(Rows.Count, 1).End(xlUp).Row
     
     wsd.Range("A1:D" & lrd).Clear
     'enable filter
    lrs = wss.Cells(Rows.Count, 2).End(xlUp).Row
    Set bul = wss.Range("A1:S" & lrs).Find("IR No.")
    
    wss.Rows("1:" & bul.Row - 1).EntireRow.Delete
    
    With wss.Range("A1:S" & lrs)
    
        .AutoFilter field:=17, Criteria1:="Latest", Operator:=xlFilterValues, VisibleDropDown:=True
        .AutoFilter field:=15, Criteria1:=Array("C", "D", "O"), Operator:=xlFilterValues, VisibleDropDown:=True
        .AutoFilter field:=6, Criteria1:=Application.Transpose(arr), Operator:=xlFilterValues, VisibleDropDown:=True
    End With
    
    wss.Range("B:B").SpecialCells(xlCellTypeVisible).Copy
    wsd.Range("A1").PasteSpecial xlPasteValues
    wss.Range("N:N").SpecialCells(xlCellTypeVisible).Copy
    wsd.Range("B1").PasteSpecial xlPasteValues
    wss.Range("O:O").SpecialCells(xlCellTypeVisible).Copy
    wsd.Range("C1").PasteSpecial xlPasteValues
    wss.Range("S:S").SpecialCells(xlCellTypeVisible).Copy
    wsd.Range("D1").PasteSpecial xlPasteValues
    
    lrd = wsd.Cells(Rows.Count, 1).End(xlUp).Row
    
    wsd.Rows(lrd & ":" & Rows.Count).EntireRow.Delete
    
    
    
    lr = wsd.Cells(Rows.Count, 1).End(xlUp).Row

arr2 = wsd.Range("A2:A" & lr).Value

Dim a As Long

For i = LBound(arr2) To UBound(arr2)
a = Len(CStr(arr2(i, 1))) - InStrRev(CStr(arr2(i, 1)), "-")
    arr2(i, 1) = Right(CStr(arr2(i, 1)), a)
    
Next i

wsd.Range("A2:A" & lr) = arr2
    
    

    wbs.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Sub


Sub labTestLogCopyPlate()
Dim wbd As Workbook
Dim wbs As Workbook
Dim wsd As Worksheet
Dim wss As Worksheet
Dim FldrPicker As FileDialog
Dim lr As Long
Dim arr As Variant
Dim a As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set wbd = ThisWorkbook
Set wsd = wbd.Worksheets("test kontrol 2")

Set wbs = Workbooks.Open("Z:\QAQC\Aytekin Bozdogan\02-MASAR PROJECT EARTHWORK\02-GRAVEL BACKFILLING\01-LOG\01-PLATE LOAD LOG.xlsx", False, True)
Set wss = wbs.Worksheets("LOG")

lr = wss.Cells(Rows.Count, 4).End(xlUp).Row

wss.Range("I10:I" & lr).Copy
wsd.Range("C2").PasteSpecial xlPasteValues
wsd.Range("C1:C" & lr).NumberFormat = "dd.mm.yyyy"

wss.Range("D10:D" & lr).Copy
wsd.Range("A2").PasteSpecial xlPasteValues

wss.Range("O10:O" & lr).Copy
wsd.Range("F2").PasteSpecial xlPasteValues



lr = wsd.Cells(Rows.Count, 1).End(xlUp).Row

arr = wsd.Range("A2:A" & lr).Value

For i = LBound(arr) To UBound(arr)
a = Len(CStr(arr(i, 1))) - InStrRev(CStr(arr(i, 1)), "-")
    arr(i, 1) = Right(CStr(arr(i, 1)), a)
    
Next i

wsd.Range("B2:B" & lr) = arr

wbs.Close

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

