Attribute VB_Name = "Link"
Sub Link()
InternalHyprlnk
DictionaryHyprlnk

End Sub

Sub InternalHyprlnk()

Dim wststcnt As Worksheet
Dim wsGnrl As Worksheet
Dim vArray1 As Variant
Dim vArray2 As Variant
Dim vArray3 As Variant
Dim dict As New Dictionary
Dim dict2 As New Dictionary

Set wsGnrl = ThisWorkbook.Sheets(1)
Set wststcnt = ThisWorkbook.Worksheets("Test Control")

Dim lr As Long
Dim i, j, k As Long
Dim adress As String
lr = wsGnrl.Cells(Rows.Count, 2).End(xlUp).Row
lr2 = wststcnt.Cells(Rows.Count, 2).End(xlUp).Row

vArray1 = wsGnrl.Range("B2:B" & lr).Value
vArray2 = wsGnrl.Range("H2:H" & lr).Value
vArray3 = wststcnt.Range("B2:B" & lr2).Value

For i = 1 To UBound(vArray1, 1)

    dict.Item(vArray1(i, 1)) = wsGnrl.Range("B" & i + 1).Address
    
    
Next i
    
For i = 1 To UBound(vArray2, 1)

    If dict.Exists(vArray2(i, 1)) And vArray2(i, 1) <> "" Then
        
        adress = dict.Item(vArray2(i, 1))
        wsGnrl.Hyperlinks.Add Anchor:=wsGnrl.Cells((i + 1), 8), Address:="", SubAddress:=adress, TextToDisplay:=CStr(vArray2(i, 1))
         
    End If
    
    
Next i


'////////////////////////////////////////////*****************************************



For i = 1 To UBound(vArray1, 1)

    dict2.Item(vArray1(i, 1)) = wsGnrl.Range("B" & i + 1).Address
    
    
Next i
    
For i = 1 To UBound(vArray3, 1)

    If dict2.Exists(vArray3(i, 1)) And vArray3(i, 1) <> "" Then
        
        adress = "'General'!" & dict2.Item(vArray3(i, 1))
        wsGnrl.Hyperlinks.Add Anchor:=wststcnt.Cells((i + 1), 2), Address:="", SubAddress:=adress, TextToDisplay:=CStr(vArray3(i, 1))
         
    End If
Next i


        


End Sub




Sub DictionaryHyprlnk()

Dim x, x2, x3, y, z
Dim a As String
Dim b As Variant
Dim dict As New Dictionary
'Set dict = CreateObject("Scripting.Dictionary")
Dim wsQC, wsEW As Worksheet
Dim wbQC, wbEW As Workbook
Dim sPath, lookupFile As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False

sPath = ThisWorkbook.path
lookupFile = sPath & "\Dosya bulma.xlsm"

Set wbQC = Workbooks.Open(lookupFile, False)
Set wbEW = ThisWorkbook
Set wsQC = wbQC.Worksheets("kontrol")
Set wsEW = ThisWorkbook.Worksheets("General")

With wsQC
    lr = .Cells(Rows.Count, "E").End(xlUp).Row
    x = .Range("B2:B" & lr).Value2
    x2 = .Range("F2:F" & lr).Value2
 
End With

For i = 1 To UBound(x, 1)

    dict.Item(x(i, 1)) = x2(i, 1)
    
    
Next i

lr2 = wsEW.Cells(Rows.Count, "B").End(xlUp).Row
y = wsEW.Range("B2:B" & lr2).Value2

'ReDim y2(1 To UBound(y, 1), 1 To 1)


For i = 1 To UBound(y, 1)

    If dict.Exists(y(i, 1)) Then
    
        'y2(i, 1) = dict(y(i, 1))
        b = y(i, 1)
        a = dict(y(i, 1))
        'wsEW.Cells(i, 10) = a
        wsEW.Hyperlinks.Add Anchor:=wsEW.Cells((i + 1), 2), Address:=a, TextToDisplay:=CStr(b)
        
    Else
    
        'y2(i, 1) = ""
        
    End If
    
Next i




'wsEW.Range("K2:K" & lr2).Value = y2

Set dict = Nothing
Application.DisplayAlerts = True
Application.ScreenUpdating = True

wbQC.Close
Call InternalHyprlnk


End Sub


Sub Link2()

Dim a, b As Integer
Dim bulunan As Range
Dim number As String
Dim adress, adress2 As String
a = 7


    Do

        number = Worksheets("Marking-Compaction2").Cells(a, 4).Value

        
        Set bulunan = Worksheets("General").Range("B:B").Find(number)
       
        
        If Not bulunan Is Nothing Then
        
         adress = bulunan.Address
        b = bulunan.Row
        adress2 = Worksheets("General").Cells(b, 16)
            
        
        Sheets("Marking-Compaction2").Hyperlinks.Add Anchor:=Sheets("Marking-Compaction2").Cells(a, 15), Address:=adress2, TextToDisplay:="PDF"
        
        
        
        End If
        
        If Worksheets("General").Cells(b, 16).Value = "IR-" & number & ".pdf" Then
        
            Cells(a, 15) = "NO PDF"
            
            End If
            
        
        a = a + 1
    
    Loop While Worksheets("Marking-Compaction2").Cells(a, 4) <> 0
    
    


End Sub

