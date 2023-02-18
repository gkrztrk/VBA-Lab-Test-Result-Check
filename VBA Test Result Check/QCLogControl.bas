Attribute VB_Name = "QCLogControl"
Sub QCLogCntrl()
DictionaryVLookupFinal_Code
DictionaryVLookup2_Reason
cpyfiltered
InternalHyprlnk
End Sub


Sub DictionaryVLookupFinal_Code()

Dim x, x2, x3, y, y2()

Dim dict As New Dictionary
'Set dict = CreateObject("Scripting.Dictionary")
Dim wsQC, wsEW As Worksheet

Dim sPath As String

Set wsQC = QCLog
Set wsEW = ThisWorkbook.Worksheets("General")

With wsQC
    lr = .Cells(Rows.Count, "A").End(xlUp).Row
    x = .Range("A2:A" & lr).Value2
    x2 = .Range("C2:C" & lr).Value2
 
End With

For i = 1 To UBound(x, 1)

    dict.Item(x(i, 1)) = x2(i, 1)
    
    
Next i

lr2 = wsEW.Cells(Rows.Count, "B").End(xlUp).Row
y = wsEW.Range("B2:B" & lr2).Value2

ReDim y2(1 To UBound(y, 1), 1 To 1)

For i = 1 To UBound(y, 1)

    If dict.Exists(y(i, 1)) Then
    
        y2(i, 1) = dict(y(i, 1))
        
        
    Else
    
        y2(i, 1) = ""
        
    End If
    
Next i

wsEW.Range("K2:K" & lr2).Value = y2

Set dict = Nothing

        


End Sub
Sub DictionaryVLookup2_Reason()

Dim x, x2, x3, y, y2()

Dim dict As New Dictionary
'Set dict = CreateObject("Scripting.Dictionary")
Dim wsQC, wsEW As Worksheet

Dim sPath As String

Set wsQC = QCLog
Set wsEW = ThisWorkbook.Worksheets("General")

With wsQC
    lr = .Cells(Rows.Count, "A").End(xlUp).Row
    x = .Range("A2:A" & lr).Value2
    x2 = .Range("D2:D" & lr).Value2
    
End With

For i = 1 To UBound(x, 1)

    dict.Item(x(i, 1)) = x2(i, 1)
    
    
Next i

lr2 = wsEW.Cells(Rows.Count, "B").End(xlUp).Row
y = wsEW.Range("B2:B" & lr2).Value2

ReDim y2(1 To UBound(y, 1), 1 To 1)

For i = 1 To UBound(y, 1)

    If dict.Exists(y(i, 1)) Then
    
        y2(i, 1) = dict(y(i, 1))
        
        
    Else
    
        y2(i, 1) = ""
        
    End If
    
Next i

wsEW.Range("L2:L" & lr2).Value = y2

Set dict = Nothing

        


End Sub

