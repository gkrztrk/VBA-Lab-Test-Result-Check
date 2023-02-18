Attribute VB_Name = "Copy_Table"
Option Explicit
Sub copyTables()
cpyfiltered
InternalHyprlnk
End Sub

Sub cpyfiltered()
Dim arr As Variant
Dim arrfltrd As Variant
Dim lastrow, lr As Long
Dim calcLayercount As Variant

Test_Control.AutoFilter.ShowAllData

lr = Test_Control.Cells(Rows.Count, 2).End(xlUp).Row
Test_Control.Range("A2:M" & lr).Clear


arr = Sheet1.Range("A1").CurrentRegion.AutoFilter(9, Array("N", "NP", "P"), xlFilterValues, True)

lastrow = Sheet1.Cells(Rows.Count, 2).End(xlUp).Row
Sheet1.Range("A2:D" & lastrow).SpecialCells(xlCellTypeVisible).Copy

Test_Control.Range("A2").PasteSpecial xlPasteValues

Sheet1.Range("I2:I" & lastrow).SpecialCells(xlCellTypeVisible).Copy

Test_Control.Range("G2").PasteSpecial xlPasteValues


Sheet1.Range("K2:K" & lastrow).SpecialCells(xlCellTypeVisible).Copy

Test_Control.Range("I2").PasteSpecial xlPasteValues


Sheet1.Range("L2:L" & lastrow).SpecialCells(xlCellTypeVisible).Copy

Test_Control.Range("J2").PasteSpecial xlPasteValues




Sheet1.Range("F2:F" & lastrow).SpecialCells(xlCellTypeVisible).Copy

Test_Control.Range("E2").PasteSpecial xlPasteValues

Test_Control.Range("A:A").NumberFormat = "dd.mm.yyyy"
arrfltrd = Test_Control.Range("A1").CurrentRegion.Value

calcLayercount = Test_Control.Range("E2:E" & Test_Control.Cells(Rows.Count, 2).End(xlUp).Row).Value
calcLayercount = LayersInIR(calcLayercount)
Test_Control.Range("F2:F" & Test_Control.Cells(Rows.Count, 2).End(xlUp).Row) = calcLayercount

Sheet1.AutoFilter.ShowAllData

End Sub

Private Function LayersInIR(ByVal CalcRange As Variant) As Variant

'to find count of layers per IR

Dim a, b As String
Dim LayerCount As Long
Dim amk As Application

Dim i As Long
    '=IFERROR(ABS((RIGHT(F6077,LEN(F6077)-FIND("o",F6077)))-(LEFT(F6077,FIND("t",F6077)-1))),-1)+1
    For i = LBound(CalcRange) To UBound(CalcRange)
        a = CalcRange(i, 1)
   
       If CalcRange(i, 1) = 0 Then CalcRange(i, 1) = 1
       
        'LayerCount = Evaluate()
        If a <> "" And InStr(a, "to") <> 0 Then
            On Error Resume Next
            If IsNumeric(a) = False Then
            
                LayerCount = Int(Trim(Right(a, Len(a) - InStrRev(a, "o")))) _
                - Int(Trim((Left(a, InStr(a, "t") - 1))))
                
                If LayerCount < 0 Then
                    LayerCount = (LayerCount * -1) + 1
                ElseIf LayerCount > 0 Then
                    LayerCount = LayerCount + 1
                Else
                    LayerCount = 1
                End If
                CalcRange(i, 1) = LayerCount
            Else
                CalcRange(i, 1) = 1
            End If
        Else
            CalcRange(i, 1) = 1
        End If
    Next i
    LayersInIR = CalcRange
End Function
