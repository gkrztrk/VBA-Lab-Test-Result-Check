Attribute VB_Name = "test_log_doldur"
Sub log_doldur()

Dim a, b, c, d, e, f As Integer
a = 6077
b = 9
c = 2
d = 4
Dim kontrol As Range
Dim ir_number As String
Dim test_cinsi As Range
Dim test_tarihi As Range
Dim test_sonuc As String
Dim ilk_kontrol As Range
Dim kontrol_yeri As Range



Do

If Worksheets("General").Cells(a, 14).Value = "ok" Then

a = a + 1

    Else
If Cells(c, d).Value <> 0 Then

c = c + 1

Else


If Worksheets("General").Cells(a, b).Value = "N" Or Worksheets("General").Cells(a, b).Value = "P" Or Worksheets("General").Cells(a, b).Value = "N,P" Then

Cells(c, d).Value = Worksheets("General").Cells(a, b - 7)
Cells(c, d - 3).Value = Worksheets("General").Cells(a, b - 8)
Cells(c, d + 1).Value = Worksheets("General").Cells(a, b - 3)
Cells(c, d + 7).Value = Worksheets("General").Cells(a, b - 6)
Cells(c, d + 6).Value = Worksheets("General").Cells(a, b - 5)
Cells(c, d + 9).Value = Worksheets("General").Cells(a, b)
Cells(c, d - 2).Value = Worksheets("General").Cells(a, b - 1)
Cells(c, d + 12).Value = Worksheets("General").Cells(a, b + 4)
Worksheets("General").Cells(a, 14).Value = "ok"

    
a = a + 1

Else
a = a + 1


End If

    End If
    

    End If

Loop While Worksheets("General").Cells(a, b - 7).Value <> 0



End Sub

