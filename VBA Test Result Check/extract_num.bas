Attribute VB_Name = "extract_num"
Function GetNumeric(CellRef As String)
Dim StringLength As Integer
StringLength = Len(CellRef)
For i = 1 To StringLength
If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
Next i
GetNumeric = result
End Function
