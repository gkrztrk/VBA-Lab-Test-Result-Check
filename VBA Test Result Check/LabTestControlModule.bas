Attribute VB_Name = "LabTestControlModule"
Sub LabTestControlDummy()

Dim Layer3 As String
Dim Layer2 As String
Dim Layer1 As String

Dim L, a, b, c, d, t As Integer



Dim kontrol, kontrol_yeri, bulunan As Range

Dim ilk_adres, test_cinsi, test_sonucu, test_cinsi_aytek, ir_no, test_tarihi As String
Dim wss As Worksheet
Dim wsd As Worksheet


Set wss = LabTestLog
Set wsd = Test_Control

a = 2
b = 4


Do


If wsd.Cells(a, 8) = "Kapali" Then

a = a + 1
Else

    ir_no = wsd.Cells(a, 2).Value
    test_cinsi = wsd.Cells(a, 7).Value
    L = wsd.Cells(a, 6).Value
    
 '///////////////////////////////NUCLEAR////////////////////
 
 If test_cinsi = "N" Then
        
        t = 11
        
        Set bulunan = wss.Range("B:B").Find(ir_no)
        
         'Cells(a, 14).Value = ""
        If bulunan Is Nothing Then
        
            'Cells(a, 14).Value = "Test bulunamadi"
                 
        Else
              
            ilk_adres = bulunan.Address
        
        
              Do
              
                  c = bulunan.Row
                  d = bulunan.Column
                  test_cinsi_aytek = wss.Cells(c, 4).Value
                  test_tarihi = wss.Cells(c, 1).Value
                  test_sonucu = wss.Cells(c, 6).Value
               
                  
                  If test_sonucu = "COMPLY" Then
                  
                      wsd.Cells(a, t) = test_tarihi
                      t = t + 1
        
                  End If
                  
                 Set bulunan = wss.Range("B:B").FindNext(bulunan)
                 
                  
              
              Loop While ilk_adres <> bulunan.Address
        
        
     End If
        
    End If
  '///////////////////////////PLATE////////////////////////////
    
    If test_cinsi = "P" Then
        
        t = 11
        
        Set bulunan = wss.Range("B:B").Find(ir_no)
        
         'Cells(a, 14).Value = ""
        If bulunan Is Nothing Then
        
            'Cells(a, 14).Value = "Test bulunamadi"
                 
        Else
              
            ilk_adres = bulunan.Address
        
        
              Do
              
                  c = bulunan.Row
                  d = bulunan.Column
                  test_cinsi_aytek = wss.Cells(c, 4).Value
                  test_tarihi = wss.Cells(c, 1).Value
                  test_sonucu = wss.Cells(c, 6).Value
               
                  
                  If test_sonucu = "COMPLY" Then
                  
                      wsd.Cells(a, t) = test_tarihi
                      t = t + 1
        
                  End If
                  
                 Set bulunan = wss.Range("B:B").FindNext(bulunan)
                 
                  
              
              Loop While ilk_adres <> bulunan.Address
        
        
     End If
        
    End If
    
    
    If test_cinsi = "P" Then
        
        t = 11
        
        Set bulunan = Worksheets("test kontrol 2").Range("B:B").Find(ir_no)
         'Cells(a, 14).Value = ""
        If bulunan Is Nothing Then
        
            'Cells(a, 14).Value = "Test bulunamadi"
                 
        Else
        
        
        
        ilk_adres = bulunan.Address
        
        
        Do
        
            c = bulunan.Row
            d = bulunan.Column
            test_cinsi_aytek = Worksheets("test kontrol 2").Cells(c, d + 3).Value
            test_tarihi = Worksheets("test kontrol 2").Cells(c, d + 1).Value
            test_sonucu = Worksheets("test kontrol 2").Cells(c, d + 4).Value
            
            If test_sonucu = "APPROVED" Then
            
                Cells(a, t) = test_tarihi
                t = t + 1
                
            
            End If
            
           Set bulunan = Worksheets("test kontrol 2").Range("B:B").FindNext(bulunan)
           
            
        
        Loop While ilk_adres <> bulunan.Address
        
        
     End If
        
    End If
'//////////////////////////////*****NUCLEAR AND PLATE*****//////////////////////////
    If test_cinsi = "NP" Then
    
   t = 11

         'Cells(a, 14).Value = ""
         
        Set bulunan = wss.Range("B:B").Find(ir_no)
        
        If bulunan Is Nothing Then
        
            'Cells(a, 14).Value = "Test bulunamadi"
        
        Else
        
        
        ilk_adres = bulunan.Address
        
        
        Do
        
            c = bulunan.Row
            d = bulunan.Column
            test_cinsi_aytek = wss.Cells(c, 4).Value
            test_tarihi = wss.Cells(c, 1).Value
            test_sonucu = wss.Cells(c, 6).Value
            
            
            
         If L = 1 Then
         
         
                If test_sonucu = "COMPLY" And Layer1 = wss.Cells(c, 3).Value Then
                
                    Cells(a, t).Value = test_tarihi
                    
                  End If
                
                    
                If test_sonucu = "COMPLY" And Layer1 = "" Then
                
                
                    
                    Layer1 = wss.Cells(c, 3).Value
                       
                  End If
                
                
       ElseIf L = 2 Then
       
                If test_sonucu = "COMPLY" And Layer1 = wss.Cells(c, 3).Value Then
                
                    Cells(a, t).Value = test_tarihi
                    
                  End If
                
                    
                If test_sonucu = "COMPLY" And Layer1 = "" Then
                
                
                    
                    Layer1 = wss.Cells(c, 3).Value
                       
                End If
                
       
                
                If test_sonucu = "COMPLY" And Layer2 = wss.Cells(c, 3).Value And Layer2 <> Layer1 Then
                
                    Cells(a, t + 1) = test_tarihi
                    
                End If
                
                If test_sonucu = "COMPLY" And Layer2 = "" And wss.Cells(c, 3).Value <> Layer1 Then
                

                    Layer2 = wss.Cells(c, 3).Value
                 End If
                        
       ElseIf L = 3 Then
                        
                        
                        
                   If test_sonucu = "COMPLY" And Layer1 = wss.Cells(c, 3).Value Then
                
                    Cells(a, t).Value = test_tarihi
                    
                  End If
                
                    
                If test_sonucu = "COMPLY" And Layer1 = "" Then
                
                
                    
                    Layer1 = wss.Cells(c, 3).Value
                       
                  End If
                
       
                
                If test_sonucu = "COMPLY" And Layer2 = wss.Cells(c, 3).Value And Layer2 <> Layer1 And Layer2 <> Layer3 Then
                
                    Cells(a, t + 1) = test_tarihi
                    
                End If
                
                If test_sonucu = "COMPLY" And Layer2 = "" And wss.Cells(c, 3).Value <> Layer1 And wss.Cells(c, 3).Value <> Layer3 Then
                

                    Layer2 = wss.Cells(c, 3).Value
                 End If
                        
                        
                        
                        
                If test_sonucu = "COMPLY" And Layer3 = wss.Cells(c, 3).Value And Layer3 <> Layer1 And Layer2 <> Layer3 Then
                
                    Cells(a, t + 2) = test_tarihi
                    
                End If
                If test_sonucu = "COMPLY" And Layer3 = "" And wss.Cells(c, 3).Value <> Layer1 And wss.Cells(c, 3).Value <> Layer2 Then
                
                    
                    Layer3 = wss.Cells(c, 3).Value
                
                    
                End If
        End If
        
                
            
            Set bulunan = wss.Range("B:B").FindNext(bulunan)
            
            
        Loop While ilk_adres <> bulunan.Address
            
            End If
    End If
        
        Layer1 = ""
        Layer2 = ""
        Layer3 = ""
        
        
        If L = 1 And wsd.Cells(a, 11) <> "" Then
        
        wsd.Cells(a, 8) = "Closed"
        
        wsd.Cells(a, 8).Interior.ColorIndex = 3
        
        ElseIf L = 2 And wsd.Cells(a, 12) <> "" And wsd.Cells(a, 11) <> "" Then
        
         wsd.Cells(a, 8) = "Closed"
         wsd.Cells(a, 8).Interior.ColorIndex = 3
        
        ElseIf L = 3 And Cells(a, 11) <> "" And Cells(a, 12) <> "" And Cells(a, 13) <> "" Then
        
         wsd.Cells(a, 8) = "Closed"
         wsd.Cells(a, 8).Interior.ColorIndex = 3
         
         Else
         
         wsd.Cells(a, 8) = "Open"
         wsd.Cells(a, 8).Interior.ColorIndex = 4
         
         End If

    a = a + 1
End If


Loop While wsd.Cells(a, 2) <> 0
    
    
    

End Sub


















































Sub SingleControl()

Dim wss As Worksheet
Dim wsd As Worksheet
Dim arr, arr2 As Variant
Dim coll As New Collection
Dim IR As IrTest
'///////////////////////////////SETTER/////////////////////////////////////

Set wsd = Test_Control
Set wss = LabTestLog
lr = wss.Cells(Rows.Count, 2).End(xlUp).Row
arr = wss.Range("A12300:F" & lr).Value

For i = LBound(arr, 1) To UBound(arr, 1)

    Set IR = New IrTest
    
    For j = LBound(arr, 2) To UBound(arr, 2)
        On Error Resume Next
        Select Case j
            Case 1
                IR.setTarih (arr(i, j))
            Case 2
                IR.setIrNo (arr(i, j))
            Case 3
                IR.setLayerNo (arr(i, j))
            Case 4
                IR.SetTestType (arr(i, j))
            Case 5
                IR.SetTestResult (arr(i, j))
            Case 6
                IR.SetIrStatus (arr(i, j))
        End Select
        
    Next j
    coll.Add IR
Next i
'//////////////////////////////////////GETTER///////////////////////////////

lr2 = wsd.Cells(Rows.Count, 2).End(xlUp).Row
arr2 = wsd.Range("B2:B" & lr).Value

Debug.Print coll.Count

For i = LBound(arr2, 1) To UBound(arr2, 1)

    For Each IR In coll
        
        If IR.getIrNo = arr2(i, 1) Then
        
            Debug.Print IR.getTestType
            Debug.Print IR.getIrNo
            Debug.Print IR.getIrStatus
            
            typ = IR.getTestType
            stts = IR.getIrStatus
        
            If typ = CStr("NUCLEAR") And stts = CStr("COMPLY") Then
            
                wsd.Cells(i + 1, 8).Value = "HALAS"
            
            End If
        
        End If
        
    Next
    
Next i

End Sub
