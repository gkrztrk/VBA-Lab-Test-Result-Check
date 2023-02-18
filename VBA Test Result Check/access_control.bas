Attribute VB_Name = "access_control"
Sub accss_cntrl()
'Or Environ("username") = "mbabar"
If Environ("username") = "goker" Or Environ("username") = "shussain" Then

    If ThisWorkbook.ReadOnly = True Then
    
        ThisWorkbook.Saved = True
        ActiveWorkbook.ChangeFileAccess Mode:=xlReadWrite
        
    End If
Else

    If ThisWorkbook.ReadOnly = False Then
        ThisWorkbook.Saved = True
        ActiveWorkbook.ChangeFileAccess Mode:=xlReadOnly
    End If
    
End If

End Sub

