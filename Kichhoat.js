Private Sub CommandButton2_Click()

End Sub

Private Sub cmdActive_Click()
Dim myLicenseKey As String, myMessage As String
Dim dem As Integer
ActiveWorkbook.Save
myLicenseKey = txtKey.Value
Sheet1.Activate
If Check_LicenseKey(myLicenseKey) Then
        lblActive.Caption = "Kich Hoat Thanh Cong :D"
        Sheet1.Range("B2").FormulaR1C1 = txtKey.Value
        ActiveWorkbook.Save
    Else
        lblActive.Caption = " Sai ma kich hoat :D"
        Sheet1.Range("B2").FormulaR1C1 = txtKey.Value
              ActiveWorkbook.Save
            
       
        
    End If

End Sub

Private Sub cmdLamlai_Click()
txtKey = ""

End Sub



Private Sub CommandButton1_Click()
Dim myLicenseKey As String
myLicenseKey = txtKey.Value

If Check_LicenseKey(myLicenseKey) Then
        Unload Me
        
    Else
       ActiveWorkbook.Save
        ThisWorkbook.Close
    End If

End Sub

Private Sub txtKey_Change()
Dim dem As Integer
dem = dem + 1

End Sub



Private Sub UserForm_Deactivate()

    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub



Private Sub UserForm_Initialize()
Dim myLicenseKey As String, myMessage As String
Dim dem As Integer
    
    dem = 1
    myLicenseKey = txtKey.Value
    txtMamay = Get_UserSerial()

End Sub
