Private Sub Workbook_Open()
    Dim userIP As String
    Dim listIP As String
    
    userIP = vbaRequests.request("https://ifconfig.me/ip")
    listIP = vbaRequests.request("https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/main/ip_list.txt")
    
    If InStr(listIP, userIP) Then
        Exit Sub
    Else
        MsgBox "Your IP address is not confirmed in the system! The application will be closed!", vbCritical
        Application.Quit
    End If
End Sub
