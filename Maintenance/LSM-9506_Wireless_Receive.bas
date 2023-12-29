Sub GetDataFromESP32OverLAN()
    Dim req As Object
    Set req = CreateObject("MSXML2.XMLHTTP")

    ' Send command to ESP32 over LAN (replace IP with ESP32's IP)
    req.Open "GET", "http://esp32_ip_address/command", False
    req.send

    If req.Status = 200 Then
        ' Get data from ESP32 and log it onto Excel
        Dim receivedData As String
        receivedData = req.responseText
        Sheets("Sheet1").Range("A1").Value = receivedData ' Modify cell as needed
    Else
        MsgBox "Error: " & req.Status & " - " & req.statusText
    End If
End Sub
