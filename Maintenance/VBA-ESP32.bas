Option Explicit

Private Sub StartServer()
    Dim objListener As Object
    Dim objClient As Object
    Dim strData As String
    Dim intPort As Integer

    intPort = 8888 ' Use the same port as defined in the Python/ESP32 code

    Set objListener = CreateObject("MSWinsock.Winsock")
    objListener.LocalPort = intPort
    objListener.Listen

    Do
        DoEvents
        If objListener.State = 7 Then ' Check if listener is listening
            If objListener.PendingEvents("DataArrival") > 0 Then
                Set objClient = objListener.Accept
                strData = objClient.GetData
                ' Modify this line to write data to your desired worksheet and cell
                Sheet1.Range("A1").Value = strData ' Change Sheet1 and A1 to your desired worksheet and cell
                objClient.Close
            End If
        End If
    Loop
End Sub
