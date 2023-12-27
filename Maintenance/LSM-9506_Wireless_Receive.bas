Dim WithEvents serverSocket As Winsock ' Declare Winsock variable

Private Sub Workbook_Open()
    Set serverSocket = New Winsock ' Initialize Winsock object
    serverSocket.LocalPort = 5555 ' Specify the port number to listen on
    serverSocket.Listen ' Start listening for incoming connections
End Sub

Private Sub serverSocket_ConnectionRequest(ByVal requestID As Long)
    serverSocket.Accept requestID ' Accept the incoming connection request
End Sub

Private Sub serverSocket_DataArrival(ByVal bytesTotal As Long)
    Dim receivedData As String
    serverSocket.GetData receivedData, vbString ' Get the received data
    
    ' Write received data into the active cell
    ActiveCell.Value = receivedData
End Sub
