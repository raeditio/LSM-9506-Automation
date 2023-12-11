Sub LSM9506_Automation()
    Dim COM As Long
    Dim SerialPort As String
    Dim receivedData As String
    Dim PinNumber As Integer
    Dim Serial_Settings As String
    Dim COMMAND As String

    'Show Serial Setting Form
    SerialSetting.Show
    'Wait for user to click OK
    COM = SerialSetting.Serial_COM
    'Set Pin Number
    PinNumber = InputBox("Please enter the number of pins to be tested", "Pin Number", 150)
    
    If START_COM_PORT(COM) = False Then
        MsgBox ("Failed to connect to device!")
    End If
    MsgBox ("Place the first pin on the LSM-9506.")
    'Set Run Command
    COMMAND = "R" & vbCr
    'Send Run Command
    receivedData = WaitForNonErrorData(COM, COMMAND)
    'Fill active cell with received data
    ActiveCell.Value = receivedData
    'Shift to next row
    ActiveCell.Offset(0, 1).Select
    'Check if user placed a different pin. The device outputs an ER0, if at any point in time no item is being measured.
    COMMAND = "D" & vbCr
    receivedData = WaitForNonErrorData(COM, COMMAND)
    If Left(receivedData, 3) <> "ER0" Then
        Do
            receivedData = READ_COM_PORT(COM, COMMAND)
        Loop Until Left(receivedData, 3) = "ER0"
    ElseIf Left(receivedData, 2) = "ER" Then
        MsgBox ("Error: " & receivedData)
    Else
    
    
    End If
    
    
    ActiveCell.Value = receivedData
    ActiveCell.Offset(1, 0).Select
    
    'Send Query Command
    'For i = 1 To PinNumber
    '    receivedData = WaitForNonErrorData(COM, "Q")
    '    'Fill active cell with received data
    '    ActiveCell.Value = receivedData
    '    ActiveCell.Offset(0, 1).Select
    'Next i
    If STOP_COM_PORT(COM) = True Then
        MsgBox ("Disconnected")
    End If
    End Sub

Function WaitForNonErrorData(COM As Long, COMMAND As String) As String
    Dim receivedData As String
    If SEND_COM_PORT(COM, COMMAND) = True Then
        If WAIT_COM_PORT(COM, 10000) = True Then
            receivedData = Mid(READ_COM_PORT(COM), 4)
        End If
    End If
    WaitForNonErrorData = receivedData
End Function
