Sub LSM9506_Automation()
    Dim COM As Integer
    Dim SerialPort As String
    Dim SerialSetting As New frmSerialSetting
    Dim receivedData As String
    Dim PinNumber As Integer

    'Show Serial Setting Form
    SerialSetting.Show
    'Wait for user to click OK
    'Set COM port number
    COM = SerialSetting.Serial_COM
    'Set Serial Port Setting
    SerialPort = SerialSetting.Serial_COM & ",""Baud=9600 Data=8 Parity=N Stop=1"
    'Set Pin Number
    PinNumber = InputBox("Please enter the number of pins to be tested", "Pin Number", 150)

    'Open Serial Port
    start_com_port(SerialPort)
    'Check serial port
    If wait_com_port(COM) != True Then
        MsgBox "Serial Communication Timeout"
        Exit Sub
    End If
    
    'Send Run Command
    WaitForNonErrorData(COM,"R")
    'Fill active cell with received data
    ActiveCell.Value = receivedData
    ActiveCell.Offset(0, 1).Select

    'Send Query Command
    for i = 1 to PinNumber
        WaitForNonErrorData(COM,"Q")
        'Fill active cell with received data
        ActiveCell.Value = receivedData
        ActiveCell.Offset(0, 1).Select
    Next i
    End Sub

    Function WaitForNonErrorData(COM As Integer, COMMAND As String) As String
        Dim receivedData As String
        receivedData = read_com_port(COM,9)
        If Left(receivedData,3) != "ER0" Then
            Do
                put_com_port(COM, COMMAND)
                receivedData = read_com_port(COM,9)
            Loop Until Left(receivedData,3) != "ER0"
        ElseIf Left(receivedData,2) = "ER" && Left(receivedData,3) != "ER0" Then
            MsgBox "Error: " & receivedData
            Exit Sub
        End If
        WaitForNonErrorData = receivedData
    End Function