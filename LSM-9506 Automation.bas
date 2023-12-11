Sub LSM9506_Automation()
    Dim COM As Long
    Dim SerialPort As String
    Dim receivedData As String
    Dim PinNumber As Integer
    Dim PinCount As Integer
    Dim Serial_Settings As String
    Dim WAIT_TIME As Long
    Dim WAIT As Boolean
    Const RUN As String = "R" & vbCr
    Const QUERY As String = "D" & vbCr
    Const ERROR As String = "ER"
    'The device outputs "ER0" for no measurement. Thus, the error indicates that the user has removed the pin for next measurement
    Const NEXT_MESUREMENT As String = "ER0"

    'Show Serial Setting Form
    SerialSetting.Show
    'Wait for user to click OK
    COM = SerialSetting.Serial_COM
    'Set Pin Number
    PinNumber = InputBox("Please enter the number of pins to be tested", "Pin Number", 3)
    
    If START_COM_PORT(COM) = False Then
        MsgBox ("Connection failure. Please reconnect the device!")
        Exit Sub
    End If
    
    PinCount = 0
    MsgBox ("Place the first pin on the LSM-9506.")
    
    'Start measurement
    WAIT_TIME = 700
    receivedData = WaitForNonErrorData(COM, RUN, WAIT_TIME)
    
    receivedData = WaitForNonErrorData(COM, QUERY, WAIT_TIME)
    'Assess measurement
    If Left(receivedData, 2) <> ERROR Then
        'Data is valid. Record data and move to the next row
        ActiveCell.Value = receivedData
        ActiveCell.Offset(1, 0).Select
        PinCount = PinCount + 1
        'Continue checking for NEXT_MEASUREMENT or ERROR
        If PinCount < PinNumber Then
            Sleep
            receivedData = WaitForNonErrorData(COM, RUN, WAIT_TIME)
            MsgBox (receivedData)
            End If
            'fLoop Until Left(receivedData, 2) = "ER"     'ERROR and NEXT_MEASUREMENT both begin with "ER"
        Else
            MsgBox ("Measurement Complete")
            Exit Sub
        End If
    'If NEXT_MEASUREMENT is receieved, wait for a measurement
    ElseIf Left(receivedData, 3) = NEXT_MEASUREMENT Then
        Do
            receivedData = WaitForNonErrorData(COM, COMMAND, WAIT_TIME)
        Loop Until Left(receivedData, 3) <> NEXT_MEASUREMENT
    'Else
    '    receivedData = READ_COM_PORT(COM, COMMAND, WAIT_TIME)
    '    MsgBox ("ERROR: " & receivedData)
    End If
    'Disconnect from Port
    If STOP_COM_PORT(COM) = True Then
        MsgBox ("Disconnected")
    End If
    End Sub

Function WaitForNonErrorData(COM As Long, COMMAND As String, WAIT_TIME As Long) As String
    Dim receivedData As String
    If SEND_COM_PORT(COM, COMMAND) = True Then
        If WAIT_COM_PORT(COM, WAIT_TIME) = True Then
            receivedData = READ_COM_PORT(COM)
        End If
    End If
    WaitForNonErrorData = Mid(receivedData, 4)
End Function

