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
    const STATUS_OK As Integer = 1
    const STATUS_NEXT As Integer = 2
    Dim err As Integer
    Dim STATUS As Integer
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
    STATUS = STATUS_NEXT
    
    receivedData = WaitForNonErrorData(COM, RUN, WAIT_TIME)

    'ER9 handling block
    For err = 0 To 2
        If Left(receivedData, 2) = ERROR Then
            receivedData = WaitForNonErrorData(COM, RUN, WAIT_TIME)
        Else
            err = 0
            Exit While
        End If
    Next err
    If err = 2 Then
        MsgBox ("Error. Please check the device!")
        Exit Sub
    End If
    
    'Assess measurement
    While PinCount < PinNumber
        If receivedData <> ERROR
            If STATUS = STATUS_NEXT Then
                ' Ready for measurement
                ActiveCell.Value = receivedData
                ActiveCell.Offset(1, 0).Select
                PinCount = PinCount + 1
                STATUS = STATUS_OK
            End If
        Elseif receivedData = NEXT_MESUREMENT
            ' Wait for the user to place next pin
            STATUS = STATUS_NEXT
        Else
            ' Error handling
            MsgBox ("Error. Please check the device!")
            Exit Sub
        End If
        receivedData = WaitForNonErrorData(COM, QUERY, 2 * WAIT_TIME)
    Wend      
    
    MsgBox ("Measurement Complete")
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

