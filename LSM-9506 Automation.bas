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
    ' const STATUS_ERROR As Integer = 3
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
    
    'Assess measurement
    Do
        If receivedData = ERROR and receivedData <> NEXT_MESUREMENT Then
            ' Error. Stop.
            MsgBox ("Error. Please check the device!")
            Exit Sub
        ElseIf receivedData = NEXT_MESUREMENT Then
            ' No data. Move on.
            STATUS = STATUS_NEXT
        Else
            If STATUS = STATUS_NEXT Then
                ' Data is valid. Record measurement and move on.
                ActiveCell.Value = receivedData
                ActiveCell.Offset(1, 0).Select
                PinCount = PinCount + 1
                STATUS = STATUS_OK
            Else
                ' User has not progressed to the next pin. Wait.
        End If
        ' Receive next data output
        receivedData = WaitForNonErrorData(COM, QUERY, 2 * WAIT_TIME)
    Loop Until PinCount = PinNumber
            
        
    
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

