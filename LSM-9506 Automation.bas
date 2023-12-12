Sub LSM9506_Automation()
    Dim COM As Long
    Dim receivedData As String
    Dim PinNumber As Integer
    Dim PinCount As Integer
    Const RUN As String = "RN" & vbCr
    Const QUERY As String = "DN" & vbCr
    Const ERROR As String = "ER"
    Const STATE_OK As Integer = 1
    Const STATE_WAIT As Integer = 0
    Dim STATE As Integer
    ' The device outputs "ER0" for no measurement. Thus, the error indicates that the user has removed the pin for next measurement
    ' Const NEXT_MESUREMENT As String = "ER0"
    Dim SENT As Boolean
    Dim err As Integer
    ' Dim disconnect_button As Boolean

    'Show Serial Setting Form
    SerialSetting.Show
    'Wait for user to click OK
    COM = SerialSetting.Serial_COM
    'Set Pin Number
    PinNumber = InputBox("Please enter the number of pins to be tested", "Pin Number", 3)
    
    If START_COM_PORT(COM) = False Then
        MsgBox ("Connection failure. Please reconnect the device!")
        GoTo DISCONNECT
    End If
    
    PinCount = 0
    MsgBox ("Place the first pin on the LSM-9506.")
    
    'If SEND_COM_PORT(COM, RUN) = False Then
    '    MsgBox ("Please Check Connection")
    'End If
    'Application.WAIT (Now + #12:00:03 AM#)
    
    PinCount = 0
    STATE = STATE_OK
    err = 0
    
    EXITMACRO.Show
    If EXITMACRO.disconnect_button = True Then
        GoTo DISCONNECT
    End If
    
    
NEXT_QUERY:
    If SEND_COM_PORT(COM, QUERY) = False Then
        MsgBox ("Please Check Connection")
    End If
    
    receivedData = READ_COM_PORT(COM)
    If Left(receivedData, 2) <> ERROR Then
        While Len(receivedData) < 7
            receivedData = READ_COM_PORT(COM)
        Wend
    End If
    
    If Left(receivedData, 2) <> ERROR Then
        If STATE = STATE_OK Then
            'Valid data. Record measurement and move on
            ActiveCell.Value = receivedData
            ActiveCell.Offset(1, 0).Select
            PinCount = PinCount + 1
            STATE = STATE_WAIT
            err = 0
        End If
    Else
        If err = 5 Then
            MsgBox ("ERROR: " & receivedData)
            GoTo DISCONNECT
        End If
        
        If receivedData = "ER0" & vbCr Then
        ' Wait for next pin to be placed
        STATE = STATE_OK
        err = 0
        Else
        ' ERROR
            err = err + 1
        End If
    End If
    
    If PinCount < PinNumber Then
        GoTo NEXT_QUERY
    End If
    
DISCONNECT:
    If STOP_COM_PORT(COM) = True Then
        MsgBox ("Disconnected")
    End If
    
End Sub
