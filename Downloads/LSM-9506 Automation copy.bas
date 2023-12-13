Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub LSM9506_Automation()
    Dim COM As Long     ' COM Port No. selected by user
    Dim receivedData As String
    Dim PinNumber As Integer    ' Total number of pins to be entered in by user
    Dim PinCount As Integer
    Const RUN As String = "RN" & vbCr       ' Single-run measurement command without program number response
    Const QUERY As String = "DN" & vbCr     ' Data request command without program number response
    Const ERROR As String = "ER"
    Const STATE_OK As Integer = 1           ' Ready to take next measurement state
    Const STATE_WAIT As Integer = 0         ' Wait for next item state
    Dim STATE As Integer
    ' The device outputs "ER0" for no measurement. Thus, the error indicates that the user has removed the pin for next measurement
    ' Const NEXT_MESUREMENT As String = "ER0"
    ' Dim SENT As Boolean
    Dim err As Integer      ' Error counter
    Dim average_count As Integer
    Dim RepeatNumber As Integer

    ' Show Serial Setting Form
    SerialSetting.Show
    ' Wait for user to click OK
    COM = SerialSetting.Serial_COM
    ' Set Pin Number
    ' PinNumber = InputBox("Please enter the number of pins to be tested", "Pin Number", 3)
    PinNumberSet.Show
    PinNumber = PinNumberSet.PinNumber
    RepeatNumber = PinNumberSet.RepeatNumber

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
    average_count = 1

    'EXITMACRO.Show
    'If EXITMACRO.disconnect_button = True Then
    '    GoTo DISCONNECT
    'End If
    
    
NEXT_QUERY:
    ' Send QUERY COMMAND
    If SEND_COM_PORT(COM, QUERY) = False Then
        MsgBox ("Please Check Connection")
    End If
    
    Sleep 200
    
    receivedData = READ_COM_PORT(COM)
    ' Wait until all characters are received
    If Left(receivedData, 2) <> ERROR Then
        While Len(receivedData) < 7
            receivedData = READ_COM_PORT(COM)
        Wend
    End If
    
    
    If Left(receivedData, 2) <> ERROR Then
        ' If at state ready for next measurement, take next measurement.
        If STATE = STATE_OK Then
            'Valid data. Record measurement and move on
            ActiveCell.Value = (ActiveCell.Value * (average_count - 1) + receivedData) / average_count
            STATE = STATE_WAIT
            err = 0
            average_count = average_count + 1
            If average_count = RepeatNumber + 1 Then
                average_count = 1
                ActiveCell.Offset(1, 0).Select
                PinCount = PinCount + 1
            End If

        End If
    Else
        ' Display error message
        If err = 5 Then
            MsgBox ("ERROR: " & receivedData)
            GoTo DISCONNECT
        End If
        ' Set state ready for next measurement
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
        MsgBox ("Measurement Complete" & vbCr & "Disconnected")
    End If
    
End Sub
