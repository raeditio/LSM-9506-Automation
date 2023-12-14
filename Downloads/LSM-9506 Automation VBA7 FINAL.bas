#If VBA7 Then
 Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems
#End If
Sub LSM9506_Automation()
    Dim COM As Long     ' COM Port No. selected by user
    Dim receivedData As String
    Dim PinNumber As Integer    ' Total number of pins to be entered in by user
    Dim PinCount As Integer
    ' Const RUN As String = "RN" & vbCr       ' Single-run measurement command without program number response
    Const QUERY As String = "DN" & vbCr       ' Data request command without program number response
    Dim CanProceed As Boolean
    Dim err As Integer      ' Error counter
    Dim average_count As Integer
    Dim RepeatNumber As Integer

    ' Show Serial Setting Form
    SerialSetting.Show
    ' Wait for user to click OK
    COM = SerialSetting.Serial_COM
    PinNumberSet.Show
    PinNumber = PinNumberSet.PinNumber
    RepeatNumber = PinNumberSet.RepeatNumber

    If START_COM_PORT(COM) = False Then
        MsgBox ("Connection failure. Please reconnect the device!")
        DISCONNECT(COM)
    End If
    
    PinCount = 0
    MsgBox ("Place the first pin on the LSM-9506.")
    
    PinCount = 0
    CanProceed = True
    err = 0
    average_count = 1
    
    
NEXT_QUERY:
    ' Send QUERY COMMAND
    If SEND_COM_PORT(COM, QUERY) = False Then
        MsgBox ("Please Check Connection")
    End If
    
    Sleep 200
    
    receivedData = FULL_READ_COM(COM)
    
    ' Measure Pin
    MeasurePin(receivedData)

    ' Check error counter
    ChcekError(err)

    ' Check if all pins are measured
    If PinCount < PinNumber Then
        GoTo NEXT_QUERY
    End If

    DISCONNECT(COM)
End Sub

' Functions
' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private  Function MeasurePin(receivedData As String)
    If Len(receivedData) = 7 Then       ' Valid data. Record measurement and move on
        ActiveCell.Value = (ActiveCell.Value * (average_count - 1) + receivedData) / average_count
        CanProceed = False      ' Wait for the user to move on to the next pin
        err = 0             ' Reset error counter
        average_count = average_count + 1       ' Increment average counter
        If average_count = RepeatNumber + 1 Then    ' Take the average of the measurements
            average_count = 1
            ActiveCell.Offset(1, 0).Select
            PinCount = PinCount + 1
        End If
    Else
        ' The device outputs "ER0" for no measurement. Thus, the error indicates that the user has removed the pin for next measurement
        If receivedData = "ER0" & vbCr Then
            CanProceed = True        ' Ready for next measurement
            err = 0        ' Reset error counter
        Else
            err = err + 1       ' Increment error counter
        End If
End Function

' Check error counter
Private  Function ChcekError(err)
    If err = 5 Then
        MsgBox ("ERROR: " & receivedData)
        DISCONNECT(COM) 
    End If
End Function

' Use this function to read COM port until all characters are received
Private  Function FULL_READ_COM(COM As Long)
    receivedData = READ_COM_PORT(COM)
    ' Wait until all characters are received
    If Left(receivedData, 2) <> ERROR Then
        While Len(receivedData) < 7
            receivedData = READ_COM_PORT(COM)
        Wend
    End If
End Function

' Disconnect the device
Public Function DISCONNECT(COM As Long)
    If STOP_COM_PORT(COM) = True Then
        MsgBox ("Measurement Complete" & vbCr & "Disconnected")
    End If
End Function