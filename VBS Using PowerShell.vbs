Dim outputDetected, skipDetected, receivedData
outputDetected = False
skipDetected = False
receivedData = ""

Do
    WScript.Sleep 2000 ' Wait for 2 seconds

    ' Read data from the serial port using PowerShell
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("powershell.exe -Command ""& { $port= new-Object System.IO.Ports.SerialPort 'COM1', 9600; $port.Open(); Start-Sleep -s 1; $receivedData = $port.ReadLine(); $port.Close(); $receivedData }""")
    
    receivedData = objExec.StdOut.ReadAll()
    
    If Len(receivedData) > 0 Then
        ' Log the received data here, for example:
        ' WScript.Echo "Received data: " & receivedData
        ' You can replace WScript.Echo with writing to a file, database, etc.
    Else
        WScript.Echo "No data received."
    End If
    
Loop