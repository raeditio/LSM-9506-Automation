Dim outputDetected, skipDetected, receivedData
outputDetected = False
skipDetected = False
receivedData = ""

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set serialPort = CreateObject("MSCommLib.MSComm")
With serialPort
    .CommPort = 1 ' Replace with your COM port number
    .Settings = "9600,N,8,1" ' Replace with your serial settings
    .InputMode = comInputModeText
    .PortOpen = True
End With

Do
    WScript.Sleep 2000 ' Wait for 2 seconds
    
    If Not outputDetected Then
        ' Record the received data if no output was detected
        skipDetected = True
        receivedData = serialPort.Input
    ElseIf skipDetected Then
        ' Record the next data after an output is skipped
        receivedData = serialPort.Input
        skipDetected = False ' Reset the skip detection flag
        
        ' Initialize Excel and log the data into Column C starting from the selected row
        objExcel.Workbooks.Open "C:\path\to\your\file.xlsx" ' Replace with your Excel file path
        objExcel.Visible = True
        
        Set objSheet = objExcel.ActiveWorkbook.Sheets(1) ' Assuming data is logged in the first sheet
        selectedRow = objExcel.ActiveCell.Row
        
        objSheet.Cells(selectedRow, 3).Value = receivedData
        selectedRow = selectedRow + 1 ' Move to the next row
        
        objExcel.ActiveWorkbook.Save
        objExcel.ActiveWorkbook.Close
        objExcel.Quit
        Set objSheet = Nothing
        Set objExcel = Nothing
    End If
    
    outputDetected = False ' Reset the output detection flag
Loop
