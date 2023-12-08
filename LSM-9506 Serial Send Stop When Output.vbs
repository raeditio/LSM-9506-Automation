Imports System.IO.Ports

Public Class Form1
    Dim WithEvents serialPort As New SerialPort("COM21", 9600) ' Replace "COM1" with your serial port and 9600 with your baud rate
    
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serialPort.Open()
        Catch ex As Exception
            MessageBox.Show("Error opening serial port: " & ex.Message)
        End Try
    End Sub
    
    Private Sub serialPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        Dim receivedData As String = serialPort.ReadExisting()
        
        ' Send the stop signal whenever any data is received
        Dim stopSignal As String = "CR" ' Replace "StopSignal" with your stop signal
        serialPort.Write(stopSignal)
    End Sub
    
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
    End Sub
End Class
