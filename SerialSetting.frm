VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SerialSetting 
   Caption         =   "Serial Port Setting"
   ClientHeight    =   3190
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2260
   OleObjectBlob   =   "SerialSetting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SerialSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Run_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    With Serial_COM
        For i = 1 To 32
            .AddItem "COM" & i
        Next i
    End With
    
    With Serial_Baud
        .AddItem "9600"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "57600"
        .AddItem "115200"
    End With

    With Serial_Parity
        .AddItem "None"
        .AddItem "Odd"
        .AddItem "Even"
    End With

    With Serial_DataBit
        .AddItem "7"
        .AddItem "8"
    End With

    With Serial_StopBit
        .AddItem "0"
        .AddItem "1"
    End With

    With Serial_Handshake
        .AddItem "None"
        .AddItem "Xon/Xoff"
        .AddItem "RTS/CTS"
        .AddItem "RTS/CTS + Xon/Xoff"
    End With

    With Serial_Delimiter
        .AddItem "CR"
        .AddItem "LF"
        .AddItem "CR/LF"
    End With
End Sub
