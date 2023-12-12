# This script intends to read data from the LSM-9506 device through serial communication using the RS-232 protocol.
# When either the Run or Query command is sent to the device, the device will return a string of data.
# The data will be parsed and the value of the data will be displayed in the Excel spreadsheet.
# If the device returns an error, the error will be displayed in the console.
# The automated data acquisition algorithm utilizes the "ER0" signal that is returned by the device when no measurement is taken
# at the moment of the query. When the user of the device removes the pin for to proceed with the next measurement, the device
# will return the "ER0" signal. The algorithm will continue to query the device until the "ER0" signal is no longer returned.

import xlwings as xw
import serial

RUN = 'R'   # Run command
QUERY = 'D'   # Query command

try:
    # open TDS
    wb = xw.Book()
except IOError:
    # return error if file not accessible
    print("File not accessible")
    exit(1)
# open sheet 1
sheet = wb.sheets['Sheet1']

# Prompt port number
port = (input("Enter port number: ") or "21")
# Open serial port
SerialPort = serial.Serial(port)
# Prompt pin number
pinNumber = int(input("Enter the total number of pins in the set: ") or 1)

# Initilize the Run command
command = RUN
s.write(command)
# Read data specified by the number of bytes
data = s.read(9)
for i in range(1, pinNumber):
    if data[:2] != "ER":    
        sheet["C"+i].value = data[-7]
    if data[:2] == "ER" and (data[3] == "0" or data[3] == "9"): # Check if the device is ready to take a measurement. "ER9" signifies a timeout error.
        while data[:2] == "ER":
            # Initilize the Query command
            command = QUERY
            s.write(command)
            # Read data specified by the number of bytes
            data = s.read(9)
        # Record measurement data
    else:
        # Return error in console
        print("Error: %s. Please reconnect the device and try again." % data)
        break
# Close serial port
close(SerialPort)
wb.save(r'C:\Users\phili\Downloads\LSM-9506 Excel Automation.xlsm')
exit(0)