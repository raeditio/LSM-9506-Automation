import serial

# Load template is determined by the esp32_gpio23 input
esp32_gpio23 = 1  # Replace with the appropriate value from ESP32 GPIO 23

# Open the serial port
ser = serial.Serial('COM1', 9600)  # Replace 'COM1' with the appropriate port and baud rate

# Send command
command = 'your_command_here'
ser.write(command.encode())  # Convert the command to bytes and send it

# Read the reply
reply = ser.read_until(b'\r')  # Read until the delimiter '\r' is encountered
reply = reply.decode()  # Convert the reply from bytes to string

# Close the serial port
ser.close()

# Print the reply
print(reply)

import openpyxl

# Load template is determined by the user button-press input
if esp32_gpio23 == 1:
    # Load the Excel template
    workbook = openpyxl.load_workbook('your_template.xlsx')

# Select the active sheet (you can also specify a sheet by name)
sheet = workbook.active

# Get the active cell
active_cell = sheet.active_cell

# Print information to the active cell
active_cell.value = 'Printed in active cell'

# Save the changes to a new file
workbook.save('updated_template.xlsx')
