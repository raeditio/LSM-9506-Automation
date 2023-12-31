## Table of Contents  
- [Debrief](#headers)
- [Method](#headers)
- [Implementation](#headers)
- [How to Use](#headers)

# Debrief
This project intends to read data from the Mitutoyo LSM-9506 laser micrometer device through serial communication using the RS-232 protocol and record the measurements on a custom MS Excel spreadsheet.
See the **Use** Section below on how to use.
When either the Run or Query command is sent to the device, the device will return a string of data. The data will be parsed and the value of the data will be displayed in the Excel spreadsheet.
- If the device returns an error, the error will be displayed in the console.

# Method
The automated data acquisition algorithm utilizes the "ER0" signal that is returned by the device when no measurement is taken at the moment of the query. When the user of the
device removes the pin for to proceed with the next measurement, the device will return the "ER0" signal. The algorithm will continue to query the device until the "ER0" signal
is no longer returned.

# Implementation
To automate the process, an excel VBA script was written to minimize the number of prerequisiste steps that needs to be achieved before the automation can be run swiftly.

# How to Use
To use the script, Download and save files from the "Download" Folder to a convenient location on your computer.
Open Excel and enter Visual Basic (Alt+F11). On the left hand project section, right click under "VBAProject" and click Import. Import all the downloaded files. Alt+Q
to close the Visual Basics Window. Either create a shortcut for the macro or run it using the macro button in Excel under "View." When executed, the macro will prompt the user
for the COM port number. Enter the correct COM port number the LSM-9506 is connected to and check that the COM port settings on the device and on the computer that is connected
to the device are matching to the settings provided in the prompt window. Press confirm. Now the user will be asked the number of pins to be tested and the number of measurements
that will be taken for each pin. Enter the desired values and press confirm. The program asks the user to place the first pin on the device. Follow the instruction and measure. To terminate the macro before all measurements are taken, either the Ctrl + Break shortcut or the ButtonStop (See <ins>Optional</ins> below) can be used.
Each time a pin is removed and the device displays a new output, the measurement is recorded onto the Excel spreadsheet. If the user has selected to take more than one
measurement per pin, the average of the measurements will be printed in the cell. After the number of pins the user has entered are measured, the program disconnects from the COM
port and terminates. <ins>Optional</ins>: Add custom macro buttons to your sheet and assign each of them to the LSM-9506 Automation and the ButtonStop macros to start and stop the macro. This will
prevent the problem where the COM port remains connected after the macro is prematurely terminated.
