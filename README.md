# MobaLedLib_pyProgGen Proof of Concept
Proof of Concept of Python based Programgenerator for the MobaLedLib


Requirements:
MobaledLib 3.1.0 has to be installed



Installation option 3: Python files - e.g. for LINUX and MAC
1. search for the MLL-subfolder LEDs_Autoprog - this folder must contain the file "LEDs_AutoProg.ino"
2. create a subfolder pyProg_Generator_MobaLedLib in the folder LEDs_Autoprog (the name of the subfolder can be any name)
3. create a subfolder Python
4. Clone MobaLedLib_PyProgGen to the folder pyProg_Generator_MobaLedLib/Python - the file pyProg_Generator_MobaLedLib.py must be in this folder
5. or download the branch as ZIP-file: unpack the ZIP file and copy the contents of the folder MobaLedLib_pyProgGen-4.0 into the folder pyProg_Generator_MobaLedLib/Python
6. open the folder pyProg_Generator_MobaLedLib/Python
7. start the Python file: pyProg_Generator_MobaLedLib.py
8. continue with first configuration described below

First configuration of PyProgGen:
1. open the tab "ARDUINO Einstellungen" - pyProgGen tries to find the connected ARDUINOs and determines the typ and COM port. If only one ARDUINO is found the com port is automatically selected. You can change this selection of the ARDUINO and the ARDUINO Type - save the changes - if the com-port is not included in the list it is possible to enter the port string by hand
2. For other OS than Windows select the Check-box "Individuellen Pfad zur ARDUINO IDE verwenden and select the path to the ARDUINO IDE - the name in Windows is "ARDUINO_DEBUG.exe" for all other OS "arduino"
3. Do not forget to save the changed parameters with the button "geänderte Einstellungen übernehmen"

Further information can be found in the MobaLedLib Wiki: https://wiki.mobaledlib.de/anleitungen/spezial/pyprogramgenerator


This Version is only a Proof of Concept for a Python based MLL-Programm Generator that simulates the UserInterface of the Excel based Program Generator.

The VBA code was translated 1 to 1 to Python using the Wedbased VB2PY-converter: http://vb2py.sourceforge.net/online_conversion.html

The Tables are based on Tkintertable: https://github.com/dmnfarrell/tkintertable

Only the "Dialog"-Button and the "Send-to-ARDUINO"-Button are implemented yet.
The main Dialogs are implemented focusing on the main way to enter data and create the MACRO definition in the table.

It is possible to use the dialog to create House/Gaslights Macros and Macros using the Generic Form UserFormOther. Editing of Macros via Dialog should be possible too.
Sending of Macros to ARDUINO is possible. The headerfile creation is using the original VBA logic. The sending to the ARDUINO is done using the Pytho based logic on the pyProgramGenerator. Check the ARDUINO Einstellungen - Page if the ARDUINO is recognized correctly.

Limitations:
There is no save or laod command for the complete data. There is also NO check if the data was saved.
It is possible to save the tables individually by right-clicking in the table. Select: File->Save.

Editing individual cells is possible. There is no automatic update of the LEDNr columns yet.

There are a lot of hidden issues with sytactical/logical differences between VBA and Python that will cause crashes or wrong behavior:
handling of global variables
handling of by_ref parameters
construction using automatic type translation: e.g. string = string + integer + string






