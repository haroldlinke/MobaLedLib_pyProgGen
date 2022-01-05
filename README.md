# MobaLedLib_pyProgGen Proof of Concept
Python based Programgenerator for the MobaLedLib

Version für den Betatest.

Requirements:
MobaledLib 2.1.3 has to be installed

Installation option 1: Via MLL Program Generator Excel file
1. Open the MLL Prog_Generator Excel file
2. Click on the "Optionen"-Button
3. Click on "LED_Farbtest starten" - this action downloads the latest pyProgen program if not already downloaded
4. pyProgGen starts with the "LED Farbtest"-page
5. you can now select the other function tabs of PyProggen

Installation option 2: Exe File
1. search for the MLL-subfolder LEDs_Autoprog - this folder must contain the file "LEDs_AutoProg.ino"
2. create a subfolder pyProg_Generator_MobaLedLib in the folder LEDs_Autoprog (the name of the subfolder can be any name)
3. download the pyProg_Generator_MobaLedLib.exe file to the Subfolder LEDs_Autoprog\pyProg_Generator_MobaLedLib
4. open the folder MobaLedLib\V2.x.xpyLEDs_Autoprog\pyProg_Generator_MobaLedLib
5. start the program file: pyProg_Generator_MobaLedLib.exe
6. continue with first configuration described below

Installation option 3: Python files - e.g. for LINUX and MAC
1. search for the MLL-subfolder LEDs_Autoprog - this folder must contain the file "LEDs_AutoProg.ino"
2. create a subfolder pyProg_Generator_MobaLedLib in the folder LEDs_Autoprog (the name of the subfolder can be any name)
3. Clone MobaLedLib_PyProgGen to the folder pyProg_Generator_MobaLedLib - the file pyProg_Generator_MobaLedLib.py must be in this folder
4. open the folder pyProg_Generator_MobaLedLib
5. start the Python file: pyProg_Generator_MobaLedLib.py
6. continue with first configuration described below

First configuration of PyProgGen:
1. open the tab "ARDUINO Einstellungen" - pyProgGen tries to find the connected ARDUINOs and determines the typ and COM port. If only one ARDUINO is found the com port is automatically selected. You can chnage this selection of the ARDUINO and the ARDUINO Type - save the changes - if the com-port is not included in the list it is possible to enter the port string by hand
2. For other OS than Windows select the Check-box "Individuellen Pfad zur ARDUINO IDE verwenden and select the path to the ARDUINO IDE - the name in Windows is "ARDUINO_DEBUG.exe" for all other OS "arduino"
3. Do not forget to save the changed parameters with the button "geänderte Einstellungen übernehmen"

Further information can be found in the MobaLedLib Wiki: https://wiki.mobaledlib.de/anleitungen/spezial/pyprogramgenerator
