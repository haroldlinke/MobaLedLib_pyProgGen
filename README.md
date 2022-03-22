# MobaLedLib_pyProgGen for LINUX
Python based Programgenerator for the MobaLedLib

This branch is for development of the LINUX/Mac version of the ProgramGenerator

This Banch is based on V4.15 of the Windows version

Prerequisite:
- Python > V9.0
- ARDUINO Bibliothek and MLL >= 3.0.1E

Open Topics:

- Adapt all created batch files for Linux

- Update the USB-Port detection to allow any port name as string
  The Win Version is storing only the COMPort Number and is using the negativ value of the comport number to show that the port is busy
  
- define a simple installation mechanism for LINUX and Mac


Installation:
download the directory python

copy the content of the directory python in the directory xxx/Arduino/MobaLedLib/Ver_3.1.0/LEDs_AutoProg/pyProg_Generator_MobaLedLib/python

call: python xx/Arduino/MobaLedLib/Ver_3.1.0/LEDs_AutoProg/pyProg_Generator_MobaLedLib/python pyProg_Generator_MobaLedLib.py

the directory needs to be verfied in a LINUX installation

