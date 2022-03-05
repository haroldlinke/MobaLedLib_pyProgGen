# -*- coding: utf-8 -*-
#
#         Write header
#
# * Version: 4.02
# * Author: Harold Linke
# * Date: January 7, 2021
# * Copyright: Harold Linke 2021
# *
# *
# * MobaLedCheckColors on Github: https://github.com/haroldlinke/MobaLedCheckColors
# *
# *  
# * https://github.com/Hardi-St/MobaLedLib
# *
# * MobaLedCheckColors is free software: you can redistribute it and/or modify
# * it under the terms of the GNU General Public License as published by
# * the Free Software Foundation, either version 3 of the License, or
# * (at your option) any later version.
# *
# * MobaLedCheckColors is distributed in the hope that it will be useful,
# * but WITHOUT ANY WARRANTY; without even the implied warranty of
# * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# * GNU General Public License for more details.
# *
# * You should have received a copy of the GNU General Public License
# * along with this program.  if not, see <http://www.gnu.org/licenses/>.
# *
# *
# ***************************************************************************

#------------------------------------------------------------------------------
# CHANGELOG:
# 2020-12-23 v4.01 HL: - Inital Version converted by VB2PY based on MLL V3.1.0
# 2021-01-07 v4.02 HL: - Else: check done, first PoC release


from vb2py.vbfunctions import *
from vb2py.vbdebug import *
from vb2py.vbconstants import *

import tkinter as tk
from tkinter import ttk
from mlpyproggen.tooltip import Tooltip
from tkcolorpicker.spinbox import Spinbox
from tkcolorpicker.limitvar import LimitVar



import proggen.M02_Public as M02
#import proggen.M03_Dialog as M03
#import proggen.M06_Write_Header as M06
#import proggen.M06_Write_Header_LED2Var as M06LED
#import proggen.M06_Write_Header_Sound as M06Sound
#import proggen.M06_Write_Header_SW as M06SW
import proggen.M07_COM_Port as M07
import proggen.M08_ARDUINO as M08
import proggen.M08_Install_FastBootLoader as M08IN
#import proggen.M09_Language as M09
#import proggen.M09_Select_Macro as M09SM
#import proggen.M09_SelectMacro_Treeview as M09SMT
#import proggen.M10_Par_Description as M10
#import proggen.m17_Import_old_Data as M17
import proggen.M18_Save_Load as M18
#import proggen.M20_PageEvents_a_Functions as M20
import proggen.M25_Columns as M25
import proggen.M27_Sheet_Icons as M27
import proggen.M28_divers as M28
import proggen.M30_Tools as M30
#import proggen.M31_Sound as M31
import proggen.M37_Inst_Libraries as M37
#import proggen.M60_CheckColors as M60
#import proggen.M70_Exp_Libraries as M70
#import proggen.M80_Create_Mulitplexer as M80

from ExcelAPI.X01_Excel_Consts import *
import ExcelAPI.P01_Workbook as P01
import proggen.Prog_Generator as PG
import proggen.M09_Language as M09

import logging

from vb2py.vbfunctions import *
from vb2py.vbdebug import *


__Disable_Set_Arduino_Typ = Boolean()

class CUserForm_Options:
    
    def __init__(self):
        self.controller = PG.get_global_controller()
        self.IsActive = False
        self.button1_txt = "Abbrechen"
        self.button2_txt = "Ok"
        self.res = False
        self.UserForm_Res = ""
        #self.__UserForm_Initialize()
        self.Controls={}
        self.rb_res = {"_R": tk.IntVar(),
                       "_L": tk.IntVar()}
        self.__Disable_Set_Arduino_Typ = True
        #*HL Center_Form(Me)                

    
    def __Close_Button_Click(self):
        #-------------------------------
        self.Hide()
    
    def __ColorTest_Button_Click(self):
        #-----------------------------------
        self.Hide()
        #Open_MobaLedCheckColors('')
    
    def __Detect_LED_Port_Button_Click(self):
        #-----------------------------------------
        self.Hide()
        M07.Detect_Com_Port_and_Save_Result(False)
        self.Show()
    
    def __Detect_Right_Port_Button_Click(self):
        #-------------------------------------------
        self.Hide()
        M07.Detect_Com_Port_and_Save_Result(True)
        self.Show()
    
    def __FastBootloader_Button_Click(self):
        #----------------------------------------
        self.Hide()
        M08IN.Install_FastBootloader()
    
    def __HardiForum_Button_Click(self):
        #------------------------------------
        self.Hide()
        if P01.MsgBox(M09.Get_Language_Str('Öffnet das Profil von Hardi im Stummi Forum.' + vbCr + vbCr + 'Dort findet man, wenn man im Forum angemeldet ist, einen Link zur ' + 'Email Adresse des Autors zum senden einer E-Mail oder PN.' + vbCr + vbCr + 'Alternativ kann auch eine Mail an \'MobaLedlib@gmx.de\' geschickt werden wenn ' + 'es Fragen oder Anregungen zu dem Programm oder zur MobaLedLib gibt.'), vbOKCancel, Get_Language_Str('Profil des Autors öffnen')) == vbOK:
            P01.Shell('Explorer "https://www.stummiforum.de/memberlist.php?mode=viewprofile&u=26419"')
    
    def __Pattern_Config_Button_Click(self):
        #----------------------------------------
        self.Hide()
        #Start_Pattern_Configurator()
    
    def __ProInstall_Button_Click(self):
        #------------------------------------
        M08.Ask_to_Upload_and_Compile_and_Upload_Prog_to_Right_Arduino()
    
    def __Nano_Normal_L_Click(self):
        self.Change_Board(True, M02.BOARD_NANO_OLD)
    
    def __Nano_New_L_Click(self):
        self.Change_Board(True, M02.BOARD_NANO_NEW)
    
    def __Nano_Full_L_Click(self):
        self.Change_Board(True, M02.BOARD_NANO_FULL)
    
    def __Uno_L_Click(self):
        self.Change_Board(True, M02.BOARD_UNO_NORM)
    
    def __Board_IDE_L_Click(self):
        self.Change_Board(True, '')
    
    def __ESP32_L_Click(self):
        self.Change_Board(True, M02.BOARD_ESP32)
        self.Autodetect_Typ_L_CheckBox.Value = False
    
    def __Pico_L_Click(self):
        self.Change_Board(True, M02.BOARD_PICO)
        self.Autodetect_Typ_L_CheckBox.Value = False
    
    def __Nano_Normal_R_Click(self):
        self.Change_Board(False, M02.BOARD_NANO_OLD)
    
    def __Nano_New_R_Click(self):
        self.Change_Board(False, M02.BOARD_NANO_NEW)
    
    def __Nano_Full_R_Click(self):
        self.Change_Board(False, M02.BOARD_NANO_FULL)
    
    def __Uno_R_Click(self):
        self.Change_Board(False, M02.BOARD_UNO_NORM)
    
    def __Board_IDE_R_Click(self):
        self.Change_Board(False, '')
    
    def __Autodetect_Typ_L_CheckBox_Click(self):
        self.__Change_Autodetect(True)
    
    def __Autodetect_Typ_R_CheckBox_Click(self):
        self.__Change_Autodetect(False)
    
    def Change_Board(self,LeftArduino, NewBrd):
        #----------------------------------------------------------------
        if self.__Disable_Set_Arduino_Typ:
            return
        M28.Change_Board_Typ(LeftArduino, NewBrd)
    
    def __Change_Autodetect(self,LeftArduino):
        Side = String()
    
        Col = Long()
        #----------------------------------------------------
        if LeftArduino:
            Col = M25.BUILDOP_COL
            Side = 'L'
            self.__Set_Autodetect_Value(Col, self.Autodetect_Typ_L_CheckBox_var.get())
        else:
            Col = M25.BUILDOpRCOL
            Side = 'R'
            self.__Set_Autodetect_Value(Col, self.Autodetect_Typ_R_CheckBox_var.get()) #Controls('Autodetect_Typ_' + Side + '_CheckBox'))
    
    def __Set_Autodetect_Value(self,BuildOpt_Col, Value):
        #--------------------------------------------------------------------------
        with_0 = P01.Cells(M02.SH_VARS_ROW, M25.BuildOpt_Col)
        if Value:
            if InStr(with_0.Value, M02.AUTODETECT_STR) == 0:
                with_0.Value = M02.AUTODETECT_STR + ' ' + Trim(with_0.Value)
        else:
            with_0.Value = M30.Replace_Multi_Space(Trim(Replace(with_0.Value, M02.AUTODETECT_STR, '')))
    
    def __Get_Arduino_Typ(self,LeftArduino):
        Side = String()
    
        Col = Integer()
    
        BuildOpt = String()
        #---------------------------------------------------
        if LeftArduino:
            Col = M25.BUILDOP_COL
            Side = 'L'
            BuildOpt = P01.Cells(M02.SH_VARS_ROW, Col)
            self.Autodetect_Typ_L_CheckBox_var.set(InStr(BuildOpt, M02.AUTODETECT_STR) > 0)
        else:
            Col = M25.BUILDOpRCOL
            Side = 'R'
            BuildOpt = P01.Cells(M02.SH_VARS_ROW, Col)
            self.Autodetect_Typ_R_CheckBox_var.set(InStr(BuildOpt, M02.AUTODETECT_STR) > 0)
        #BuildOpt = P01.Cells(M02.SH_VARS_ROW, Col)
        #self.Controls['Autodetect_Typ_' + Side + '_CheckBox'] = ( InStr(BuildOpt, M02.AUTODETECT_STR) > 0 )
        if InStr(BuildOpt, M02.BOARD_NANO_OLD) > 0:
            self.Controls['Nano_Normal_' + Side].Value = True
            return
        if InStr(BuildOpt, M02.BOARD_NANO_FULL) > 0:
            self.Controls['Nano_Full_' + Side].Value = True
            return
            # 28.10.20:
        if InStr(BuildOpt, M02.BOARD_NANO_NEW) > 0:
            self.Controls['Nano_New_' + Side].Value = True
            return
        if InStr(BuildOpt, M02.BOARD_UNO_NORM) > 0:
            self.Controls['Uno_' + Side].Value = True
            return
        if InStr(BuildOpt, M02.BOARD_NANO_EVERY) > 0:
            return
            # currently no option in GUI, but that's ok, as ATMEGA4809 is currently unsupported 28.10.20: Jürgen
        if InStr(BuildOpt, M02.BOARD_ESP32) > 0 and Side == 'L' and M37.ESP32_Lib_Installed():
            self.Controls['ESP32_L'].Value = True
            return
            # 11.11.20:
        if InStr(BuildOpt, M02.BOARD_PICO) > 0 and Side == 'L' and M37.PICO_Lib_Installed():
            self.Controls['PICO_L'].Value = True
            return
            # 18.04.21: Juergen
        if InStr(BuildOpt, '--board ') > 0:
            self.Controls['Nano_Normal_' + Side].Value = False
            self.Controls['Nano_New_' + Side].Value = False
            self.Controls['Uno_' + Side].Value = False
            self.Controls['Board_IDE_' + Side].Value = False
            if Side == 'L':
                self.Controls['ESP32_L'].Value = False
            if Side == 'L':
                self.Controls['PICO_L'].Value = False
            P01.MsgBox(M09.Get_Language_Str('Unbekannte Board Option: ') + vbCr + BuildOpt, vbInformation, M09.Get_Language_Str('Unbekanntes Board'))
        self.Controls['Board_IDE_' + Side].Value = True
    
    def __Test_Get_Arduino_Typ():
        #UT-------------------------------
        self.__Get_Arduino_Typ(True)
    
    def __Import_Button_Click(self):
        #--------------------------------
        self.Hide()
        M17.Import_from_Old_Version()
    
    def __Save_Button_Click(self):
        #------------------------------
        self.Hide()
        M18.Save_Data_to_File()
    
    def __Load_Button_Click(self):
        #------------------------------
        M18.Load_Data_from_File()
    
    def __Copy_Page_Button_Click(self):
        #-----------------------------------
        self.Hide()
        M18.Copy_from_Sheet_to_Sheet()
    
    def __Update_to_Arduino_Button_Click(self):
        #-------------------------------------------
        self.Hide()
        M37.Update_MobaLedLib_from_Arduino_and_Restart_Excel()
    
    def __Update_Beta_Button_Click(self):
        #-------------------------------------
        self.Hide()
        M27.Update_MobaLedLib_from_Beta_and_Restart_Excel()
    
    def __Show_Lib_and_Board_Page_Button_Click(self):
        #-------------------------------------------------
        self.Hide()
        with_1 = P01.ThisWorkbook.Sheets(M02.LIBRARYS__SH)
        with_1.Visible = True
        with_1.Select()
        
    def Button_Setup(self,button_frame,Text,Command,Accelerator,Row=0):
        Text = Trim(Text)
        Button = tk.Button(button_frame, text=Text, command=Command,width=20,font=("Tahoma", 8))
        Button.grid(row=Row,column=0,sticky="e",padx=10,pady=10)
        self.top.bind(Accelerator, Command)
        return
        
        
    def create_arduinopage(self,frame,LeftArduino):
        
        self.radiobuttons = {"Nano_Normal_L": {"text":"Nano Normal (old Bootloader)", "value": 1},
                             "Nano_New_L"   : {"text":"Nano (neue Version)", "value": 2},
                             "Nano_Full_L"  : {"text":"Nano (Full memory)", "value": 3},
                             "Uno_L"        : {"text":"Uno", "value": 4},
                             "Board_IDE_L"  : {"text":"Typ von Arduino IDE benutzen", "value": 5},
                             "ESP32_L"      : {"text":"ESP32 Wroom", "value": 6},
                             "Pico_L"       : {"text":"Raspberry Pico (Experimental)", "value": 7},
                             "Nano_Normal_R": {"text":"Nano Normal (old Bootloader)", "value": 1},
                             "Nano_New_R"   : {"text":"Nano (neue Version)", "value": 2},
                             "Nano_Full_R"  : {"text":"Nano (Full memory)", "value": 3},
                             "Uno_R"        : {"text":"Uno", "value": 4},
                             "Board_IDE_R"  : {"text":"Typ von Arduino IDE benutzen", "value": 5},
                             }
        
        self.Button_Setup(frame,M09.Get_Language_Str("USB Port erkennen"),self.__Detect_LED_Port_Button_Click,"U",Row=0)
        
        if LeftArduino:
            #self.Button_Setup(M09.Get_Language_Str("Version lesen"),self.CommandButton4_,"U",Row=2)
            side="_L"
            self.Autodetect_Typ_L_CheckBox_var = tk.IntVar(master=self.top)
            self.Autodetect_Typ_L_CheckBox_var.set(0)
            self.Autodetect_Typ_L_CheckBox = tk.Checkbutton(frame, text=M09.Get_Language_Str("Automatisch erkennen"),width=30,wraplength = 200,anchor="w",variable=self.Autodetect_Typ_L_CheckBox_var,font=("Tahoma", 8),onvalue = 1, offvalue = 0)
            self.Autodetect_Typ_L_CheckBox.grid(row=0, column=1,sticky="nw", padx=2, pady=2)                    
        else:
            side="_R"
            self.Button_Setup(frame,M09.Get_Language_Str("Prog. Installieren"),self.__ProInstall_Button_Click,"U",Row=2)    
            self.Autodetect_Typ_R_CheckBox_var = tk.IntVar(master=self.top)
            self.Autodetect_Typ_R_CheckBox_var.set(0)
            self.Autodetect_Typ_R_CheckBox = tk.Checkbutton(frame, text=M09.Get_Language_Str("Automatisch erkennen"),width=30,wraplength = 200,anchor="w",variable=self.Autodetect_Typ_R_CheckBox_var,font=("Tahoma", 8),onvalue = 1, offvalue = 0)
            self.Autodetect_Typ_R_CheckBox.grid(row=0, column=1,sticky="nw", padx=2, pady=2)                  
        
        for rb in self.radiobuttons.keys():
            rb_text = self.radiobuttons[rb]["text"]
            if rb.endswith(side):
                if rb.startswith("ESP32"):
                    if not M37.ESP32_Lib_Installed():
                        continue
                if rb_text.startswith("Pico"):
                    if not M37.PICO_Lib_Installed():
                        continue                    
                rbutton=tk.Radiobutton(frame, text=rb_text, variable=self.rb_res.get(side), value=self.radiobuttons[rb]["value"])
                row = self.radiobuttons[rb]["value"]
                if row > 3:
                    row+=1
                rbutton.grid(row=row,column=1,sticky="nw",padx=10,pady=0)
                self.radiobuttons[rb]["button"]=rbutton
                self.Controls[rb]=P01.CControl(False)
        
        subtitle_txt = M09.Get_Language_Str("Für andere Hauptplatine")
        self.subtitle_Label = ttk.Label(frame, text=subtitle_txt,font=("Tahoma", 8),width=40,wraplength=350,relief=tk.FLAT, borderwidth=1)
        self.subtitle_Label.grid(row=4,column=1,sticky="w",padx=10,pady=0)
   
    def __UserForm_Initialize(self):
        #--------------------------------
        # Is called once to initialice the form
        #Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
        #Change_Language_in_Dialog(Me)
        #Center_Form(Me)
        
        self.top = tk.Toplevel(self.controller)
        self.top.transient(self.controller)

        self.top.grab_set()
        
        self.top.resizable(True, True)  # This code helps to disable windows from resizing
        
        window_height = 500
        window_width = 800
        
        winfo_x = PG.global_controller.winfo_x()
        winfo_y = PG.global_controller.winfo_y()
        
        screen_width = PG.global_controller.winfo_width()
        screen_height = PG.global_controller.winfo_height()
        
        x_cordinate = winfo_x+int((screen_width/2) - (window_width/2))
        y_cordinate = winfo_y+int((screen_height/2) - (window_height/2))
        
        #self.top.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))                 
        self.top.geometry("+{}+{}".format(x_cordinate, y_cordinate))                   
        
        
        self.top.title(M09.Get_Language_Str("Optionen und spezielle Funktionien")) 
         
        self.container = ttk.Notebook(self.top)
        self.container.grid(row=0,column=0,sticky="nesw")
        
        self.tabdict = {}
        LED_Arduino_frame = tk.Frame(self.container)
        LED_Arduino_frame_Name = "LED Arduino"
        self.tabdict[LED_Arduino_frame_Name] = LED_Arduino_frame
        self.container.add(LED_Arduino_frame, text=LED_Arduino_frame_Name)
        
        self.create_arduinopage(LED_Arduino_frame, True)
        
        if  M25.Page_ID != 'CAN':
            DCC_Arduino_frame = tk.Frame(self.container)
            DCC_Arduino_frame_Name = M25.Page_ID + ' Arduino'
            self.tabdict[DCC_Arduino_frame_Name] = DCC_Arduino_frame
            self.container.add(DCC_Arduino_frame, text=DCC_Arduino_frame_Name)
            self.create_arduinopage(DCC_Arduino_frame, False)
        
        File_frame = tk.Frame(self.container)
        File_frame_Name = M09.Get_Language_Str("Dateien")
        self.tabdict[File_frame_Name] = File_frame
        self.container.add(File_frame, text=File_frame_Name)        
        
        Update_frame = tk.Frame(self.container)
        Update_frame_Name = M09.Get_Language_Str("Update")
        self.tabdict[Update_frame_Name] = Update_frame
        self.container.add(Update_frame, text=Update_frame_Name)
        
        Bootloader_frame = tk.Frame(self.container)
        Bootloader_frame_Name = M09.Get_Language_Str("Bootloader")
        self.tabdict[Bootloader_frame_Name] = Bootloader_frame
        self.container.add(Bootloader_frame, text=Bootloader_frame_Name)
                
    def UserForm_Activate(self):
        #------------------------------
        # Is called every time when the form is shown
        M25.Make_sure_that_Col_Variables_match()
        
        self.__Disable_Set_Arduino_Typ = True
        self.__Get_Arduino_Typ(True)
        self.__Get_Arduino_Typ(False)
        self.__Disable_Set_Arduino_Typ = False
        
    def Hide(self):
        self.top.destroy()
        
    def control_to_radiobutton(self):
        for rb in self.radiobuttons.keys():
            control = self.Controls.get(rb,P01.CControl(False))
            if control.Value==True:
                rb_var=self.rb_res.get(rb[-2:])
                rb_var.set(self.radiobuttons[rb]["value"])
        
        
    def Show(self):
        self.IsActive = True
        self.__UserForm_Initialize()
        self.UserForm_Activate()
        self.control_to_radiobutton()
     
        self.controller.wait_window(self.top)
        
    
    # VB2PY (UntranslatedCode) Option Explicit
