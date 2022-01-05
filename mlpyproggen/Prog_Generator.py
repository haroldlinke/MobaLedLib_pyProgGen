# -*- coding: utf-8 -*-
#
#         MobaLedCheckColors: Color checker for WS2812 and WS2811 based MobaLedLib
#
#         ProgramGen2
#
# * Version: 1.00
# * Author: Harold Linke
# * Date: December 25th, 2019
# * Copyright: Harold Linke 2019
# *
# *
# * MobaLedCheckColors on Github: https://github.com/haroldlinke/MobaLedCheckColors
# *
# *
# * History of Change
# * V1.00 25.12.2019 - Harold Linke - first release
# *
# *
# * MobaLedCheckColors supports the MobaLedLib by Hardi Stengelin
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
# * along with this program.  If not, see <http://www.gnu.org/licenses/>.
# *
# * MobaLedCheckColors is based on tkColorPicker by Juliette Monsel
# * https://sourceforge.net/projects/tkcolorpicker/
# *
# * tkcolorpicker - Alternative to colorchooser for Tkinter.
# * Copyright 2017 Juliette Monsel <j_4321@protonmail.com>
# *
# * tkcolorpicker is free software: you can redistribute it and/or modify
# * it under the terms of the GNU General Public License as published by
# * the Free Software Foundation, either version 3 of the License, or
# * (at your option) any later version.
# *
# * tkcolorpicker is distributed in the hope that it will be useful,
# * but WITHOUT ANY WARRANTY; without even the implied warranty of
# * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# * GNU General Public License for more details.
# *
# * You should have received a copy of the GNU General Public License
# * along with this program.  If not, see <http://www.gnu.org/licenses/>.
# *
# * The code for changing pages was derived from: http://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter
# * License: http://creativecommons.org/licenses/by-sa/3.0/
# ***************************************************************************

import tkinter as tk
from tkinter import ttk,messagebox

from mlpyproggen.DefaultConstants import ARDUINO_WAITTIME,LARGE_FONT, SMALL_FONT, VERY_LARGE_FONT, PROG_VERSION,ARDUINO_LONG_WAITTIME
from mlpyproggen.configfile import ConfigFile
from locale import getdefaultlocale
from tkintertable import TableCanvas, TableModel
from collections import OrderedDict
from mlpyproggen.P01_Workbook import create_workbook

import os
import serial
import sys
import threading
import subprocess
import queue
import time
import logging
import platform
from scrolledFrame.ScrolledFrame import VerticalScrolledFrame,HorizontalScrolledFrame,ScrolledFrame
from mlpyproggen.M06_Write_Header import Create_HeaderFile
#from mlpyproggen.P01_Workbook import global_tablemodel

from datetime import datetime
from mlpyproggen.T01_exceltable import get_globaltabelmodel
import mlpyproggen.F00_mainbuttons as F00
import mlpyproggen.M20_PageEvents_a_Functions as M20

# --- Translation - not used
EN = {}
FR = {"Red": "Rouge", "Green": "Vert", "Blue": "Bleu",
      "Hue": "Teinte", "Saturation": "Saturation", "Value": "Valeur",
      "Cancel": "Annuler", "Color Chooser": "Sélecteur de couleur",
      "Alpha": "Alpha"}
DE = {"Red": "Rot", "Green": "Grün", "Blue": "Blau",
      "Hue": "Farbton", "Saturation": "Sättigung", "Value": "Helligkeit",
      "Cancel": "Beenden", "Color Chooser": "Farbwähler",
      "Alpha": "Alpha", "Configuration": "Einstellungen"}

try:
    TR = EN
    if getdefaultlocale()[0][:2] == 'fr':
        TR = FR
    else:
        if getdefaultlocale()[0][:2] == 'de':
            TR = DE
except ValueError:
    TR = EN

def _(text):
    """Translate text."""
    return TR.get(text, text)

global_controller = None
dialog_parent = None

def set_global_controller(controller):
    global global_controller
    global_controller=controller

def get_global_controller():
    return global_controller

def get_dialog_parent():
    return dialog_parent

ThreadEvent = None

BUTTONLABELWIDTH = 10
            
class Prog_GeneratorPage(tk.Frame):
    def __init__(self, parent, controller):
        global global_tablemodel
        global dialog_parent
        dialog_parent = self
        self.tabClassName = "ProgGeneratorPage"
        tk.Frame.__init__(self,parent)
        self.controller = controller
        set_global_controller(controller)
        macrodata = self.controller.MacroDef.data.get(self.tabClassName,{})
        self.tabname = macrodata.get("MTabName",self.tabClassName)
        self.title = macrodata.get("Title",self.tabClassName)
        
        #button1_text = macrodata.get("Button_1",self.tabClassName)
        #button2_text = macrodata.get("Button_2",self.tabClassName)
        
        #self.fontlabel = self.controller.get_font("FontLabel")
        #self.fontspinbox = self.controller.get_font("FontSpinbox")
        #self.fonttext = self.controller.get_font("FontText")
        #self.fontbutton = self.controller.get_font("FontLabel")
        #self.fontentry = self.controller.get_font("FontEntry")
        #self.fonttext = self.controller.get_font("FontText")
        #self.fontscale = self.controller.get_font("FontScale")
        self.fonttitle = self.controller.get_font("FontTitle")        
        
        #self.grid_columnconfigure(0,weight=1)
        #self.grid_rowconfigure(0,weight=1)
        
        self.frame=ttk.Frame(self,relief="ridge", borderwidth=1)
        #self.frame.grid_columnconfigure(0,weight=1)
        #self.frame.grid_rowconfigure(0,weight=1)        
        
        #self.scroll_main_frame = ScrolledFrame(self.frame)
        #self.scroll_main_frame.grid_columnconfigure(0,weight=1)
        #self.scroll_main_frame.grid_rowconfigure(0,weight=1)
        
        #self.main_frame = ttk.Frame(self.scroll_main_frame.interior, relief="ridge", borderwidth=2)
        #self.main_frame.grid_columnconfigure(0,weight=1)
        #self.main_frame.grid_rowconfigure(2,weight=1)         

        
        #config_frame = self.controller.create_macroparam_frame(self.main_frame,self.tabClassName, maxcolumns=1,startrow =1,style="CONFIGPage")        

        self.parent = parent
        
        title_frame = ttk.Frame(self.frame, relief="ridge", borderwidth=2)
        self.button_frame = ttk.Frame(self.frame, borderwidth=0)
        self.workbook_frame = ttk.Frame(self.frame, relief="ridge", borderwidth=2)
        filedir = os.path.dirname(os.path.realpath(__file__))
        self.filedir2 = os.path.dirname(filedir)
        self.workbook_frame.rowconfigure(0,weight=1)
        self.workbook_frame.columnconfigure(0,weight=1)
        
        self.workbook = create_workbook(frame=self.workbook_frame,path=self.filedir2)
        
        for sheet in self.workbook.sheets:
            sheet.SetChangedCallback(self.wschangedcallback)
            sheet.SetSelectedCallback(self.wsselectedcallback)
        
        # Tabframe
        self.frame.grid(row=0,column=0,sticky="nesw")

        label = ttk.Label(title_frame, text=self.title, font=self.fonttitle)
        label.pack(padx=5,pady=(5,5))
        
        # create buttonlist
        self.create_button_list()
        
        title_frame.grid(row=0, column=0, sticky="n",pady=(10, 10), padx=10)
        self.button_frame.grid(row=1, column=0, sticky="nw",pady=(10, 10), padx=10)
        self.workbook_frame.grid(row=2,column=0,sticky="nesw",pady=(10, 10), padx=10)
    
        #config_frame.grid(row=1, columnspan=2, pady=(20, 30), padx=10)        
        #in_button_frame.grid(row=2, column=0, sticky="n", padx=4, pady=4)
        
        # ----------------------------------------------------------------
        # Standardprocedures for every tabpage
        # ----------------------------------------------------------------

    def tabselected(self):
        #self.controller.currentTabClass = self.tabClassName
        logging.debug("Tabselected: %s",self.tabname)
        #self.controller.send_to_ARDUINO("#END")
        #time.sleep(ARDUINO_WAITTIME)        
        pass
    
    def tabunselected(self):
        logging.debug("Tabunselected: %s",self.tabname)
        #self.controller.send_to_ARDUINO("#BEGIN")
        #time.sleep(ARDUINO_WAITTIME)            
        pass
    
    def TabChanged(self,_event=None):
        logging.debug("Tabchanged: %s",self.tabname)
        pass
    
    def cancel(self,_event=None):
        pass

    def getConfigData(self, key):
        return self.controller.getConfigData(key)
    
    def readConfigData(self):
        self.controller.readConfigData()
        
    def setConfigData(self,key, value):
        self.controller.setConfigData(key, value)

    def setParamData(self,key, value):
        self.controller.setParamData(key, value)

    def MenuUndo(self,_event=None):
        pass
    
    def MenuRedo(self,_event=None):
        pass

    def connect(self):
        pass

    def disconnect(self):
        pass
    
    # ----------------------------------------------------------------
    # End of Standardprocedures for every tabpage
    # ---------------------------------------------------------------- 
    
    def wschangedcallback(self,changedcell):
        
        print ("wschangedcallback ",changedcell.Row,":",changedcell.Column)
        M20.Global_Worksheet_Change(changedcell)
        
    def wsselectedcallback(self,changedcell):
        
        print ("wsselectedcallback ",changedcell.Row,":",changedcell.Column)
        M20.Global_Worksheet_SelectionChange(changedcell)    
        
        
    def create_button_list(self):
        
        button_list=(
                        {"Icon_name": "btn_dialog.png",
                         "command"  : F00.Dialog_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "Dialog aufrufen"},
                        {"Icon_name": "Btn_Send_to_ARDUINO.png",
                         "command"  : F00.Arduino_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "ARDUINO aufrufen"},
                        {"Icon_name": "btn_insert_row.png",
                         "command"  : F00.Insert_Button_Click,
                         "padx"     : 10,
                         "tooltip"  : "Zeile einfügen"},
                        {"Icon_name": "btn_delete_row.png",
                         "command"  : F00.Del_Button_Click,
                         "padx"     : 10,
                         "tooltip"  : "Zeile löschen"},
                        {"Icon_name": "btn_move_row.png",
                         "command"  : F00.Move_Button_Click,
                         "padx"     : 10,
                         "tooltip"  : "Zeile verschieben"},
                        {"Icon_name": "btn_copy_row.png",
                         "command"  : F00.Copy_Button_Click,
                         "padx"     : 10,
                         "tooltip"  : "Zeile kopieren"},
                        {"Icon_name": "btn_hide_unhide.png",
                         "command"  : F00.Hide_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "Zeile verstecken"},
                        {"Icon_name": "btn_unhide_all.png",
                         "command"  : F00.UnHideAll_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "Zeile einfügen"},
                        {"Icon_name": "btn_delete_table.png",
                         "command"  : F00.ClearSheet_Button_Click,
                         "padx"     : 10,
                         "tooltip"  : "Zeile einfügen"},
                        {"Icon_name": "btn_options.png",
                         "command"  : F00.Options_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "Zeile einfügen"},
                        {"Icon_name": "btn_help.png",
                         "command"  : F00.Help_Button_Click,
                         "padx"     : 20,
                         "tooltip"  : "Zeile einfügen"}                        
                    )
        
        filedir = os.path.dirname(os.path.realpath(__file__))
        self.filedir2 = os.path.dirname(filedir)
        
        self.icon_dict = {}
        
        for button_desc in button_list:
            self.create_button(button_desc)
            
    def create_button(self, button_desc):
        filename = r"/images/"+button_desc["Icon_name"]
        filepath = self.filedir2 + filename
        self.icon_dict[button_desc["Icon_name"]] = tk.PhotoImage(file=filepath)
        button=ttk.Button(self.button_frame, text="Dialog", image=self.icon_dict[button_desc["Icon_name"]], command=button_desc["command"])
        button.pack( side="left",padx=button_desc["padx"])
        self.controller.ToolTip(button, text=button_desc["tooltip"])                
            
    def dialog_button_cmd(self):
        print("Dialog Button")
        return
    
    def send_to_ARDUINO_button_cmd(self):
        print("Send to ARDUINO Button")
        Create_HeaderFile()
        return
    
    def get_param_config_dict(self, paramkey):
        paramconfig_dict = self.controller.MacroParamDef.data.get(paramkey,{})
        return paramconfig_dict    

    
    def start_ARDUINO_program_Run(self):
        result = subprocess.run(self.startfile, stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
        
        if result.returncode == 0:
            self.arduinoMonitorPage.add_text_to_textwindow("\n"+self.ARDUINO_message2+"\n*******************************************************\n",highlight="OK")
        else:
            self.arduinoMonitorPage.add_text_to_textwindow("\n******** "+self.ARDUINO_message3+" ********\n",highlight="Error")
            self.arduinoMonitorPage.add_text_to_textwindow("\n"+result.stdout,highlight="Error")
            self.arduinoMonitorPage.add_text_to_textwindow("\n*******************************************************\n\n\n",highlight="Error")
    
    def write_stdout_to_text_window(self):
        if self.continue_loop:
            output = self.process.stdout.readline()
    
            if output != '' and self.continue_loop:
                try:
                    self.arduinoMonitorPage.add_text_to_textwindow(output.decode('utf-8').strip())
                except BaseException as e:
                    logging.debug(e)
                    logging.debug("ERROR: Write_stdout_to_text_window: %s",output)
                    pass            
            
            if self.process.poll() is not None:
                self.continue_loop=False
                self.rc = self.process.poll()
                if self.rc==1:
                    self.arduinoMonitorPage.add_text_to_textwindow("\n******** "+self.ARDUINO_message3+" ********\n",highlight="Error")
                else:
                    self.arduinoMonitorPage.add_text_to_textwindow("\n"+self.ARDUINO_message2+"\n*******************************************************\n",highlight="OK")
    
        if self.continue_loop:
            self.after(10,self.write_stdout_to_text_window)
    
            
    def start_ARDUINO_program_Popen(self):
        try:
            self.process = subprocess.Popen(self.startfile, stdout=subprocess.PIPE,stderr=subprocess.STDOUT,stdin = subprocess.DEVNULL)
            self.continue_loop=True
            self.write_stdout_to_text_window()
        except BaseException as e:
            #logging.error("Exception in start_ARDUINO_program_Popen %s - %s",e,self.startfile[0])
            self.arduinoMonitorPage.add_text_to_textwindow("\n*****************************************************\n",highlight="Error")
            self.arduinoMonitorPage.add_text_to_textwindow("\n* Exception in start_ARDUINO_program_Popen "+ e + "-" + self.startfile[0]+ "\n",highlight="Error")
            self.arduinoMonitorPage.add_text_to_textwindow("\n*****************************************************\n",highlight="Error")
    
    def upload_to_ARDUINO(self,_event=None,arduino_type="LED",init_arduino=False):
    # send effect to ARDUINO
        if self.controller.ARDUINO_status == "Connecting":
            tk.messagebox.showwarning(title="Zum ARDUINO Hochladen", message="PC versucht gerade sich mit dem ARDUINO zu verbinden, bitte warten Sie bis der Vorgang beendet ist!")
            return
        self.controller.disconnect()
        self.controller.set_connectstatusmessage("Kompilieren und Hochladen ...",fg="green")
        
        #if arduino_type == "LED":
        #    self.create_ARDUINO_CMD(init_arduino=init_arduino)
                
        private_startfile = self.getConfigData("startcmdcb")
        
        macrodata = self.controller.MacroDef.data.get("ARDUINOMonitorPage",{})
        
        self.arduinoMonitorPage=self.controller.getFramebyName ("ARDUINOMonitorPage")
        self.arduinoMonitorPage.delete_text_from_textwindow()
        
        file_not_found = True
        
        system_platform = platform.platform()
        
        if not "Windows" in system_platform:
            private_startfile = True
        
        self.ARDUINO_message4=""
        
        if private_startfile == True:
            filename = self.getConfigData("startcmd_filename")
            logging.debug("upload_to_ARDUINO - Individual Filename: %s",filename)
            if filename == " " or filename == "":
                filename = "No Filename provided"
            logging.debug("send to ARDUINO - Platform: %s",platform.platform())
            
            macos = "macOS" in system_platform
            macos_fileending = "/Contents/MacOS/Arduino" 
            if macos:
                logging.debug("This is a MAC")
                if not filename.endswith(macos_fileending):
                    filename = filename + "/Contents/MacOS/Arduino"
                file_not_found = False
            else:
                if os.path.isfile(filename):
                    file_not_found = False
                if file_not_found:
                    self.ARDUINO_message4 = macrodata.get("Message_4","") + filename
    
        else:
            file_not_found = True
            #check if arduino_debug.exe exists in the program dirs
            Win_ARDUINO_searchlist = macrodata.get("Win_ARDUINOIDE","")
            
            for IDE_filename in Win_ARDUINO_searchlist:
                if os.path.isfile(IDE_filename):
                    file_not_found = False
                    filename = IDE_filename
                    break
            if file_not_found:
                self.ARDUINO_message4 = macrodata.get("Message_4","") + repr(Win_ARDUINO_searchlist)
        logging.debug("upload_to_ARDUINO - Filename: %s",filename)
        self.controller.showFramebyName("ARDUINOMonitorPage")
        self.controller.set_connectstatusmessage("Kompilieren und Hochladen ...",fg="green")
        self.update()
        
        if file_not_found:
            self.arduinoMonitorPage.add_text_to_textwindow("\n*******************************************************\n"+self.ARDUINO_message4+"\n*******************************************************\n")            
        else:
            serport = self.getConfigData("serportname")
            arduinotypenumber = self.getConfigData("ArduinoTypeNumber")
            ArduinoTypeNumber_config_dict = self.get_param_config_dict("ARDUINO Type")
            ArduinotypeList = ArduinoTypeNumber_config_dict["Values2Params"]
            ArduinoType = ArduinotypeList[arduinotypenumber]
            
            self.ARDUINO_message1 = macrodata.get("Message_1","")
            self.ARDUINO_message2 = macrodata.get("Message_2","")
            self.ARDUINO_message3 = macrodata.get("Message_3","")
            
            #h_filedir = os.path.dirname(filename)
            filedir = os.path.dirname(os.path.realpath(__file__))
            h_filedir1 = os.path.dirname(filedir)
            h_filedir = os.path.dirname(h_filedir1)            
            #os.chdir(h_filedir)
            os.chdir(self.controller.mainfile_dir)
            
            #filedirname = os.path.join(h_filedir, filename)
            #if platform.platform == "darwin":
            #    logging.debug("This is a MAC")
            #    filedirname = filename + "/Contents/MacOS/Arduino"
            #else:
            filedirname = filename
            
            if arduino_type=="DCC":
                ino_filename = self.controller.get_macrodef_data("StartPage","InoName_DCC") #"../../examples/23_A.DCC_Interface/23_A.DCC_Interface.ino"
            elif arduino_type=="Selectrix":
                ino_filename = self.controller.get_macrodef_data("StartPage","InoName_SX") #"../../examples/23_A.DCC_Interface/23_A.DCC_Interface.ino"
            elif arduino_type=="LED":
                ino_filename = self.controller.get_macrodef_data("StartPage","InoName_LED") #"LEDs_AutoProg.ino"
            
            if ArduinoType == " ":
                self.startfile = [filedirname,ino_filename,"--upload","--port",serport,"--pref","programmer=arduino:arduinoisp","--pref","build.path=../Arduino_Build_LEDs_AutoProg","--preserve-temp-files"]
            else:
                self.startfile = [filedirname,ino_filename,"--upload","--port",serport,"--board",ArduinoType,"--pref","programmer=arduino:arduinoisp","--pref","build.path=../Arduino_Build_LEDs_AutoProg","--preserve-temp-files"]
            
            logging.debug(repr(self.startfile))
            
            self.arduinoMonitorPage.add_text_to_textwindow("\n\n*******************************************************\n"+self.ARDUINO_message1+"\n*******************************************************\n\n")
            
            if arduino_type=="LED":
                self.arduinoMonitorPage.add_text_to_textwindow("***************************************************************************\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*    Zum                                                                  *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*     PC                    Prog_Generator " + PROG_VERSION+ " by Harold    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*      \\\\                                                                 *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*       \\\\                                                                *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*    ____\\\\___________________                                            *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  | [_] | | [_] |[oo]    |  Achtung: Es muss der LINKE               *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  Arduino mit dem PC verbunden             *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  sein.                                    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  | LED | |     |        |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  | Nano| |     |        |  Wenn alles gut geht, dann wird           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  hier eine Erfolgsmeldung angezeigt       *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |_____| |_____| [O]    |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |    [@] [@] [@]          |  Falls Probleme auftreten, dann werden    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |__________________[:::]__|  die Fehler hier aufgelistet              *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*                                                                         *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("***************************************************************************\n")
            elif arduino_type=="DCC":
                self.arduinoMonitorPage.add_text_to_textwindow("***************************************************************************\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*             Zum                                                         *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*              PC           Prog_Generator " + PROG_VERSION+ " by Harold    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*              \\\\                                                         *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*               \\\\                                                        *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*    ____________\\\\____________                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  | [_] | | [_] |[oo]    |  Achtung: Es muss der RECHTE              *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  Arduino mit dem PC verbunden             *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  sein.                                    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | | DCC |        |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | | Nano|        |  Wenn alles gut geht, dann wird           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |  hier eine Erfolgsmeldung angezeigt       *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |     | |     |        |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |  |_____| |_____| [O]    |                                           *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |    [@] [@] [@]          |  Falls Probleme auftreten, dann werden    *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*   |__________________[:::]__|  die Fehler hier aufgelistet              *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("*                                                                         *\n")
                self.arduinoMonitorPage.add_text_to_textwindow("***************************************************************************\n")                
            else:
                pass
            
            self.controller.showFramebyName("ARDUINOMonitorPage")
            start_with_realtime_logging = True
            use_start_cmd = False
            if use_start_cmd:
                filedir = os.path.dirname(os.path.realpath(__file__))
                h_filedir1 = os.path.dirname(filedir)
                h_filedir = os.path.dirname(h_filedir1)
                filename = "Start_Arduino.cmd"
                
                os.chdir(h_filedir)
                self.startfile = filename
                os.startfile(self.startfile)
            elif start_with_realtime_logging:
                self.after(500, self.start_ARDUINO_program_Popen)
            else:
                self.after(500, self.start_ARDUINO_program_Run)
    
    
    
        


