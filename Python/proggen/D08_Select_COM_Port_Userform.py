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



#import proggen.M02_Public as M02
#import proggen.M03_Dialog as M03
#import proggen.M06_Write_Header as M06
#import proggen.M06_Write_Header_LED2Var as M06LED
#import proggen.M06_Write_Header_Sound as M06Sound
#import proggen.M06_Write_Header_SW as M06SW
import proggen.M07_COM_Port as M07
#import proggen.M08_ARDUINO as M08
#import proggen.M09_Language as M09
#import proggen.M09_Select_Macro as M09SM
#import proggen.M09_SelectMacro_Treeview as M09SMT
#import proggen.M10_Par_Description as M10
#import proggen.M20_PageEvents_a_Functions as M20
#import proggen.M25_Columns as M25
#import proggen.M27_Sheet_Icons as M27
#import proggen.M28_divers as M28
import proggen.M30_Tools as M30
#import proggen.M31_Sound as M31
#import proggen.M37_Inst_Libraries as M37
#import proggen.M60_CheckColors as M60
#import proggen.M70_Exp_Libraries as M70
#import proggen.M80_Create_Mulitplexer as M80

from ExcelAPI.X01_Excel_Consts import *
import ExcelAPI.P01_Workbook as P01
import proggen.Prog_Generator as PG
import proggen.M09_Language as M09

import logging

LocalComPorts = vbObjectInitialize(objtype=Byte)
__OldL_ComPorts = vbObjectInitialize(objtype=Byte)
PortNames = vbObjectInitialize(objtype=String)
__OldSpinButton = Long()
__Pressed_Button = Long()
LocalPrintDebug = Boolean()
__LocalShow_ComPort = Boolean()


class CSelect_COM_Port_UserForm:
    def __init__(self):

        self.controller = PG.get_global_controller()
        self.IsActive = False
        self.button1_txt = "Abbrechen"
        self.button2_txt = "Ok"
        self.res = False
        self.UserForm_Res = ""
        self.ParList = Variant()
        self.FuncName = String()
        self.NamesA = Variant()
        self.Show_Channel_Type = Long()
        self.CurWidth = Long()
        self.CurHeight = Long()
        self.MinFormHeight = Long()
        self.MinFormWidth = Long()
        self.MAX_PAR_CNT = 14
        self.TypA = vbObjectInitialize((self.MAX_PAR_CNT,), String)
        self.MinA = vbObjectInitialize((self.MAX_PAR_CNT,), Variant)
        self.MaxA = vbObjectInitialize((self.MAX_PAR_CNT,), Variant)
        self.ParName = vbObjectInitialize((self.MAX_PAR_CNT,), String)
        self.Invers = vbObjectInitialize((self.MAX_PAR_CNT,), Boolean)
        self.DEFAULT_PAR_WIDTH = 48
        self.ParamVar = {}
        self.Controls={}
        self.__UserForm_Initialize()
        #*HL Center_Form(Me)                
 
    def ok(self, event=None):
        self.IsActive = False
        self.OK_Button_Click()
        
        #self.Userform_res = value
        self.top.destroy()
        P01.ActiveSheet.Redraw_table()
        self.res = True
 
    def cancel(self, event=None):
        self.UserForm_Res = '<Abort>'
        self.IsActive = False
        self.top.destroy()
        P01.ActiveSheet.Redraw_table()
        self.res = False

    def show(self):
        
        self.IsActive = True
        self.controller.wait_window(self.top)

        return self.res
    
    def __Check_Button_Click():
        #-------------------------------
        # Left Button
        __Pressed_Button = 1
        Me.Hide()
        CheckCOMPort = 0
    
    def __Abort_Button_Click():
        #-------------------------------
        # Middle Button
        __Pressed_Button = 2
        Me.Hide()
        CheckCOMPort = 0
    
    def __Default_Button_Click():
        #---------------------------------
        # Right Button
        __Pressed_Button = 3
        Me.Hide()
        CheckCOMPort = 0
    
    def __SpinButton_Change():
        #------------------------------
        #Debug.Print "Update_SpinButton"
        Update_SpinButton(0)
    
    # VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: DefaultPort - ByVal 
    def Update_SpinButton(self, DefaultPort):
        #return #*HL
    
    
        global CheckCOMPort_Txt,CheckCOMPort,COM_Port_Label,PortNames,LocalPrintDebug,LocalComPorts
        #------------------------------------------------------
        # Is also called by the OnTime proc which checks the available ports
        Show_Unknown_CheckBox = True #*HL
        LocalComPorts,PortNames = M07.EnumComPorts(Show_Unknown_CheckBox, PortNames, PrintDebug= LocalPrintDebug)
        if M30.isInitialised(LocalComPorts):
            SpinButton.Max = UBound(LocalComPorts)
            if DefaultPort > 0:
                for i in vbForRange(0, UBound(LocalComPorts)):
                    if DefaultPort == LocalComPorts(i):
                        SpinButton = i
            else:
                if M30.isInitialised(__OldL_ComPorts):
                    if UBound(LocalComPorts) > UBound(__OldL_ComPorts):
                        for ix in vbForRange(0, UBound(__OldL_ComPorts)):
                            if LocalComPorts(ix) != __OldL_ComPorts(ix):
                                SpinButton = ix
                                break
                        if ix > UBound(__OldL_ComPorts):
                            SpinButton = ix
            if SpinButton > SpinButton.Max:
                SpinButton = SpinButton.Max
            CheckCOMPort_Txt = PortNames(SpinButton)
            CheckCOMPort = LocalComPorts(SpinButton)
            COM_Port_Label = ' COM' + CheckCOMPort
            if SpinButton != __OldSpinButton:
                self.Show_Status(False, M09.Get_Language_Str('Aktualisiere Status ...'))
                __OldSpinButton = SpinButton
            for Port in LocalComPorts:
                PortsStr = PortsStr + Port + ' '
            self.AvailPorts_Label.configure(text=M30.DelLast(PortsStr))
            __OldL_ComPorts = LocalComPorts
        else:
            CheckCOMPort = 999
            elf.AvailPorts_Label.configure(text='')
            COM_Port_Label = ' -'
    
    def Show_Status(self,ErrBox, Msg):
        global Error_Label,Status_Label
        #-------------------------------------------------------
        if ErrBox:
            #if Error_Label != Msg:
            self.Error_Label.configure(text=Msg)
                # "If" is used to prevent flickering
        else:
            self.Status_Label.configure(text=Msg)
            #if Status_Label != Msg:
            #    Status_Label = Msg
        #if Error_Label.Visible != ErrBox:
        if ErrBox:
            self.Error_Label.grid()
            self.Status_Label.grid()
        else:
            self.Error_Label.grid_remove()
            self.Status_Label.grid_remove()            
    
    # VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ComPort_IO - ByRef 
    def ShowDialog(self, Caption, Title, Text, Picture, Buttons, FocusButton, Show_ComPort, Red_Hint, ComPort_IO, PrintDebug=False):
        fn_return_value = None
            
        c = Variant()
    
        Found = Boolean()
        #----------------------------------------------------------------------------------------------
        # Variables:
        #  Caption     Dialog Caption
        #  Title       Dialog Title
        #  Text        Message in the text box on the top left side
        #  Picture     Name of the picture to be shown. Available pictures: "LED_Image", "CAN_Image", "Tiny_Image", "DCC_Image"
        #  Buttons     List of 3 buttons with Accelerator. Example "H Hallo; A Abort; O Ok"  Two Buttons: " ; A Abort; "O Ok"
        #  ComPort_IO  is used as input and output
        # Return:
        #  1: If the left   Button is pressed  (Install, ...)
        #  2: If the middle Button is pressed  (Abort)
        #  3: If the right  Button is pressed  (OK)
        self.title = M09.Get_Language_Str(Caption)
        
        self.top = tk.Toplevel(self.controller)
        self.top.transient(self.controller)

        self.top.grab_set()
        
        self.top.resizable(True, True)  # This code helps to disable windows from resizing
        
        window_height = 700
        window_width = 500
        
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        
        x_cordinate = int((screen_width/2) - (window_width/2))
        y_cordinate = int((screen_height/2) - (window_height/2))
        
        self.top.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))        
        
        if len(self.title) > 0: 
            self.top.title(self.title)        
        
        self.Title_Label = ttk.Label(self.top, text=Title,font=("Tahoma", 11),width=30,wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Title_Label.focus_set()
        self.Title_Label.grid(row=0,column=0,columnspan=2,sticky="nesw",padx=10,pady=10)
        
        self.Text_Label = ttk.Label(self.top, text=Text,font=("Tahoma", 11),width=30, wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Text_Label.grid(row=1,column=0,columnspan=2,sticky="nesw",padx=10,pady=10)
        
        self.Status_Label = ttk.Label(self.top, text="",font=("Tahoma", 11),width=15,wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Status_Label.grid(row=2,column=1,rowspan=2,sticky="nesw",padx=10,pady=10)
        
        self.Error_Label = ttk.Label(self.top, text="",font=("Tahoma", 11),width=15, wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Error_Label.grid(row=2,column=0,rowspan=1,sticky="nesw",padx=10,pady=10)        
                
        self.AvailPortsTxt_Label = ttk.Label(self.top, text="",font=("Tahoma", 11),width=15, wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.AvailPortsTxt_Label.grid(row=3,column=0,columnspan=2,sticky="nesw",padx=10,pady=10)        
                
        self.Image_Label = ttk.Label(self.top, text="",font=("Tahoma", 11),width=30, wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Image_Label.grid(row=0,column=2,rowspan=7,sticky="nesw",padx=10,pady=10)
        
        #self.Show_Unknown_CheckBox = 
        
        self.Hint_Label = ttk.Label(self.top, text="Zur Identifikation des Arduinos blinken die LEDs des ausgewählten Arduinos schnell.\nEin anderer COM Port kann über die Pfeiltasten ausgewählt werden.\nDer Arduino kann auch nachträglich angesteckt werden.",font=("Tahoma", 11),width=30,wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.Hint_Label.grid(row=7,column=0,columnspan=2,rowspan=2,sticky="nesw",padx=10,pady=10)
        
        self.AvailPorts_Label = ttk.Label(self.top, text="",font=("Tahoma", 11),width=30, wraplength=window_width-20,relief=tk.SUNKEN, borderwidth=1)
        self.AvailPorts_Label.grid(row=5,column=0,columnspan=2,sticky="nesw",padx=10,pady=10)
        
        self.button_frame = ttk.Frame(self.top)
        
        self.b_cancel = tk.Button(self.button_frame, text=self.button1_txt, command=self.cancel,width=10,font=("Tahoma", 11))
        self.b_ok = tk.Button(self.button_frame, text=self.button2_txt, command=self.ok,width=10,font=("Tahoma", 11))

        self.b_cancel.grid(row=0,column=0,sticky="e",padx=10,pady=10)
        self.b_ok.grid(row=0,column=1,sticky="e",padx=10,pady=10)
        
        self.button_frame.grid(row=8,column=2,sticky="e",padx=10,pady=10)
        
        self.top.bind("<Return>", self.ok)
        self.top.bind("<Escape>", self.cancel)                   
        self.show()
        
        return
        
        #Me.Caption = Caption
        #Title_Label = Title
        #Text_Label = Text
        #Error_Label = ''
        #Status_Label = ''
        #ButtonArr = Split(Buttons, ';')
        #if UBound(ButtonArr) != 2:
        #    MsgBox('Internal Error in Select_COM_Port_UserForm: \'Buttons\' must be a string with 3 buttons separated by \';\'' + vbCr + 'Wrong: \'' + Buttons + '\'', vbCritical, 'Internal Error (Wrong translation?)')
        #    EndProg()
        #Button_Setup(Check_Button, ButtonArr(0))
        #Button_Setup(Abort_Button, ButtonArr(1))
        #Button_Setup(Default_Button, ButtonArr(2))
        #if FocusButton != '':
        #    Controls(FocusButton).setFocus()
        __LocalPrintDebug = PrintDebug
        __OldSpinButton = - 1
        __Pressed_Button = 0
        Update_SpinButton(ComPort_IO)
        SpinButton.Visible = Show_ComPort
        if Show_ComPort:
            SpinButton.setFocus()
        __LocalShow_ComPort = Show_ComPort
        # Show / Hide the COM Port
        COM_Port_Label.Visible = Show_ComPort
        Error_Label.Visible = Show_ComPort
        Status_Label.Visible = Show_ComPort
        AvailPortsTxt_Label.Visible = Show_ComPort
        Available_Ports_Label.Visible = Show_ComPort
        Show_Unknown_CheckBox.Visible = Show_ComPort
        Hint_Label.Visible = Show_ComPort
        if Show_ComPort:
            Text_Label.Height = Error_Label.Top - Text_Label.Top
        else:
            Text_Label.Height = Hint_Label.Top + Hint_Label.Height - Text_Label.Top
        for c in Me.Controls:
            if Right(c.Name, Len('Image')) == 'Image':
                if Picture == c.Name:
                    c.Visible = True
                    Found = True
                else:
                    c.Visible = False
        if not Found:
            MsgBox('Internal Error: Unknown picture: \'' + Picture + '\'', vbCritical, 'Internal Error')
        Red_Hint_Label = Red_Hint
        Me.Show()
        # Store the results
        if Show_ComPort:
            if isInitialised(LocalComPorts):
                ComPort_IO = LocalComPorts(SpinButton)
        fn_return_value = __Pressed_Button
        return fn_return_value
    
    def __UserForm_Initialize(self):
        #--------------------------------
        #Debug.Print vbCr & me.Name & ": UserForm_Initialize"
        #Change_Language_in_Dialog(Me)
        #P01.Center_Form(Me)
        pass
        
    
    def __UserForm_QueryClose(self, CloseMode, Cancel):
        __Abort_Button_Click()
    
    # VB2PY (UntranslatedCode) Option Explicit



        
    
        