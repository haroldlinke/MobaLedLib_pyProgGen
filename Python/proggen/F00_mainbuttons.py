# -*- coding: utf-8 -*-
#
#         Write header
#
# * Version: 1.21
# * Author: Harold Linke
# * Date: January 1st, 2020
# * Copyright: Harold Linke 2020
# *
# *
# * MobaLedCheckColors on Github: https://github.com/haroldlinke/MobaLedCheckColors
# *
# *
# * History of Change
# * V1.00 10.03.2020 - Harold Linke - first release
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

from vb2py.vbfunctions import *
from vb2py.vbdebug import *
import tkinter as tk

import proggen.M06_Write_Header as M06
import proggen.M03_Dialog as M03
import proggen.M01_Gen_Release_Version as M01
import ExcelAPI.P01_Workbook as P01
import proggen.Prog_Generator as PG
import proggen.M22_Hide_UnHide as M22
import proggen.M23_Add_Move_Del_Row as M23
import proggen.M20_PageEvents_a_Functions as M20
import proggen.M30_Tools as M30
import proggen.M32_DCC as M32

import proggen.D02_Userform_Select_Typ_DCC as D02
import proggen.D02_Userform_Select_Typ_SX as D02SX
import proggen.D09_StatusMsg_Userform as D09
import proggen.D08_Select_COM_Port_Userform as D08
import proggen.D10_UserForm_Options as D10

def Arduino_Button_Click(event=None):
    #---------------------------------
    __Button_Pressed_Proc()
    M06.Create_HeaderFile()
    
def Arduino_Button_Shift_Click(event=None):
    #---------------------------------
    __Button_Pressed_Proc()
    P01.shift_key = True
    M06.Create_HeaderFile()


def ClearSheet_Button_Click():
    #------------------------------------
    __Button_Pressed_Proc()
    M20.ClearSheet()

def Dialog_Button_Click():
    #-------------------------------
    __Button_Pressed_Proc()
    M03.Dialog_Guided_Input()

def Help_Button_Click():
    #------------------------------
    __Button_Pressed_Proc()
    M20.Show_Help()

def __Button_Pressed_Proc():
    #--------------------------------
    #Selection.Select()
    #Application.EnableEvents = True
    #__Correct_Buttonsizes()
    pass

def __Name_with_LF():
    Hide_Button.Caption = 'Aus- oder' + vbLf + 'Einblenden'

def __Correct_Create_Buttonsize(obj):
    #-----------------------------------------
    obj.Height = 160
    obj.Width = 100
    obj.Height = 76
    obj.Width = 60

def __Correct_Buttonsizes():
    OldScreenupdating = Boolean()
    #------------------------
    # There is a bug in excel which changes the size of the buttons
    # if the resolution of the display is changed. This happens
    # fore instance if the computer is connected to a beamer.
    # To prevent this the buttons are resized with this function.
    OldScreenupdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    __Correct_Create_Buttonsize(Arduino_Button)
    __Correct_Create_Buttonsize(Dialog_Button)
    __Correct_Create_Buttonsize(Insert_Button)
    __Correct_Create_Buttonsize(Del_Button)
    __Correct_Create_Buttonsize(Move_Button)
    __Correct_Create_Buttonsize(Copy_Button)
    __Correct_Create_Buttonsize(Hide_Button)
    __Correct_Create_Buttonsize(UnHideAll_Button)
    __Correct_Create_Buttonsize(ClearSheet_Button)
    __Correct_Create_Buttonsize(Options_Button)
    __Correct_Create_Buttonsize(Help_Button)
    Application.ScreenUpdating = OldScreenupdating

def EnableDisableAllButtons(Enab):
    #------------------------------------------------------
    Arduino_Button.Enabled = Enab
    Dialog_Button.Enabled = Enab
    Insert_Button.Enabled = Enab
    Del_Button.Enabled = Enab
    Move_Button.Enabled = Enab
    Copy_Button.Enabled = Enab
    Hide_Button.Enabled = Enab
    UnHideAll_Button.Enabled = Enab
    ClearSheet_Button.Enabled = Enab
    Options_Button.Enabled = Enab

def __EnableAllButtons():
    #-----------------------------
    # Could be called manually after a crash
    EnableDisableAllButtons(True)

def Hide_Button_Click():
    #------------------------------
    __Button_Pressed_Proc()
    #notimplemented("Hide/Unhide")
    M22.Proc_Hide_Unhide()

def Insert_Button_Click():
    __Button_Pressed_Proc()
    P01.ActiveSheet.addrow_after_current_row()

def Del_Button_Click():
    #-----------------------------
    __Button_Pressed_Proc()
    P01.ActiveSheet.deleterows()

def Move_Button_Click():
    #------------------------------
    __Button_Pressed_Proc()
    #notimplemented("Move Row")
    M23.Proc_Move_Row()

def Copy_Button_Click():
    #------------------------------
    __Button_Pressed_Proc()
    #notimplemented("Copy Row")
    P01.ActiveSheet.addrow_after_current_row(copy=True)
    #Proc_Copy_Row()

def Options_Button_Click():
    #----------------------------------
    __Button_Pressed_Proc()
    #notimplemented("Options")
    M20.Option_Dialog()

def UnHideAll_Button_Click():
    #-----------------------------------
    __Button_Pressed_Proc()
    #notimplemented("UnHide All")
    M22.Proc_UnHide_All()

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Target - ByVal 
def __Worksheet_Change(Target):
    #--------------------------------------------------------
    # This function is called if the worksheet is changed.
    # It performs several checks after a user input depending form the column of the changed cell:
    Global_Worksheet_Change(Target)

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Target - ByVal 
def __Worksheet_SelectionChange(Target):
    #-----------------------------------------------------------------
    # Is called by event if the worksheet selection has changed
    Global_Worksheet_SelectionChange(Target)

def __Worksheet_Activate():
    #-------------------------------
    Global_Worksheet_Activate()

def __Worksheet_Calculate():
    #--------------------------------
    # This proc is also called if an other workbook (from a mail/internet) is opened    24.02.20:
    if DEBUG_CHANGEEVENT:
        Debug.Print('Worksheet_Calculate() in sheet \'DCC\' called')
    if ActiveSheet is None:
        return
    if Cells.Parent.Name == ActiveSheet.Name:
        Global_Worksheet_Calculate()
        
def workbook_init(workbook):
    M01.__Release_or_Debug_Version(True)
    init_UserForms()
    for sheet in workbook.sheets:
        worksheet_init(sheet)
        
    P01.Application.set_canvas_leftclickcmd(M32.DCCSend)    
    
def worksheet_init(worksheet):
    return

    first_call=True
    if worksheet.Datasheet:
        P01.ActiveSheet=worksheet
        for row in range(3,M30.LastUsedRow()):
            M20.Update_TestButtons(row,First_Call=first_call)
            M20.Update_StartValue(row)
            first_call=False
    
def init_UserForms():
    global StatusMsg_UserForm, UserForm_Select_Typ_DCC, UserForm_Select_Typ_SX, Select_COM_Port_UserForm,UserForm_Options
    StatusMsg_UserForm = D09.CStatusMsg_UserForm()
    UserForm_Select_Typ_DCC = D02.UserForm_Select_Typ_DCC()
    UserForm_Select_Typ_SX = D02SX.UserForm_Select_Typ_SX()
    Select_COM_Port_UserForm = D08.CSelect_COM_Port_UserForm()
    UserForm_Options = D10.CUserForm_Options()

    
def notimplemented(command):
    n = tk.messagebox.showinfo(command,
                           "Not implemented yet",
                            parent=PG.dialog_parent)    

# VB2PY (UntranslatedCode) Option Explicit

#***************************
#* UserForms
#***************************

StatusMsg_UserForm = None
UserForm_Select_Typ_DCC = None
UserForm_Select_Typ_SX = None
Select_COM_Port_UserForm = None
