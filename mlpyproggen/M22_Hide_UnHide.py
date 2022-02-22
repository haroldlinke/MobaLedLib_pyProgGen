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
# 2021-01-07 v4.02 HL: - Else:, ByRef check done, first PoC release


from vb2py.vbfunctions import *
from vb2py.vbdebug import *
from vb2py.vbconstants import *

import mlpyproggen.M02_Public as M02
import mlpyproggen.M02_global_variables as M02GV
import mlpyproggen.M03_Dialog as M03
import mlpyproggen.M06_Write_Header as M06
import mlpyproggen.M06_Write_Header_LED2Var as M06LED
import mlpyproggen.M06_Write_Header_Sound as M06Sound
import mlpyproggen.M06_Write_Header_SW as M06SW
import mlpyproggen.M07_COM_Port as M07
import mlpyproggen.M08_ARDUINO as M08
import mlpyproggen.M09_Language as M09
import mlpyproggen.M09_Select_Macro as M09SM
import mlpyproggen.M09_SelectMacro_Treeview as M09SMT
import mlpyproggen.M10_Par_Description as M10
import mlpyproggen.M20_PageEvents_a_Functions as M20
import mlpyproggen.M25_Columns as M25
import mlpyproggen.M27_Sheet_Icons as M27
import mlpyproggen.M28_divers as M28
import mlpyproggen.M30_Tools as M30
import mlpyproggen.M31_Sound as M31
import mlpyproggen.M37_Inst_Libraries as M37
import mlpyproggen.M60_CheckColors as M60
import mlpyproggen.M70_Exp_Libraries as M70
import mlpyproggen.M80_Create_Mulitplexer as M80

import mlpyproggen.Prog_Generator as PG

import mlpyproggen.P01_Workbook as P01

from mlpyproggen.X01_Excel_Consts import *


from vb2py.vbfunctions import *
from vb2py.vbdebug import *

"""----------------------------
---------------------------
"""


def Proc_Hide_Unhide():
    OldScreenupdating = Boolean()

    Row = Variant()

    DetectedHidden = Boolean()

    FirstHiddenRow = Long()

    LastHiddenRow = Long()
    #----------------------------
    # If the selected range is not in the Data range an error message is generated
    # If there are hidden rows in the selected range the are unhidden
    # Otherwise the selected range is hidden
    OldScreenupdating = P01.Application.ScreenUpdating
    P01.Application.ScreenUpdating = False
    for Row in P01.Selection.Rows():
        if Row.Row >= M02.FirstDat_Row:
            if Row.EntireRow.Hidden:
                DetectedHidden = True
                break
    for Row in P01.Selection.Rows():
        if Row.Row >= M02.FirstDat_Row:
            if DetectedHidden and Row.EntireRow.Hidden:
                if FirstHiddenRow == 0:
                    FirstHiddenRow = Row.Row
                LastHiddenRow = Row.Row
            Row.EntireRow.Hidden = not DetectedHidden
    if DetectedHidden:
        P01.Rows(str(FirstHiddenRow) + ':' + str(LastHiddenRow)).Select()
    M20.Update_Start_LedNr()
    P01.Application.ScreenUpdating = OldScreenupdating

def Proc_UnHide_All():
    OldScreenupdating = Boolean()

    FirstHiddenRow = Long()

    LastHiddenRow = Long()

    Row = Variant()
    #---------------------------
    OldScreenupdating = P01.Application.ScreenUpdating
    P01.Application.ScreenUpdating = False
    for Row in P01.ActiveSheet.UsedRange_Rows():
        if Row.Row >= M02.FirstDat_Row:
            if Row.EntireRow.Hidden:
                if FirstHiddenRow == 0:
                    FirstHiddenRow = Row.Row
                LastHiddenRow = Row.Row
            Row.EntireRow.Hidden = False
    if FirstHiddenRow > 0:
        P01.Rows(str(FirstHiddenRow) + ':' + str(LastHiddenRow)).Select()
    M20.Update_Start_LedNr()
    P01.Application.ScreenUpdating = OldScreenupdating

# VB2PY (UntranslatedCode) Option Explicit
