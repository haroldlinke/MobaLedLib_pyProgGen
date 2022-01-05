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
from vb2py.vbconstants import *

#from mlpyproggen.M02_Public import *
#from mlpyproggen.M06_Write_Header_LED2Var import *
#from mlpyproggen.M06_Write_Header_Sound import *
#from mlpyproggen.M06_Write_Header_SW import *
#from mlpyproggen.M08_ARDUINO import *
#from mlpyproggen.M09_Language import *
#from mlpyproggen.M09_Select_Macro import *
#from mlpyproggen.M20_PageEvents_a_Functions import *
#from mlpyproggen.M25_Columns import *
#from mlpyproggen.M28_divers import *
#from mlpyproggen.M30_Tools import *
#from mlpyproggen.M80_Create_Mulitplexer import *

#from mlpyproggen.P01_Workbook import *

import mlpyproggen.M02_Public as M02
import mlpyproggen.M03_Dialog as M03
#import mlpyproggen.M06_Write_Header_LED2Var as M06LED
import mlpyproggen.M06_Write_Header_Sound as M06Sound
import mlpyproggen.M06_Write_Header_SW as M06SW
import mlpyproggen.M06_Write_Header as M06
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

import mlpyproggen.P01_Workbook as P01

""" Header file generation for the LED_to_Var function
"""

__LED2Var_Tab = String()

def Init_HeaderFile_Generation_LED2Var():
    global __LED2Var_Tab
    fn_return_value = None
    #--------------------------------------------------------------
    __LED2Var_Tab = ''
    fn_return_value = True
    return fn_return_value

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Cmd - ByRef 
def Add_LED2Var_Entry(Cmd, LEDNr):
    fn_return_value = None
    Parts = vbObjectInitialize(objtype=String)

    Typ = String()

    Offset = Integer()
    #-------------------------------------------------------------------------------
    Parts = Split(Replace(Replace(Replace(Trim(Cmd), M02.SF_LED_TO_VAR, ''), ')', ''), ' ', ''), ',')
    select_0 = Trim(Parts(2))
    if (select_0 == '='):
        Typ = 'T_EQUAL_THEN'
    elif (select_0 == '!='):
        Typ = 'T_NOT_EQUAL_THEN'
    elif (select_0 == '<'):
        Typ = 'T_LESS_THEN'
    elif (select_0 == '>'):
        Typ = 'T_GREATER_THAN'
    elif (select_0 == '&'):
        Typ = 'T_BIN_MASK'
    elif (select_0 == '!&'):
        Typ = 'T_NOT_BIN_MASK'
    else:
        P01.MsgBox(Replace(M09.Get_Language_Str('Fehler: Falscher Typ \'#1#\' in der \'LED_to_Var\' Funktion'), "#1#", Parts(2)), vbCritical, M09.Get_Language_Str('Fehler: Falscher Typ in \'LED_to_Var\' Funktion'))
        return fn_return_value
    Offset = P01.val(Parts(1))
    __LED2Var_Tab = __LED2Var_Tab + '        { ' + M30.AddSpaceToLen(Parts(0) + ',', 20) + M30.AddSpaceToLen(LEDNr + Offset // 3 + ',', 7) + M30.AddSpaceToLen('(' + Offset % 3, 5) + '<< 3) | ' + M30.AddSpaceToLen(Typ + ', ', 19) + M30.AddSpaceToLen(Parts(3), 4) + '},' + vbCr
    Cmd = '// ' + Cmd
    fn_return_value = True
    return fn_return_value

def Write_Header_File_LED2Var(fp):
    global __LED2Var_Tab
    
    fn_return_value = None
    #------------------------------------------------------------------
    if __LED2Var_Tab != '':
        VBFiles.writeText(fp, '// ----- LED to Var -----', '\n')
        VBFiles.writeText(fp, '  #define USE_LED_TO_VAR', '\n')
        VBFiles.writeText(fp, '', '\n')
        VBFiles.writeText(fp, '  #define T_EQUAL_THEN     0', '\n')
        VBFiles.writeText(fp, '  #define T_NOT_EQUAL_THEN 1', '\n')
        VBFiles.writeText(fp, '  #define T_LESS_THEN      2', '\n')
        VBFiles.writeText(fp, '  #define T_GREATER_THAN   3', '\n')
        VBFiles.writeText(fp, '  #define T_BIN_MASK       4', '\n')
        VBFiles.writeText(fp, '  #define T_NOT_BIN_MASK   5', '\n')
        VBFiles.writeText(fp, '', '\n')
        VBFiles.writeText(fp, '  typedef struct', '\n')
        VBFiles.writeText(fp, '      {', '\n')
        VBFiles.writeText(fp, '      uint8_t  Var_Nr;', '\n')
        VBFiles.writeText(fp, '      uint8_t  LED_Nr;', '\n')
        VBFiles.writeText(fp, '      uint8_t  Offset_and_Typ; // ---oottt    Offset: 0..2', '\n')
        VBFiles.writeText(fp, '      uint8_t  Val;', '\n')
        VBFiles.writeText(fp, '      } __attribute__ ((packed)) LED2Var_Tab_T;', '\n')
        VBFiles.writeText(fp, '', '\n')
        VBFiles.writeText(fp, '  const PROGMEM LED2Var_Tab_T LED2Var_Tab[] =', '\n')
        VBFiles.writeText(fp, '      {', '\n')
        VBFiles.writeText(fp, '        // Var name           LED_Nr LED Offset   Typ                Compare value', '\n')
        VBFiles.writeText(fp, M30.DelLast(__LED2Var_Tab), '\n')
        VBFiles.writeText(fp, '      };', '\n')
        VBFiles.writeText(fp, '', '\n')
        VBFiles.writeText(fp, '', '\n')
    fn_return_value = True
    return fn_return_value

# VB2PY (UntranslatedCode) Option Explicit
