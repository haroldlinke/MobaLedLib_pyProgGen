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



import mlpyproggen.M02_Public as M02
#import mlpyproggen.M03_Dialog as M03
#import mlpyproggen.M06_Write_Header as M06
#import mlpyproggen.M06_Write_Header_LED2Var as M06LED
#import mlpyproggen.M06_Write_Header_Sound as M06Sound
#import mlpyproggen.M06_Write_Header_SW as M06SW
#import mlpyproggen.M07_COM_Port as M07
#import mlpyproggen.M08_ARDUINO as M08
import mlpyproggen.M09_Language as M09
#import mlpyproggen.M09_Select_Macro as M09SM
#import mlpyproggen.M09_SelectMacro_Treeview as M09SMT
#import mlpyproggen.M10_Par_Description as M10
#import mlpyproggen.M20_PageEvents_a_Functions as M20
#import mlpyproggen.M25_Columns as M25
#import mlpyproggen.M27_Sheet_Icons as M27
#import mlpyproggen.M28_divers as M28
import mlpyproggen.M30_Tools as M30
#import mlpyproggen.M31_Sound as M31
#import mlpyproggen.M37_Inst_Libraries as M37
#import mlpyproggen.M60_CheckColors as M60
#import mlpyproggen.M70_Exp_Libraries as M70
#import mlpyproggen.M80_Create_Mulitplexer as M80

from vb2py.vbfunctions import *
from vb2py.vbdebug import *
from vb2py.vbconstants import *


from mlpyproggen.X01_Excel_Consts import *
import mlpyproggen.P01_Workbook as P01

__ParName_COL = 1
__Par_Cnt_COL = 2
__ParType_COL = 3
__Par_Min_COL = 4
__Par_Max_COL = 5
__Par_Def_COL = 6
__Par_Opt_COL = 7
ParInTx_COL = 8
ParHint_COL = 9
CHAN_TYPE_NONE = 1
CHAN_TYPE_LED = 2
CHAN_TYPE_SERIAL = 3
__FirstDatRow = 2

def __Get_ParDesc_Row(Sh, Name):
    global __ParName_COL
    fn_return_value = None
    #r = Range()

    #f = Variant()
    #------------------------------------------------------------------------
    with_0 = Sh
    r = with_0.Range(with_0.Cells(1, __ParName_COL), with_0.Cells(M30.LastUsedRowIn(Sh), __ParName_COL))
    #f = r.Find(What= Name, after= r.Cells(__FirstDatRow, 1), LookIn= xlFormulas, LookAt= xlWhole, SearchOrder= xlByRows, SearchDirection= xlNext, MatchCase= True, SearchFormat= False)
    f_row=r.Cells.index(Name)+1 #*HL
    if f_row is None:
        Debug.Print('Fehlender Parameter: ' + Name)
        P01.MsgBox('Fehler: Der Parameter Name \'' + Name + '\' wurde nicht im Sheet \'' + Sh.Name + '\' gefunden!', vbCritical, 'Internal Error')
        M30.EndProg()
    else:
        fn_return_value = f_row
    return fn_return_value

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ParName - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Typ - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Min - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Max - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Def - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Opt - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: InpTxt - ByRef 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Hint - ByRef 
def Get_Par_Data(ParName):
    DeltaCol = 2

    Row = int()

    #Sh = Worksheet()

    ActLanguage = Integer()

    Offs = int()
    #------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ActLanguage = M09.Get_ExcelLanguage()
    if ActLanguage != 0:
        Offs = 1
    Sh = P01.Sheets(M02.PAR_DESCR_SH)
    Row = __Get_ParDesc_Row(Sh, ParName)
    with_1 = Sh
    Typ = with_1.Cells(Row, __ParType_COL)
    Min = with_1.Cells(Row, __Par_Min_COL)
    Max = with_1.Cells(Row, __Par_Max_COL)
    Def = with_1.Cells(Row, __Par_Def_COL)
    Opt = with_1.Cells(Row, __Par_Opt_COL)
    InpTxt = with_1.Cells(Row, ParInTx_COL + ActLanguage * DeltaCol + Offs)
    if InpTxt == '':
        InpTxt = ParName
    Hint = with_1.Cells(Row, ParHint_COL + ActLanguage * DeltaCol + Offs)
    
    return Typ, Min, Max, Def, Opt, InpTxt, Hint

def __Test_Get_Par_Data():

    #UT----------------------------
    Typ, Min, Max, Def, Opt, InpTxt, Hint = Get_Par_Data('Pin_List')
    Debug.Print('Typ:' + Typ, 'Min:' + Min + ' Max:' + Max + ' Def:' + Def + ' Opt:' + Opt + vbCr + 'InpTxt:' + InpTxt + vbCr + 'Hint:' + Hint)

# VB2PY (UntranslatedCode) Option Explicit
