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

""" https://wellsr.com/vba/2019/excel/vba-playsound-to-play-system-sounds-and-wav-files/
# VB2PY (CheckDirective) VB directive took path 1 on VBA7
-------------------------------------------------------------------------
"""

__SND_SYNC = 0x0
__SND_ASYNC = 0x1
__SND_NODEFAULT = 0x2
__SND_NOSTOP = 0x10
__SND_ALIAS = 0x10000
__SND_FILENAME = 0x20000

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ThisSound='Beep' - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ThisValue=VBMissingArgument - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ThisCount=1 - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Wait=False - ByVal 
def BeepThis2(ThisSound='Beep', ThisValue=VBMissingArgument, ThisCount=1, Wait=False):
    return #*HL
    fn_return_value = None
    sPath = String()
    flags = int()
    sMedia = '\\Media\\'
    if IsMissing(ThisValue):
        ThisValue = ThisSound
    fn_return_value = ThisValue
    if ThisCount > 1:
        Wait = True
    flags = __SND_ALIAS
    sPath = StrConv(ThisSound, vbProperCase)
    if (sPath == 'Beep'):
        Beep()
        return fn_return_value
    elif (sPath == 'Asterisk') or (sPath == 'Exclamation') or (sPath == 'Hand') or (sPath == 'Notification') or (sPath == 'Question'):
        sPath = 'System' + sPath
    elif (sPath == 'Connect') or (sPath == 'Disconnect') or (sPath == 'Fail'):
        sPath = 'Device' + sPath
    elif (sPath == 'Mail') or (sPath == 'Reminder'):
        sPath = 'Notification.' + sPath
    elif (sPath == 'Text'):
        sPath = 'Notification.SMS'
    elif (sPath == 'Message'):
        sPath = 'Notification.IM'
    elif (sPath == 'Fax'):
        sPath = 'FaxBeep'
    elif (sPath == 'Select'):
        sPath = 'CCSelect'
    elif (sPath == 'Error'):
        sPath = 'AppGPFault'
    elif (sPath == 'Close') or (sPath == 'Maximize') or (sPath == 'Minimize') or (sPath == 'Open'):
        # ok
        pass
    elif (sPath == 'Default'):
        sPath = '.' + sPath
    elif (sPath == 'Chimes') or (sPath == 'Chord') or (sPath == 'Ding') or (sPath == 'Notify') or (sPath == 'Recycle') or (sPath == 'Ringout') or (sPath == 'Tada'):
        sPath = Environ('SystemRoot') + sMedia + sPath + '.wav'
        flags = __SND_FILENAME
    else:
        if LCase(Right(ThisSound, 4)) != '.wav':
            ThisSound = ThisSound + '.wav'
        sPath = ThisSound
        if Dir(sPath) == '':
            sPath = ActiveWorkbook.Path + '\\' + ThisSound
            if Dir(sPath) == '':
                sPath = Environ('SystemRoot') + sMedia + ThisSound
        flags = __SND_FILENAME
    flags = flags + IIf(Wait, __SND_SYNC, __SND_ASYNC)
    while ThisCount > 0:
        PlaySound(sPath, 0, flags)
        ThisCount = ThisCount - 1
    return fn_return_value

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ThisSound='Beep' - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ThisValue=VBMissingArgument - ByVal 
def BeepThis1(ThisSound='Beep', ThisValue=VBMissingArgument):
    fn_return_value = None
    #-------------------------------------------------------------------------
    if IsMissing(ThisValue):
        ThisValue = ThisSound
    fn_return_value = ThisValue
    Beep()
    return fn_return_value

def __Test_BeepThis1():
    #BeepThis2 "Default"
    #BeepThis2 "Asterisk"
    #BeepThis2 "Fax"
    BeepThis2()('Windows Information Bar.wav', VBGetMissingArgument(BeepThis2(), 1), VBGetMissingArgument(BeepThis2(), 2), True)

# VB2PY (UntranslatedCode) Option Explicit
# VB2PY (UntranslatedCode) Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
