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

"""-------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------
UT----------------------------
------------------------------------------------------------------------
"""


# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Path - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ExpectedFilesLst - ByVal 
def __Check_Expected_Files(Path, ExpectedFilesLst):
    fn_return_value = None
    Name = Variant()
    #-------------------------------------------------------------------------------------------------------
    for Name in Split(ExpectedFilesLst, ' '):
        if Dir(Path + Name) == '':
            return fn_return_value
    fn_return_value = True
    return fn_return_value

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: Expected_DirName - ByVal 
# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ExpectedFilesLst - ByVal 
def __Make_Sure_that_GitHub_Library_Exists(Expected_DirName, ExpectedFilesLst):
    fn_return_value = None
    DestName_for_ZIP = String()

    ExtractedDirName = String()

    DestName = Variant()

    Path = String()
    #-----------------------------------------------------------------------------------------------------------------------------------
    DestName_for_ZIP = Expected_DirName + '.zip'
    ExtractedDirName = Expected_DirName + '-master'
    Path = Get_Ardu_LibDir()
    CreateFolder(Path)
    if Dir(Path + Expected_DirName, vbDirectory) != '':
        Debug.Print('Directory already exists: ' + Path + Expected_DirName)
        if __Check_Expected_Files(Path + Expected_DirName + '\\', ExpectedFilesLst):
            fn_return_value = True
            return fn_return_value
        else:
            MsgBox(Replace(Get_Language_Str('Fehler: Das Verzeichnis \'#1#\' existiert, es enthält aber ' + 'nicht alle der erwarteten Dateien:'), "#1#", Path + Expected_DirName) + vbCr + '  \'' + ExpectedFilesLst + '\'' + vbCr + vbCr + Get_Language_Str('Das Verzeichnis muss manuell gelöscht werden!'), vbCritical, Get_Language_Str('Fehler: Einige Dateien Fehlen'))
            return fn_return_value
    DestName = Get_Ardu_LibDir() + DestName_for_ZIP
    if WIN7_COMPATIBLE_DOWNLOAD:
        F_shellExec('powershell Invoke-WebRequest "' + 'https://github.com/merose/AnalogScanner/archive/master.zip" -o:' + DestName + '"')
    else:
        if Check_if_curl_is_Available_and_gen_Message_if_not('AnalogScanner', 'https://github.com/merose/AnalogScanner/archive/master.zip') == False:
            return fn_return_value
            # 05.06.20:
        F_shellExec('powershell curl "' + 'https://github.com/merose/AnalogScanner/archive/master.zip" -o:' + DestName + '"')
    #   ToDo:
    #   - Erkennung von Fehlern
    # Unzip
    if not UnzipAFile(DestName, Path):
        return fn_return_value
        # Wenn die Datei bereits existiert, dann wird eine Windows Meldung angezeigt
    if Dir(Path + ExtractedDirName, vbDirectory) == '':
        MsgBox(Get_Language_Str('Fehler beim entpacken der ZIP-Datei:') + vbCr + '  \'' + DestName + '\'', vbCritical, Get_Language_Str('Fehler: Zip-Datei konnte nicht entpackt werden'))
        return fn_return_value
    else:
        # VB2PY (UntranslatedCode) On Error GoTo Error_Rename
        Name(Path + ExtractedDirName + Expected_DirName)
        # VB2PY (UntranslatedCode) On Error GoTo 0
        if Dir(Path + Expected_DirName, vbDirectory) == '':
            # VB2PY (UntranslatedCode) GoTo Error_Rename
            pass
    fn_return_value = True
    # VB2PY (UntranslatedCode) On Error Resume Next
    Kill(DestName)
    # VB2PY (UntranslatedCode) On Error GoTo 0
    return fn_return_value
    MsgBox(Replace(Replace(Get_Language_Str('Fehler beim umbenennen des Verzeichnisses' + vbCr + '  \'#1#\'' + vbCr + 'nach' + vbCr + '  \'#2#\''), "#1#", ExtractedDirName), '#2#', Expected_DirName))
    return fn_return_value
    return fn_return_value

def __Test_Download_Exe():
    #UT----------------------------
    #Const Link_to_Exe_ZipFile = "https://www.hlinke.de/dokuwiki/lib/exe/fetch.php?media=de:mobaledcheckcolors_exe_v01.00.zip"
    # Download ins "Downloads" Verzeichnis. Die Web Seite bleibt offen
    #Shell "Explorer """ & Link_to_Exe_ZipFile & """"
    # Damit kann die Datei heruntergeladen werden ohne das eine Explorer Fenster offen bleibt             (Getestet mit Win7)
    #Shell "powershell Invoke-WebRequest """ & Link_to_Exe_ZipFile & """ -o:C:\Temp\TestDownload.zip"""
    # mit F_shellExec wird gewartet bis der download beendet ist
    # Das geht auch mit einer Exe auf GitHub: ("curl" ist eine Abkürzung für "Invoke-WebRequest")
    F_shellExec('powershell curl "' + 'https://github.com/Hardi-St/MobaLedLib_Docu/blob/master/Tools/CheckColors/MobaLedCheckColors.exe?raw=true" -o:C:\\Temp\\DownloadTest\\MobaLedCheckColors.exe"')
    #   ToDo:
    #   - Erkennung von Fehlern
    ## VB2PY (CheckDirective) VB directive took path 1 on False
    # Unzip
    # VB2PY (UntranslatedCode) On Error Resume Next
    MkDir('C:\\Temp\\TestUnZip')
    # VB2PY (UntranslatedCode) On Error GoTo 0
    UnzipAFile('C:\\Temp\\TestDownload.zip', 'C:\\Temp\\TestUnZip')

def Make_Sure_that_AnalogScanner_Library_Exists():
    fn_return_value = None
    #------------------------------------------------------------------------
    fn_return_value = __Make_Sure_that_GitHub_Library_Exists('AnalogScanner', 'AnalogScanner.cpp AnalogScanner.h')
    return fn_return_value

# VB2PY (UntranslatedCode) Option Explicit
