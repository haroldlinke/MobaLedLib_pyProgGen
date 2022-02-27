Attribute VB_Name = "M80_Create_Multiplexer"
'----------------------------------------------------------------
' Module made by Misha 21-4-2020, Bug 'Brightness' fixed 3-5-2020 by Misha.
' Updated 26-5-2020 by Misha. Now compatible with Multiplexer 0.99a.
'
' This code is used by the Multiplexer macros and needs integration in the MobaLedLib Program Generator.
'
'----------------------------------------------------------------
'
' ToDo:
' - Support the LED_Channel
' - Control the Multiplexer by DCC/Switch/Variable
'   - Change the displayed pattern
'   - Enable/Disable the whole multiplexer



Option Explicit

#If VBA7 Then 'For 64 Bit Systems
    Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Declare PtrSafe Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpSectionNames As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
#Else 'For 32 Bit Systems
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpSectionNames As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
#End If
    
Public Const Version As String = "v0.99"
Public Const Multiplexer_INI_FILE_NAME = "Multiplexer.ini"
Public MltPlxr As Long
Private FirstOneInGroup As Boolean


'----------------------------------------------------------------
Function Get_Multiplexer_Group(Res As String, Description As String, Row As Long)
'----------------------------------------------------------------
    '   Added by Misha 29-03-2020
    '   Needed for multipling group of Leds with 'Multiplexer' commands
 
    ' Multiplexer_<Value>(#LED, #InCh, #LocInCh, Brightness, Groups, Options, RndMinTime, RndMaxTime, #CtrMode, ControlNr, NumOfLEDs)
    ' Param:                0      1      2         3           4       5          6           7          8         9        10
   
    Dim LStr As String, RStr As String, Str As String, Parts() As String, LedsInGroup As Integer, Groups As Integer, Cmd As String, i As Integer, LEDCnt As Integer, Options As Integer
    
    LStr = Left(Res, InStr(Res, ")"))
    RStr = Mid(Res, InStr(Res, ")") + 1)
    Cmd = Left(Res, InStr(Res, "("))
    Parts = Split(Mid(LStr, Len(Cmd) + 1, Len(Res) - Len(RStr) - 1), ",")
    LEDCnt = val(Parts(0))
    Groups = val(Parts(4))
    Options = val(Parts(5))
    LedsInGroup = val(ReadIniFileString("Multiplexer_" & Cells(Row, Descrip_Col).Value, "Number_Of_LEDs")) ' Added: "val" to prevent crash
    FirstOneInGroup = True
    For i = 1 To Groups
        
        Str = Str & Create_Multiplexer(Trim(LStr), LEDCnt + (i - 1) * LedsInGroup, Description, Row)
        FirstOneInGroup = False
    Next i
    Get_Multiplexer_Group = Str
    
End Function


'---------------------------------------------------------------------------------------------------------------
Private Function Create_Multiplexer(Res As String, LEDCnt As Integer, Description As String, Row As Long) As String
'---------------------------------------------------------------------------------------------------------------
    Dim ProgDir, Map As String, FileName, IniFileName As String
    Dim Nr As Integer, LStr As String, RStr As String, Cmd As String, RdCmd As String
    Dim Parts() As String, MltplxrOptions As Integer, binOptions As String, ParOpt As Integer, DstVar As String
    Dim InCh As String, LocInCh As String, Brightness As Integer, RndMinTime As Long, RndMaxTime As Long
    Dim Tmp As String
    Dim ReadStr As String, RandomDescription As String
    '  Dim INCH_RND As String, ReadLines() As String, Line As Variant, LastLine As Long, FoundCmd As Boolean
    
'    IniFileName = Get_MyExampleDir() & "\" & Multiplexer_INI_FILE_NAME
    Map = Environ("USERPROFILE") & "\Documents\" & "MyPattern_Config_Examples"
    IniFileName = Map & "\" & Multiplexer_INI_FILE_NAME
    
    ProgDir = IniFileName
    If Dir(ProgDir, vbDirectory) = "" Then
       MsgBox Get_Language_Str("Fehler das Verzeichnis existiert nicht:") & vbCr & _
              "  '" & ProgDir & "'", vbCritical, Get_Language_Str("Multiplexer Verzeichnis nicht vorhanden")
       Exit Function
    End If
    
    If Not Dir(IniFileName) <> "" Then
      MsgBox Get_Language_Str("Fehler die Datei existiert nicht:") & vbCr & _
              "  '" & IniFileName & "'", vbCritical, Get_Language_Str("Multiplexer Datei nicht gefunden!")
      Exit Function
    End If
  
    LStr = Left(Res, InStr(Res, ")"))
    RStr = Mid(Res, InStr(Res, ")") + 1)
    Cmd = Left(Res, InStr(Res, "(") - 1)
    
    Parts = Split(Mid(Left(LStr, InStr(Res, ")") - 1), Len(Cmd) + 1, Len(Res) - Len(RStr) - 1), ",")
    ' Multiplexer_<Value>(#LED, #InCh, #LocInCh, Brightness, Groups, Options, RndMinTime, RndMaxTime, CtrMode, ControlNr, NumOfLEDs)
    '                Param: 0      1      2         3           4       5          6           7         8         9        10
    
    InCh = Parts(1)                         ' Input Channel
    If InCh = " [Multiplexer]" Then InCh = " SI_1"                          ' 10.02.21: Misha
    LocInCh = Parts(2)                      ' local Input Channel
    Brightness = val(Trim(Parts(3)))
    MltplxrOptions = val(Trim(Parts(5)))        ' Which patterns should be seen on the LED display of the Multiplexer command. Decimal number represents a binary value.
    ParOpt = Count_Ones(val(Parts(5)))      ' Number off 1's (binary number off patterns to display) needed to add the correct number of INCH parameters to counter.
    RndMinTime = val(Trim(Parts(6)))        ' Minimum Time for the Random Function to switch to next pattern
    RndMaxTime = val(Trim(Parts(7)))        ' Maximum Time for the Random Function to switch to next pattern
    
    If MltplxrOptions <= 0 Then                 ' If there are zero Options then no patterns can be displayed.
        Create_Multiplexer = "  // No Patterns for command : " & Res & vbCrLf
        Exit Function
    End If
    
    ' Syntax for reading INI file
    ' Section = "Test_Multiplexer_RGB_Ext4"              ' Cmd
    ' KeyName = "Option 1 Name"                     ' Variable Name
    ' Value   = ReadIniFileString(Section, KeyName) ' Variable Value
    
    '---------------------------------------------------------------------------------------------------------------
    ' Create Multiplexer Line with Random() and Counter()
    '---------------------------------------------------------------------------------------------------------------
    Description = ReadIniFileString("Multiplexer_" & Cells(Row, Descrip_Col).Value, "Description")
    ReadStr = ReadStr & vbCrLf & Add_Description("  // " & Res, "- Excel row " & Row & " - " & Description, True)          ' Comment about Multiplexer
    
    If FirstOneInGroup Then
        ' Random( DstVar, InCh, RandMode, MinTime, MaxTime, MinOn, MaxOn)
        ' Parts(x)  0       1      2         3        4       5      6
        
'        DstVar = "MltPlxr" & Int((999 - 300 + 1) * Rnd + 300)                                        ' Random Number for DstVar between 300-999
        DstVar = "MltPlxr" & MltPlxr
        MltPlxr = MltPlxr + 1
        
        Tmp = Add_Variable_to_DstVar_List(DstVar)
        
        RandomDescription = "Trigger for Counter in " & Cmd & " with Destination Variable : " & DstVar
'        ReadStr = ReadStr & Add_Description("  Random(" & DstVar & ", SI_1, RF_SEQ," & Parts(6) & "," & Parts(7) & ", 5 Sec, 5 Sec)", RandomDescription, True)
        ReadStr = ReadStr & Add_Description("  Random(" & DstVar & "," & InCh & ", RF_SEQ," & Parts(6) & "," & Parts(7) & ", 5 Sec, 5 Sec)", RandomDescription, True) ' 10.02.21: Misha
        
        ' Counter(CtrMode, InCh, Enable, TimeOut, ...)
        ' Parts(x)  0        1      2       3      4
        ' Counter(CF_ROTATE|CF_SKIP0, #INCH_RND, SI_1, 0 Sek, #LOC_INCH)     Opties: CF_ROTATE|CF_SKIP0 and CF_RANDOM|CF_SKIP0
        ReadStr = ReadStr & Add_Description("  Counter(" & Trim(Parts(8)) & "," & DstVar & "," & InCh & ", 0 Sek" & Options_INCH(LocInCh, ParOpt) & ")", RandomDescription, False)  ' 10.02.21: Misha: Using InCh instead og SI_1

        ' How to implement the DCC InCh ?????       !!!!! 31-3-2020 Hardi HELP !!!!!
'        ReadStr = ReadStr & Add_Description("  Counter(" & Trim(Parts(8)) & "," & "INCH_DCC_22_GREEN" & ", SI_1, 0 Sek" & Options_INCH(LocInCh, ParOpt) & ")", RandomDescription, False)

    End If
    
    
    '---------------------------------------------------------------------------------------------------------------
    ' Create Pattern Lines
    '---------------------------------------------------------------------------------------------------------------
    
    Dim OptionName, OptionPattern, RestPartsFrom6 As String, PartsCount, PartNr As Integer, OptionNr As Integer
    
    ' PatternT(x)(Start LED, Brightness Level (bits), InCh, Number Output Channels, Min Brightness, Max Brightness, SwitchMode, CtrMode, < Patternconfig >)
    ' Parts(y)        0               1                 2             3                   4               5             6          7          8 ...>
    
    ' XPatternT1(9,4,LOC_INCH0+0,12,0,128,0,PM_NORMAL,1 sec,0,0,0)   /* 0_Pattern_to_stop_Multiplexer (pc)      ' 10.02.21: Misha: New block                                                                                                                                                    */
    OptionName = "0_Pattern_to_stop_Multiplexer (pc)"
    OptionPattern = "XPatternT1(#LED,4,LOC_INCH0+0,12,0,128,0,PM_NORMAL,1 sec,0,0,0)"

    RdCmd = Left(OptionPattern, InStr(OptionPattern, "("))
    Parts = Split(Mid(OptionPattern, InStr(OptionPattern, "(") + 1, Len(OptionPattern)), ",")

    PartsCount = UBound(Parts())
    RestPartsFrom6 = ""
    For PartNr = 6 To PartsCount
        RestPartsFrom6 = RestPartsFrom6 & Parts(PartNr)
        If PartNr < PartsCount Then RestPartsFrom6 = RestPartsFrom6 & ","
    Next PartNr

    ReadStr = ReadStr & Add_Description("  /* " & OptionName & " */ ", Description, False)

    ReadStr = ReadStr & Add_Description(("  " & RdCmd _
                                              & LEDCnt & "," & Parts(1) & "," _
                                              & LocInCh & "+" & OptionNr & "," _
                                              & Parts(3) & "," _
                                              & Parts(4) & "," & Brightness & "," _
                                              & RestPartsFrom6), Description, False)
    OptionNr = OptionNr + 1



    For Nr = 1 To 8
        OptionName = ReadIniFileString("Multiplexer_" & Cells(Row, Descrip_Col).Value, "Option " & Nr & " Name")
        OptionPattern = ReadIniFileString("Multiplexer_" & Cells(Row, Descrip_Col).Value, "Option " & Nr & " Pattern")
        
        RdCmd = Left(OptionPattern, InStr(OptionPattern, "("))
        Parts = Split(Mid(OptionPattern, InStr(OptionPattern, "(") + 1, Len(OptionPattern)), ",")
        
        PartsCount = UBound(Parts())
        RestPartsFrom6 = ""
        For PartNr = 6 To PartsCount
            RestPartsFrom6 = RestPartsFrom6 & Parts(PartNr)
            If PartNr < PartsCount Then RestPartsFrom6 = RestPartsFrom6 & ","
        Next PartNr
        
        binOptions = DecToBin(MltplxrOptions)
        If Right(binOptions, 1) = 1 Then    ' LSB = 1
            ReadStr = ReadStr & Add_Description("  /* Option " & Nr & " - " & OptionName & " */ ", Description, False)
            
            ReadStr = ReadStr & Add_Description(("  " & RdCmd _
                                                      & LEDCnt & "," & Parts(1) & "," _
                                                      & LocInCh & "+" & OptionNr & "," _
                                                      & Parts(3) & "," _
                                                      & Parts(4) & "," & Brightness & "," _
                                                      & RestPartsFrom6), Description, False)
            OptionNr = OptionNr + 1
        Else
            ReadStr = ReadStr & Add_Description("  /* Option " & Nr & " - " & OptionName & " */ ", Description, False)
            ReadStr = ReadStr & Add_Description("  /* Option " & Nr & " - NOT selected! */ ", Description, False)
        End If
        MltplxrOptions = Application.WorksheetFunction.Bitrshift(MltplxrOptions, 1) ' Shift to right for next bit (next Multiplexer option)
    Next Nr
    
    Create_Multiplexer = ReadStr
  
End Function


'--------------------------------------------------------------------------------------------
Function Add_Description(Cmd As String, Description As String, AddDescription As Boolean)
'--------------------------------------------------------------------------------------------
    If AddDescription Then ' The description is only added to the first line
        Cmd = AddSpaceToLen(Cmd, 109) & " /* " & Description
    ElseIf Description <> "" Then
        Cmd = AddSpaceToLen(Cmd, 109) & " /*     """
    End If
    Cmd = AddSpaceToLen(Cmd, 300) & " */"
    Add_Description = Cmd & vbCr

End Function

'--------------------------------------------------------------------------------------------
Function Count_Ones(Waarde As Integer) As Integer
'--------------------------------------------------------------------------------------------

    Dim t As Integer, binOptions As String, Count As Integer

    For t = 1 To 8
        binOptions = DecToBin(Waarde)
'        Debug.Print Waarde, " / ", binOptions
        If Right(binOptions, 1) = 1 Then
            ' LSB = 1
            Count = Count + 1
        End If
        Waarde = Application.WorksheetFunction.Bitrshift(Waarde, 1)
    Next t
    Count_Ones = Count
'    Debug.Print "Count_Ones    = ", Count
    
End Function


'--------------------------------------------------------------------------------------------
Private Sub Test_Count_Ones()
'--------------------------------------------------------------------------------------------

    Dim Options As Integer, binOptions As String, Count As Integer

    Options = 225
    binOptions = DecToBin(Options)
    Debug.Print "Options       = ", Options
    Debug.Print "binOptions    = ", binOptions
    Count = Count_Ones(Options)
'    Debug.Print "Count         = ", Count

End Sub


'--------------------------------------------------------------------------------------------
Function DecToBin(ByVal DecimalIn As Variant) As String
'--------------------------------------------------------------------------------------------
' The DecimalIn argument is limited to 79228162514264337593543950245
' (approximately 96-bits) - large numerical values must be entered
' as a String value to prevent conversion to scientific notation.

  DecToBin = ""
  DecimalIn = CDec(DecimalIn)
  Do While DecimalIn <> 0
    DecToBin = Trim$(Str$(DecimalIn - 2 * Int(DecimalIn / 2))) & DecToBin
    DecimalIn = Int(DecimalIn / 2)
  Loop
End Function


'--------------------------------------------------------------------------------------------
Private Function Options_INCH(InCh As String, Options As Integer)
'--------------------------------------------------------------------------------------------
    Dim i As Integer, Str As String
    
'    For i = 0 To Options - 1
    For i = 0 To Options        ' 10.02.21: 20201028 Misha. Change for adding extra Pattern (Zero position) to hold the Multiplexer.
        Str = Str & "," & (InCh & "+" & CStr(i))
    Next i
    Options_INCH = Str
'    Debug.Print "Options_INCH = ", Str
    
End Function


'--------------------------------------------------------------------------------------------
Function Special_Multiplexer_Ext(ByVal Res As String, ByRef LEDs As String) As String
'--------------------------------------------------------------------------------------------
'   Added by Misha 2020-3-29
'   Calculate number of used LEDs for the Multiplexer command

' Multiplexer_<Value>(#LED, #InCh, #LocInCh, Brightness, Groups, Options, RndMinTime, RndMaxTime, CtrMode, ControlNr, NumOfLEDs)
'                Param: 0      1      2         3           4       5          6           7         8         9        10

    Dim Parts As Variant, Param As Variant, Ret As String, Cmd As String, LedsInGroup As Integer, LedType As String, Temp As Integer
    
    Parts = Split(Replace(Res, ")", ""), "(")
    Param = Split(Parts(1), ",")
'    Param = Split(Parts(1), " ")                                           ' 10.02.21: Misha
    Cmd = Left(Res, InStr(Res, "("))
    Ret = Cmd & Trim(Param(0)) & ", " & _
                                 Trim(Param(1)) & ", " & _
                                 Trim(Param(2)) & ", " & _
                                 Trim(Param(3)) & ", " & _
                                 Trim(Param(4)) & ", " & _
                                 Trim(Param(5)) & ", " & _
                                 Trim(Param(6)) & ", " & _
                                 Trim(Param(7)) & ", " & _
                                 Trim(Param(8)) & ", " & _
                                 Trim(Param(9)) & ", " & _
                                 Trim(Param(10)) & ")"
    Special_Multiplexer_Ext = Ret
    
    If Cells(ActiveCell.Row, DCC_or_CAN_Add_Col).Value = "" Then
        Cells(ActiveCell.Row, DCC_or_CAN_Add_Col).Value = "[Multiplexer]"
    End If
    
    LedsInGroup = val(ReadIniFileString("Multiplexer_" & Cells(ActiveCell.Row, Descrip_Col).Value, "Number_Of_LEDs")) ' 14.06.20: Added "val" to prevent crash
    
    ' 10.02.21: 20201025 Misha. Assingning the right number of LEDs for Single and RGB LEDs.
    LedType = ReadIniFileString("Multiplexer_" & Cells(ActiveCell.Row, Descrip_Col).Value, "LED_Type")
    If LedType = "Single LEDs" Then
        ' LedType = "Single LEDs"
        LEDs = "C1-" & Trim(Param(4)) * LedsInGroup       '   Calculate number of used Single LEDs for this command
    Else
        ' LedType = "RGB LEDs"
        LEDs = Trim(Param(4)) * LedsInGroup               '   Calculate number of used RGB LEDs for this command
    End If
  
End Function


'--------------------------------------------------------------------------------------------
Function LedCount(Cmd As String)
'--------------------------------------------------------------------------------------------
'   Added by Misha 2020-3-29
'   Get number of LEDs used in this Multiplexer command
    
    Dim OldSheet As Worksheet, SelRow As Long, Row As Long
    Set OldSheet = ActiveSheet
    Worksheets(LIBMACROS_SH).Activate
    Row = 3
    With ThisWorkbook.Sheets(LIBMACROS_SH)
        Dim r As Range
        Set r = .Range(.Cells(Row, 1), .Cells(ActiveCell.SpecialCells(xlLastCell).Row, 13))
        
        Dim f As Variant
        Set f = r.Find(What:=Cmd, after:=r.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        
        If f Is Nothing Then
             MsgBox Get_Language_Str("Fehler: Die Spalte '") & Cmd & Get_Language_Str("' wurde nicht im Sheet '") & ActiveSheet.Name & Get_Language_Str("' gefunden!" & vbCr & _
                    vbCr & _
                    "Die Spaltennamen dürfen nicht verändert werden"), vbCritical, Get_Language_Str("Fehler Spaltenname nicht gefunden")
             EndProg
        Else
             SelRow = f.Row
        End If
    
        LedCount = val(.Cells(SelRow, SM_SngLEDCOL))
    End With
    OldSheet.Activate

End Function


'--------------------------------------------------------------------------------------------
Public Function IniFileName() As String
'--------------------------------------------------------------------------------------------
  
'    IniFileName = ThisWorkbook.Path & "\" & Multiplexer_DIR & "\Multiplexer.ini"
    Dim Dir As String
  
    Dir = Environ("USERPROFILE") & "\Documents\" & "MyPattern_Config_Examples"
    IniFileName = Dir & "\" & Multiplexer_INI_FILE_NAME

End Function


'--------------------------------------------------------------------------------------------
Sub Test_ReadIniFileString()
'--------------------------------------------------------------------------------------------

Dim Value, Section, KeyName As String

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Description"
Value = ReadIniFileString(Section, KeyName)
Debug.Print Value

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 1 Name"
Value = ReadIniFileString(Section, KeyName)
Debug.Print Value

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 7 Pattern"
Value = ReadIniFileString(Section, KeyName)
Debug.Print Value

End Sub


'--------------------------------------------------------------------------------------------
Function ReadIniFileString(ByVal Section As String, ByVal KeyName As String) As String
'--------------------------------------------------------------------------------------------

Dim iNoOfCharInIni As Long
Dim sIniString, sProfileString As String
Dim Worked As Long
Dim RetStr As String * 1500
Dim StrSize As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Section = "" Or KeyName = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(1500)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Section, KeyName, "", RetStr, StrSize, IniFileName())
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function


'--------------------------------------------------------------------------------------------
Sub Test_WriteIniFileString()
'--------------------------------------------------------------------------------------------

Dim Test, Section, KeyName, Value As String

Section = "Multiplexer_Macro"
KeyName = "INI File Production Date"
Value = Now()
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Macro Syntax"
Value = "Multiplexer_RGB_Ext4(#LED, #InCh, #LocInCh, Brightness, Groups4, #Options, RndMinTime, RndMaxTime, #CtrMode)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 1 Name"
Value = "RGB_Multiplexer_3_4_Running_Blue (pc)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 2 Pattern"
Value = "PatternT2(LED,4,#LOC_INCH+1,12,0,Brightness,0,PM_NORMAL,0.1 sec,0.1 sec,0,0,0,195,48,12)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 3 Pattern"
Value = "PatternT1(LED,4,#LOC_INCH+2,12,0,Brightness,0,PM_NORMAL,100 ms,0,0,48,0,192,0,0,3,0,12,0,0)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 4 Pattern"
Value = "PatternT1(LED,4,#LOC_INCH+3,12,0,Brightness,0,PM_NORMAL,100 ms,0,240,255,255,15,0)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 5 Pattern"
Value = "To be filled!"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 6 Pattern"
Value = "To be filled!"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 7 Pattern"
Value = "PatternT1(LED,8,#LOC_INCH+6,24,0,Brightness,0,PM_PINGPONG,108 ms,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,224,0,0,0,0,0,0,0,112,224,0,0,0,0,0,0,56,112,224,0,0,0,0,0,28,56,112,224,0,0,0,0,14,28,56,112,224,0,0,0,7,14,28,56,112,224,0,128,3,7,14,28,56,112,224,192,129,3,7,14,28,56,112,224,192,129,3,7,14,28,56,112,0,192,129,3,7,14,28,56,0,0,192,129,3,7,14,28,0,0,0,192,129,3,7,14,0,0,0,0,192,129,3,7,0,0,0,0,0,192,129,3,0,0,0,0,0,0,192,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
Test = WriteIniFileString(Section, KeyName, Value)

Section = "Test_Multiplexer_RGB_Ext4"
KeyName = "Option 8 Pattern"
Value = "To be filled!"
Test = WriteIniFileString(Section, KeyName, Value)

End Sub


'--------------------------------------------------------------------------------------------
Function WriteIniFileString(ByVal Section As String, ByVal KeyName As String, ByVal Wstr As String) As String
'--------------------------------------------------------------------------------------------

Dim iNoOfCharInIni, Worked As Long
Dim sIniString As String

  iNoOfCharInIni = 0
  sIniString = ""
  If Section = "" Or KeyName = "" Then
    MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Section, KeyName, Wstr, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
End Function


