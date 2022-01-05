Attribute VB_Name = "M82_Create_Pattern_H_File"
Option Explicit

Private Const TimeCnt = 64
Private Const Width_Define = 76

'-------------------------------------------------------------------------------------------------------------------------------
Private Sub Write_one_Block(fp As Integer, MName As String, Params1 As String, MTyp As String, Ram As String, Params2 As String)
'-------------------------------------------------------------------------------------------------------------------------------
  Dim Line As String, i As Long, TNr As Long, Define As String
  MName = RTrim(MName)
  Define = "#define "
  If Left(MName, Len("XPatternT")) = "XPatternT" Then
    If MName = "XPatternT" Then Print #fp, "#define USE_XFADE // 220 Bytes Flash"
    Print #fp, "#ifdef USE_XFADE"
    Print #fp, "// Drittes Makro bei dem fuer jede LED ein Byte RAM Reserviert wird. Das wird benoetigt wenn das flag _PF_XFADE gesetzt ist."
    Print #fp, "// Dummerweise kann bei dieser Funktion keine Berechnung beim Parameter LED gemacht werden ;-("
    Define = " " & Define
  End If
  
  If MName = " PatternTE" Then
    Print #fp, "// Same macros with an enable input 05.11.18:"
    Print #fp, ""
  End If
  
  For TNr = 1 To TimeCnt
    Line = AddSpaceToLen(Define & MName & TNr & "(", 22) & Params1
    For i = 1 To TNr
        Line = Line & "T" & i & ","
    Next i
    Line = AddSpaceToLen(Line & "...)", Width_Define + TimeCnt * 4) & AddSpaceToLen(MTyp & TNr & "_T,", 15) & "_CHKL(LED)+" & Ram & Params2
    
    For i = 1 To TNr
        Line = Line & "_T2B(T" & i & "),"
    Next i
    Line = Line & "_W2B(COUNT_VARARGS(__VA_ARGS__)), __VA_ARGS__,"
    Print #fp, Line
  Next TNr
  
  If Left(MName, Len("XPatternT")) = "XPatternT" Then
    Print #fp, "#endif"
  End If
  
  Print #fp, ""
End Sub



'----------------------------------
Private Sub Create_Pattern_H_File()
'----------------------------------
Dim fp As Integer, FileName As String

  FileName = ThisWorkbook.Path & "\Pattern_Macros.h"
  fp = FreeFile
  On Error GoTo WriteError
  Open FileName For Output As #fp
  Print #fp, "// Created by Create_Pattern_H_File() from " & ThisWorkbook.Name
  Print #fp, ""
  Write_one_Block fp, " PatternT ", "LED,NStru,InCh,       LEDs,Val0,Val1,Off,Mode,", " PATTERNT", "RAM5,           ", "(NStru)&0xFF,_ChkIn(InCh),SI_1,  LEDs,Val0,Val1,Off,Mode,          "
  Write_one_Block fp, "APatternT ", "LED,NStru,InCh,       LEDs,Val0,Val1,Off,Mode,", "APATTERNT", "RAM7,           ", "(NStru)&0xFF,_ChkIn(InCh),SI_1,  LEDs,Val0,Val1,Off,Mode,          "
  Write_one_Block fp, "XPatternT ", "LED,NStru,InCh,       LEDs,Val0,Val1,Off,Mode,", "APATTERNT", "RAM7+RAMN(LEDs),", "(NStru)&0xFF,_ChkIn(InCh),SI_1,  LEDs,Val0,Val1,Off,Mode|_PF_XFADE,"
  Write_one_Block fp, " PatternTE", "LED,NStru,InCh,Enable,LEDs,Val0,Val1,Off,Mode,", " PATTERNT", "RAM5,           ", "(NStru)&0xFF,_ChkIn(InCh),Enable,LEDs,Val0,Val1,Off,Mode,          "
  Write_one_Block fp, "APatternTE", "LED,NStru,InCh,Enable,LEDs,Val0,Val1,Off,Mode,", "APATTERNT", "RAM7,           ", "(NStru)&0xFF,_ChkIn(InCh),Enable,LEDs,Val0,Val1,Off,Mode,          "
  Write_one_Block fp, "XPatternTE", "LED,NStru,InCh,Enable,LEDs,Val0,Val1,Off,Mode,", "APATTERNT", "RAM7+RAMN(LEDs),", "(NStru)&0xFF,_ChkIn(InCh),Enable,LEDs,Val0,Val1,Off,Mode|_PF_XFADE,"
  Close #fp
  On Error GoTo 0
  Exit Sub
  
WriteError:
  MsgBox "Fehler beim schreiben der Macro Datei"
End Sub
