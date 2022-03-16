Attribute VB_Name = "M09_Languages_Add"
Option Explicit

' Add missing language specific strings in the VBA code to the "Languages" sheet
' The program parses all VBA Modules for the command "Get_Language_Str("
'
' The following files are processed:
'  - *.bas: normal VBA files
'  - *.cls: Classes (all sheets and "DieseArbeitsmappe")
'  - *.frm: VBA code of the dialoges
' The *.frx files are not processed because they contaim
'
' The source code files are generated by "Export_Code.xlsm"
'
' The function "Add_All_VBA_Strings_to_the_Languages_Sheet()" must be called
' manually to update the list of language strings.
'

Private VBA_Modul_Name As String
Private AddedCnt As Long
Private Internal_Lines As Long   ' *.Cls and *.frm files contain some internal ines which are not shown in the VB Editor

'-------------------------------------------------------------------------------
Private Function Load_File_in_String(FileName As String, ByRef ResStr As String)
'-------------------------------------------------------------------------------
  Dim fp As Integer: fp = FreeFile
  On Error GoTo ErrProc
  Open FileName For Input As #fp
  ResStr = Input(LOF(fp), fp)
  Close #fp
  On Error GoTo 0
  Load_File_in_String = True
  Exit Function
  
ErrProc:
  MsgBox "Error reading '" & FileName & "'", vbCritical, "Error reading file in 'Load_File_in_String'"
End Function

'-------------------------------------------------------
Private Function Get_LineStr(Str As String, Pos As Long)
'-------------------------------------------------------
  Dim Start As Long, EndLine As Long
  Start = InStrRev(Str, vbCr, Pos) + 1
  EndLine = InStr(Pos, Str, vbCr)
  If EndLine = 0 Then EndLine = Len(Str)
  Get_LineStr = Mid(Str, Start, EndLine - Start)
End Function

'------------------------------------------------------------------------
Private Function Get_LineNumber(Str As String, ByVal Pos As Long) As Long
'------------------------------------------------------------------------
  Dim LineCnt As Long, Cnt2 As Long
  LineCnt = 0
  Do
    Dim Start As Long
    Start = InStrRev(Str, vbCr, Pos)
    If Start > 0 Then
         Dim Line As String
         LineCnt = LineCnt + 1
         Pos = Start - 1
         If Pos = 0 Then Exit Do
    Else
         Exit Do
    End If
  Loop While True
  Get_LineNumber = LineCnt
End Function

'------------------------------------------------------------------
Private Sub Set_Internal_Lines(Str As String, ByVal Name As String)
'------------------------------------------------------------------
  Dim Pos As Long
  If LCase(right(Name, 4)) = ".cls" Or LCase(right(Name, 4)) = ".frm" Then
        Pos = InStrRev(Str, "Attribute VB_Exposed")
        Internal_Lines = Get_LineNumber(Str, Pos) ' Stimmt nicht immer ganz genau => Egal
  Else: Internal_Lines = 0
  End If
End Sub

'-------------------------------------------------------------------------------
Private Sub Add_String_if_Missing(LangStr As String, Str As String, Pos As Long)
'-------------------------------------------------------------------------------
  If LangStr <> "" Then
     Dim Row As Long
     Row = Find_Language_Str_Row(LangStr)
     If Row = 0 Then
        Dim Sh As Worksheet, DstRow As Long
        Set Sh = ThisWorkbook.Sheets(LANGUAGES_SH)                          ' 29.04.20: Added: ThisWorkbook to hopefully prevent problems at startup
        DstRow = LastUsedRowIn(LANGUAGES_SH) + 1
        With Sh
          .Cells(DstRow, LangType_Col) = "VBA"
          .Cells(DstRow, LangParamCol) = VBA_Modul_Name & " " & Get_LineNumber(Str, Pos) - Internal_Lines
          .Cells(DstRow, FirstLangCol) = "'" & LangStr
        End With
        AddedCnt = AddedCnt + 1
        Debug.Print "Adding string: '" & LangStr & "'" ' Debug
     End If
  End If
End Sub

' Das zerlegen ist gar nicht so einfach ;-(
' - Das erste Argument der "Get_Language_Str()" Funktion ist ein String.
' - Ein String kann aus mehreren �ber "&" zusammengesetzten Strings besten
' - Ein String kann direkt angegeben werden: "Hallo"
'   oder eine String Konstante sein. Hier wird aber nur "vbcr" unterst�tzt.
' - Die Liste kann sich auch �ber mehrere Code Zeilen erstrecken wenn ein
'   "_" ganz am Ende der Zeile steht.
' - In einer Zeile k�nnen mehrere "Get_Language_Str()" Aufrufe stehen

' Ablauf:
' 1. String verarbeiten (Der String kann auch erst in der n�chsten Zeile kommen '_')
' 2. Pr�fen ob ein weiterer String kommt '&' Ja => 1
' 3. Das Ende muss erreicht sein: ',' oder ')'

'----------------------------------------------------------------------------------------
Private Function Proc_Quotation_Mark_String(ByRef Start As Long, Str As String) As String
'----------------------------------------------------------------------------------------
  Do
    Dim QPos As Long
    QPos = InStr(Start, Str, """")
    If QPos = 0 Then
       MsgBox "Error: Expected ending quotation mark in line:" & vbCr & Get_LineStr(Str, QPos), vbCritical
       Exit Function
    End If
    If Mid(Str, QPos + 1, 1) = """" Then ' Second quotation mark ? => Quotation mark within a string ?
       Proc_Quotation_Mark_String = Proc_Quotation_Mark_String & Mid(Str, Start, QPos - Start + 1)
       Start = QPos + 2
    Else ' End of the string
       Proc_Quotation_Mark_String = Proc_Quotation_Mark_String & Mid(Str, Start, QPos - Start)
       Start = QPos + 1
       Exit Function
    End If
  Loop While True
End Function

'---------------------------------------------------------------------------------
Private Function Get_Constant_String(ByRef Start As Long, Str As String) As String
'---------------------------------------------------------------------------------
  ' List of known constants which is replaced. Each entry has two elements.
  ' 1. The search string
  ' 2. The replace text
  ' They are separaed by vbTab.
  ' vbCr is used to depatate the lines.
  Const Known_Constant_Strings = vbCr & _
                                 "vbCr" & vbTab & "|" & vbLf & vbCr & _
                                 "vbLf" & vbTab & vbLf & vbCr & _
                                 "StdDescStart" & vbTab & "Mit diesem Blatt kann die Konfiguration" & vbCr & _
                                 "o.ControlTipText" & vbTab & "" & vbCr

  
  Dim EndStr1 As Long, EndStr2 As Long
  EndStr1 = InStr(Start, Str, " ")
  EndStr2 = InStr(Start, Str, ")")
  If EndStr1 = 0 And EndStr2 = 0 Then
      Debug.Print "Wrong line:" & vbCr & Get_LineStr(Str, Start), "Press Ctrl+Break to debug"
      MsgBox "Error: End not found in 'Get_Constant_String()", vbCritical ' Should never happen
      Start = Start + 1
  Else
      ' Set EndStr1 to the first occourence of ' ' or ')'
      If EndStr1 = 0 Then EndStr1 = EndStr2
      If EndStr2 <> 0 And EndStr2 < EndStr1 Then EndStr1 = EndStr2
      Dim Name As String
      Name = Mid(Str, Start, EndStr1 - Start)
      Start = Start + Len(Name)
      Dim Pos As Long
      Pos = InStr(Known_Constant_Strings, vbCr & Name & vbTab)
      If Pos > 0 Then
           Dim StartTxt As Long
           StartTxt = Pos + 1 + Len(Name) + 1
           EndStr1 = InStr(StartTxt, Known_Constant_Strings, vbCr)
           If EndStr1 = 0 Then
              MsgBox "Error: 'vbCr' missing at the end of 'Known_Constant_Strings'", vbCritical, "Internal Error"
           Else
              Get_Constant_String = Mid(Known_Constant_Strings, StartTxt, EndStr1 - StartTxt)
           End If
      Else
           Debug.Print "Wrong line:" & vbCr & Get_LineStr(Str, Start)
           MsgBox "Unknown constant string: '" & Name & "'", vbCritical, "Press Ctrl+Break to debug"
      End If
  End If
End Function


'-------------------------------------------------------------------------
Private Function Proc_String(ByRef Start As Long, Str As String) As String
'-------------------------------------------------------------------------
  If Mid(Str, Start, 1) = "_" Then
     Do
       Dim c As String
       Start = Start + 1
       c = Mid(Str, Start, 1)
     Loop While InStr("_ " & vbCr & vbLf & vbTab, c) > 0
  End If
  
  If Mid(Str, Start, 1) = """" Then
     Start = Start + 1
     Proc_String = Proc_Quotation_Mark_String(Start, Str)
  Else
     Proc_String = Get_Constant_String(Start, Str)
  End If
End Function

'-------------------------------------------------------------------------------
Private Function Proc_Get_Language_Str(ByVal Pos As Long, Str As String) As Long
'-------------------------------------------------------------------------------
  Dim LangStr As String
  
  Do
     If Mid(Str, Pos, 1) = " " Then Pos = Pos + 1 ' If the first character is a "_" (See: M65_Special_Modules)
     LangStr = LangStr & Proc_String(Pos, Str)
     If Mid(Str, Pos, 1) = " " Then Pos = Pos + 1
     Select Case Mid(Str, Pos, 1)
          Case "&":  Pos = Pos + 2
          Case ",", _
               ")":  ' The "Get_Language_Str()" could have additional optional parameters => End of the string
                     ' ")" = End of the function detected
                     Add_String_if_Missing LangStr, Str, Pos
                     Proc_Get_Language_Str = Pos + 1
                     Exit Function
          Case Else: Debug.Print "Wrong line:" & vbCr & _
                                 Get_LineStr(Str, Pos)
                     MsgBox "Error: Unexpected character '" & Mid(Str, Pos, 1) & "' detected after ending quotation mark (Press break and check Debug output)", vbCritical
                     Pos = Pos + 1 ' In case the program is not stopped
       End Select
  Loop While Pos < Len(Str)
End Function



'----------------------------------------------------------
Private Function Process_File(VBAName As String) As Boolean
'UT--------------------------------------------------------
  Dim Str As String
  If Not Load_File_in_String(VBAName, Str) Then Exit Function
  VBA_Modul_Name = FileName(VBAName)
  
  Set_Internal_Lines Str, VBAName
    
  Dim Start As Long, FPos As Long
  Start = 1
  Do
     FPos = InStr(Start, Str, "Get_Language_Str(", vbBinaryCompare)
     If FPos <= 0 Then Exit Do
     If Mid(Str, FPos - Len("Function "), Len("Function ")) = "Function " Then ' Skip the function definition
        Start = Start + Len("Function Get_Language_Str(")
     Else
        Start = Proc_Get_Language_Str(FPos + Len("Get_Language_Str("), Str)
     End If
  Loop While True
  
End Function

'-------------------------------------------------------
Private Sub Add_All_VBA_Strings_to_the_Languages_Sheet()
'UT-----------------------------------------------------
  Dim SrcDir As String, DateStr As String
  DateStr = Format(Now, "DD_MM_YYYY")
  SrcDir = ThisWorkbook.Path & "\Code_" & FileName(ThisWorkbook.Name) & "_" & DateStr & "\"
  ' Debug with a single file
  'Process_File ThisWorkbook.Path & SrcDir & "M65_Special_Modules.bas":      Exit Sub
  'Process_File ThisWorkbook.Path & SrcDir & "M15_Par_Description.bas":      Exit Sub
  'Process_File ThisWorkbook.Path & SrcDir & "M65_Special_Modules.bas":      Exit Sub
  'Process_File ThisWorkbook.Path & SrcDir & "M07_COM_Port.bas":             Exit Sub
  'Process_File ThisWorkbook.Path & SrcDir & "Test_Main.bas":                Exit Sub
  'Process_File ThisWorkbook.Path & SrcDir & "M08_Load_Sheet_Data.bas":      Exit Sub
  
  
  MsgBox "Add all missing text constants in the VBA program to the languages sheet." & vbCr & _
         vbCr & _
         "Attention: The program 'Export_Code.xlsm' must be called to export the source code modules " & _
         "to the directory:" & vbCr & _
         "  '" & SrcDir & "'", vbInformation
  
  AddedCnt = 0
  Dim Res As String, Skipped As String, cnt As Long, Ext As Variant
  Const Extentions = "*.bas *.cls *.frm"

  For Each Ext In Split(Extentions, " ")
    Res = Dir(ThisWorkbook.Path & SrcDir & Ext)
    Do
      If Res <> "" Then
        If Res = "M09_Language.bas" Or Res = "M09_Languages_Add.bas" Then
          Skipped = Skipped & Res & vbCr
        Else
          Debug.Print "File: " & Res
          Process_File ThisWorkbook.Path & SrcDir & Res
          cnt = cnt + 1
        End If
        Res = Dir() ' Mit Excel f�r Mac 2016 wird der urspr�ngliche Dir-Funktionsaufruf erfolgreich ausgef�hrt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses f�hren jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
      Else
        Exit Do
      End If
    Loop While True
  Next Ext
  If cnt = 0 Then
       MsgBox "Error: Directory doesn't exist or it dosn't contain files:" & vbCr & _
              "  '" & SrcDir & "'" & vbCr & _
              vbCr & _
              "The program 'Export_Code.xlsm' must be called to export the source code modules " & _
              "to the directory.", vbCritical, "No source files found"
  Else
       MsgBox AddedCnt & " strings added to the 'languages' sheet" & vbCr & _
              cnt & " modules processed" & vbCr & _
              vbCr & _
              "Following modules have been skipped:" & vbCr & _
              Skipped & _
              "because they contain special 'Get_Language_Str' calls", vbInformation
  End If
End Sub


