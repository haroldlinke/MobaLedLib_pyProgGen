Attribute VB_Name = "M06_Write_Header_LED2Var"
Option Explicit

' Header file generation for the LED_to_Var function
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private LED2Var_Tab As String

'--------------------------------------------------------------
Public Function Init_HeaderFile_Generation_LED2Var() As Boolean
'--------------------------------------------------------------
  LED2Var_Tab = ""
  Init_HeaderFile_Generation_LED2Var = True
End Function

'-------------------------------------------------------------------------------
Public Function Add_LED2Var_Entry(ByRef Cmd As String, LEDNr As Long) As Boolean
'-------------------------------------------------------------------------------
  Dim Parts() As String, Typ As String
  
  Parts = Split(Replace(Replace(Replace(Trim(Cmd), SF_LED_TO_VAR, ""), ")", ""), " ", ""), ",")
  Select Case Trim(Parts(2))
     Case "=":  Typ = "T_EQUAL_THEN"
     Case "!=": Typ = "T_NOT_EQUAL_THEN"
     Case "<":  Typ = "T_LESS_THEN"
     Case ">":  Typ = "T_GREATER_THAN"
     Case "&":  Typ = "T_BIN_MASK"
     Case "!&": Typ = "T_NOT_BIN_MASK"
     Case Else: MsgBox Replace(Get_Language_Str("Fehler: Falscher Typ '#1#' in der 'LED_to_Var' Funktion"), "#1#", Parts(2)), vbCritical, _
                       Get_Language_Str("Fehler: Falscher Typ in 'LED_to_Var' Funktion")
                Exit Function
  End Select
  Dim Offset As Integer
  Offset = val(Parts(1))
  LED2Var_Tab = LED2Var_Tab & "        { " & AddSpaceToLen(Parts(0) & ",", 20) & _
                                             AddSpaceToLen(LEDNr + Offset \ 3 & ",", 7) & _
                                             AddSpaceToLen("(" & Offset Mod 3, 5) & "<< 3) | " & _
                                             AddSpaceToLen(Typ & ", ", 19) & _
                                             AddSpaceToLen(Parts(3), 4) & "}," & vbCr

  Cmd = "// " & Cmd
  Add_LED2Var_Entry = True
End Function

'------------------------------------------------------------------
Public Function Write_Header_File_LED2Var(fp As Integer) As Boolean
'------------------------------------------------------------------
  If LED2Var_Tab <> "" Then
     Print #fp, "// ----- LED to Var -----"
     Print #fp, "  #define USE_LED_TO_VAR"
     Print #fp, ""
     Print #fp, "  #define T_EQUAL_THEN     0"
     Print #fp, "  #define T_NOT_EQUAL_THEN 1"
     Print #fp, "  #define T_LESS_THEN      2"
     Print #fp, "  #define T_GREATER_THAN   3"
     Print #fp, "  #define T_BIN_MASK       4"
     Print #fp, "  #define T_NOT_BIN_MASK   5"
     Print #fp, ""
     Print #fp, "  typedef struct"
     Print #fp, "      {"
     Print #fp, "      uint8_t  Var_Nr;"
     Print #fp, "      uint8_t  LED_Nr;"
     Print #fp, "      uint8_t  Offset_and_Typ; // ---oottt    Offset: 0..2"
     Print #fp, "      uint8_t  Val;"
     Print #fp, "      } __attribute__ ((packed)) LED2Var_Tab_T;"           ' 05.11.20: Added: __attribute__ ((packed)) to be able to use it on oa 32 Bit platform
     Print #fp, ""
     Print #fp, "  const PROGMEM LED2Var_Tab_T LED2Var_Tab[] ="
     Print #fp, "      {"
     Print #fp, "        // Var name           LED_Nr LED Offset   Typ                Compare value"
     Print #fp, DelLast(LED2Var_Tab)
     Print #fp, "      };"
     Print #fp, ""
     Print #fp, ""
  
  
  End If
  Write_Header_File_LED2Var = True
End Function



