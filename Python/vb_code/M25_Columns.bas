Attribute VB_Name = "M25_Columns"
' Contains Variables which contain the column numbers for
' the date sheets (DCC/Selectrix)

Option Explicit
Option Compare Binary ' Use case sensitive compare.




Public Col_from_Sheet As String ' The following ...Col variables have been read from this sheet

Public Filter__Col As Long
Public Inp_Typ_Col As Long
Public Start_V_Col As Long
Public Descrip_Col As Long
Public Dist_Nr_Col As Long
Public Conn_Nr_Col As Long
Public MacIcon_Col As Long  ' Macro icon column                             ' 20.10.21:
Public LanName_Col As Long  ' Language specific macro name                      "
Public Config__Col As Long
Public LED_Nr__Col As Long
Public LEDs____Col As Long
Public InCnt___Col As Long
Public LocInCh_Col As Long
Public LED_Cha_Col As Long
Public LED_TastCol As Long  ' First additional LEDs column


Public COMPort_COL As Long  ' This cell contains the COM Port of the LED Arduino
Public BUILDOP_COL As Long  ' This cell contains additional build options like "--board arduino:avr:nano:cpu=atmega328old"
Public R_UPLOD_COL As Long  ' Contains "OK" if the program of the right arduino was uploaded
Public COMPrtR_COL As Long  ' This cell contains the COM Port of the Right Arduino (DCC/Selectrix)
Public BUILDOpRCOL As Long  ' Additional build options for the right Arduino like "--board arduino:avr:nano:cpu=atmega328old"

Public Const COMPrtT_COL = 0 ' Not used in this program (It's only used in the Pattern_Configurator)


'Only valid if the DCC or CAN Page is active
Public DCC_or_CAN_Add_Col As Long

'Only valid if the Selectrix Page is active
Public SX_Channel_Col As Long
Public SX_Bitposi_Col As Long

Public Page_ID As String

#If False Then
    '--------------------------------
    Private Sub Add_Icons_and_Lines()                                           ' 20.10.21:
    '--------------------------------
      Dim Sh As Variant
      Application.EnableEvents = False
      For Each Sh In ThisWorkbook.Sheets
          If Is_Data_Sheet(Sh) Then
             Make_sure_that_Col_Variables_match Sh
             With Sh
               If .Cells(Header_Row, MacIcon_Col) = "Beleuchtung, Sound, oder andere Effekte" Then
                  With .Columns(ColumnLettersFromNr(MacIcon_Col) & ":" & ColumnLettersFromNr(MacIcon_Col))
                     '.Select ' Debug
                     .Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                     .Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                  End With
                  .Columns(ColumnLettersFromNr(MacIcon_Col) & ":" & ColumnLettersFromNr(MacIcon_Col)).ColumnWidth = 1.78
                  With .Cells(Header_Row, MacIcon_Col)
                     .FormulaR1C1 = "Icon"
                     .Orientation = 90
                     .Offset(0, 1).FormulaR1C1 = "Name"
                  End With
                  .Range(.Cells(Header_Row, LanName_Col), .Cells(LastUsedRowIn(Sh), LanName_Col)).HorizontalAlignment = xlLeft
                  .Cells(SH_VARS_ROW, Conn_Nr_Col) = ""     ' delete old value for the right com port ("Com?")
                  .Cells(SH_VARS_ROW, COMPrtR_COL) = "Com?"
               End If
             End With
          End If
      Next Sh
      Application.EnableEvents = True
    End Sub
    
    '--------------------------------
    Private Sub Del_Icons_and_Lines()                                           ' 20.10.21:
    '--------------------------------
      Dim Sh As Variant
      Application.EnableEvents = False
      For Each Sh In ThisWorkbook.Sheets
          If Is_Data_Sheet(Sh) Then
             Make_sure_that_Col_Variables_match Sh
             
             Dim First_Col As Long
             First_Col = MacIcon_Col
             With Sh
               If .Cells(Header_Row, First_Col) = "Icon" Then
                   .Columns(ColumnLettersFromNr(First_Col) & ":" & ColumnLettersFromNr(First_Col + 1)).Delete Shift:=xlToLeft
               End If
             End With
          End If
       Next Sh
      Application.EnableEvents = True
      
    End Sub
#End If

'------------------------------------------------------------------------------------------
Private Function Has_Macro_and_LanguageName_Column(Sh As Worksheet, Col As Long) As Boolean  ' 26.10.21:
'------------------------------------------------------------------------------------------
  With Sh.Cells(Header_Row, Col)
      If .Orientation = xlUpward Then
         Has_Macro_and_LanguageName_Column = True
      End If
  End With
End Function


'----------------------------------------------------------------------------------------------------------------------------------
Public Sub Make_sure_that_Col_Variables_match(Optional ByVal Sh As Worksheet = Nothing, Optional Switch_back_Target As Excel.Range)
'----------------------------------------------------------------------------------------------------------------------------------
' Fills the global variables which contain the column numbers
  If Sh Is Nothing Then Set Sh = ActiveSheet

  If Sh.Name = Col_from_Sheet Then Exit Sub ' Already read in => exit
  
  'Debug.Print "Updating the Col_Variables"
  
  If Not Switch_back_Target Is Nothing Then ' Check if the page has been changed while a cell was changed
     If Switch_back_Target.Parent.Name <> Sh.Name Then
        Debug.Print "Switching back to " & Switch_back_Target.Parent.Name & " in Make_sure_that_Col_Variables_match()"
        On Error GoTo ErrSwitchBack
        Switch_back_Target.Parent.Select
        On Error GoTo 0
     End If
     Exit Sub
  End If
  
  Page_ID = Sh.Cells(SH_VARS_ROW, PAGE_ID_COL)
  
  If Page_ID = "" Then Exit Sub                                             ' 07.10.21:
          
  ' Sheet specific columns
  SX_Channel_Col = 0
  SX_Bitposi_Col = 0
  DCC_or_CAN_Add_Col = 0
  
  Select Case Page_ID
     Case "Selectrix":  SX_Channel_Col = 4                                                    ' 24:02.20: Old: FindHeadCol(sh, Header_Row, "SX Channel [0..99]")
                        SX_Bitposi_Col = 5                                                    ' 24:02.20: Old: FindHeadCol(sh, Header_Row, "Bitposition [1..8]")
     Case "DCC":        DCC_or_CAN_Add_Col = 4                                                ' 24:02.20: Old: FindHeadCol(sh, Header_Row, "DCC Adresse")
     Case "CAN":        DCC_or_CAN_Add_Col = 4                                                ' 24:02.20: Old: FindHeadCol(sh, Header_Row, "CAN Adresse")
     Case Else:         Debug.Print "Seitenname: " & Sh.Name & " Page_ID: '" & Page_ID & "'"
                        MsgBox Get_Language_Str("Fehler: Die Excel Seite wurde gewechselt während einer Änderung in einer Zelle. " & vbCr & _
                               "Die Änderungen können nicht überprüft werden ;-(" & vbCr & _
                               vbCr & _
                               "Die Eingaben in einer Zelle müssen mit Enter abgeschlossen werden bevor die Seite gewechselt wird."), _
                               vbCritical, Get_Language_Str("Fehler: Seite gewechselt während der Eingabe in einer Zelle")
                        ' Normalerweise sollte diese Fehlermeldung nicht mehr kommen wenn Switch_back_to_Last_Sheet aktiv ist.
                        ' Wenn Col_from_Sheet = "" ist, dann kann es immer noch passiern
                        EndProg
  End Select
  
  Filter__Col = 3 ' "Filter"
  If Page_ID = "Selectrix" Then
        Inp_Typ_Col = 6 ' F: "Typ"
  Else: Inp_Typ_Col = 5 ' E: "Typ"
  End If
  Dim Ref_Col As Long
  Ref_Col = Inp_Typ_Col
  
  Start_V_Col = Ref_Col + 1  ' "Start- wert"
  Descrip_Col = Ref_Col + 2  ' "Beschreibung"
  Dist_Nr_Col = Ref_Col + 3  ' "Verteiler- Nummer"
  Conn_Nr_Col = Ref_Col + 4  ' "Stecker- Nummer"
  If Has_Macro_and_LanguageName_Column(Sh, Conn_Nr_Col + 1) Then ' 26.10.21:
       MacIcon_Col = Ref_Col + 5 ' Macro icon column
       LanName_Col = Ref_Col + 6 ' Language specific macro name
       Ref_Col = Ref_Col + 2
  Else
       MacIcon_Col = 0
       LanName_Col = 0
  End If
  Config__Col = Ref_Col + 5  ' "Beleuchtung, Sound, oder andere Effekte"
  LED_Nr__Col = Ref_Col + 6  ' "Start LedNr"
  LEDs____Col = Ref_Col + 7  ' "LEDs"
  InCnt___Col = Ref_Col + 8  ' "InCnt"
  LocInCh_Col = Ref_Col + 9  ' "Loc InCh"
  LED_Cha_Col = Ref_Col + 10 ' "LED Channel"
  LED_TastCol = Ref_Col + 11 ' "LED Taster"
  
  COMPort_COL = Inp_Typ_Col
  BUILDOP_COL = Descrip_Col
  R_UPLOD_COL = Dist_Nr_Col
  COMPrtR_COL = LanName_Col
  BUILDOpRCOL = Config__Col

  
  Col_from_Sheet = Sh.Name
  Exit Sub
  
ErrSwitchBack:
  MsgBox "Interner Fehler: Die letzte Seite '" & Col_from_Sheet & "' konnte nicht aktiviert werden", vbCritical, "Interner Fehler"
  EndProg
End Sub


'-----------------------------------------------------------------------------
Public Function Get_First_Number_of_Range(Row As Long, Col As Long) As Variant
'-----------------------------------------------------------------------------
' Accepts also a address which contains two adressed separated by '-'
' Example: '1 - 3'
  Dim Addr As Variant
    
  Addr = Replace(Replace(Cells(Row, Col), vbLf, ""), " ", "")
  If Addr = "" Then
     Get_First_Number_of_Range = ""
     Exit Function
  End If
    
  If InStr(Addr, "-") > 0 Then
       Dim Parts As Variant
       Parts = Split(Addr, "-")
       If UBound(Parts) > 1 Or IsNumeric(Parts(0)) = False Or IsNumeric(Parts(1)) = False Then
             Get_First_Number_of_Range = -9            ' Generate an error message in the following routines
       Else: Get_First_Number_of_Range = Int(val(Parts(0)))
       End If
  Else
       If IsNumeric(Addr) Then
            Get_First_Number_of_Range = Int(val(Addr))
       Else: Get_First_Number_of_Range = ""
       End If
  End If
End Function

'----------------------------------------
Public Function Get_Address_Col() As Long
'----------------------------------------
    If Page_ID = "Selectrix" Then
        Get_Address_Col = SX_Channel_Col
  Else: Get_Address_Col = DCC_or_CAN_Add_Col
  End If
End Function

'--------------------------------------------------------
Public Function Get_Address_String(Row As Long) As String                   ' 03.04.20:
'--------------------------------------------------------
' Return true if the string in the address/selectrix channel column
  Dim s As String, AddrCol As Long
  AddrCol = Get_Address_Col()
  Get_Address_String = Trim(Cells(Row, AddrCol))
End Function


'-------------------------------------------------------------------
Public Function Address_starts_with_a_Number(Row As Long) As Boolean        ' 03.04.20:
'-------------------------------------------------------------------
' Return true if the first character of the address/channel column is a number
  Dim s As String
  s = Get_Address_String(Row)
  If s <> "" Then
    Address_starts_with_a_Number = IsNumeric(left(s, 1))
  End If
End Function


