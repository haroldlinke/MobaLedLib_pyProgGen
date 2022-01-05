Attribute VB_Name = "M23_Add_Move_Del_Row"
Option Explicit

Private Move_Info_Shown As Boolean
Private Del_Row_Msg_Shown As Boolean

'-------------------------------------
Public Sub Used_Rows_All_Borderlines()
'-------------------------------------
  All_Borderlines Range(Cells(FirstDat_Row, Enable_Col), Cells(LastUsedRow, LastColumnDatSheet())) ' 27.11.21: Old: LastUsedColumn
End Sub



'--------------------------------------------
Public Sub Proc_Insert_Line_Above(c As Range)
'--------------------------------------------
  Dim OldUpdating As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  
  Dim OldMode As Variant                                                    ' 27.04.20:
  OldMode = Application.CutCopyMode
  Application.CutCopyMode = xlCut

  If c.Row = FirstDat_Row Then
      c.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  ElseIf c.Row > FirstDat_Row Then
     c.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  End If
  
  Application.CutCopyMode = OldMode                                         ' 27.04.20:
  Application.ScreenUpdating = OldUpdating
End Sub

'----------------------------
Public Sub Proc_Insert_Line()
'----------------------------
  Proc_Insert_Line_Above ActiveCell
End Sub

'------------------------
Public Sub Proc_Del_Row()
'------------------------
  If Not Del_Row_Msg_Shown Then
    If MsgBox(Get_Language_Str("Mit dem 'Lösche Zeilen' Knopf können eine oder mehrere Zeilen gelöscht werden." & vbCr & _
              vbCr & _
              "Die zu löschenden Zeilen markiert man mit der Maus oder Tastatur und der Großschreibetaste und betätigt den 'Löschen' Knopf. " & vbCr & _
              vbCr & _
              "Tipp:" & vbCr & _
              "Mit einem Klick auf den Haken an Anfang der Zeile können diese deaktiviert werden ohne sie gleich zu löschen." & vbCr & _
              "Alternativ können Zeilen über den 'Aus- und Einblenden' Knopf versteckt werden. Unsichtbare Zeilen werden nicht in die Arduino Konfiguration übernommen." & vbCr & _
              vbCr & _
              "Soll die Zeile tatsächlich gelöscht werden?"), _
              vbYesNo + vbQuestion, Get_Language_Str("Zeile löschen?")) = vbNo Then Exit Sub
  End If
  Del_Row_Msg_Shown = True
  
  Dim OldUpdating As Boolean, OldEvents As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  OldEvents = Application.EnableEvents
  Application.EnableEvents = False
  
  
  If ActiveCell.Row >= FirstDat_Row Then
     Selection.EntireRow.Delete Shift:=xlUp
     Update_Start_LedNr
     All_Borderlines Range(Cells(FirstDat_Row, Enable_Col), Cells(LastUsedRow + SPARE_ROWS, LastColumnDatSheet()))  ' 27.11.21: Old: LastUsedColumn
     Format_Cells_to_Row LastUsedRow
  End If
  
  Application.ScreenUpdating = OldUpdating
  Application.EnableEvents = OldEvents
End Sub

'-------------------------
Public Sub Proc_Move_Row()
'-------------------------
  Dim ActSh As String
  ActSh = ActiveSheet.Name
  
  If Not Move_Info_Shown Then
     If MsgBox(Get_Language_Str("Mit dem 'Verschiebe Zeilen' Knopf können eine oder mehrere Zeilen verschoben werden." & vbCr & _
            "Damit kann die Reihenfolge der Beleuchtungen oder der anderen Effekte an die physikalisch vorgegebene " & _
            "Anschlussreihenfolge angepasst werden." & vbCr & _
            vbCr & _
            "Die zu verschiebenden Zeilen markiert man mit der Maus oder Tastatur und der Großschreibetaste und betätigt den 'Verschieben' Knopf. " & _
            "Dann wählt man mit der Maus die neue Position der Zeilen. Eine Grüne Linie markiert dabei die Zielposition." & vbCr & _
            "Mit der 'ESC' Taste kann die Aktion abgebrochen werden." & vbCr & _
            vbCr & _
            "Achtung: Diese Meldung wird nur einmal pro Programmstart angezeigt."), _
            vbOKCancel, Get_Language_Str("Funktionsweise des 'Verschiebe Zeilen' Knopfes")) = vbCancel Then Exit Sub
     Move_Info_Shown = True
  End If
  If ActiveCell.Row < FirstDat_Row Then
      MsgBox Get_Language_Str("Achtung: Zum verschieben von Zeilen müssen eine oder mehrere Zellen im Datenbereich der Tabelle markiert " & _
             "sein. Die entsprechenden Zeilen können dann per Maus verschoben werden." & vbCr & _
             vbCr & _
             "Der gewählte Bereich liegt (teilweise) außerhalb des Datenbereichs"), vbInformation, Get_Language_Str("Zu verschiebende Zeile muss im Datenbereich der Tabelle liegen")
      Exit Sub
  End If
  If Selection.Row >= FirstDat_Row Then
     On Error GoTo EnabButtons
     
     ActiveSheet.EnableDisableAllButtons False
     Dim src As Range
     Set src = Selection.EntireRow
     src.Select
     Application.StatusBar = Get_Language_Str("Zeilen verschieben: Bitte Zielposition mit der Maus oder der Tastatur wählen")
     Dim DestRow As Long
     DestRow = Select_Move_Dest_by_Mouse(Enable_Col, LastColumnDatSheet()) ' Show move Cursor and wait until button is pressed   ' 27.11.21: Old: LastUsedColumn
     If DestRow > 0 Then
        Dim OldUpdating As Boolean, OldEvents As Boolean
        OldUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        OldEvents = Application.EnableEvents
        Application.EnableEvents = False                                    ' 26.09.19:
        
        If DestRow <> src.Row Then
            src.EntireRow.Cut
            Rows(DestRow & ":" & DestRow + src.Rows.Count).Insert Shift:=xlDown
            Rows(DestRow & ":" & DestRow + src.Rows.Count - 1).Select
            Update_Start_LedNr
        End If
        Used_Rows_All_Borderlines  ' Sometimes the borders get corrupted ;-(
        'Format_Cells_to_Row LastUsedRow
        Format_All_Rows
        Update_Sum_Func
        Application.ScreenUpdating = OldUpdating
        Application.EnableEvents = OldEvents                                ' 26.09.19:
     End If
     Application.StatusBar = ""
  End If
  
EnabButtons:
  Sheets(ActSh).EnableDisableAllButtons True
End Sub

'-------------------------
Public Sub Proc_Copy_Row()
'-------------------------
  Make_sure_that_Col_Variables_match
  Dim OldUpdating As Boolean, OldEvents As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  OldEvents = Application.EnableEvents
  Application.EnableEvents = False
  
  If Selection.Row >= FirstDat_Row Then
     Dim DestRow As Long, i As Long, EndDestRow As Long, RowCnt As Long
     DestRow = Selection.Row + Selection.Rows.Count
     RowCnt = Selection.Rows.Count
     
     For i = 1 To RowCnt ' Insert lines
         Rows(DestRow).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
     Next i
     
     EndDestRow = DestRow + Selection.Rows.Count - 1
     Rows(DestRow & ":" & EndDestRow) = Selection.EntireRow.Value
     Range(Cells(DestRow, Selection.Column), Cells(EndDestRow, Selection.Column + Selection.Columns.Count - 1)).Select
     Used_Rows_All_Borderlines
     Format_Cells_to_Row DestRow + Selection.Rows.Count
     Update_Sum_Func
     
     Dim Row As Long ' Copy the icons                                       ' 04.11.21:
     For Row = DestRow To DestRow + RowCnt
         Dim s As String
         s = Cells(Row, Config__Col)
         If s <> "" Then FindMacro_and_Add_Icon_and_Name s, Row, ActiveSheet
     Next Row
  End If
  Update_Start_LedNr
  Application.ScreenUpdating = OldUpdating
  Application.EnableEvents = OldEvents
End Sub
