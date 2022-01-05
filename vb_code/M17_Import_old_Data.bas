Attribute VB_Name = "M17_Import_old_Data"
Option Explicit

' Import data from an old version
'
' ToDo:
' - Sollen mehrere Sheets gespeichert / geladen werden oder nur eins
'   => Beim speichern werden alle Sheets gespeichert
' - Was ist wenn mehrere DCC Sheets vorhanden sind
'   => Die Sheets werden angelegt
' - Save und Load

Private ImportWB As Workbook

'-----------------------------------------------------------------
Private Function Close_and_Delete_Temp_Prog_Gen(TempName As String) As Boolean
'-----------------------------------------------------------------
      If Same_Name_already_open(FileNameExt(TempName)) Then Workbooks(FileNameExt(TempName)).Close Savechanges:=False
      On Error GoTo Error_Kill
      If Dir(TempName) <> "" Then Kill (TempName)
      On Error GoTo 0
      Close_and_Delete_Temp_Prog_Gen = True
      Exit Function
      
Error_Kill:
   MsgBox Get_Language_Str("Fehler beim löschen der Datei") & vbCr & _
          "  '" & TempName & "' ", vbCritical, Get_Language_Str("Fehler beim löschen des temporären version alten Programm Generators")
End Function

'---------------------------------------------------------
Private Function Select_and_Open_Old_Version() As Workbook
'---------------------------------------------------------
' Select the old version of the Prog_Generator and open it
   Dim Name As Variant, Path As String
   Path = Get_MobaUserDir()
   ChDrive Path
   ChDir Path
   Do
     Name = Application.GetOpenFilename("Program generator  (*.xlsm),Prog_Generator_*.xlsm", _
            Title:=Get_Language_Str("Altes Prog_Generator Programm auswählen von der importiert werden soll"))
     If Name <> False Then
        If InStr(FileName(Name), "Prog_Generator") = 0 Then
           If MsgBox(Get_Language_Str("Fehler: Der Dateiname muss 'Prog_Generator' enthalten." & vbCr & _
                     vbCr & _
                     "Auswahl wiederholen?"), vbQuestion + vbOKCancel, _
                     Get_Language_Str("Fehler: Falsche Datei ausgewählt")) = vbCancel Then
                 Name = ""
           Else: Name = False
           End If
        End If
        If Name = ThisWorkbook.FullName Then
           Name = False
           If MsgBox(Get_Language_Str("Fehler: Die Daten können nicht aus der aktuellen Datei importiert werden." & vbCr & _
                   vbCr & _
                   "Auswahl wiederholen?"), vbQuestion + vbOKCancel, _
                   Get_Language_Str("Fehler: Aktuelle Datei ausgewählt")) = vbCancel Then
              Name = ""
           End If
        End If
     Else: Name = ""
     End If
   Loop Until Name <> False
   
   If Name <> "" Then
      Dim TempName As String
      TempName = FilePath(Name) & "~" & FileName(Name) & "~Temp.xlsm"
      If Close_and_Delete_Temp_Prog_Gen(TempName) Then
         On Error GoTo Error_Copy
         FileCopy Name, TempName
         On Error GoTo 0
         Set Select_and_Open_Old_Version = Workbooks.Open(TempName, ReadOnly:=True)
      End If
   End If
   
   ChDir ThisWorkbook.Path
   ChDrive ThisWorkbook.Path
   Exit Function

Error_Copy:
   MsgBox Get_Language_Str("Fehler beim kopieren der Datei") & vbCr & _
          "  '" & Name & "' " & Get_Language_Str("nach") & vbCr & _
          "  '" & TempName & "'", vbCritical, Get_Language_Str("Fehler beim kopieren der alten Programm Generators")
   ChDir ThisWorkbook.Path
   ChDrive ThisWorkbook.Path
End Function

'--------------------------------------------------------------------------------------------------
Private Sub Import_from_Old_Version_CallBack(Do_Import As Boolean, Import_FromAllSheets As Boolean)
'--------------------------------------------------------------------------------------------------
  Debug.Print "Import_from_Old_Version_CallBack(" & Do_Import & ", " & Import_FromAllSheets & ")"
  
  If Do_Import Then
        Dim PGF_Name As String, Res As Boolean
        PGF_Name = ActiveWorkbook.Path & "\Import_From_old_Prog.MLL_pgf"
        Res = Save_Sheets_to_pgf(PGF_Name, Import_FromAllSheets)
        ImportWB.Close Savechanges:=False
        ThisWorkbook.Activate
        If Res Then Res = Read_PGF(PGF_Name)
  Else: ImportWB.Close Savechanges:=False
  End If
  
  Set ImportWB = Nothing
End Sub

'----------------------------------------------------
Public Sub Remove_Selection_in_Sheet(Sh As Worksheet)
'----------------------------------------------------
  If Is_Data_Sheet(Sh) And Sh.Visible = xlSheetVisible Then                 ' 25.04.20: Added: And Sh.Visible = xlSheetVisible
     Sh.Select
     On Error Resume Next ' In case something else (Textbox, ...) has been selected    ' 07.10.20:
     If Selection.Cells.Count > 1 Then
        Cells(Selection.Row, Selection.Column).Select
     End If
     On Error GoTo 0
  End If
End Sub


'------------------------------------------------
Public Sub Remove_Selections_in_all_Data_Sheets()
'------------------------------------------------
  Dim OldSheet As String, Sh As Worksheet, ScrUpd As Boolean
  ScrUpd = Application.ScreenUpdating
  Application.ScreenUpdating = False
  OldSheet = ActiveSheet.Name
  For Each Sh In Sheets
      Remove_Selection_in_Sheet Sh
  Next Sh
  
  Sheets(OldSheet).Select
  Application.ScreenUpdating = ScrUpd
End Sub

'-----------------------------------
Public Sub Import_from_Old_Version()
'-----------------------------------
  
  If MsgBox(Get_Language_Str("Mit dem folgenden Dialog wird die alte Version des Prog_Generatos ausgewählt " & _
                             "von der die Daten importiert werden sollen."), vbOKCancel, Get_Language_Str("Import der Daten von alter Programm Version")) = vbOK Then
                          
     Set ImportWB = Select_and_Open_Old_Version
     If Not ImportWB Is Nothing Then
        Application.Visible = True
        Remove_Selections_in_all_Data_Sheets
        Import_Hide_Unhide.Start "Import_from_Old_Version_CallBack"
     End If
  End If
End Sub

'-----------------------------------------------
Private Function Old_Version_exists() As Boolean
'-----------------------------------------------
   Dim Name As Variant, Path As String, ActDir As String
   ActDir = FileName(ThisWorkbook.Path)
   Path = Get_MobaUserDir() & "MobaLedLib_*"
   Name = Dir(Path, vbDirectory)
   While Name <> ""
      If FileName(Name) <> ActDir Then
         Old_Version_exists = True
         Exit Function
      End If
      Name = Dir
   Wend
   ' Since Version 1.9.6 the data are stored in the directory "MobaLedlib" and use the prefix "Ver_"  ' 26.10.20:
   Path = Get_MobaUserDir() & "MobaLedLib\Ver_*"
   Name = Dir(Path, vbDirectory)
   While Name <> ""
      If FileName(Name) <> ActDir Then
         Old_Version_exists = True
         Exit Function
      End If
      Name = Dir
   Wend
End Function

'---------------------------------------------
Public Sub Import_from_Old_Version_If_exists(ByVal CheckExisting As Boolean)                ' 12.11.21: Juergen make OldVersion check optional
'---------------------------------------------
  If CheckExisting = False Or Old_Version_exists() Then
     If MsgBox(Get_Language_Str("Sollen die Daten aus der alten Programm Version importiert werden?" & vbCr & _
               "Dieser Schritt kann auch Später über die 'Optionen/Dateien' durchgeführt werden." & vbCr & _
               vbCr & _
               "Daten jetzt importieren?"), vbQuestion + vbYesNo, Get_Language_Str("Importiren von Daten aus alter Version")) = vbYes Then
        Import_from_Old_Version
     End If
  End If
End Sub
