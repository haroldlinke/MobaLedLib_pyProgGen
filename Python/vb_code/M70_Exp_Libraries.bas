Attribute VB_Name = "M70_Exp_Libraries"
Option Explicit

'-------------------------------------------------------------------------------------------------------
Private Function Check_Expected_Files(ByVal Path As String, ByVal ExpectedFilesLst As String) As Boolean
'-------------------------------------------------------------------------------------------------------
  Dim Name As Variant
  For Each Name In Split(ExpectedFilesLst, " ")
     If Dir(Path & Name) = "" Then Exit Function
  Next Name
  Check_Expected_Files = True
End Function

'-----------------------------------------------------------------------------------------------------------------------------------
Private Function Make_Sure_that_GitHub_Library_Exists(ByVal Expected_DirName As String, ByVal ExpectedFilesLst As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------
  Dim DestName_for_ZIP As String, ExtractedDirName As String
  DestName_for_ZIP = Expected_DirName & ".zip"
  ExtractedDirName = Expected_DirName & "-master"
  
  
  Dim DestName As Variant, Path As String
  Path = Get_Ardu_LibDir()
  CreateFolder Path ' Create the Arduino library directory
  
  If Dir(Path & Expected_DirName, vbDirectory) <> "" Then
     Debug.Print "Directory already exists: " & Path & Expected_DirName
     If Check_Expected_Files(Path & Expected_DirName & "\", ExpectedFilesLst) Then
        Make_Sure_that_GitHub_Library_Exists = True
        Exit Function
     Else
        MsgBox Replace(Get_Language_Str("Fehler: Das Verzeichnis '#1#' existiert, es enthält aber " & _
                                        "nicht alle der erwarteten Dateien:"), "#1#", Path & Expected_DirName) & vbCr & _
                                        "  '" & ExpectedFilesLst & "'" & vbCr & _
                                        vbCr & _
                       Get_Language_Str("Das Verzeichnis muss manuell gelöscht werden!"), vbCritical, _
                       Get_Language_Str("Fehler: Einige Dateien Fehlen")
        Exit Function
     End If
  End If
  
  DestName = Get_Ardu_LibDir() & DestName_for_ZIP
  If WIN7_COMPATIBLE_DOWNLOAD Then
       F_shellExec "powershell Invoke-WebRequest """ & "https://github.com/merose/AnalogScanner/archive/master.zip"" -o:" & DestName & """"   ' 20.06.20:
  Else
       If Check_if_curl_is_Available_and_gen_Message_if_not("AnalogScanner", "https://github.com/merose/AnalogScanner/archive/master.zip") = False Then Exit Function ' 05.06.20:

       F_shellExec "powershell curl """ & "https://github.com/merose/AnalogScanner/archive/master.zip"" -o:" & DestName & """"
  End If
  '   ToDo:
  '   - Erkennung von Fehlern
  
  ' Unzip
  If Not UnzipAFile(DestName, Path) Then Exit Function ' Wenn die Datei bereits existiert, dann wird eine Windows Meldung angezeigt
   
  If Dir(Path & ExtractedDirName, vbDirectory) = "" Then
       MsgBox Get_Language_Str("Fehler beim entpacken der ZIP-Datei:") & vbCr & _
                    "  '" & DestName & "'", vbCritical, Get_Language_Str("Fehler: Zip-Datei konnte nicht entpackt werden")
       Exit Function
  Else
       On Error GoTo Error_Rename
       Name Path & ExtractedDirName As Path & Expected_DirName
       On Error GoTo 0
       If Dir(Path & Expected_DirName, vbDirectory) = "" Then GoTo Error_Rename
  End If
  Make_Sure_that_GitHub_Library_Exists = True
  
  On Error Resume Next
  Kill DestName  ' Delete the ZIP file
  On Error GoTo 0
  
  Exit Function
  
Error_Rename:
  MsgBox Replace(Replace(Get_Language_Str("Fehler beim umbenennen des Verzeichnisses" & vbCr & _
         "  '#1#'" & vbCr & _
         "nach" & vbCr & _
         "  '#2#'"), "#1#", ExtractedDirName), "#2#", Expected_DirName)
  Exit Function
End Function


'UT----------------------------
Private Sub Test_Download_Exe()
'UT----------------------------
  'Const Link_to_Exe_ZipFile = "https://www.hlinke.de/dokuwiki/lib/exe/fetch.php?media=de:mobaledcheckcolors_exe_v01.00.zip"
  ' Download ins "Downloads" Verzeichnis. Die Web Seite bleibt offen
  'Shell "Explorer """ & Link_to_Exe_ZipFile & """"
  
  ' Damit kann die Datei heruntergeladen werden ohne das eine Explorer Fenster offen bleibt             (Getestet mit Win7)
  'Shell "powershell Invoke-WebRequest """ & Link_to_Exe_ZipFile & """ -o:C:\Temp\TestDownload.zip"""
  ' mit F_shellExec wird gewartet bis der download beendet ist
  
  ' Das geht auch mit einer Exe auf GitHub: ("curl" ist eine Abkürzung für "Invoke-WebRequest")
  F_shellExec "powershell curl """ & "https://github.com/Hardi-St/MobaLedLib_Docu/blob/master/Tools/CheckColors/MobaLedCheckColors.exe?raw=true"" -o:C:\Temp\DownloadTest\MobaLedCheckColors.exe"""
  
  '   ToDo:
  '   - Erkennung von Fehlern
  
  #If False Then
    ' Unzip
    On Error Resume Next ' In case the directory exists
    MkDir "C:\Temp\TestUnZip"
    On Error GoTo 0
    UnzipAFile "C:\Temp\TestDownload.zip", "C:\Temp\TestUnZip" ' Wenn die Datei bereits existiert, dann wird eine Meldung angezeigt
  #End If
End Sub


'------------------------------------------------------------------------
Public Function Make_Sure_that_AnalogScanner_Library_Exists() As Boolean
'------------------------------------------------------------------------
  Make_Sure_that_AnalogScanner_Library_Exists = Make_Sure_that_GitHub_Library_Exists("AnalogScanner", "AnalogScanner.cpp AnalogScanner.h")
End Function

