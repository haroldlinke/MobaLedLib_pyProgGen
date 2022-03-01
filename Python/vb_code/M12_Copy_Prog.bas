Attribute VB_Name = "M12_Copy_Prog"
Option Explicit

' Copy the program to %USERPROFILE%\Documents\Arduino\MobaLedLib_<Version>\


'---------------------------------------------------------------------------------------------------------------
Public Function FileCopy_with_Check(DestDir As String, Name As String, Optional SourceName As String) As Boolean
'---------------------------------------------------------------------------------------------------------------
' Could also copy a whole folder with all sub directories
' If SourceName is empty the file/directory "Name" is copied from the program dir to DestDir
' If a SourceName is given the file the source dir is extracted from the name.
  Dim CopyDir As Boolean, SrcPath As String, SrcName As String
  SrcPath = ThisWorkbook.Path & "\"
  SrcName = Name
  If SourceName <> "" Then
    If FilePath(SourceName) <> "" Then SrcPath = FilePath(SourceName)
    SrcName = FileNameExt(SourceName)
  End If
  If Dir(SrcPath & SrcName) = "" Then
     If Dir(SrcPath & SrcName, vbDirectory) <> "" Then ' Check if it's a directory
        CopyDir = True
     Else
        MsgBox Get_Language_Str("Fehler: Die Datei oder das Verzeichnis '") & SrcName & Get_Language_Str("' ist nicht im Verzeichnis") & vbCr & _
        "  '" & SrcPath & "'" & vbCr & _
        Get_Language_Str("vorhanden."), vbCritical, Get_Language_Str("Fehler beim kopieren")
        Exit Function
     End If
  End If
  
  On Error GoTo CopyError
  If CopyDir Then
        Dim fsoObj As Scripting.FileSystemObject 'You need to set a reference to Microsoft Scripting Runtime via the Tools | References... in the VB-editor.
        Set fsoObj = New Scripting.FileSystemObject
        fsoObj.CopyFolder SrcPath & SrcName, DestDir & Name, OverWriteFiles:=True  ' Copy directory
  Else: FileCopy SrcPath & SrcName, DestDir & Name                                 ' Copy single file
  End If
  On Error GoTo 0
  FileCopy_with_Check = True
  Exit Function
  
CopyError:
  MsgBox Get_Language_Str("Es ist ein Fehler beim kopieren der Datei oder des Verzeichnisses '") & SrcName & Get_Language_Str("' vom Verzeichnis") & vbCr & _
  "  '" & SrcPath & "'" & vbCr & _
  Get_Language_Str("nach '") & DestDir & Get_Language_Str("' aufgetreten"), vbCritical, Get_Language_Str("Fehler beim kopieren")
End Function

'----------------------------------------------------------------------------
Private Function Copy_File_With_new_Name_If_Exists(Name As String, ByRef DidCopy As Boolean) As Boolean  ' 14.06.20:
                                                                                                         ' 12.11.21: set DidCopy to true in case of copy occured
'----------------------------------------------------------------------------
  Dim ProgName As String, FullDestDir As String
  FullDestDir = FilePath(Name)
  ProgName = FullDestDir & FileNameExt(Name)

  ' Check if the program already exists
  If Dir(ProgName) <> "" Then
     Dim Nr As Long, CopyName As String
     Do
       Nr = Nr + 1
       CopyName = FullDestDir & FileName(Name) & "_Old_" & Nr & ".xlsm"
     Loop While Dir(CopyName) <> ""
     
     If Not FileCopy_with_Check(FullDestDir, FileNameExt(CopyName), FullDestDir & FileNameExt(Name)) Then Exit Function
     DidCopy = True
     MsgBox Get_Language_Str("Achtung: Die Datei '") & FileNameExt(Name) & Get_Language_Str("' existiert bereits im Verzeichnis") & vbCr & _
            "  '" & FullDestDir & "'" & vbCr & _
            vbCr & _
            Get_Language_Str("Das existierende Programm wurde unter dem Namen '") & FileName(CopyName) & Get_Language_Str("' gesichert.") & vbCr & _
            vbCr & _
            Get_Language_Str("Dieser Fehler tritt auf wenn das Programm mehrfach aus dem 'extras' Verzeichnis " & _
            "der Bibliothek gestartet wird.") & vbCr & _
            Get_Language_Str("Das Programm muss nur beim ersten mal nach der Installation der MobaLedLib aus dem " & _
            "Bibliotheksverzeichnis gestartet werden. Dabei wird es in das oben genannte Verzeichnis " & _
            "kopiert damit die Bibliothek unverändert bleibt." & vbCr & _
            "Im folgenden wird ein Link auf dem Desktop kreiert über den das Programm in Zukunft gestartet wird."), _
            vbInformation, Get_Language_Str("Programm existiert bereits (Mehrfacher Start aus 'extras' Verzeichnis)")
  End If
  Copy_File_With_new_Name_If_Exists = True
End Function

'--------------------------------------------------
Public Function Copy_Prog_If_in_LibDir() As Boolean
    Dim Result As Boolean
    Copy_Prog_If_in_LibDir = Copy_Prog_If_in_LibDir_WithResult(Result)
End Function
'--------------------------------------------------

'--------------------------------------------------
Public Function Copy_Prog_If_in_LibDir_WithResult(ByRef DidCopy As Boolean) As Boolean      ' 12.11.21: set DidCopy to true in case of copy occured
'--------------------------------------------------
' Return true if the programm was stored in the LibDir
  If InStr(UCase(ThisWorkbook.Path & "\"), UCase(Get_SrcDirInLib())) = 0 Then ' The program is not stored in the library directory
     Exit Function                                        ' => We don't have to copy it
  End If
  Copy_Prog_If_in_LibDir_WithResult = True
  DidCopy = False
  
  ' 30.05.20: Disabled
  'Dim UserProfile As String
  'UserProfile = Environ(Env_USERPROFILE)
  'If UserProfile = "" Then
  '   MsgBox Get_Language_Str("Fehler: Die Variable 'USERPROFILE' ist nicht definiert ;-(" & vbCr & _
  '            vbCr & _
  '            "Das Programm wird nach 'C:") & Get_DestDir_All() & Get_Language_Str("' kopiert."), vbInformation, _
  '          Get_Language_Str("Fehler beim kopieren des LED Programmiertool")
  '   UserProfile = "C:"
  'End If
   
  ' Create the destination directory if it doesn't exist
  Dim FullDestDir As String
  FullDestDir = Get_DestDir_All()
  On Error Resume Next ' In case the directory exists
  Dim Parts As Variant, p As Variant, ActDir As String
  Parts = Split(FullDestDir, "\")
  For Each p In Parts
    ActDir = ActDir & p & "\"
    If Len(ActDir) > 3 Then MkDir ActDir
  Next
  On Error GoTo 0
  
  ' 14.06.20: Copy both programms
  If Not FileCopy_with_Check(FullDestDir, "Pattern_Config_Examples") Then Exit Function
  If Not FileCopy_with_Check(FullDestDir, "LEDs_AutoProg") Then Exit Function
  If Not FileCopy_with_Check(FullDestDir, "Prog_Generator_Examples") Then Exit Function   ' 05.10.20:
  
  If Not FileCopy_with_Check(FullDestDir, "Icons") Then Exit Function ' Is used in both programms
  
  If Not Copy_File_With_new_Name_If_Exists(FullDestDir & SECOND_PROG & ".xlsm", DidCopy) Then Exit Function   ' 14.06.20: Copy also the second program
  If Not FileCopy_with_Check(FullDestDir, SECOND_PROG & ".xlsm") Then Exit Function
  CreateDesktopShortcut SECOND_LINK, FullDestDir & SECOND_PROG & ".xlsm", SECOND_ICON                ' 14.06.20: Create Link to the second programm
  
  Dim ProgName As String
  ProgName = FullDestDir & ThisWorkbook.Name
  If Not Copy_File_With_new_Name_If_Exists(ProgName, DidCopy) Then Exit Function  ' 14.06.20: Moved in seperate function
  
  On Error GoTo ErrSaveAs
  Application.DisplayAlerts = False
  ThisWorkbook.SaveAs ProgName
  Application.DisplayAlerts = True
  On Error GoTo 0
  
  Dim CopyMsg As String
  CopyMsg = Get_Language_Str("Das Programm wurde in das Verzeichnis") & vbCr & _
              "  '" & FullDestDir & "'" & vbCr & _
              Get_Language_Str("kopiert.") & vbCr & _
              vbCr
  'If Create_Desktoplink_w_Powershell(DSKLINKNAME, ProgName) Then
  If CreateDesktopShortcut(DSKLINKNAME, ProgName) Then
        MsgBox CopyMsg & _
               Get_Language_Str("In Zukunft kann das Programm über den Link") & vbCr & _
               "   " & Split(DSKLINKNAME, " ")(0) & vbCr & _
               "   " & Split(DSKLINKNAME, " ")(1) & vbCr & _
               Get_Language_Str("auf dem Desktop gestartet werden"), vbInformation, _
               Get_Language_Str("Programm kopiert und Link auf Desktop erzeugt")
  Else: MsgBox CopyMsg & _
               Get_Language_Str("Beim Anlegen des Links gab es Probleme ;-(" & vbCr & _
                                "Das Programm kann trotzdem von der oben angegebenen Position aus gestartet werden."), vbInformation, _
               Get_Language_Str("Programm kopiert")
  End If
  
  CreateDesktopShortcut "Wiki MobaLedLib", WikiPg_Link, WikiPg_Icon ' Create Link to the Wiki
  
  
  Exit Function

ErrSaveAs:
  Application.DisplayAlerts = True
  MsgBox Get_Language_Str("Fehler beim Speichern des Programms im Verzeichnis:") & vbCr & _
         "  '" & FilePath(ProgName) & "'", vbCritical, _
         Get_Language_Str("Fehler beim Speichen des Excel Programms")
End Function

'UT------------------------------------
Private Sub TestCreateDesktopShortcut()
'UT------------------------------------
  CreateDesktopShortcut "Aber Hallo3", ThisWorkbook.FullName  ', "Icons\02_Sarah.ico"
End Sub

'UT----------------------------------------
Private Sub TestCreateWikiDesktopShortcut()
'UT----------------------------------------
  CreateDesktopShortcut "Wiki MobaLedLib", WikiPg_Link, WikiPg_Icon
End Sub
' To update the icon Cache in windows enter
'   ie4uinit -show
' in the command line

 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CreateDesktopShortcut(LinkName As String, BookFullName As String, Optional IconName As String = DefaultIcon) As Boolean ' 14.02.20:
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' Create a custom icon shortcut on the users desktop
    
     
  ' Constant string values, you can replace "Desktop"
  ' with any Special Folders name to create the shortcut there
  Const location As String = "Desktop"
  Const LinkExt As String = ".lnk"
     
  ' Object variables
  Dim oWsh As Object, oShortcut As Object
     
  ' String variables
  Dim Sep As String, Path As String
  Dim DesktopPath As String, Shortcut As String
     
  ' Initialize variables
  Sep = Application.PathSeparator
  Path = ThisWorkbook.Path
    
  On Error GoTo ErrHandle
  ' The WScript.Shell object provides functions to read system
  ' information and environment variables, work with the registry
  ' and manage shortcuts
  Set oWsh = CreateObject("WScript.Shell")
  DesktopPath = oWsh.SpecialFolders(location)
     
     
  ' Get the path where the shortcut will be located
  Shortcut = DesktopPath & Sep & LinkName & LinkExt
     
  Set oShortcut = oWsh.CreateShortcut(Shortcut)
     
  ' Link it to this file
  With oShortcut
      .TargetPath = BookFullName
      If IconName <> "" Then
         If Dir(Path & Sep & IconName) = "" Then
             MsgBox Get_Language_Str("Fehler: Das Icon '") & IconName & Get_Language_Str("' existiert nicht"), vbCritical, Get_Language_Str("Fehler: Icon nicht im Programm Verzeichnis")
         Else
             .IconLocation = Path & Sep & IconName
         End If
      End If
      .Save
  End With
     
  ' Explicitly clear memory
  Set oWsh = Nothing
  Set oShortcut = Nothing
     
  CreateDesktopShortcut = True
  Exit Function
     
ErrHandle:
End Function

