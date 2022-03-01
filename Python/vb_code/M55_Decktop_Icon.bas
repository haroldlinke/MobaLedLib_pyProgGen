Attribute VB_Name = "M55_Decktop_Icon"
Option Explicit
 
'UT------------------------------------
Private Sub TestCreateDesktopShortcut()
'UT------------------------------------
  CreateDesktopShortcut "Aber Hallo", ThisWorkbook.FullName, "mll_platine_ausschnitt3_icon.ico"
End Sub

 
'---------------------------------------------------------------------------------------------------------------
Public Function CreateDesktopShortcut(LinkName As String, BookFullName As String, IconName As String) As Boolean ' 14.02.20:
'---------------------------------------------------------------------------------------------------------------
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
      .IconLocation = Path & Sep & IconName
      .Save
  End With
     
  ' Explicitly clear memory
  Set oWsh = Nothing
  Set oShortcut = Nothing
     
  CreateDesktopShortcut = True
  Exit Function
     
ErrHandle:
End Function



