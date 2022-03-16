Attribute VB_Name = "M27_Sheet_Icons"
Option Explicit

' ToDo: Untersuchen wie das bei anderen Skaliereungen aussieht
'       - Excel          O.K.
'       - Windows
'

Private Const Icon_Size = 11
Private Const Icon_Ext = ".bmp"
Private Const Icon_Left = 2
Private Const Icon_Top = 1

'-------------------------------------
Private Function Icon_Path() As String
'-------------------------------------
  Icon_Path = ThisWorkbook.Path & "\Icons\"
End Function


'-------------------------------------------------------------------------
Public Sub Add_Icon(Name As String, Row As Long, Optional Sh As Worksheet)
Attribute Add_Icon.VB_ProcData.VB_Invoke_Func = " \n14"
'-------------------------------------------------------------------------
    Make_sure_that_Col_Variables_match Sh
    Dim r As Range, Pic As Variant
    If Sh Is Nothing Then
       Set Sh = ActiveSheet
    End If
    If MacIcon_Col <= 0 Then Exit Sub
    If Sh.Columns(MacIcon_Col).Hidden Then Exit Sub

    On Error GoTo ErrProc
    Set Pic = Sh.Pictures.Insert(Icon_Path & Name & Icon_Ext)
    On Error GoTo 0
    With Pic
        .Locked = True
        .Placement = xlMoveAndSize
        With .ShapeRange
          If .Width > .Height Then
                .Width = Icon_Size
          Else: .Height = Icon_Size
          End If
          .left = Sh.Cells(Row, MacIcon_Col).left + Icon_Left + (Icon_Size - .Width) / 2
          .top = Sh.Cells(Row, MacIcon_Col).top + Icon_Top
        End With
        .OnAction = "SelectMacros_from_Icon"
    End With
    Exit Sub

ErrProc:
End Sub


'UT------------------------
Private Sub Test_Add_Icon()
'UT------------------------
  Add_Icon "Ampel", 10
  Add_Icon "BlueLight", 11
  Add_Icon "Andreaskreuz", 12
End Sub

'--------------------------------
Public Sub Del_Icons(r As Range)
'--------------------------------
  Dim Pic As Variant, MinTop As Double, MaxTop As Double, MinLeft As Double, MaxLeft As Double, Sh As Worksheet
  Set Sh = r.Parent
  With Sh
    MinTop = r.top
    MaxTop = MinTop + r.Height
    MinLeft = r.left
    MaxLeft = MinLeft + r.Width
    For Each Pic In .Shapes
        If Pic.top > MinTop And Pic.top < MaxTop And Pic.left >= MinLeft And Pic.left <= MaxLeft Then
           Pic.Delete
        End If
    Next Pic
  End With
End Sub

'----------------------------------------------
Public Sub Del_one_Icon_in_IconCol(Row As Long, Optional Sh As Worksheet)
'----------------------------------------------
  If Sh Is Nothing Then Set Sh = ActiveSheet
  
  Make_sure_that_Col_Variables_match Sh
  With Sh
    Del_Icons .Cells(Row, MacIcon_Col)
  End With
End Sub

'UT---------------------------------------
Private Sub Test_Del_one_Icon_in_IconCol()
'UT---------------------------------------
  Add_Icon "Ampel", 10
  Del_one_Icon_in_IconCol 10
End Sub

'---------------------------------------------------------
Private Sub Del_Icons_in_Col(Col As Long, Sh As Worksheet)
'---------------------------------------------------------
  With Sh
    Del_Icons .Range(.Cells(FirstDat_Row, Col), .Cells(MAX_ROWS, Col))
  End With
End Sub


'------------------------------------
Private Sub Del_All_Icons_in_TypCol()
'------------------------------------
  Dim Sh As Worksheet: Set Sh = ActiveSheet
  
  Make_sure_that_Col_Variables_match Sh
  Del_Icons_in_Col Inp_Typ_Col, Sh
End Sub


'--------------------------------
Public Sub Del_Icons_in_IconCol()
'--------------------------------
  Dim Sh As Worksheet: Set Sh = ActiveSheet
  
  Make_sure_that_Col_Variables_match Sh
  Del_Icons_in_Col MacIcon_Col, Sh
End Sub

'UT-----------------------------
Private Sub Test_Add_All_Icons()
'UT-----------------------------
  Dim File As String, Row As Long
  Row = 10
  File = Dir(Icon_Path & "*" & Icon_Ext)
  While File <> ""
    Add_Icon FileName(File), Row
    Cells(Row, LanName_Col) = FileName(File)
    Row = Row + 1
    File = Dir
  Wend
End Sub

'---------------------------------------------------------------------------------------------------
Private Function Show_Hide_Column_in_Sheet(Show As Boolean, Col As Long, Sh As Worksheet) As Boolean
'---------------------------------------------------------------------------------------------------
' Return true if the state has been chenged
  With Sh.Columns(ColumnLettersFromNr(Col) & ":" & ColumnLettersFromNr(Col)).EntireColumn
    If .Hidden = Show Then
       .Hidden = Not Show
       Show_Hide_Column_in_Sheet = True
    End If
  End With
End Function



'------------------------------------------------------
Private Sub Hide_Icons_Column_in_Sheet(Sh As Worksheet)
'------------------------------------------------------
  Make_sure_that_Col_Variables_match Sh
  If Show_Hide_Column_in_Sheet(False, MacIcon_Col, Sh) Then
       Del_Icons_in_Col MacIcon_Col, Sh
  End If
End Sub

'UT------------------------------------------
Private Sub Test_Hide_Icons_Column_in_Sheet()
'UT------------------------------------------
  Hide_Icons_Column_in_Sheet ActiveSheet
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------
Public Sub FindMacro_and_Add_Icon_and_Name(ByVal MacroStr As String, ByVal Row As Long, Sh As Worksheet, Optional NameOnly As Boolean)
'-------------------------------------------------------------------------------------------------------------------------------------
  Dim LibMacRow As Long
  LibMacRow = Find_Macro_in_Lib_Macros_Sheet(MacroStr)
  If LibMacRow > 0 Then
     Add_Icon_and_Name LibMacRow, Row, Sh, NameOnly:=NameOnly
  Else ' Special treatement if the line was not found in the "Lib_Macros" sheet
     If InStr(MacroStr, "Pattern") > 0 Then
        Dim OldEvents As Boolean
        OldEvents = Application.EnableEvents
        Application.EnableEvents = False
        Sh.Cells(Row, LanName_Col) = Get_Language_Str("Muster") & " Pattern_Configurator"
        Application.EnableEvents = OldEvents
        If NameOnly = False Then
           Del_one_Icon_in_IconCol Row, Sh                                  ' 31.10.21: Moved into If statemant. Otherwise the Pattern icon is deleted at program start ;-(
           Add_Icon "Pattern", Row, Sh
        End If
     End If
  End If
End Sub



'------------------------------------------------------
Private Sub Show_Icons_Column_in_Sheet(Sh As Worksheet)
'------------------------------------------------------
  Dim OldUpdating As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  Make_sure_that_Col_Variables_match Sh
  If Show_Hide_Column_in_Sheet(True, MacIcon_Col, Sh) Then
    Dim Row As Long
    For Row = FirstDat_Row To LastUsedRowIn(Sh)
        Dim s As String
        s = Sh.Cells(Row, Config__Col)
        If s <> "" Then
           FindMacro_and_Add_Icon_and_Name s, Row, Sh
        End If
    Next Row
  End If
  Application.ScreenUpdating = OldUpdating
End Sub

'UT------------------------------------------
Private Sub Test_Show_Icons_Column_in_Sheet()
'UT------------------------------------------
  Show_Icons_Column_in_Sheet ActiveSheet
End Sub


'--------------------------------------------------
Private Sub Show_Hide_Icons_Column(Show As Boolean)
'--------------------------------------------------
  Dim Sh As Worksheet
  For Each Sh In ThisWorkbook.Sheets
      If Is_Data_Sheet(Sh) Then
         Make_sure_that_Col_Variables_match Sh
         If Show Then
               Show_Icons_Column_in_Sheet Sh
         Else: Hide_Icons_Column_in_Sheet Sh
         End If
      End If
  Next Sh
End Sub

'----------------------------------------------------------------
Private Sub Update_Language_Name_Column_in_Sheet(Sh As Worksheet)
'----------------------------------------------------------------
  Dim OldUpdating As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  Make_sure_that_Col_Variables_match Sh
  'If Show_Hide_Column_in_Sheet(True, LanName_Col, Sh) Then
    Dim Row As Long
    For Row = FirstDat_Row To LastUsedRowIn(Sh)
        Dim s As String
        s = Sh.Cells(Row, Config__Col)
        If s <> "" Then
           FindMacro_and_Add_Icon_and_Name s, Row, Sh, NameOnly:=True
        End If
    Next Row
  'End If
  Application.ScreenUpdating = OldUpdating
End Sub

'UT----------------------------------------------------
Private Sub Test_Update_Language_Name_Column_in_Sheet()
'UT----------------------------------------------------
  Update_Language_Name_Column_in_Sheet ActiveSheet
End Sub


'-----------------------------------------------------
Public Sub Update_Language_Name_Column_in_all_Sheets()
'-----------------------------------------------------
  Dim Sh As Worksheet, Col As Long
  Dim OldUpdating As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  For Each Sh In ThisWorkbook.Sheets
      If Is_Data_Sheet(Sh) Then
         Update_Language_Name_Column_in_Sheet Sh
      End If
  Next Sh
  Application.ScreenUpdating = OldUpdating
End Sub

'----------------------------------
Public Sub SelectMacros_from_Icon()
'----------------------------------
    Dim Button As Object, Row As Long, top As Double
    Make_sure_that_Col_Variables_match
    On Error GoTo NotFound
    Set Button = ActiveSheet.Shapes(Application.caller)
    top = Button.top
    For Row = LastUsedRow To FirstDat_Row Step -1
         If Cells(Row, 1).top < top Then
            Cells(Row, MacIcon_Col).Select
            SelectMacros
            Exit Sub
         End If
    Next Row
NotFound:
End Sub

'UT---------------------------------------------------------------------------------------------
Private Sub Test_Hide_MacIcon_Column(): Show_Hide_Column_in_all_Sheets 0, "MacIcon_Col": End Sub
Private Sub Test_Show_MacIcon_Column(): Show_Hide_Column_in_all_Sheets 1, "MacIcon_Col": End Sub

Private Sub Test_Hide_LanName_Column(): Show_Hide_Column_in_all_Sheets 0, "LanName_Col": End Sub
Private Sub Test_Show_LanName_Column(): Show_Hide_Column_in_all_Sheets 1, "LanName_Col": End Sub

Private Sub Test_Hide_Config__Column(): Show_Hide_Column_in_all_Sheets 0, "Config__Col": End Sub
Private Sub Test_Show_Config__Column(): Show_Hide_Column_in_all_Sheets 1, "Config__Col": End Sub
'UT---------------------------------------------------------------------------------------------



'----------------------------------------------------------------------------
Public Sub Show_Hide_Column_in_all_Sheets(Show As Boolean, ColName As String)
'----------------------------------------------------------------------------
  If ColName = "MacIcon_Col" Then
       ShowHourGlassCursor True
       Show_Hide_Icons_Column Show ' Macro Icons are deleted/created to speed up the program
       ShowHourGlassCursor False
  Else
       Dim Sh As Worksheet, Col As Long
       For Each Sh In ThisWorkbook.Sheets
           If Is_Data_Sheet(Sh) Then
              Make_sure_that_Col_Variables_match Sh
              Select Case ColName
                Case "LanName_Col": Col = LanName_Col
                Case "Config__Col": Col = Config__Col
                Case Else
                     MsgBox "Unknown ColName: '" & ColName & "'", vbCritical, "Internal Error"
                     Stop
              End Select
              Show_Hide_Column_in_Sheet Show, Col, Sh
           End If
       Next Sh
  End If
End Sub


