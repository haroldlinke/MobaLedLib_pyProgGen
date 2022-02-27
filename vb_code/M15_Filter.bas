Attribute VB_Name = "M15_Filter"
' Module Description:
' ~~~~~~~~~~~~~~~~~~~
' This module contains functions to load and save the filters
' The filters are internaly stored in the structure FilterButton_T
'
'
' Die Filter werden in der Struktur "FilterButton_T" gespeichert
'
' Zur Umwandlung werden folgende
' Funktionen benötigt:
' - Excel    => Struct          Function  Get_Excel_Filter(s As Worksheet, DestFilter As FilterButton_T) as boolean
' - Struct   => Excel           Sub       Activate_Filter (FB As FilterButton_T, s As Worksheet)

' Die Daten werden in folgender Struktur gespeichert werden:
'   FilterButton1
'     + Name                    ' Name defined by the user in the dialog
'     + Description
'     + Sheet
'     + Workbook
'     + AutoEnable
'     + EnableFilters
'     + OverwriteExisting
'     + FirstRow
'     + FirstColumn
'     + Filter1                 ' Save_One_Column_Filter_to_Registry()
'        + Enabled
'        + Column
'        + Name
'        + Count
'        + Operator
'        + Criteria1
'        + Criteria2
'     + Filter2
'     ...
'     + HideColumn1
'         + Column
'         + Name
'     + HideColumn2
'
Option Explicit

Public Const APP_NAME = "FilterTest"

Private Const Separator = "¦" ' Chr(254)

Private Enum Debug_Flags_E
  D_EVENT = 1
  D_RIBBON = 2
  D_DEL_UPDATE = 4
  D_LOAD = 8
  D_ERROR = 16
End Enum

Public Type Filter_T
  Enabled As Boolean
  Column As Long
  Name As String
  Count As Long
  Criteria1 As Variant
  Criteria2 As Variant
  Operator As XlAutoFilterOperator
End Type

Public Type HiddenCol_T
  Enabled As Boolean
  Column As Long
  Name As String
End Type

Public Type FilterButton_T
  FilterName As String                  ' Name in the registry
  Name As String                        ' User defined name
  Description As String                 ' User entered description
  Sheet As String                       ' Excel sheet name (could include wild cards '?', '*')
  Workbook As String                    ' Excel wokksheet name without path (could include wild cards '?', '*')
  AutoEnable As Boolean
  EnableFilters As Boolean
  OverwriteExisting As Boolean
  FirstRow As Long
  FirstColumn As Long
  Filters() As Filter_T
  EnableColumnHiding As Boolean
  ShowAllOtherColumns As Boolean
  HiddenCols() As HiddenCol_T
  Save_ColWidth As Boolean
  ColWidthList As String
End Type

Private Debug_Mode As Integer
Private Print_Part_Act As Boolean

Private Const DEBUG_MASK = D_EVENT + D_RIBBON + D_DEL_UPDATE + D_LOAD + D_ERROR

'---------------------------
Private Function isProgDev()
'---------------------------
    isProgDev = UCase(Environ("USERNAME")) = "Hardi"
End Function


'-----------------------------------------------------------------
Private Function Check_DebugMode(Flag As Debug_Flags_E) As Boolean
'-----------------------------------------------------------------
#If True Then  ' Disable this line to use the debug mode for all users
  If Debug_Mode = 0 Then
     If isProgDev() Then
           Debug_Mode = 1
     Else: Debug_Mode = -1
     End If
  End If
#Else
  Debug_Mode = 1
#End If
  
  If Debug_Mode = 1 Then
     If Flag And DEBUG_MASK > 0 Then Check_DebugMode = True
  End If
End Function

'--------------------------------------------------------
Private Sub D_Print(Flag As Debug_Flags_E, Txt As String)
'--------------------------------------------------------
  If Check_DebugMode(Flag) Then
     If Print_Part_Act Then
           Print_Part_Act = False
           Debug.Print Txt
     Else: Debug.Print "AF: " & Format(Now, "hh:mm:ss ") & Txt
     End If
  End If
End Sub


'---------------------------------------------------------------------------------------------------------------------------------
Private Sub Get_General_Excel_Column_Filters(s As Worksheet, af As AutoFilter, ByRef FB As FilterButton_T, SetDefaults As Boolean)
'---------------------------------------------------------------------------------------------------------------------------------
  With FB
    If SetDefaults Then
        .Name = ""        ' Has to be entered later by the user
        .Description = "" ' Has to be entered later by the user
        .Sheet = s.Name
        .Workbook = s.Parent.Name
        .AutoEnable = False
        .OverwriteExisting = True
        .ShowAllOtherColumns = True
        If Not af Is Nothing Then
           .FirstRow = af.Range.Row
           .FirstColumn = af.Range.Column
           .EnableFilters = True                                            ' 24.09.19
        End If
        
        .Save_ColWidth = False
    End If
    
    ' Store the column width
    Dim Col As Long, List As String
    For Col = 1 To LastUsedColumnIn(s)
       List = List & s.Columns(Col).ColumnWidth & Separator
    Next Col
    List = Left(List, Len(List) - Len(Separator))
    .ColWidthList = List
  End With
End Sub

'------------------------------------------------------------
Private Function Get_Filtered_Count_Str(r As Range) As String
'------------------------------------------------------------
  Get_Filtered_Count_Str = r.SpecialCells(xlCellTypeVisible).Count - 1
End Function

'------------------------------------------------------------------
Private Function Enter_Percent(f As Filter, DestFilter As Filter_T)
'------------------------------------------------------------------
' ToDo: Evtl. kann man das auch automatisch berechen.
'       Die genaue Prozentzahl zu treffen wird aber nicht möglich sein.
  Dim TypeStr As String
  With f
    If .Operator = xlTop10Percent Then
          TypeStr = "top"
    Else: TypeStr = "bottom"
    End If
    Enter_Percent = InputBox("Enter the percentage for the '" & TypeStr & " n percent'" & vbCr & _
                             "filter in column " & ColumnLettersFromNr(DestFilter.Column) & ": '" & DestFilter.Name & "'", _
                             "Enter percentage for filter", 10)
  End With
End Function


'--------------------------------------------------------------------------------------------------------------
Private Sub Store_one_Column_Filter(f As Filter, Head As Range, ByRef DestFilter As Filter_T, af As AutoFilter)
'--------------------------------------------------------------------------------------------------------------
' Store the excel autofilter of one column to the struct Filter_T
'
' The following filters need special treatement
' - Cell Color
' - Top10 / Bottom10     The Top10 filter show the 10 greatest numbers. In total 10 rows are shown even if the same number exists several times.
'
  With f
    DestFilter.Enabled = True
    DestFilter.Column = Head.Column
    DestFilter.Name = Head.Value
    DestFilter.Count = .Count
    DestFilter.Operator = .Operator '  XlAutoFilterOperator: (xlAnd = 1), (xlOr = 2), ..., (xlFilterValues = 7), ..
    
    If .Count <> 0 Then
        'If TypeName(.Criteria1) = "Variant()" Then ' Operator = (xlFilterValues = 7) oder ... ToDo
        If isVariantArray(.Criteria1) Then
           DestFilter.Criteria1 = .Criteria1
        Else ' Criteria1 is no variant
           Select Case .Operator
              Case xlTop10Items, _
                   xlBottom10Items: DestFilter.Criteria1 = Get_Filtered_Count_Str(af.Range.Columns(Head.Column))
              Case xlTop10Percent, _
                   xlBottom10Percent: DestFilter.Criteria1 = Enter_Percent(f, DestFilter)
              Case xlFilterNoFill, _
                   xlFilterCellColor: DestFilter.Criteria1 = .Criteria1.Color
              Case Else:              DestFilter.Criteria1 = .Criteria1
           End Select
           Select Case .Count
              Case 1:    ' Nothing
              Case 2:    DestFilter.Criteria2 = .Criteria2
              Case Else: MsgBox "Unexpected number of filter criterias (" & .Count & ") in column " & ColumnLetters(Head) & ": '" & Head.Value & "'", vbCritical, "Error writing filter to registry"
           End Select
        End If
    Else
        MsgBox "Unsupported filter criteria in column " & ColumnLetters(Head) & ": '" & Head.Value & "'", vbCritical, "Error writing filter to registry"  ' Must be store some were else => ToDo:
    End If
  End With
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Function Get_All_Excel_Column_Filters(s As Worksheet, af As AutoFilter, ByRef Dest As FilterButton_T, SetDefaults As Boolean) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
  Get_General_Excel_Column_Filters s, af, Dest, SetDefaults
  
  Dim i As Long, Head As Range, FilterNr As Long ' , F As Filter
  With af
    For i = 1 To .Filters.Count
        If .Filters(i).On Then
            Set Head = s.Cells(.Range.Row, .Range.Column + i - 1)
            ReDim Preserve Dest.Filters(FilterNr)
            Store_one_Column_Filter .Filters(i), Head, Dest.Filters(FilterNr), af
            FilterNr = FilterNr + 1
        End If
    Next i
  End With
  Dest.EnableFilters = (FilterNr > 0)
  Get_All_Excel_Column_Filters = Dest.EnableFilters
End Function


'-------------------------------------------------------------------------------------
Private Function Get_HiddenCols(s As Worksheet, ByRef FB As FilterButton_T) As Boolean
'-------------------------------------------------------------------------------------
  Dim Col As Long, Cnt As Long
  With FB
    Erase .HiddenCols
    Dim Row As Long
    Row = .FirstRow
    If Row = 0 Then Row = 1
    For Col = 1 To LastUsedColumnIn(s)
        If s.Columns(Col).Hidden Then
           ReDim Preserve .HiddenCols(Cnt)
           With .HiddenCols(Cnt)
             .Enabled = True
             .Column = Col
             .Name = s.Cells(Row, Col)
             If .Name = "" Then .Name = "Column " & Col
           End With
           Cnt = Cnt + 1
        End If
    Next Col
  End With
  
  FB.EnableColumnHiding = Cnt > 0
  Get_HiddenCols = FB.EnableColumnHiding
End Function

'UT-------------------------------------------------
Private Sub Test_Get_HiddenCols()
'UT-------------------------------------------------
  Dim FB As FilterButton_T
  Debug.Print Get_HiddenCols(ActiveSheet, FB)
End Sub


'---------------------------------------------------------------------------------------------------------------
Public Function Get_Excel_Filter(s As Worksheet, ByRef DestFilter As FilterButton_T, SetDefaults As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------
' Store the excel autofilter of all columns to an Filter_T array
  If Not s.AutoFilter Is Nothing Then
        If s.AutoFilter.FilterMode Then
              Get_Excel_Filter = Get_All_Excel_Column_Filters(s, s.AutoFilter, DestFilter, SetDefaults)
        Else: Get_General_Excel_Column_Filters s, s.AutoFilter, DestFilter, SetDefaults
        End If
  Else
        ' (see Save_Active_Filters_Direct_to_Registry)
        Dim obj As Variant
        'FilterNr = 0
        For Each obj In s.ListObjects
            If Not obj.AutoFilter Is Nothing Then
               Get_Excel_Filter = Get_All_Excel_Column_Filters(s, obj.AutoFilter, DestFilter, SetDefaults)
            End If
        Next obj
  End If
  If DestFilter.Sheet = "" Then Get_General_Excel_Column_Filters s, Nothing, DestFilter, SetDefaults
  
  If Get_HiddenCols(s, DestFilter) Then Get_Excel_Filter = True

End Function


'---------------------------------------------------------------
Public Sub Activate_Filter(FB As FilterButton_T, s As Worksheet)
'---------------------------------------------------------------
' Activates the filter which is defined in the structure FilterButton_T
  Dim OldUpdate As Boolean
  OldUpdate = Application.ScreenUpdating
  Application.ScreenUpdating = False
  
  Dim NormalFilter As Boolean
  If s.AutoFilterMode Then
      If s.FilterMode Then
         NormalFilter = True
      End If
  End If
      
  Dim OldActCell As Range, RestoreActiveCell As Boolean
  If Not NormalFilter Then ' If a table is used the active cell must be inside the table
     Dim obj As Variant
     For Each obj In s.ListObjects
         If Not obj.AutoFilter Is Nothing Then
            Set OldActCell = ActiveCell
            s.Cells(obj.Range.Row, obj.Range.Column).Select
            RestoreActiveCell = True
            Exit For
         End If
     Next obj
  End If
      
  '*** Filters ***
  With FB
      If .EnableFilters Then
          Dim r As Range, i As Long
          If s.AutoFilter Is Nothing Then
               On Error GoTo ErrCreateFilter
               Set r = Range(Cells(FB.FirstRow, FB.FirstColumn), Cells(LastUsedRow(), LastColumnDatSheet()))  ' 27.11.21: Old: LastUsedColumn
               'r.Select ' Debug
               r.AutoFilter
               On Error GoTo 0
          End If
          Set r = s.Range(s.AutoFilter.Range.Address)
          If FB.OverwriteExisting Then DisableFiltersInSheet s  ' Disable the active filter
          On Error GoTo Err_UBound_EndFilterLoop
          For i = 0 To UBound(.Filters)
              On Error GoTo 0
              With .Filters(i)
                 If .Enabled Then
                    If .Operator = 0 Then .Operator = xlOr        ' Achtung Operator = 0 darf nicht vorkommen. Das passiert aber z.B. beim Datumsfilter
                    On Error GoTo ErrSetFilter
                    If IsEmpty(.Criteria2) Then
                          r.AutoFilter Field:=.Column - FB.FirstColumn + 1, Criteria1:=.Criteria1, Operator:=.Operator
                    Else: r.AutoFilter Field:=.Column - FB.FirstColumn + 1, Criteria1:=.Criteria1, Operator:=.Operator, Criteria2:=.Criteria2  ' Geht das auch mit arrays
                    End If
                    On Error GoTo 0
                 End If
              End With
          Next i
      End If
  End With
EndFilterLoop:
  
  '*** Hidden Columns ***
  With FB
     If .EnableColumnHiding Then
        On Error GoTo ErrHiding
        If FB.ShowAllOtherColumns Then s.Cells.EntireColumn.Hidden = False
        On Error GoTo Err_EndHiddenColsLoop
        For i = 0 To UBound(.HiddenCols)
            On Error GoTo 0
            With .HiddenCols(i)
               If .Enabled Then
                  s.Columns(.Column).EntireColumn.Hidden = True
                  ' ToDo: Evtl. prüfung der Spaltenüberschrift
               End If
            End With
        Next i
     End If
  End With
EndHiddenColsLoop:

  If FB.Save_ColWidth Then
     Dim Col As Long, ColWidth As Variant
     Col = 1
     For Each ColWidth In Split(FB.ColWidthList, Separator)
       s.Columns(Col).ColumnWidth = Replace(ColWidth, ",", ".") ' 28.02.18: Added: Replace(...) to prevent problems with a german komma
       Col = Col + 1
    Next ColWidth
  End If
  
  If RestoreActiveCell Then
     OldActCell.Select
  End If
  
  Application.ScreenUpdating = OldUpdate
  Exit Sub
  
'eeeeeeeeeeeeeeeeeeeeeeeeeee Error routines eeeeeeeeeeeeeeee

ErrCreateFilter:
  MsgBox "Error: The current sheet doesn't contain an autofilter" & vbCr & _
         "and the program wasn't able to define a filter for row:" & FB.FirstRow & ", col:" & FB.FirstColumn & vbCr & _
         "Check if the range is empty." & vbCr & _
         vbCr & _
         "The filter section is skipped. ", vbCritical, "No auto filter available and not able to create it"
  On Error GoTo 0
  Resume EndFilterLoop ' DON'T use goto at this point because then the next "On Error goto" doesn't work


Err_UBound_EndFilterLoop:
  D_Print D_ERROR, "Error in UBound(.Filters) detected ;-( Aborting loop"
  On Error GoTo 0
  Resume EndFilterLoop


ErrSetFilter:
  With FB.Filters(i)
    MsgBox "Error: setting filter for column " & ColumnLettersFromNr(.Column - FB.FirstColumn + 1) & ": '" & .Name & "'" & vbCr & _
           vbCr & _
           "Check if the filter conditions match with the stored filter." & vbCr & _
           "Check if the sheet is protected." & vbCr & _
           vbCr & _
           "The filter section is skipped. ", vbCritical, "Error setting filter"
  End With
  On Error GoTo 0
  Resume EndFilterLoop  ' DON'T use goto at this point because then the next "On Error goto" doesn't work
    
Err_EndHiddenColsLoop:
  D_Print D_ERROR, "Error in UBound(.HiddenCols) detected ;-( Aborting loop"
  On Error GoTo 0
  Resume EndHiddenColsLoop
    
ErrHiding:
  MsgBox "Error: hiding / unhiding columns !" & vbCr & _
         vbCr & _
         "Check if the filter conditions match with the stored filter." & vbCr & _
         "Check if the sheet is protected." & vbCr & _
         vbCr & _
         "The column hidung section is skipped. ", vbCritical, "Error hiding columns"
  On Error GoTo 0
  Resume EndHiddenColsLoop ' DON'T use goto at this point because then the next "On Error goto" doesn't work
  
End Sub


'UT------------------------------------
Private Sub Test_Get_Get_Excel_Filter()
'UT------------------------------------
  Dim ActFilter As FilterButton_T
  If Get_Excel_Filter(ActiveSheet, ActFilter, True) Then                    ' Read the active filter from the excel sheet
        DisableFiltersInSheet ActiveSheet
        
        ActiveSheet.Cells.AutoFilter
        MsgBox "Filters are disabled"
        
        Activate_Filter ActFilter, ActiveSheet   ' Activate the filter again
        
        MsgBox "And enabled again"
  Else: MsgBox "No filters active"
  End If
End Sub

'------------------------------------------
Public Sub Expand_Filter_to_all_used_Rows()
'------------------------------------------
  Dim ActFilter As FilterButton_T
  If Get_Excel_Filter(ActiveSheet, ActFilter, True) Then  ' Read the active filter from the excel sheet
       DisableFiltersInSheet ActiveSheet        ' Show all lines
       ActiveSheet.Cells.AutoFilter             ' Disable the auto filter filter
       Activate_Filter ActFilter, ActiveSheet   ' Activate the filter again using all rows
  Else ' No filter active
       If ActiveSheet.AutoFilterMode Then
          ActiveSheet.Cells.AutoFilter             ' Disable the auto filter filter
          Range(Cells(ActFilter.FirstRow, ActFilter.FirstColumn), Cells(LastUsedRow, LastColumnDatSheet())).AutoFilter  ' 27.11.21: Old: LastUsedColumn
       End If
  End If
End Sub

