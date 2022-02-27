Attribute VB_Name = "M22_Hide_UnHide"
Option Explicit


'----------------------------
Public Sub Proc_Hide_Unhide()
'----------------------------
' If the selected range is not in the Data range an error message is generated
' If there are hidden rows in the selected range the are unhidden
' Otherwise the selected range is hidden
  Dim OldScreenupdating As Boolean
  OldScreenupdating = Application.ScreenUpdating
  Application.ScreenUpdating = False


  Dim Row As Variant, DetectedHidden As Boolean
  For Each Row In Selection.Rows
     If Row.Row >= FirstDat_Row Then
        If Row.EntireRow.Hidden Then
           DetectedHidden = True
           Exit For
        End If
     End If
  Next Row
  
  Dim FirstHiddenRow As Long, LastHiddenRow As Long
  For Each Row In Selection.Rows
     If Row.Row >= FirstDat_Row Then
        If DetectedHidden And Row.EntireRow.Hidden Then
           If FirstHiddenRow = 0 Then FirstHiddenRow = Row.Row
           LastHiddenRow = Row.Row
        End If
        Row.EntireRow.Hidden = Not DetectedHidden
      End If
  Next Row
  
  If DetectedHidden Then
    Rows(FirstHiddenRow & ":" & LastHiddenRow).Select
  End If
  Update_Start_LedNr
  Application.ScreenUpdating = OldScreenupdating
End Sub

'---------------------------
Public Sub Proc_UnHide_All()
'---------------------------
  Dim OldScreenupdating As Boolean
  OldScreenupdating = Application.ScreenUpdating
  Application.ScreenUpdating = False

  Dim FirstHiddenRow As Long, LastHiddenRow As Long
  Dim Row As Variant
  For Each Row In ActiveSheet.UsedRange.Rows
     If Row.Row >= FirstDat_Row Then
        If Row.EntireRow.Hidden Then
           If FirstHiddenRow = 0 Then FirstHiddenRow = Row.Row
           LastHiddenRow = Row.Row
        End If
        Row.EntireRow.Hidden = False
     End If
  Next Row
      
  If FirstHiddenRow > 0 Then
    Rows(FirstHiddenRow & ":" & LastHiddenRow).Select
  End If
  
  Update_Start_LedNr
  Application.ScreenUpdating = OldScreenupdating
End Sub

