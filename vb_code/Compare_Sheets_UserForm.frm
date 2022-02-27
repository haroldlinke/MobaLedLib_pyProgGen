VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Compare_Sheets_UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9555
   OleObjectBlob   =   "Compare_Sheets_UserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Compare_Sheets_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const WB1_Name = "Prog_Generator_MobaLedLib.xlsm"
Private Const WB2_Name = "Prog_Generator_MobaLedLib copie.xlsm"
Private Const FirstCol = 6 ' F
Private Const Last_Col = 6 ' F

Private sh1 As Worksheet
Private sh2 As Worksheet

Private r1 As Range
Private r2 As Range

Private MaxCol As Long
Private MaxRow As Long

Private EndReached As Boolean

Private Direction As Integer

'---------------------------------------------
Private Sub Make_sure_that_Var_are_defined()
'---------------------------------------------
  Dim Old_WB As String
  Old_WB = ActiveWorkbook.Name
  If r1 Is Nothing Then
    Workbooks(WB1_Name).Activate
    Set r1 = ActiveCell
    MaxCol = LastUsedColumn()
    MaxRow = LastUsedRow
    
    Workbooks(WB2_Name).Activate
    Set r2 = ActiveCell
    If LastUsedColumn() > MaxCol Then MaxCol = LastUsedColumn()
    If LastUsedRow > MaxRow Then MaxRow = LastUsedRow
    
    If MaxCol > Last_Col Then MaxCol = Last_Col
    
    Workbooks(Old_WB).Activate
  End If
End Sub

'-----------------------------
Private Sub Next_Cell(ByRef r)
'-----------------------------
  If r.Column < MaxCol Then
       Set r = r.Offset(0, 1)
  Else
       Set r = r.Offset(1, -r.Column + FirstCol)
  End If
  EndReached = r.Row >= MaxRow
End Sub

'-----------------------------
Private Sub Prev_Cell(ByRef r)
'-----------------------------
  If r.Column > FirstCol Then
       Set r = r.Offset(0, -1)
  Else
       Dim RowOffs As Long
       If r.Row > 1 Then RowOffs = -1
       Set r = r.Offset(RowOffs, MaxCol - 1)
       'EndReached = r.Row >= MaxRow And r.Column >= MaxCol
  End If
  EndReached = r.Row = 1 And r.Column = FirstCol
End Sub

'-------------------------------------
Private Sub Show_Act_Cells_in_Dialog()
'-------------------------------------
  Dim p As Long
  p = InStr(r1, "|")
  If p > 0 Then
     Debug.Print "Nach |" & Asc(Mid(r1, p + 1, 1))
  End If
  Cell1Box = "{" & r1 & "}"   ' Add {} to see space characters at the end
  Cell2Box = "{" & r2 & "}"
  
  Diff1Box = ""
  Diff2Box = ""
  Dim i As Long
  For i = 1 To Len(r1) + 1
      If Mid(r1, i, 1) <> Mid(r2, i, 1) Then
         If Len(r1) > i Then
               Diff1Box = "ASC(" & Asc(Mid(r1, i, 1)) & "):"
         Else: Diff1Box = "Len(" & Len(r1) & "):"
         End If
         Diff1Box = Diff1Box & Mid(r1, i)
         
         If Len(r2) > i Then
               Diff2Box = "ASC(" & Asc(Mid(r2, i, 1)) & "):"
         Else: Diff2Box = "Len(" & Len(r2) & "):"
         End If
         Diff2Box = Diff2Box & Mid(r2, i)
         Exit For
      End If
  Next
  If r1 = r2 Then Diff1Box = ">>> Equal <<<"
  
  Addr1 = Replace(r1.Address, "$", "")
  Addr2 = Replace(r2.Address, "$", "")
End Sub


'----------------------------------------------
Private Sub Show_Next_Prev_Diff(Nxt As Boolean)
'----------------------------------------------
  Make_sure_that_Var_are_defined
  Dim ActWb As String
  ActWb = ActiveWorkbook.Name
  If r1 = r2 Then
    Do
      If Nxt Then
           Next_Cell r1
           Next_Cell r2
           Direction = 1
      Else
           Prev_Cell r1
           Prev_Cell r2
           Direction = -1
      End If
    Loop Until r1 <> r2 Or EndReached
  End If
  Show_Act_Cells_in_Dialog
  
  r1.Parent.Parent.Activate ' Switch to the workbook
  r1.Select
  r2.Parent.Parent.Activate ' Switch to the workbook
  r2.Select

  Workbooks(ActWb).Activate
 
End Sub


'------------------------------
Private Sub AbortButton_Click()
'------------------------------
  Me.Hide
End Sub

'--------------------------------
Private Sub Reload_Button_Click()
'--------------------------------
  Set r1 = Nothing
  Set r2 = Nothing
  Make_sure_that_Var_are_defined
  Show_Act_Cells_in_Dialog
End Sub

'----------------------------------------
Private Sub Show_Next_Diff_Button_Click()
'----------------------------------------
  Show_Next_Prev_Diff True
End Sub

'----------------------------------------
Private Sub Show_Prev_Diff_Button_Click()
'----------------------------------------
  Show_Next_Prev_Diff False
End Sub

'-----------------------------------
Private Sub Use_Lower_Button_Click()
'-----------------------------------
' Copy the lower text to the upper
  If Left(r2, 1) = "'" Then
        r1 = "'" & r2
  Else: r1 = r2
  End If
  Show_Next_Prev_Diff Direction >= 0
End Sub



'-----------------------------------
Private Sub Use_Upper_Button_Click()
'-----------------------------------
' Copy the upper text to the lower
  If Left(r1, 1) = "'" Then
        r2 = "'" & r1
  Else: r2 = r1
  End If
  Show_Next_Prev_Diff Direction >= 0

End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
' Center the dialog if it's called the first time.
' On a second call without me.close this function is not called
' => The last position is used
  Center_Form Me
  Make_sure_that_Var_are_defined
  Show_Act_Cells_in_Dialog
End Sub

