Attribute VB_Name = "M10_Par_Description"
Option Explicit


Private Const ParName_COL = 1
Private Const Par_Cnt_COL = 2
Private Const ParType_COL = 3
Private Const Par_Min_COL = 4
Private Const Par_Max_COL = 5
Private Const Par_Def_COL = 6
Private Const Par_Opt_COL = 7
Public Const ParInTx_COL = 8
Public Const ParHint_COL = 9

Public Const CHAN_TYPE_NONE = 1
Public Const CHAN_TYPE_LED = 2
Public Const CHAN_TYPE_SERIAL = 3

Private Const FirstDatRow = 2

'------------------------------------------------------------------------
Private Function Get_ParDesc_Row(Sh As Worksheet, Name As String) As Long
'------------------------------------------------------------------------
  Dim r As Range, f As Variant
  With Sh
    Set r = .Range(.Cells(1, ParName_COL), .Cells(LastUsedRowIn(Sh), ParName_COL))
  End With
    
  Set f = r.Find(What:=Name, after:=r.Cells(FirstDatRow, 1), LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

  If f Is Nothing Then
       Debug.Print "Fehlender Parameter: " & Name
       MsgBox "Fehler: Der Parameter Name '" & Name & "' wurde nicht im Sheet '" & Sh.Name & "' gefunden!", _
              vbCritical, "Internal Error"
       EndProg
  Else
       Get_ParDesc_Row = f.Row
  End If
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Get_Par_Data(ByVal ParName As String, ByRef Typ As String, ByRef Min As String, ByRef Max As String, ByRef Def As String, ByRef Opt As String, ByRef InpTxt As String, ByRef Hint As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Const DeltaCol = 2
  Dim Row As Long, Sh As Worksheet, ActLanguage As Integer, Offs As Long
  ActLanguage = Get_ExcelLanguage()                                         ' 24.02.20:
  If ActLanguage <> 0 Then Offs = 1
  
  Set Sh = Sheets(PAR_DESCR_SH)
  Row = Get_ParDesc_Row(Sh, ParName)
  With Sh
    Typ = .Cells(Row, ParType_COL)
    Min = .Cells(Row, Par_Min_COL)
    Max = .Cells(Row, Par_Max_COL)
    Def = .Cells(Row, Par_Def_COL)
    Opt = .Cells(Row, Par_Opt_COL)
    InpTxt = .Cells(Row, ParInTx_COL + ActLanguage * DeltaCol + Offs)   ' 24.02.20: Added: + ActLanguage * DeltaCol  + Offs
    If InpTxt = "" Then InpTxt = ParName
    Hint = .Cells(Row, ParHint_COL + ActLanguage * DeltaCol + Offs)   ' Inserting a LF seames to be not possible ;-(   Test with: Replace(.Cells(Row, ParHint_COL), "|", vbCrLf)
  End With
End Sub


'UT----------------------------
Private Sub Test_Get_Par_Data()
'UT----------------------------
  Dim Typ As String, Min As String, Max As String, Def As String, Opt As String, InpTxt As String, Hint As String
  Get_Par_Data "Pin_List", Typ, Min, Max, Def, Opt, InpTxt, Hint
  Debug.Print "Typ:" & Typ, "Min:" & Min & " Max:" & Max & " Def:" & Def & " Opt:" & Opt & vbCr & "InpTxt:" & InpTxt & vbCr & "Hint:" & Hint
End Sub
