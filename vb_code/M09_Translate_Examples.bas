Attribute VB_Name = "M09_Translate_Examples"
Option Explicit
' Translate the example strings


'UT------------------------------------------------
Private Sub Test_Translate_Example_Texts_in_Sheet()
'UT------------------------------------------------
  Translate_Example_Texts_in_Sheet ActiveSheet
End Sub

'-----------------------------------------------------------------
Public Sub Translate_Example_Texts_in_Sheet(ByVal Sh As Worksheet)
'-----------------------------------------------------------------
  Dim r As Long
  Dim Transl As String
  
  Make_sure_that_Col_Variables_match Sh
  
  ' Description
  For r = FirstDat_Row To LastUsedRow
      With Cells(r, Descrip_Col)
          If .Value <> "" Then
             Transl = Get_Language_Str(.Value)
             If .Value <> Transl Then .Value = Transl
          End If
      End With
  Next r

  ' Typ column (Red/Green/...)                                              ' 07.03.20:
  For r = FirstDat_Row To LastUsedRow
      With Cells(r, Inp_Typ_Col)
          If .Value <> "" Then
             Transl = Get_Language_Str(.Value)
             If .Value <> Transl Then .Value = Transl
          End If
      End With
  Next r
End Sub


