Attribute VB_Name = "M99_Test"
Option Explicit

'-------------------------------------
Private Sub Test_Normal_Color_Dialog()
'-------------------------------------
  Dim Res As Variant
  Res = Application.Dialogs(xlDialogEditColor).Show(1, 26, 82, 48)
End Sub

'------------------------
Private Sub ColorDialog()
'------------------------
'Create variables for the color codes
Dim FullColorCode As Long
Dim RGBRed As Integer
Dim RGBGreen As Integer
Dim RGBBlue As Integer

'Get the color code from the cell named "RGBColor"
FullColorCode = Range("A3").Interior.Color

'Get the RGB value for each color (possible values 0 - 255)
RGBRed = FullColorCode Mod 256
RGBGreen = (FullColorCode \ 256) Mod 256
RGBBlue = FullColorCode \ 65536

'Open the ColorPicker dialog box, applying the RGB color as the default
If Application.Dialogs(xlDialogEditColor).Show _
    (1, RGBRed, RGBGreen, RGBBlue) = True Then

    'Set the variable RGBColorCode equal to the value
    'selected the DialogBox
    FullColorCode = ActiveWorkbook.Colors(1)
    
    'Set the color of the cell named "RGBColor"
    Range("A2").Interior.Color = FullColorCode

Else
   
    'Do nothing if the user selected cancel

End If

End Sub

'-----------------------
Private Sub Get_AllPar()
'-----------------------
' Generates a list of all MobaledLib Macro parameters and how often the parameter is used
  Dim c As Variant, Res As String, f As Variant, Sh As Worksheet
  Res = " "
  Set Sh = Sheets("Tabelle3")
  For Each c In ActiveSheet.UsedRange.Cells
      If c <> "" Then
      Set f = Sh.Cells.Find(What:=c, after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
      Dim r As Long
      With Sh
        If f Is Nothing Then
           r = LastUsedRowIn(Sh) + 1
           .Cells(r, 1) = c
        Else: r = f.Row
        End If
        .Cells(r, 2) = val(.Cells(r, 2)) + 1
      End With
      'If InStr(Res, " " & c & " ") = 0 Then Res = Res & c & " "
      End If
  Next c
  'Debug.Print Res
  
End Sub

'-----------------------------------
Private Sub Buttons_with_vblf_Text()
'-----------------------------------
  ActiveSheet.Shapes("Insert_Button").DrawingObject.Object.Caption = "Zeile" & vbLf & "einfügen"
  ActiveSheet.Shapes("Del_Button").DrawingObject.Object.Caption = "Lösche" & vbLf & "Zeilen"
  ActiveSheet.Shapes("Move_Button").DrawingObject.Object.Caption = "Verschiebe" & vbLf & "Zeilen"
  ActiveSheet.Shapes("ClearSheet_Button").DrawingObject.Object.Caption = "Lösche" & vbLf & "Tabelle"
End Sub

'-----------------------------------
Private Sub TestSingleLineInput()
'-----------------------------------
    Dim frm As New UserForm_SingleInput
    frm.ShowForm "Cap", "Lab", "DefaultText"

End Sub

Private Sub Add_Translation_from_Theo()
  Const DstCol = 5
  Dim src As Worksheet, Row As Long, DstRow As Long
  Dim Dst As Worksheet
  
  Set src = Workbooks("Vertaling_Duits_Nederlands-1.xlsx").Sheets("Blad1")
  Set Dst = ThisWorkbook.Sheets("Languages")
  With src
    For Row = 3 To LastUsedRowIn(src)
        DstRow = .Cells(Row, 1)
        Dst.Activate
        Dst.Cells(DstRow, DstCol).Select
        If Dst.Cells(DstRow, DstCol) <> .Cells(Row, 2) Then
           DstRow = DstRow + 2
           If Dst.Cells(DstRow, DstCol) <> .Cells(Row, 2) Then
              Debug.Print "Ungleich:" & vbCr & _
                          "  " & Dst.Cells(DstRow, DstCol) & vbCr & _
                          "  " & .Cells(Row, 2)
           End If
        End If
        Dst.Cells(DstRow, DstCol) = .Cells(Row, 3)
    Next Row
  End With
End Sub


