VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'-----------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'-----------------------------------------------------------------
' Is called by event if the worksheet selection has changed
  If Is_Libraries_Select_Column(Target) Then
     If Target.Value = "" Then
           Target.Value = ChrW(Hook_CHAR)
     Else: Target.Value = ""
     End If
     Range(ActiveCell.Address).Offset(0, 1).Activate
     BeepThis2 "Windows Balloon.wav"
  End If
End Sub

