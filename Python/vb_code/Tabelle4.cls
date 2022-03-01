VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

#If False Then
'--------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
'--------------------------------------------------------
' This function is called if the worksheet is changed.
  Debug.Print "Languages changed"
End Sub
#End If
