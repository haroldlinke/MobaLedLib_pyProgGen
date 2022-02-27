VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 03.04.21: Juergen in case the LEDNr_Display_Type is changed we need to modify the headers and contents
'--------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
'--------------------------------------------------
    If Not Is_Named_Range(Target) Then Exit Sub
    If Target.Name.Name = "LEDNr_Display_Type" Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Dim curSheet As Worksheet, s As Worksheet
        Set curSheet = Application.ActiveSheet
        For Each s In Worksheets ' Find the last data sheet
            If Is_Data_Sheet(s) Then
                s.Select
                Make_sure_that_Col_Variables_match
                Update_Start_LedNr
            End If
        Next s
        curSheet.Select
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    ElseIf Target.Name.Name = "Show_Icon_Column" Then
        'Application.EnableEvents = False
        Show_Hide_Column_in_all_Sheets Get_Bool_Config_Var(Target.Name.Name), "MacIcon_Col"
        Make_Sure_That_One_Macro_Column_is_Visible "Show_Simple_Names", "LanName_Col"
        'Application.EnableEvents = True
    ElseIf Target.Name.Name = "Show_Simple_Names" Then
        'Application.EnableEvents = False
        Show_Hide_Column_in_all_Sheets Get_Bool_Config_Var(Target.Name.Name), "LanName_Col"
        Make_Sure_That_One_Macro_Column_is_Visible "Show_Icon_Column", "MacIcon_Col"
        'Application.EnableEvents = True
    ElseIf Target.Name.Name = "Show_Macros_Column" Then
        'Application.EnableEvents = False
        Show_Hide_Column_in_all_Sheets Get_Bool_Config_Var(Target.Name.Name), "Config__Col"
        Make_Sure_That_One_Macro_Column_is_Visible "Show_Simple_Names", "LanName_Col"
        'Application.EnableEvents = True
   End If
End Sub

'------------------------------------------------------------------------------------------------------
Private Sub Make_Sure_That_One_Macro_Column_is_Visible(Config_Var As String, Prefered_Column As String)
'------------------------------------------------------------------------------------------------------
  If Get_Bool_Config_Var("Show_Icon_Column") = False And _
     Get_Bool_Config_Var("Show_Simple_Names") = False And _
     Get_Bool_Config_Var("Show_Macros_Column") = False Then
       Set_String_Config_Var Config_Var, "1"
       Show_Hide_Column_in_all_Sheets 1, Prefered_Column
  End If
End Sub
