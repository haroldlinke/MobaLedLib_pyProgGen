VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectMacrosTreeForm 
   Caption         =   "Auswahl des Makros"
   ClientHeight    =   10236
   ClientLeft      =   45
   ClientTop       =   -945
   ClientWidth     =   18885
   OleObjectBlob   =   "SelectMacrosTreeForm.frx":0000
End
Attribute VB_Name = "SelectMacrosTreeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
' This modul is copied and modified ftom the TreeView example from JKP Application Development Services
'

'Build 026
'***************************************************************************
'
' Authors:  JKP Application Development Services, info@jkp-ads.com, http://www.jkp-ads.com
'           Peter Thornton, pmbthornton@gmail.com
'
' (c)2013-2015, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************

Option Explicit


'''''''''/Treeview Code'''''''''''
Private WithEvents mcTree As clsTreeView
Attribute mcTree.VB_VarHelpID = -1
'''''''''/Treeview Code'''''''''''

'''' for stress testing this demo
Private mlCntChildren As Long

Private ActKey As String

Public AppName As String

Private StartTime_Update_TextBox As Double

Private InitializeTreeFrom_Activ As Boolean

Private ActLanguage As Integer ' 27.11.21:

#Const USE_LNAME = True                ' Use language specific names in the dialog
#Const DELAYED_FILTER_FUNCTION = False ' Delayed filter function: The filter is updated after one second (Not sure it it's better)


#If Mac Then
    Const mcPtPixel As Long = 1
#Else
    Const mcPtPixel As Single = 0.75
#End If



'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
'see the Compile constant DebugMode in tools, VBAProject properties
'DebugMode = 1 will enable the #If to break in Error handlers
    Center_Form Me
    Change_Language_in_Dialog Me
  
    ActKey = ""
    
    ' Hide the Image container
    Me.frmImageBox.Visible = False
    Me.frmImageBox.Enabled = False
    Me.Width = Me.frmImageBox.Left
    
    Please_Wait.Visible = False
    Please_Wait.Top = frTreeControl.Top + (frTreeControl.Height - Please_Wait.Height) / 2
    Please_Wait.Left = frTreeControl.Left + (frTreeControl.Width - Please_Wait.Width) / 2
    
    If Me.frTreeControl.Font.Size < 4 Then
         Me.frTreeControl.Font.Size = 4
    End If
    
    #If DebugMode = 1 Then
        gFormInit = gFormInit + 1
    #End If
    
    #If Mac Then ' Mac is not supported by the Prog_Generator at the moment, but we keep this line in case it will be supported sometimes
        Dim objCtl As MSForms.Control
        
        With Me
            .Font.Size = 10
            .Width = .Width * 4 / 3
            .Height = .Height * 4 / 3
            .BackColor = .labInfo.BackColor
        End With

        For Each objCtl In Me.Controls
            With objCtl
                .Left = Int(.Left * 4 / 3)
                .Top = Int(.Top * 4 / 3)
                .Width = Int(.Width * 4 / 3)
                .Height = Int(.Height * 4 / 3)
                Select Case TypeName(objCtl)
                Case "Image", "SpinButton"
                Case "TextBox", "Frame"
                    .Font.Size = 10
                Case Else
                    .Font.Size = 10
                End Select
            End With
        Next
    #End If
    
    InitializeTreeFrom_Activ = True
    Expert_CheckBox = Get_Bool_Config_Var("Expert_Mode_aktivate")
    InitializeTreeFrom_Activ = False
    
    InitializeTreeView
    'If Not mcTree Is Nothing Then    Disabled because it it is set in the
    '    Me.frTreeControl.setFocus
    'End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Make sure all objects are destroyed
  Abort_Button_Click ' 10.11.21:
  Cancel = 1 ' Don't call the standard close functions what ever they are   ' 10.11.21:
  
  #If False Then     ' 10.11.21:
    If Not mcTree Is Nothing Then
        mcTree.TerminateTree ' *important* to call TerminateTree to clear circular references
        Set mcTree = Nothing
    End If
  #End If
End Sub

Private Sub userform_terminate()
    #If DebugMode = 1 Then
        gFormTerm = gFormTerm + 1
        ClassCounts
    #End If
End Sub

#If False Then
    '---------------------------
    Private Sub InitializeDemo()
    '---------------------------
        Dim cRoot As clsNode, cNode As clsNode
        With mcTree
            ' In the AddChild calls below we pass string Keys to identify the images in the collection, but could use numeric key indexes (1 based)
    
            ' add a Root node with main and expanded icons and make it bold
            Set cRoot = .AddRoot("Root", "Licht", "FolderClosed", "FolderOpen")
            cRoot.Bold = True
            cRoot.ControlTipText = "Makros mit denen Lichteffekte generiert werden"
            
            'Add branches with child nodes to the root:
            'Keys are optional but if using them they must be unique,
            'attempting to add a node with a duplicate key will cause a runtime error.
            '(below we will include unique keys with all the nodes)
            Set cNode = cRoot.AddChild("1", "House", "House")
            cNode.Bold = True
            cNode.ControlTipText = "Mit dieser Funktion wird ein „belebtes“ Haus nachgebildet. In diesem Haus sind zufällig nur einige der Räume beleuchtet."
            
            'Add a 2nd branch to the root and make it bold:
            Set cNode = cRoot.AddChild("2", "Gaslight", "Streetlight")
            cNode.Bold = True
    
    
            ' and add a 4th branch to the root
            Set cNode = cRoot.AddChild("4", "Lichteffekte", "LightBulb")
            cNode.Bold = True
            cNode.AddChild "_4.1", "Ampel" & vbTab & "Erklährung", "Ampel"
            cNode.AddChild "_4.2", "Bluelight" & vbTab & "Erklährung", "BlueLight"
            
    
            
            Set cRoot = .AddRoot("Root2", "Steuerung", "FolderClosed", "FolderOpen")
            cRoot.Bold = True
            cRoot.AddChild "_4.3", "Taster" & vbTab & vbTab & "Erklährung", "Taster"
            'create the node controls and display the tree
            .Refresh
        End With
    
    End Sub
#End If

'----------------------------------------------------------
Private Function Find_Root(ByVal Name As String) As clsNode
'----------------------------------------------------------
  Dim N As Variant, RNodes As Collection
  On Error GoTo ErrProc
  Set RNodes = mcTree.RootNodes
  On Error GoTo 0
  For Each N In RNodes
     If N.Caption = Name Then
        Set Find_Root = N
        Exit Function
     End If
  Next N
ErrProc:
End Function

'-----------------------------------------------------------------------------
Private Function Find_Child(cRoot As clsNode, ByVal Name As String) As clsNode
'-----------------------------------------------------------------------------
  If cRoot.ChildNodes Is Nothing Then Exit Function
  Dim cChild As Variant
  For Each cChild In cRoot.ChildNodes
      If cChild.Caption = Name Then
         Set Find_Child = cChild
         Exit Function
      End If
  Next cChild
End Function


'-------------------------------
Private Sub Add_Node(c As Range)
'-------------------------------
  Dim Row As Long, Name As String, GroupNames As String, Description As String, DescForNode As String, Sh As Worksheet
  Set Sh = c.Parent
  Row = c.Row
#If USE_LNAME Then
  Name = Get_Language_Text(Row, SM_LName_COL, ActLanguage)
  If Name = "" Then Name = c.Value
#Else
  Name = c.Value
#End If
  With Sh
    GroupNames = Get_Language_Text(Row, SM_Group_COL, ActLanguage)  '  .Cells(Row, SM_Group_COL).Value
    Description = Get_Language_Text(Row, SM_ShrtD_COL, ActLanguage) '  .Cells(Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value
    If GroupNames = "" Then GroupNames = "Not grouped"
    Dim GroupNArr() As String, g As Variant, IsRoot As Boolean, Level As Integer
    GroupNArr = Split(GroupNames, "|")
    Dim PicNamesArrInp() As String, PicNamesArr() As String, i As Integer
    PicNamesArrInp = Split(.Cells(Row, SM_Pic_N_COL).Value, "|")
  End With
  ReDim PicNamesArr(UBound(GroupNArr) + 1)
  For i = 0 To UBound(PicNamesArrInp)
      If i > UBound(GroupNArr) + 1 Then Exit For
      PicNamesArr(i) = NoExt(Trim(PicNamesArrInp(i)))
  Next i
  IsRoot = True
  With mcTree
    Dim cNode As clsNode
    For Each g In GroupNArr
        g = Trim(g)
        If Level = UBound(GroupNArr) And Name = "" Then
           DescForNode = Description ' Last node without child => Use the description for the node
        End If
        If IsRoot Then
           Set cNode = Find_Root(g)
           If cNode Is Nothing Then
              If PicNamesArr(Level) <> "" Then ' No Picture given
                    Set cNode = .AddRoot(Row & " Root", g, PicNamesArr(Level), vCaption2:=DescForNode)
              Else: Set cNode = .AddRoot(Row & " Root", g, "FolderClosed", "FolderOpen", vCaption2:=DescForNode)
              End If
           End If
        Else
           Dim cChild As clsNode
           Set cChild = Find_Child(cNode, g)
           If cChild Is Nothing Then
              Set cNode = cNode.AddChild(Row & " " & Level, g, PicNamesArr(Level), vCaption2:=DescForNode)
           Else
              Set cNode = cChild
           End If
        End If
        Level = Level + 1
        cNode.Bold = True
        'cNode.BackColor = rgb(255, 255, 220) ' pale yellow
        If IsRoot Then
              cNode.ForeColor = rgb(0, 128, 0)      ' green
        Else: cNode.ForeColor = rgb(0, 0, 255)      ' blue
        End If
        IsRoot = False
    Next g
    If Name <> "" Then
       cNode.AddChild CStr(Row), Name, PicNamesArr(Level), vCaption2:=Description
    End If
  End With
End Sub


'-----------------------------------------------------------------------------
Private Function InitializeTreeFrom_Lib_Macros_Sheet(Filter As String) As Long
'-----------------------------------------------------------------------------
' Add elements to the Treeview and return the number of the listbox entry which matches with the given string
    ShowHourGlassCursor True                                                ' 29.10.21:
    If Not Me.Visible Then
       StatusMsg_UserForm.ShowDialog "Initialisiere Dialog...", "" ' Can't be used when the dialog is visible
    Else
       Please_Wait.Visible = True ' This message is only shown when the dialog is visible
       DoEvents ' Otherwise the screen is not updated
    End If
    
    If Not mcTree Is Nothing Then
        mcTree.NodesClear      'clears the physical node controls then TerminateTree to clear internal 2-way references.
                               'Some Treeview properties are retained for another session
    End If
    
    
    Dim r As Range, c As Range, ListDataSh As Worksheet, Cnt As Long, Found, Debug_Language As Integer
    Set ListDataSh = ThisWorkbook.Sheets(LIBMACROS_SH)
    Debug_Language = ListDataSh.Range("Test_Language")
    If Debug_Language = -1 Then
          ActLanguage = Get_ExcelLanguage()
    Else: ActLanguage = Debug_Language
          Debug.Print "Using Debug_Language:" & Debug_Language
    End If
    
    'Debug.Print "InitializeTreeFrom_Lib_Macros_Sheet(" & Filter & "," & TextBoxFilter.Text & ")" ' Debug
    InitializeTreeFrom_Lib_Macros_Sheet = -1 ' Nothing found
    
    With ListDataSh
       Set r = .Range(.Cells(SM_DIALOGDATA_ROW1, SM_Name__COL), .Cells(LastUsedRowIn(ListDataSh), SM_Name__COL))
       On Error GoTo ErrMsg
            For Each c In r
              'If c.Row = 80 Then
              '   Debug.Print "Row " & c.Row
              'End If
              If (c.Value <> "" Or .Cells(c.Row, SM_ShrtD_COL) <> "" Or .Cells(c.Row, SM_DetailCOL) <> "") And (.Cells(c.Row, SM_Mode__COL) = "" Or Expert_CheckBox) Then
                 Found = False
                 If Trim(TextBoxFilter.Text) = "" Then ' Use seperate if/elseif statements to speed up the process because excel evalualtes al parts in an or expression
                     Found = True
                 ElseIf InStr(1, c.Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then
                     Found = True
                 ElseIf InStr(1, .Cells(c.Row, SM_Group_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then
                     Found = True
                 ElseIf InStr(1, .Cells(c.Row, SM_LName_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then ' 25.10.21:
                     Found = True
                 ElseIf InStr(1, .Cells(c.Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then
                     Found = True
                 End If
                 If Found Then
                    Add_Node c
                    Cnt = Cnt + 1
                 End If
              End If
            Next c
    End With
    If Cnt > 0 Then
        mcTree.Refresh
        If Trim(TextBoxFilter.Text) = "" Then ' No Filter activ
            mcTree.ExpandToLevel 0                ' close nodes
            mcTree.RootNodes(1).Expanded = True   ' Open the first root node
        End If
        mcTree.Refresh
    End If
    Please_Wait.Visible = False                                             ' 29.10.21:
    If Not Me.Visible Then Unload StatusMsg_UserForm
    ShowHourGlassCursor False
    Exit Function
    
ErrMsg:
    Please_Wait.Visible = False                                             ' 29.10.21:
    If Not Me.Visible Then Unload StatusMsg_UserForm
    ShowHourGlassCursor False
    
    Dim Row As Long
    If Not c Is Nothing Then Row = c.Row
    MsgBox "Error adding line " & Row & " to the TreeView Dialog" & vbCr & _
           vbCr & _
           Err.Number & ": " & Err.Description & vbCr & _
           "Prog. Line: " & Erl, _
           vbCritical, "Error creating the treeView Dialog"
    ListDataSh.Activate
    ListDataSh.Cells(Row, SM_ShrtD_COL).Select
           
    EndProg
    Resume Next ' Could be used to find the line
End Function

'-------------------------------------------------------------------------
Private Sub InitializeTreeView()
'-------------------------------------------------------------------------
' Procedure : Initialize
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Initializes the userform,
'             adds the VBA treeview to the container frame on the userform
'             and populates the treeview.
'-------------------------------------------------------------------------

    Dim cRoot As clsNode
    Dim cNode As clsNode
    Dim cExtraNode As clsNode
    Dim i As Long
    
    Set mcTree = New clsTreeView

    On Error GoTo errH

    With mcTree
        Set .TreeControl = Me.frTreeControl

        .AppName = Me.AppName

        'Add some tree properties:
        .CheckBoxes = False
        .RootButton = True
        .EnableLabelEdit(bAutoSort:=False, bMultiLine:=False) = False
        .FullWidth = True  ' Show bar with the maximal width when selected. Otherwise the length of the bar has the length of the text
        .Indentation = 12
        .NodeHeight = 12
        .ShowLines = True
        .ShowExpanders = True
        
        Set .Images = Me.frmImageBox
    End With
    
    Display_Navigation_Keys
    
    InitializeTreeFrom_Lib_Macros_Sheet TextBoxFilter.Value

    Exit Sub

errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
    If Not mcTree Is Nothing Then
        MsgBox Err.Description, , AppName
        mcTree.NodesClear
        mcTree.TerminateTree
    End If

End Sub

Private Function GetIcons(colImages As Collection, Optional ImageNames) As Long
'-------------------------------------------------------------------------
' Procedure : GetIcons
' Company   : -
' Author    : Peter Thornton
' Created   : 28-01-2013
' Purpose   : Creates a collection of StdPicture objects from images in the frame frmImageBox
'-------------------------------------------------------------------------
    Dim v
    Dim img As MSForms.Image

    Set colImages = New Collection

    If IsMissing(ImageNames) Then
        ' get all available images
        For Each img In Me.frmImageBox.Controls
            colImages.Add img.Picture, img.Name
        Next
    Else
        ' only get specified images
        For Each v In ImageNames
            colImages.Add Me.frmImageBox.Controls(v).Picture, v
        Next
    End If

    GetIcons = colImages.Count
End Function

'''''''''' /Treeview Demo ''''''''''


''''''''''' Treeview container frame events ''''''''''

' Enter/Exit events are not trapped with 'WithEvents' in the treeclass
' so if needed they can be trapped in the form. (These will toggle the active node's highlight states)

Private Sub frTreeControl_Enter()
    If Not mcTree Is Nothing Then
        mcTree.EnterExit False
    End If
End Sub

Private Sub frTreeControl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mcTree Is Nothing Then
        mcTree.EnterExit True
    End If
End Sub

'''''''''' /Treeview container frame events ''''''''''




''''''''''' Treeview Events, raised in clsTreeView ''''''''''

'-----------------------------------------
Private Sub Set_Description(Txt As String)                                  ' 04.11.21:
'-----------------------------------------
  Dim ActFocus As Variant
  On Error Resume Next
  Set ActFocus = Me.ActiveControl
  On Error GoTo 0
  Me.Description = Txt
  With Description ' Display the scroll bars if necessary
    .setFocus
    .SelStart = 0
  End With
  If Not ActFocus Is Nothing Then ActFocus.setFocus
End Sub


'------------------------------------
Private Sub Display_Navigation_Keys()
'------------------------------------
  Set_Description Get_Language_Str("Tasten:" & vbCr & _
                                   "Pfeiltasten Auf/Ab, Seite Auf/Ab, Home/Ende zum Navigieren." & vbCr & _
                                   "Pfeiltasten links/rechts zum Ein-/Ausklappen." & vbCr & _
                                   "Zifferntasten 0 bis 9 zum Auf- und Zuklappen auf eine Ebene." & vbCr & _
                                   "Buchstaben zum aktivieren des Filters." & vbCr & _
                                   vbCr & _
                                   "Treeview by JKP Application Development Services, info@jkp-ads.com, http://www.jkp-ads.com")

End Sub

'This gets fired when a node is clicked
'-----------------------------------------
Private Sub mcTree_Click(cNode As clsNode)
'-----------------------------------------
    'Debug.Print "mcTree_Click " & cNode.Key
    
    With cNode
       Dim Sh As Worksheet, Desc As String, Row As Long
       Set Sh = ThisWorkbook.Sheets(LIBMACROS_SH)
       ActKey = .key
       Row = val(Split(.key, " ")(0))
       Desc = Replace(Sh.Cells(Row, SM_DetailCOL + ActLanguage * DeltaCol_Lib_Macro_Lang), "|", vbLf)
       If IsNumeric(.key) Then ' Entry without childs have a numeric key. All others have some additional texts
             If Desc = "" Then Desc = Replace(Sh.Cells(Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang), "|", vbLf)
             Set_Description Desc
             Me.Detail = Replace_Multi_Space(Sh.Cells(Row, SM_Macro_COL))
       Else:
             If Desc <> "" Then
                   Set_Description Desc
             Else: Display_Navigation_Keys
             End If
             Me.Detail = Replace(Sh.Cells(Row, SM_Group_COL).Value, "|", " > ")
       End If
    End With
End Sub

'------------------------------------------------------------------------------------------------------------------------------
Private Sub mcTree_MouseAction(cNode As clsNode, Action As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) ' Hardi
'------------------------------------------------------------------------------------------------------------------------------
' Activated by a double click
' See: NodeEventRouter
  'Debug.Print "mcTree_DblClick:" & cNode.Key & " Action:" & Action
  
  'If IsNumeric(cNode.key) Then Select_Button_Click
  If IsNumeric(cNode.key) And InStr(cNode.key, " ") = 0 Then Select_Button_Click           ' 15.1.21 Juergen workaround for Excel 2007 isNumericBug
End Sub

'------------------------------------------------------------------------------------------------------------------
Private Function Send_Letters_to_TextBoxFilter(KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) As Boolean
'------------------------------------------------------------------------------------------------------------------
    Select Case KeyCode
        Case Asc("A") To Asc("Z"):
             If Shift < 2 Then
                TextBoxFilter.setFocus
                If Shift = 0 Then KeyCode = KeyCode + 32 ' Convert to lower case
                TextBoxFilter = TextBoxFilter & Chr(KeyCode)
             ElseIf Shift = 4 And KeyCode = Asc("F") Then ' Alt+F => Filter (Attention: The accelerator character can't be displayed in the text because then the excel Toolbox dialog is opened for some reasons)
                TextBoxFilter.setFocus
             End If
             Send_Letters_to_TextBoxFilter = True
       'Case Else: Debug.Print "KeyCode: " & KeyCode.Value & " chr: " & Chr(KeyCode.Value) & " Shift:" & Shift
    End Select
End Function

'-----------------------------------------------------------------------------------------------------------
Private Sub mcTree_KeyDown(cNode As clsNode, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'-----------------------------------------------------------------------------------------------------------
'This gets fired when a key is pressed down
    Dim bMove As Boolean
    Dim sMsg As String
    Dim cSource As clsNode
    
    If Send_Letters_to_TextBoxFilter(KeyCode, Shift) Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, _
             48 To 57, 96 To 105, vbKeyF2, 20, 93, _
             vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
            ' these keys are already trapped in clsTreeView for navigation, expand/collapse, edit mode
#If 0 Then
        Case Asc("A") To Asc("Z"):
             If Shift < 2 Then
                TextBoxFilter.setFocus
                If Shift = 0 Then KeyCode = KeyCode + 32 ' Convert to lower case
                TextBoxFilter = TextBoxFilter & Chr(KeyCode)
             ElseIf Shift = 4 And KeyCode = Asc("F") Then ' Alt+F => Filter (Attention: The accelerator character can't be displayed in the text because then the excel Toolbox dialog is opened for some reasons)
                TextBoxFilter.setFocus
             End If
#End If
        Case vbKeyReturn: Select_Button_Click
       'Case Else: Debug.Print "KeyCode: " & KeyCode.Value & " chr: " & Chr(KeyCode.Value) & " Shift:" & Shift
    End Select
End Sub

' Keyboard event functions to send key press to the TextBoxFilter
Private Sub Expert_CheckBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):  Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub

Private Sub Abort_Button_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):     Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub

Private Sub Select_Button_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):    Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub

'Private Sub LabelFilter_Click()    ' Is not working
'  TextBoxFilter.setFocus
'End Sub



'''''''''''/ Treeview Events, raised in clsTreeView ''''''''''

'---------------------------------
Private Sub Calc_SelectMacro_Res()
'---------------------------------
' Return the name and the row number in the ListDataSheet
  Dim Res As String, Row As Long
  If ActKey <> "" And IsNumeric(ActKey) Then
    Row = val(ActKey)
    With ThisWorkbook.Sheets(LIBMACROS_SH)
         Res = .Cells(Row, SM_Name__COL) & "," & Row  ' Return "Name,Nr"  Nr: row nr in the data sheet
         Last_SelectedNr_Valid = True
         Last_SelectedNr = Row
    End With
        SelectMacro_Res = Res
  Else: SelectMacro_Res = ""
  End If
End Sub


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  UnhookFormScroll ' Deactivate the mouse wheel scroll function
  Me.Hide  ' no "Unload Me" to keep the entered data and dialog position
  SelectMacro_Res = ""
  'Enable_Listbox_Changed = False
End Sub

'--------------------------------
Private Sub Select_Button_Click()
'--------------------------------
  UnhookFormScroll ' Deactivate the mouse weel scrol function
  Calc_SelectMacro_Res
  If SelectMacro_Res = "" Then
        'Show_Status_for_a_while Get_Language_Str "Please selec
        BeepThis2 "Windows Balloon.wav"
  Else: Me.Hide
        Debug.Print "Result: " & SelectMacro_Res
  End If
End Sub


'----------------------------------
Private Sub Expert_CheckBox_Click()
'----------------------------------
  If Not InitializeTreeFrom_Activ Then                                    ' 12.04.20:
     InitializeTreeFrom_Lib_Macros_Sheet TextBoxFilter.Value
     Set_String_Config_Var "Expert_Mode_aktivate", IIf(Expert_CheckBox, "1", "0")
  End If
End Sub


'---------------------------------
Public Sub Update_TextBoxFilter()
'---------------------------------
   If Not InitializeTreeFrom_Activ Then
      InitializeTreeFrom_Lib_Macros_Sheet TextBoxFilter.Value
   End If

End Sub

'---------------------------------
Private Sub TextBoxFilter_Change()
'---------------------------------
   If Not InitializeTreeFrom_Activ Then
      #If DELAYED_FILTER_FUNCTION Then ' Delayed filter function
        If StartTime_Update_TextBox <> 0 Then
           On Error Resume Next
           Application.OnTime StartTime_Update_TextBox, "Update_TextBoxFilter_onTime", Schedule:=False
        End If
        StartTime_Update_TextBox = Now + TimeValue("00:00:01")
        Application.OnTime StartTime_Update_TextBox, "Update_TextBoxFilter_onTime"
      #Else
        DoEvents ' To update the input box text
        InitializeTreeFrom_Lib_Macros_Sheet TextBoxFilter.Value
      #End If
   End If
End Sub



'-----------------------------------------------
Public Sub MouseWheel(ByVal lngRotation As Long)
'-----------------------------------------------
' Process the mouse wheel changes
  'Debug.Print "MouseWheel " & lngRotation & " " & Me.frTreeControl.ScrollTop
  
  If Description Is Me.ActiveControl Then                                   ' 05.11.21:
    Dim ActiveNode As clsNode ' Scroll the description text box
    Set ActiveNode = mcTree.ActiveNode
    With Description
      .setFocus
      If lngRotation < 0 Then
            Application.SendKeys "{HOME}{DOWN}{DOWN}{DOWN}" ' {HOME} to move the cursor to the start of the line
      Else
           If .SelStart > 0 Then Application.SendKeys "{HOME}{UP}{UP}{UP}"
      End If
      
      ' Problem beim Scrollen in der Description box:
      ' Wenn der Cursor ganz oben ist, dann verliert die Box den Focus und die Simulierten Tasten werden von dem TreeView
      ' verarbeitet. Das führt dazu, dass ein anderer Eintrag ausgewählt wird.
      ' Das wird mit dem folgenden Block abgefangen
      Dim i As Long
      For i = 1 To 7
         DoEvents ' Process the simulated keys
         Sleep 10
      Next i
      If Not Description Is Me.ActiveControl Then ' Check if we have lost the focus by the {up} key
         'mcTree.ScrollToView ActiveNode, 0
         mcTree.NodeEventRouter ActiveNode, "Caption", 1
         .setFocus
      End If
    End With
  
  Else ' Scroll the treeView
    Const Step = 36 ' One line is 12
    With Me.frTreeControl
      If lngRotation > 0 Then
          If .ScrollTop > 0 Then
              If .ScrollTop > Step Then
                  .ScrollTop = .ScrollTop - Step
              Else
                  .ScrollTop = 0
              End If
          End If
      Else
          .ScrollTop = .ScrollTop + Step
      End If
    End With
  End If
End Sub

'---------------------------------------------------------
Private Function Find_Node(ByVal key As String) As clsNode
'---------------------------------------------------------
  On Error GoTo ErrProc
  Set Find_Node = mcTree.Nodes(key)  ' So geht es auch
  Exit Function
ErrProc:
  Debug.Print "Not Found " & key & " ;-("
End Function


'----------------------------------------------------------------
Public Sub Show_SelectMacros_TreeView(ByVal SelectName As String)
'----------------------------------------------------------------
' Use this function to show the dialog.
        
  ' Selct the node by Selectname
  Dim ActiveNode As clsNode, Row As Long
  If SelectName <> "" Then
     Row = Find_Macro_in_Lib_Macros_Sheet(SelectName)
     If Row > 0 Then
        Set ActiveNode = Find_Node(Row)
        If ActiveNode Is Nothing And TextBoxFilter <> "" Then ' The macro is not fond ;-( Maybe because it's filtered
           TextBoxFilter = ""  ' This will automatically update the dialog because TextBoxFilter_Change is triggered
           Set ActiveNode = Find_Node(Row)
        End If
        If ActiveNode Is Nothing And Expert_CheckBox = False Then ' Still not found => Enable the expret mode
           InitializeTreeFrom_Activ = True ' Don't update the Tree View because this is done below
           TextBoxFilter = ""
           InitializeTreeFrom_Activ = False
           Expert_CheckBox = True ' This will automatically update the dialog because Expert_CheckBox_Click is triggered
           Set ActiveNode = Find_Node(Row)
        End If
     End If
  End If
  
  ' Selct the node Last_SelectedNr
  If ActiveNode Is Nothing And Last_SelectedNr_Valid And Last_SelectedNr > 0 Then
    Row = Last_SelectedNr
    Set ActiveNode = Find_Node(Row)
  End If
  If Not ActiveNode Is Nothing Then
     mcTree.NodeEventRouter ActiveNode, "Caption", 1
     mcTree.ScrollToView ActiveNode, 0
  End If
  
  HookFormScroll Me, "frTreeControl"   ' Initialize the mouse wheel scroll function
  
  frTreeControl.setFocus
  
  Me.Show
End Sub

'-----------------------------------------------------------
Public Function Get_Icon(vKey, Pic As StdPicture) As Boolean
'-----------------------------------------------------------
  Dim bFullWidth As Boolean
  Get_Icon = mcTree.GetNodeIcon(vKey, Pic, bFullWidth)
End Function
