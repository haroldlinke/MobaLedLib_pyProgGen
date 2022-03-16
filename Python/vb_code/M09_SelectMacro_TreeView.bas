Attribute VB_Name = "M09_SelectMacro_TreeView"
' ********************************************************************************
' Select macro using the treeView Dialog from JKP Application Development Services
'
' This Modul is based on the demo from JKP Application Development Services
'

'Build 026.4
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

' Revision History:
' - Filter function
' - Added description to the intermidiate nodes
'   - Separate lines must be zsed for the intermidiate nodes
'     - Icons must be added also to this lines
'   - Short discription is shown in the treeView
'   - Detailed discription is shown in the window below
'     If not available the Keys help is shown
'   - In filter mode the descriptions are not shown because the separate lines are not available
' - Generate result
' - Scroll wheel support
' > Send to Michael
' - DoubleCkick opens/closes child
' - Horizontal scroll bar corrected
' - TreeView settings are keeped when the dialog is closed to get the same view when it's opened again
' - Show HourGlassCursor when building/updating the TreeView
' - New dialog is used in the DCC, ... sheets

Option Explicit

Private ufForm As SelectMacrosTreeForm

' To switch to debug mode and break in error handlers:
'   In the VBAProject properties, right-click Treeview_26
'   Set the Conditional Compilation Argument -
'       DebugMode = 1
'   See: https://bettersolutions.com/vba/debugging/conditional-compiling.htm

#If DebugMode = 1 Then
    ' PT counters for class Init & Term events
    Public gClsTreeViewInit As Long
    Public gClsTreeViewTerm As Long
    Public gClsNodeInit As Long
    Public gClsNodeTerm As Long
    Public gFormInit As Long
    Public gFormTerm As Long
#End If
Public gbCellsChanged As Boolean

#If 0 Then
    ' Geht nicht so richtig
    ' Es werden zwar irgend welche (zufälligen) Pixel Muster eingebunden, aber nicht das Bild
    ' https://www.pcreview.co.uk/threads/copy-shape-image-into-image-control.982606/
    
    ' 21.10.21: Test ToDo Adapt oo 64 Bit
    Private Declare Function OpenClipboard& Lib "user32" (ByVal hWnd&)
    Private Declare Function EmptyClipboard& Lib "user32" ()
    Private Declare Function GetClipboardData& Lib "user32" (ByVal wFormat%)
    Private Declare Function SetClipboardData& Lib "user32" (ByVal wFormat&, ByVal hMem&)
    Private Declare Function CloseClipboard& Lib "user32" ()
    Private Declare Function CopyImage& Lib "user32" (ByVal handle&, ByVal un1&, ByVal n1&, ByVal n2&, ByVal un2&)
    Private Declare Function DestroyIcon& Lib "user32" (ByVal hIcon&)

    '----------------------
    Private Sub Test_Icon()
    '----------------------
      Dim Pic As StdPicture, Res As Boolean
      If ufForm Is Nothing Then
         Set ufForm = New SelectMacrosTreeForm
      End If
      Res = ufForm.Get_Icon("Andreaskreuz", Pic)
      If Not Res Then
         MsgBox "Error reading the picture from the UserForm"
         Exit Sub
      End If
      
      Dim iPic As StdPicture, hCopy&
      'Set iPic = Me.Picture
      Set iPic = Pic ' Me.Image1.Picture
      OpenClipboard 0&: EmptyClipboard
      hCopy = SetClipboardData(2, iPic.handle)
      CloseClipboard
      If hCopy Then
        ActiveSheet.Cells(15, 5).Select
        ActiveSheet.Paste
        ' Save picture as file (Metafichier -> c:\xxx.wmf)
        'SavePicture iPic, "c:\xxx.bmp"
      End If
      'DestroyIcon iPic.handle
      Set iPic = Nothing
      
      
    '  Cells(26,1) pic.
    '  pic.CopyPicture xlScreen, xlBitmap
      'Range("J10").Select
      'ActiveSheet.Image1.Picture
      'ActiveSheet.Pictures.Insert pic
    End Sub
#End If



'UT-----------------------------------
Public Sub SelectMacro_TreeView_Test()
'UT-----------------------------------
  SelectMacro_TreeView "" ' "RGB_Heartbeat("
  End ' To reload the dialog the next time
End Sub


'----------------------------------------------------------
Public Sub SelectMacro_TreeView(ByVal SelectName As String)
'----------------------------------------------------------
    If ufForm Is Nothing Then
       Set ufForm = New SelectMacrosTreeForm
    End If
    If SelectName <> "" Then
       If InStr(SelectName, "(") > 0 Then
          SelectName = Split(SelectName, "(")(0) & "("
       End If
    End If
    With ufForm
        .Show_SelectMacros_TreeView SelectName
    End With
End Sub

#If DebugMode = 1 Then
    Sub ClassCounts()
    ' PT, If making any code modifications it's important to ensure all class instances have been properly terminated.
    '     Classes are counted when created and when destroyed in respective Initialize & Terminate events,
    '     when done the totals should be the same.

        If gClsTreeViewInit <> gClsTreeViewTerm Or _
           gClsNodeInit <> gClsNodeTerm Or _
           gFormInit <> gFormTerm Then
            Debug.Print "clsTreeView", gClsTreeViewInit, gClsTreeViewTerm, gClsTreeViewInit - gClsTreeViewTerm
            Debug.Print "clsNode", gClsNodeInit, gClsNodeTerm, gClsNodeInit - gClsNodeTerm
            Debug.Print "gFormInit", gFormInit, gFormTerm, gFormInit - gFormTerm
            MsgBox "NOT all Classes were terminated !" & vbCr & "see Immediate window"
        End If
        gClsTreeViewInit = 0
        gClsTreeViewTerm = 0
        gClsNodeInit = 0
        gClsNodeTerm = 0
        gFormInit = 0
        gFormTerm = 0
    End Sub
#End If

'----------------------------------------
Private Sub Update_TextBoxFilter_onTime()
'----------------------------------------
  If Not ufForm Is Nothing Then
     ufForm.Update_TextBoxFilter
  End If
End Sub


'---------------------------------------------------------
Private Sub Helper_Replace_Empty_Lines_in_Front_of_Cells() ' 27.11.21:
'---------------------------------------------------------
' Some entries start with an additional empty line. I don't want to change them manualy...
  Sheets(LIBMACROS_SH).Select
  Dim c As Variant, cnt As Long
  For Each c In Range(Cells(4, 1), Cells(LastUsedRow, LastUsedColumn()))
      If c <> "" Then
         While Len(c) > 1 And left(LTrim(c), 1) = vbLf
            c.Value = Mid(LTrim(c), 2)
            cnt = cnt + 1
         Wend
      End If
  Next c
  Debug.Print "Replaced " & cnt & " entries"
End Sub

'----------------------------------------------------------------
Public Sub Set_Lib_Macros_Test_Language(Test_Language As Integer) ' 28.11.21:
'----------------------------------------------------------------
    Dim ListDataSh As Worksheet
    Set ListDataSh = ThisWorkbook.Sheets(LIBMACROS_SH)
    ListDataSh.Range("Test_Language") = Test_Language
End Sub


'UT--------------------------------------------
Private Sub Test_Set_Lib_Macros_Test_Language() ' 28.11.21:
'UT--------------------------------------------
 Set_Lib_Macros_Test_Language -1
End Sub

