Attribute VB_Name = "M26_Add_Icons_to_TreeView"
''Build 026
''***************************************************************************
''
'' Authors:  JKP Application Development Services, info@jkp-ads.com, http://www.jkp-ads.com
''           Peter Thornton, pmbthornton@gmail.com
''
'' (c)2013-2015, all rights reserved to the authors
''
'' You are free to use and adapt the code in these modules for
'' your own purposes and to distribute as part of your overall project.
'' However all headers and copyright notices should remain intact
''
'' You may not publish the code in these modules, for example on a web site,
'' without the explicit consent of the authors
''***************************************************************************
'
'Option Explicit
'

'-------------------------------------------------------------------------
Public Sub ImageBoxAdder()
'-------------------------------------------------------------------------
    Dim sPath As Variant
    With ThisWorkbook.Sheets(LIBMACROS_SH)
        sPath = .Cells(1, SM_Pic_N_COL)
        sPath = InputBox("Read all pictures from the directory." & vbCr & _
                         "The pictures are added to the dialog tree view dialog." & vbCr & _
                         "They should have a size of 16x16 pixel", "Read pictures for the tree view dialog", sPath)
        If sPath = "" Then Exit Sub
        .Cells(1, SM_Pic_N_COL) = sPath
    End With
    AddImagesToTreeForm (sPath)
End Sub

'-------------------------------------------------------------------------
Public Sub AddImagesToTreeForm(ByRef sPath As String, Optional Silent As Boolean = False)
'-------------------------------------------------------------------------
' Procedure : ImageBoxAdder
' Author    : Peter Thornton
' Created   : 26-01-2013
' Purpose   : A macro to add icon images at design time to a Frame
'-------------------------------------------------------------------------
' This function has to be called manualy if new icons are used

' ToDo: Prüfung einbauen ob alle Bilder in der Spalte "Picture Names" vorhanden sind

    Dim i As Long
    Dim lt As Single, tp As Single
    Dim arrIcons() As String
    Dim uf As UserForm
    Dim fm As MSForms.Frame
    Dim img As MSForms.Image
    Dim sFile As String

    ' In Windows ideally the images should be 16x16 pixels to fit 12x12 points
    ' ensure the images are "masked" in transparent areas, or make background colour same as
    ' the treeview's background, typically white
    '
    ' if after adding images with this code the images masked areas do not appear transparent,
    ' manually re-add from file, from the image control's Picture property
    '
    ' In Mac ideally the images should be about 20x20.
    ' They must be 8 bit/256 colours or less and not masked with a transparent colour
    ' (to make the images in this demo portable between Windows & Mac they have not been given transparent atributes)


    
    If Left(sPath, 2) = ".\" Then sPath = Application.ActiveWorkbook.Path & Mid(sPath, 2)       ' 19.11.21 Juergen: allow relative path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    arrIcons = FileNames(sPath)     ' get the file names

    If UBound(arrIcons) = -1 Then
       MsgBox "Error: No pictures found ;-(", vbCritical, "Error"
       Exit Sub
    End If

    Set uf = ThisWorkbook.VBProject.VBComponents("SelectMacrosTreeForm").designer
    Set fm = uf.Controls("frmImageBox")    ' ensure the form has a similarly named Frame

    ' remove existing images
    For i = fm.Controls.Count - 1 To 0 Step -1
           'On Error Resume Next                       ' 28.11.21: Vor some reasons we got a crash here
           fm.Controls.Remove i
           'On Error GoTo 0
    Next

    tp = 1.5
    Dim lf As Single                                                        ' Hardi
    lf = 1.5

    For i = 0 To UBound(arrIcons)
        sFile = sPath & arrIcons(i)

        Set img = fm.Controls.Add("forms.image.1")

        With img
            .BackStyle = fmBackStyleTransparent    ' needed to see through transparent if transparent gifs are used
            .Left = lf                                                      ' Hardi
            .Top = tp
            .Width = 12
            .Height = 12
            .Picture = LoadPicture(sFile)
            .BackStyle = fmBackStyleTransparent
            .Name = Left$(arrIcons(i), Len(arrIcons(i)) - 4)
        End With
        tp = tp + 15
        If tp + 15 > fm.Height Then                                         ' Hardi
           tp = 1.5
           lf = lf + 15
        End If
    Next
    
    If Silent Then                                                          ' Juergen 19.11.21
        Debug.Print UBound(arrIcons) + 1 & " pictures imported."
    Else
        MsgBox UBound(arrIcons) + 1 & " pictures imported." & vbCr & _
           vbCr & _
           "Check with Alt+F11 in the form 'SelectMacrosTreeForm'", vbInformation, UBound(arrIcons) + 1 & " pictures imported "
    End If
End Sub

'-----------------------------------------------
Private Function FileNames(ByVal Path As String)
'-----------------------------------------------
' Read the file names fron a directory
  Dim Files As String, Name As String                                       ' Hardi
  Name = Dir$(Path & "*.*")
  Do While Name <> ""
    If InStr(".bmp .gif .jpg", LCase(Right(Name, 4))) > 0 Then
       Files = Files & Name & vbCr
    End If
    Name = Dir$
  Loop
  
  FileNames = Split(DelLast(Files), vbCr)
End Function


