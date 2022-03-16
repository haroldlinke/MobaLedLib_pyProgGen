Attribute VB_Name = "M09_Translate"
Option Explicit
' Module to translate cells by Misha

Private Const SrcLanguage = "en"

#If PROG_GENERATOR_PROG Then
  Private Const DestLanguages = "nl fr it es da" ' 16.06.20: Added da = Danish
#Else
  Private Const DestLanguages = "nl fr"
#End If

'-------------------------------------------
Private Function ConvertToGet(val As String)
'-------------------------------------------
    val = Replace(val, " ", "+")
    val = Replace(val, "#", "_HASH") ' 18.04.20: 09.05.20: Old "_HASH_", but Google adds an space ;-(  "_HASH _"
    'val = Replace(val, vbNewLine, "+")
    val = Replace(val, vbNewLine, "+VBNL+")  ' 17.06.20:
    val = Replace(val, vbCr, "+VBCR+")       '    "
    val = Replace(val, vbLf, "+VBLF+")       '    "
    val = Replace(val, "&", "&amp")
    val = Replace(val, "%", "_PERCENT")      ' 27.11.21:
    'val = Replace(val, "=", "&eq;")
    
    Const ReplaceTxt = "()="
    Dim i As Integer, c As String
    For i = 1 To Len(ReplaceTxt)
        c = Mid(ReplaceTxt, i, 1)
        val = Replace(val, c, "%" & Hex(Asc(c)))
    Next i
    
    ConvertToGet = val
End Function


'------------------------------------
Private Function Clean(val As String)
'------------------------------------
    val = Replace(val, "&quot;", """")
    val = Replace(val, "&gt;", ">")
    val = Replace(val, "&lt;", "<")
    val = Replace(val, "%2C", ",")
    val = Replace(val, "&#39;", "'")
    val = Replace(val, "_HASH", "#") ' 18.04.20: 09.05.20: Old "_HASH_", but Google adds an space ;-(  "_HASH _"
    val = Replace(val, "VBNL", vbNewLine) ' 17.06.20:
    val = Replace(val, "VBCR", vbCr)      '    "
    val = Replace(val, "VBLF", vbLf)      '    "
    val = Replace(val, "_PERCENT", "%")   ' 27.11.21:
    
    If left(val, 1) = "=" Then val = "'" & val
    Clean = val
End Function

'---------------------------------------------------------------------------------------------------------------------------------
Private Function RegexExecute(Str As String, reg As String, Optional matchIndex As Long, Optional subMatchIndex As Long) As String
'---------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandl
    Dim regex As Object, Matches As Object
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = reg
    regex.Global = Not (matchIndex = 0 And subMatchIndex = 0) 'For efficiency
    If regex.Test(Str) Then
        Set Matches = regex.Execute(Str)
        RegexExecute = Matches(matchIndex).submatches(subMatchIndex)
        Exit Function
    End If
ErrHandl:
    RegexExecute = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------------------------------------------
Public Function Translate_Text(Txt As String, ByVal translateFrom As String, ByVal translateTo As String) As String
'-------------------------------------------------------------------------------------------------------------------
' Example parameters:
' translateFrom = "de"
' translateTo = "en"
  Dim Trans As String, objHTTP As Object, URL As String, getParam As String
  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  getParam = ConvertToGet(Txt)
  ' Hier kann auch "https://translate.google.de" verwendet werden   oder URL = "https://translate.google.pl
  URL = "https://translate.google.pl/m?hl=" & translateFrom & "&sl=" & translateFrom & "&tl=" & translateTo & "&ie=UTF-8&prev=_m&q=" & getParam
  objHTTP.Open "GET", URL, False
  objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.send ("")
  If InStr(objHTTP.responseText, "div dir=""ltr""") > 0 Then
        ' "<html><head><title>Google Übersetzer</title><style>body{font:normal small arial,sans-serif,helvetica}body,html,form,div,p,img,input{margin:0padding:0}body{padding:3px}.nb{border:0}.s1{padding:5px}.ub{border-top:1px solid #36c}.db{border-bottom:1px so
        Trans = RegexExecute(objHTTP.responseText, "div[^""]*?""ltr"".*?>(.+?)</div>")
  ElseIf InStr(objHTTP.responseText, "<div class=""result-container"">") > 0 Then                      ' 19.01.21:
        Trans = RegexExecute(objHTTP.responseText, "<div class=""result-container"">(.+?)</div>")      '    "
  Else
        MsgBox "Error translating '" & Txt & "'" & vbCr & _
                vbCr & _
                "Response: '" & objHTTP.responseText & "'", vbCritical, "Error translating text"
  End If
  Translate_Text = Clean(Trans)
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub TranslateOneCell(src As Range, Optional DeltaCol As Long = 1, _
                             Optional StartOffs As Long = 0, _
                             Optional SrcLng As String = SrcLanguage, _
                             Optional DstLng As String = DestLanguages, _
                             Optional FirstCharUppercase As Boolean)
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Translate the current cell to all languages
  
  Dim Txt As String, DestCol As Long, DestLang As Variant, Row As Long, Res As String
  If src = "" Then Exit Sub
  Row = src.Row
  Txt = src
  DestCol = src.Column + DeltaCol + StartOffs
  For Each DestLang In Split(DstLng, " ")
    If Cells(Row, DestCol) = "" Then
       If left(src.Formula, 2) = "=$" Then ' Special translation if the the source starts with is a formula like =$H$8 (Used in the Group names)
          Dim Addr As String, NewAddr As String, DstRow As Long, TailingTxt As String
          Addr = Split(Mid(src.Formula, 2), " ")(0)                         ' 07.10.21: Added "Splitt..." to be able to process lines with mixed content (Link and text)
          DstRow = Range(Addr).Row
          NewAddr = Cells(DstRow, DestCol).Address
          TailingTxt = Trim(Mid(src.Formula, Len(Addr) + 2))
          If left(TailingTxt, 2) = "& " Then
             TailingTxt = Replace(Mid(TailingTxt, 3), """", "")
             TailingTxt = " & "" " & Translate_Text(TailingTxt, SrcLng, DestLang) & """"
          End If
          Res = "=" & NewAddr & TailingTxt                                  ' 07.10.21:
          With Cells(Row, DestCol).Font
              .ThemeColor = xlThemeColorDark1
              .TintAndShade = -0.499984740745262
          End With
       ElseIf left(src.Formula, 2) = "=C" Then ' Don't translate it if the source  points to column C like =C
          Res = src.Formula                    ' Used in the 'Languages' sheet
          With Cells(Row, DestCol).Font
              .ThemeColor = xlThemeColorDark1
              .TintAndShade = -0.499984740745262
          End With
       Else
          Res = Translate_Text(Txt, SrcLng, DestLang)  ' Translat the cell
       End If
       If FirstCharUppercase Then Res = left(UCase(Res), 1) & Mid(Res, 2) ' 26.10.21:
       Cells(Row, DestCol) = Res
    End If
    DestCol = DestCol + DeltaCol
  Next DestLang
End Sub
 
#If PROG_GENERATOR_PROG Then
'-----------------------------------------------
Private Sub Change_Language_in_Par_Description()
'-----------------------------------------------
' Translate the selected range in the "Par_Description" sheet
' The columns G and H could be selected at once
'
  Const DeltaCol = 2
  Dim StartOffs As Long, SrcLng As String, DstLng As String
  Select Case Selection.Column ' First column of the selection
      Case ParInTx_COL, _
           ParHint_COL: '   ' German to English to be able to check the english translation which is used for all other translations
                            SrcLng = "de"
                            DstLng = "en"
                            StartOffs = 1
      Case ParInTx_COL + 3, _
           ParHint_COL + 3: ' English to all other languages
                            SrcLng = "en"
                            DstLng = DestLanguages
                            StartOffs = 0
      Case Else: MsgBox "Translation for this column is not defined in 'Change_Language_in_Par_Description()"
                 Exit Sub
  End Select
  
  Dim c As Range
  
  For Each c In Selection
    TranslateOneCell c, DeltaCol, StartOffs, SrcLng:=SrcLng, DstLng:=DstLng
  Next c
End Sub

'------------------------------------------
Private Sub Change_Language_in_Lib_Macros()
'------------------------------------------
' Translate the selected range in the "Lib_Macros" sheet
  Const DeltaCol = DeltaCol_Lib_Macro_Lang                                  ' 07.10.21: Old: 2
  Const StartOffs = 0
  Dim SrcLng As String, DstLng As String
  Select Case Selection.Column ' First column of the selection
      Case SM_Group_COL, _
           SM_LName_COL, _
           SM_ShrtD_COL, _
           SM_DetailCOL: '   ' German to English to be able to check the english translation which is used for all other translations
                             SrcLng = "de"
                             DstLng = "en"
      Case SM_Group_COL + DeltaCol, _
           SM_LName_COL + DeltaCol, _
           SM_ShrtD_COL + DeltaCol, _
           SM_DetailCOL + DeltaCol:   ' English to all other languages
                             SrcLng = "en"
                             DstLng = DestLanguages
      Case Else: MsgBox "Translation for this column is not defined in 'Change_Language_in_Lib_Macros()"
                 Exit Sub
  End Select
  
  Dim c As Range
  
  For Each c In Selection
    TranslateOneCell c, DeltaCol, StartOffs, SrcLng:=SrcLng, DstLng:=DstLng, FirstCharUppercase:=(Selection.Column = SM_Group_COL) ' 26.10.21: Added: FirstCharUppercase ...
  Next c
End Sub
#End If ' PROG_GENERATOR_PROG

'-----------------------------------------------
Private Sub Change_Language_in_Languages_Sheet()
'-----------------------------------------------
' Translate the selected range in the "Languages" sheet
  Const DeltaCol = 1
  Const StartOffs = 0
  Dim SrcLng As String, DstLng As String
  Select Case Selection.Column ' First column of the selection
      Case FirstLangCol:    ' German to English to be able to check the english translation which is used for all other translations
                             SrcLng = "de"
                             DstLng = "en"
      Case FirstLangCol + 1: ' English to all other languages
                             SrcLng = "en"
                             DstLng = DestLanguages
      Case Else: MsgBox "Translation for this column is not defined in 'Change_Language_in_Lib_Macros()"
                 Exit Sub
  End Select
  
  Dim c As Range
  
  For Each c In Selection
    TranslateOneCell c, DeltaCol, StartOffs, SrcLng:=SrcLng, DstLng:=DstLng
  Next c

#If 0 Then
  Const DeltaCol = 1
  Const StartOffs = 0
  Dim c As Range
  
  For Each c In Selection
    TranslateOneCell c, DeltaCol, StartOffs
  Next c
#End If
End Sub


'-------------------------------
Public Sub Translate_Selection()
Attribute Translate_Selection.VB_ProcData.VB_Invoke_Func = "T\n14"
'-------------------------------
' This macro is called when CTRL+Shift+T is pressed
  Select Case ActiveSheet.Name
#If PROG_GENERATOR_PROG Then
    Case PAR_DESCR_SH: Change_Language_in_Par_Description
    Case LIBMACROS_SH: Change_Language_in_Lib_Macros
#End If
    Case LANGUAGES_SH: Change_Language_in_Languages_Sheet
  End Select
End Sub



