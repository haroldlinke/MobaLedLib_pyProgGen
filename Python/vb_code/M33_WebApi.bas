Attribute VB_Name = "M33_WebApi"
'-------------------------------------------------------------------
' VBA JSON Parser
' see https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------
Option Explicit
Private p&, token, dic

Function ParseJSON(json$, Optional Key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj Key Else ParseArr Key
    Set ParseJSON = dic
End Function

Function ParseObj(Key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr Key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add Key, "null"
                       Else
                           ParseObj Key
                       End If
                
            Case "}":  Key = ReducePath(Key): Exit Do
            Case ":":  Key = Key & "." & token(p - 1)
            Case ",":  Key = ReducePath(Key)
            Case Else: If token(p + 1) <> ":" Then dic.Add Key, token(p)
        End Select
    Loop
End Function

Function ParseArr(Key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj Key & ArrayID(e)
            Case "[":  ParseArr Key
            Case "]":  Exit Do
            Case ":":  Key = Key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add Key & ArrayID(e), token(p)
        End Select
    Loop
End Function '-------------------------------------------------------------------
' Support Functions
'-------------------------------------------------------------------
Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function
Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .Test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function
Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

Function ReducePath$(Key$)
    If InStr(Key, ".") Then ReducePath = left(Key, InStrRev(Key, ".") - 1) Else ReducePath = Key
End Function

Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function
Function GetFilteredValues(dic, Match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like Match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function
Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function


Function ListAllBetas()

  Dim Trans As String, objHTTP As Object, URL As String, getParam As String
  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  URL = "https://api.github.com/repos/Hardi-St/MobaLedLib_Docu/commits?path=Betatest/MobaLedLib-master.zip"
  objHTTP.Open "GET", URL, False
  objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.send ("")
  
  Dim Response
  Set Response = ParseJSON(objHTTP.responseText)
  
   Dim i, obj
   i = 0
   Do
    obj = "obj(" & i & ")"
    If Response.Exists(obj & ".url") Then
        Debug.Print Response(obj & ".commit.committer.date") & ": " & Response(obj & ".commit.committer.name") & ": " & Response(obj & ".commit.message")
    Else
        Exit Do
    End If
    i = i + 1
   Loop
   'Debug.Print ListPaths(Response)

End Function

