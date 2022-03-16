Attribute VB_Name = "M38_Extensions"
Option Explicit

Public Extensions As Collection
Public Const ExtensionKey = "EX."

Private ExtensionsActive As Scripting.Dictionary
Private ExtensionLines As Scripting.Dictionary
Private JSONData As Object

Sub Load_Extensions()

    CollectExtensions
    RemoveExistingExtensions
  
    Dim OldEvents, ParamSheet, MacroSheet
    OldEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    Set ParamSheet = ThisWorkbook.Sheets(PAR_DESCR_SH)
    Set MacroSheet = ThisWorkbook.Sheets(LIBMACROS_SH)
    
    Dim Extension As clsExtension, Macro As clsExtensionMacro, Parameter As clsExtensionParameter
    Dim Constructor As clsExtensionConstructor, Argument
    Dim MacroRow As Integer, ParamRow As Integer, MacroStr As String
    
    MacroRow = LastUsedRowIn(MacroSheet) + 1
    ParamRow = LastUsedRowIn(ParamSheet) + 1
    
    For Each Extension In Extensions
      For Each Constructor In Extension.Constructors
        ' TODO translate
        MacroStr = ""
        For Each Argument In Constructor.Arguments
            If MacroStr <> "" Then MacroStr = MacroStr & ", "
            
            ' fully qualify the argument name if argument is defined withon the plugin
            If Extension.IsExtensionParameter(Argument) Then MacroStr = MacroStr & ExtensionKey & Extension.Name & "."
            
            MacroStr = MacroStr & Argument
        Next
        MacroStr = ExtensionKey & Constructor.TypeName & "(" & MacroStr & ")"
        
        MacroSheet.Cells(MacroRow, SM_Typ___COL) = "EX.Constructor"
        MacroSheet.Cells(MacroRow, SM_Pic_N_COL) = "Puzzle | IDE"
        MacroSheet.Cells(MacroRow, SM_LEDS__COL) = Constructor.LEDs
        MacroSheet.Cells(MacroRow, SM_InCnt_COL) = Constructor.InCnt
        MacroSheet.Cells(MacroRow, SM_Macro_COL) = MacroStr
        MacroSheet.Cells(MacroRow, SM_FindN_COL) = ExtensionKey & Constructor.TypeName & "("
        MacroSheet.Cells(MacroRow, SM_Name__COL) = ExtensionKey & Constructor.TypeName
        MacroSheet.Cells(MacroRow, SM_Group_COL) = "Erweiterungen"
        MacroSheet.Cells(MacroRow, SM_LName_COL) = Constructor.Texts.GetText(DE, "DisplayName")
        MacroSheet.Cells(MacroRow, SM_ShrtD_COL) = Constructor.Texts.GetText(DE, "ShortDescription")
        MacroSheet.Cells(MacroRow, SM_DetailCOL) = Constructor.Texts.GetText(DE, "DetailedDescription")
        MacroRow = MacroRow + 1
      Next
    
      For Each Parameter In Extension.Parameters
        ParamSheet.Cells(ParamRow, ParName_COL) = "EX." & Extension.Name & "." & Parameter.Name
        ParamSheet.Cells(ParamRow, ParType_COL) = Parameter.TypeName
        ParamSheet.Cells(ParamRow, Par_Min_COL) = Parameter.Min
        ParamSheet.Cells(ParamRow, Par_Max_COL) = Parameter.Max
        ParamSheet.Cells(ParamRow, Par_Def_COL) = Parameter.Default
        ParamSheet.Cells(ParamRow, Par_Opt_COL) = Parameter.Options
        ParamSheet.Cells(ParamRow, ParInTx_COL) = Parameter.Texts.GetText(DE, "DisplayName")
        ParamSheet.Cells(ParamRow, ParHint_COL) = Parameter.Texts.GetText(DE, "ShortDescription")
        
        ParamRow = ParamRow + 1
      Next
    
      For Each Macro In Extension.Macros
        ' TODO translate
        MacroStr = ""
        For Each Argument In Macro.Arguments
            If MacroStr <> "" Then MacroStr = MacroStr & ", "
            MacroStr = MacroStr & Argument
        Next
        If MacroStr = "" Then
            MacroStr = Macro.Name
        Else
            MacroStr = Macro.Name & "(" & MacroStr & ")"
        End If
        
        MacroSheet.Cells(MacroRow, SM_Typ___COL) = "EX.Macro"
        MacroSheet.Cells(MacroRow, SM_Pic_N_COL) = "Puzzle | IDE"
        MacroSheet.Cells(MacroRow, SM_LEDS__COL) = Macro.LEDs
        MacroSheet.Cells(MacroRow, SM_InCnt_COL) = Macro.InCnt
        MacroSheet.Cells(MacroRow, SM_Macro_COL) = MacroStr
        MacroSheet.Cells(MacroRow, SM_FindN_COL) = MacroStr
        If Macro.Arguments.Count <> 0 Then MacroSheet.Cells(MacroRow, SM_FindN_COL) = MacroSheet.Cells(MacroRow, SM_FindN_COL) & "("
        
        MacroSheet.Cells(MacroRow, SM_Name__COL) = Macro.Name
        MacroSheet.Cells(MacroRow, SM_Group_COL) = "Erweiterungen"
        MacroSheet.Cells(MacroRow, SM_LName_COL) = Macro.Texts.GetText(DE, "DisplayName")
        MacroSheet.Cells(MacroRow, SM_ShrtD_COL) = Macro.Texts.GetText(DE, "ShortDescription")
        MacroSheet.Cells(MacroRow, SM_DetailCOL) = Macro.Texts.GetText(DE, "DetailedDescription")
        MacroRow = MacroRow + 1
      Next
    Next
    Application.EnableEvents = OldEvents
End Sub

Sub CollectExtensions()
  Set Extensions = New Collection
  Dim oFSO
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  If Not CheckArduinoHomeDir() Then Exit Sub
  
  Dim DirName, Id As Integer
  Dim Libs As Collection
  Set Libs = New Collection
  
  DirName = Dir(Sketchbook_Path & "\libraries\*.*", vbDirectory)
  While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
        If oFSO.FolderExists(Sketchbook_Path & "\libraries\" & DirName) Then
            If oFSO.FileExists(Sketchbook_Path & "\libraries\" & DirName & "\MobaLedLib.properties") Then
                Libs.Add Sketchbook_Path & "\libraries\" & DirName & "\MobaLedLib.properties"
            End If
        End If
    End If
    DirName = Dir
  Wend
  
  Id = 1
  For Each DirName In Libs
    Dim FileStr As String, LibName As String
    LibName = DirName
    Set JSONData = Nothing
    On Error Resume Next
    Set JSONData = ParseJSON(Read_File_to_String(LibName))
    On Error GoTo 0
    If Not JSONData Is Nothing Then
        Dim Extension As clsExtension
        Set Extension = New clsExtension
        Set Extension.Platforms = ToList("obj.platforms(#)")
        Extension.Name = JSONData("obj.id")
        Extension.Path = Replace(LibName, "\MobaLedLib.properties", "")
        Extension.Includes = JSONData("obj.includes")
        Extension.MacroIncludes = JSONData("obj.macroIncludes")
        Dim typeList
        Dim Index, Index2
        For Each Index In ToIndexList("obj.types(#).TypeName")
            Dim ctr
            Set ctr = New clsExtensionConstructor
            ctr.TypeName = GetDataAtIndex("obj.types(#).TypeName", Index)
            ctr.LEDs = GetDataAtIndex("obj.types(#).LEDs", Index)
            ctr.InCnt = GetDataAtIndex("obj.types(#).InCount", Index)
            ctr.Texts.SetText DE, "DisplayName", GetDataAtIndex("obj.types(#).DisplayName", Index)
            ctr.Texts.SetText DE, "ShortDescription", GetDataAtIndex("obj.types(#).ShortDescription", Index)
            ctr.Texts.SetText DE, "DetailedDescription", GetDataAtIndex("obj.types(#).DetailedDescription", Index)
            For Each Index2 In ToIndexList("obj.types(" & Index & ").Arguments(#)")
                ctr.Arguments.Add GetDataAtIndex("obj.types(" & Index & ").Arguments(#)", Index2)
            Next
            Extension.Constructors.Add ctr
        Next
        For Each Index In ToIndexList("obj.macros(#).Macro")
            Dim mac
            Set mac = New clsExtensionMacro
            mac.Name = GetDataAtIndex("obj.macros(#).Macro", Index)
            mac.LEDs = GetDataAtIndex("obj.macros(#).LEDs", Index)
            mac.InCnt = GetDataAtIndex("obj.macros(#).InCount", Index)
            mac.Texts.SetText DE, "DisplayName", GetDataAtIndex("obj.macros(#).DisplayName", Index)
            mac.Texts.SetText DE, "ShortDescription", GetDataAtIndex("obj.macros(#).ShortDescription", Index)
            mac.Texts.SetText DE, "DetailedDescription", GetDataAtIndex("obj.macros(#).DetailedDescription", Index)
            For Each Index2 In ToIndexList("obj.macros(" & Index & ").Arguments(#)")
                mac.Arguments.Add GetDataAtIndex("obj.macros(" & Index & ").Arguments(#)", Index2)
            Next
            Extension.Macros.Add mac
        Next
        For Each Index In ToIndexList("obj.parameters(#).ParameterName")
            Dim par
            Set par = New clsExtensionParameter
            par.Name = GetDataAtIndex("obj.parameters(#).ParameterName", Index)
            par.TypeName = GetDataAtIndex("obj.parameters(#).Type", Index)
            par.Min = GetDataAtIndex("obj.parameters(#).Min", Index)
            par.Max = GetDataAtIndex("obj.parameters(#).Max", Index)
            par.Default = GetDataAtIndex("obj.parameters(#).Default", Index)
            par.Options = GetDataAtIndex("obj.parameters(#).Options", Index)
            par.Texts.SetText DE, "DisplayName", GetDataAtIndex("obj.parameters(#).DisplayName", Index)
            par.Texts.SetText DE, "ShortDescription", GetDataAtIndex("obj.parameters(#).ShortDescription", Index)
            Extension.Parameters.Add par
        Next
        Dim Message As String
        Message = Extension.CheckValid
        If Message = "" Then
            Extensions.Add Extension
            Id = Id + 1
        Else
            MsgBox Replace("Extension '#1#' load error", "#1#", Extension.Name) & vbCrLf$ & Message, vbExclamation, "MobaLedLib Extensions"
        End If
    End If
  Next

End Sub

Private Function ToList(Pattern As String) As Collection
    Dim search As String
    Dim cnt As Integer
    cnt = 0
    Set ToList = New Collection
    
    Do
        search = Replace(Pattern, "#", cnt)
        If Not JSONData.Exists(search) Then Exit Do
        ToList.Add JSONData(search)
        cnt = cnt + 1
    Loop
End Function

Private Function GetDataAtIndex(Pattern As String, ByVal Index As Integer) As String
    Dim search As String
    GetDataAtIndex = ""
    search = Replace(Pattern, "#", Index)
    If Not JSONData.Exists(search) Then Exit Function
    GetDataAtIndex = JSONData(search)
End Function

Private Function ToIndexList(Pattern As String) As Collection
    Dim search As String
    Dim cnt As Integer
    cnt = 0
    Set ToIndexList = New Collection
    
    Do
        search = Replace(Pattern, "#", cnt)
        If Not JSONData.Exists(search) Then Exit Do
        ToIndexList.Add cnt
        cnt = cnt + 1
    Loop
End Function

Public Function GetExtension(ByVal Id As String) As clsExtension
    Dim Extension
    For Each Extension In Extensions
        If Extension.Id = Id Then
            Set GetExtension = Extension
            Exit Function
        End If
    Next
    Set GetExtension = Nothing
End Function

Public Function GetConstructor(ByVal Key As String) As clsExtensionConstructor
    Dim Splits
    Splits = Split(Key, ".")
    If UBound(Splits) = 3 And Splits(0) = "PI" And Splits(1) = "Constructor" Then
        Dim Extension
        Set Extension = GetExtension(Splits(2))
        If Not Extension Is Nothing Then
            Set GetConstructor = Extension.GetConstructor(Splits(3))
        End If
    End If
End Function

Public Function GetExtensionByTypeName(ByVal TypeName As String) As clsExtension
    Dim Extension, Constructor
    For Each Extension In Extensions
        For Each Constructor In Extension.Constructors
            If Constructor.TypeName = TypeName Then
                Set GetExtensionByTypeName = Extension
                Exit Function
            End If
        Next
    Next
    Set GetExtensionByTypeName = Nothing
End Function


Public Function IsExtensionKey(Key As String) As Boolean
    IsExtensionKey = InStr(Key, ExtensionKey) = 1
End Function

'-------------------------------------------------------------------------------
Public Function Init_HeaderFile_Generation_Extension() As Boolean
  Set ExtensionsActive = New Scripting.Dictionary
  Set ExtensionLines = New Scripting.Dictionary
  CollectExtensions
  Init_HeaderFile_Generation_Extension = True
End Function

'-------------------------------------------------------------------------------
Public Function Add_Extension_Entry(ByRef Cmd As String) As Boolean

  Dim Str As String, Extension As clsExtension
  Cmd = Mid(Cmd, Len(ExtensionKey) + 1)
  Str = Cmd
  If InStr(Cmd, "(") > 0 Then
     Str = Split(Str, "(")(0) ' Cut of the text after the "("
  End If
  Set Extension = GetExtensionByTypeName(Str)
  If Extension Is Nothing Then
    MsgBox Replace("Extension '#1#' not found", "#1#", Str), vbCritical, "MobaLedLib Extensions"
    Add_Extension_Entry = False
    Exit Function
  End If
  
  If Not ExtensionsActive.Exists(Extension.Id) Then
    ExtensionsActive.Add Extension.Id, Extension
  End If
  
  ExtensionLines.Add ExtensionLines.Count + 1, Cmd
  Add_Extension_Entry = True
End Function

'------------------------------------------------------------------
Public Function Write_Header_File_Extension_Before_Config(fp As Integer) As Boolean
'------------------------------------------------------------------
' here we write include all the header files, that have definitions already needed inside the config block
' e.g. macros

  Dim Extension As clsExtension, Str, HeaderWritten As Boolean
  HeaderWritten = False
  If ExtensionsActive.Count > 0 Then
    For Each Extension In Extensions
      If Extension.MacroIncludes <> "" Then
      
        If HeaderWritten = False Then
          Print #fp, "// ----- dynamic extensions section begin -----"
          Print #fp, "#define MLL_EXTENSIONS_ACTIVE"
          Print #fp, "#include <MLLExtension.h>"
          Print #fp, ""
          HeaderWritten = True
        End If

        Print #fp, "// Extension " & Extension.Name
        For Each Str In Split(Extension.MacroIncludes, ",")
          Print #fp, "#include <" & Str & ">";
        Next
        Print #fp, ""
      End If
    Next
    If HeaderWritten Then
      Print #fp, "// ----- dynamic Extensions section end -----"
    End If
  End If
  Write_Header_File_Extension_Before_Config = True
End Function

'------------------------------------------------------------------
Public Function Write_Header_File_Extension_After_Config(fp As Integer) As Boolean
'------------------------------------------------------------------
  Dim Extension, Str, Index As Integer

  If ExtensionsActive.Count > 0 Then
    Print #fp, "// ----- dynamic extensions section begin -----"
    Print #fp, "#include <MLLExtension.h>"
    Print #fp, ""
    For Each Extension In Extensions
      If Extension.Includes <> "" Then
        Print #fp, "// Extension " & Extension.Name
        For Each Str In Split(Extension.Includes, ",")
            Print #fp, "#include <" & Str & ">";
        Next
      End If
      Print #fp, ""
    Next
    Print #fp, ""

    If ExtensionLines.Count > 0 Then
      Print #fp, "// ----- dynamic extensions section begin -----"
      Print #fp, "MLLExtension* mllExtensions[] = {"
      For Index = 1 To ExtensionLines.Count
        If Index = ExtensionLines.Count Then
          Print #fp, "  new " & ExtensionLines(Index)
        Else
          Print #fp, "  new " & ExtensionLines(Index) & ","
        End If
      Next
      Print #fp, "} ;"
      Print #fp, "#define MLL_EXTENSIONS_COUNT " & ExtensionLines.Count
      Print #fp, ""
    End If
    Print #fp, "// ----- dynamic Extensions section end -----"
    Print #fp, ""
  End If
  Write_Header_File_Extension_After_Config = True
End Function

Public Function Write_PIO_Extension(fp As Integer) As Boolean
  Dim Extension, Index As Integer
  
  For Index = 0 To ExtensionsActive.Count - 1
    Print #fp, "    " & ExtensionsActive(Index).Name & "=file://" & GetShortPath(ExtensionsActive(Index).Path)
  Next
  Write_PIO_Extension = True

End Function

Public Sub RemoveExistingExtensions()
    Dim r As Integer, MacroSheet, ParamSheet
    Set MacroSheet = ThisWorkbook.Sheets(LIBMACROS_SH)
    Set ParamSheet = ThisWorkbook.Sheets(PAR_DESCR_SH)
    For r = 1 To LastUsedRowIn(ParamSheet)
      If IsExtensionKey(ParamSheet.Cells(r, ParName_COL)) Then
           ParamSheet.Rows(r).EntireRow.Delete
         r = r - 1
      End If
    Next
    
    For r = SM_DIALOGDATA_ROW1 To LastUsedRowIn(MacroSheet)
      If IsExtensionKey(MacroSheet.Cells(r, SM_Typ___COL)) Then
           MacroSheet.Rows(r).EntireRow.Delete
         r = r - 1
      End If
    Next
End Sub
