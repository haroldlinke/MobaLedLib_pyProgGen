VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTexts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Language
  DE = 0
  EN = 1
End Enum

Private Texts As Scripting.Dictionary

Private Sub class_initialize()
    Set Texts = New Scripting.Dictionary
End Sub

Public Sub SetText(Lang As Language, Key As String, Value As String)
    Texts(Lang & "." & Key) = Value
End Sub

Public Function GetText(Lang As Language, Key As String) As String
    If Texts.Exists(Lang & "." & Key) Then
        GetText = Texts(Lang & "." & Key)
    Else
        GetText = Key
    End If
End Function

