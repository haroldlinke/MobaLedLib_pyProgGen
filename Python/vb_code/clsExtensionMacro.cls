VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExtensionMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Texts As clsTexts
Public Arguments As Collection
Public LEDs As String
Public InCnt As String

Private Sub class_initialize()
    Set Texts = New clsTexts
    Set Arguments = New Collection
End Sub
