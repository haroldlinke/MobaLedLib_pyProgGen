VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExtensionParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Min As String
Public Max As String
Public TypeName As String
Public Default As String
Public Options As String
Public Texts As clsTexts

Private Sub class_initialize()
    Set Texts = New clsTexts
End Sub
