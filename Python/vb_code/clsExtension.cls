VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Platforms As Collection
Public Id As Integer
Public Name As String
Public Path As String
Public Includes As String
Public MacroIncludes As String
Public Constructors As Collection
Public Macros As Collection
Public Parameters As Collection

' the name of the extension must contain only basic letters (A-Z or a-z) and numbers (0-9), underscores (_)and dashes (-).
' They must start with a letter or number. They must contain at least one letter.
'
Private Const NamePattern = "^[A-Za-z0-9][A-Za-z0-9_\-]*$"

Private Sub class_initialize()
    Set Constructors = New Collection
    Set Macros = New Collection
    Set Parameters = New Collection
End Sub

Public Function GetConstructor(ByVal TypeName As String)
    Dim Constructor
    For Each Constructor In Constructors
        If Constructor.TypeName = TypeName Then
            Set GetConstructor = Constructor
            Exit Function
        End If
    Next
    Set GetConstructor = Nothing
End Function

'-----------------------------------------------------------
'return an error message in case of check fails, otherwise empty string
'-----------------------------------------------------------
Public Function CheckValid() As String
    If Not Matches(Name, NamePattern) Then
        CheckValid = "Extension Name is invalid"
        Exit Function
    End If
    CheckValid = ""
End Function

'-----------------------------------------------------------
'return an error message in case of check fails, otherwise empty string
'-----------------------------------------------------------
Public Function IsExtensionParameter(ByVal Argument As String) As Boolean
    Dim Parameter
    For Each Parameter In Parameters
        If Parameter.Name = Argument Then
            IsExtensionParameter = True
            Exit Function
        End If
    Next
End Function

