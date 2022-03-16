Attribute VB_Name = "Serial_Test"
Option Explicit

' https://www.ozgrid.com/forum/forum/help-forums/excel-vba-macros/147023-vba-button-to-send-command-to-serial-port-to-get-data-and-send-to-excel


' Das Empfangen von seriellen Zeichen funktioniert prinzipiel, aber
' Die Input und Get Funktion warten so lange bis ein Zeichen kommt.
' Das führt dazu, dass nichts mehr geht bis das nächste Zeichen kommt ;-(
' Der Abbruch Button geht nur wenn ein Zeichen kommt

' Threads:
' https://codereview.stackexchange.com/questions/185212/a-new-approach-to-multithreading-in-excel
' https://analystcave.com/excel-vba-multithreading-tool/
' => C:\Users\Hardi\Downloads\VBA-Multithreading-Tool_20141221\VBA Multithreading Tool_20141221
' Noch nicht ausprobiert
'
' http://mikejuniperhill.blogspot.com/2017/06/excelvba-multi-threading-example.html

Public End_Serial_Test As Boolean

Private Test_Cnt As Long

'UT---------------------
Private Sub Debug_Text()
'UT---------------------
  Debug.Print "**" & Test_Cnt
  If Test_Cnt < 15 Then Application.OnTime Now + TimeValue("00:00:01"), "Debug_Text"
  Test_Cnt = Test_Cnt + 1
End Sub

'---------------------
Private Sub Test_Com()
'---------------------
  Dim c As String
  Dim fp As Integer
  fp = FreeFile

  End_Serial_Test = False
  Test_Cnt = 0
  UserForm_SerialTest.Show
  
    Open "COM7:115200,N,8,1" For Binary Access Read Write As #fp 'Open the com port
    'Put #fp, , sendVar$ 'write string to interface
    Application.OnTime Now + TimeValue("00:00:01"), "Debug_Text"
    c = ""
    While Not End_Serial_Test And Test_Cnt < 10
        'If Not EOF(fp) Then ' Geht nicht
            'c = Input(1, #fp)        ' Wartet bis ein Zeichen kommt ;-(
            c = "###"
            On Error Resume Next
            Get #fp, , c
            On Error GoTo 0
            If c <> "" And c <> "###" Then
                'Answer = Answer + c   'add, if printable char
                Debug.Print c;
            End If
        'End If
        DoEvents
        Sleep 100
    Wend
    Close #fp
    Debug.Print "End Serial Test"
End Sub

