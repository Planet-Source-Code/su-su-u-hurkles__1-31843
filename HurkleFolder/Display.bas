Attribute VB_Name = "StandAloneSubs"
'Project ONE for Year 10 Software Design
'HURKLES
'Feb 2002

'Daniel Ho
'10(1) Computer Studies


Public XMax As Integer
Public YMax As Integer
Public XPosition As Integer
Public YPosition As Integer
Public XPlayer As Integer
Public YPlayer As Integer
Public Counter As Integer
Public HurkleFound As Boolean
Public NumGuessAllowed As Integer
Public StartSize As Integer





'Sub Main()
''MASTER CONTROLLER
'
'  Randomize ' Initialize random-number generator.
'  Hurkles.Show 1
'
'End Sub
'
'
''----------------------------------------------------------------
'Public Sub CheckThePosition(Message As String)
'
'Dim Response As Boolean
'Dim NorthWord, EastWord
'
'  If YPosition > YPlayer Then
'    NorthWord = " North"
'  ElseIf YPosition < YPlayer Then
'    NorthWord = " South"
'  Else
'    NorthWord = ""
'    CNorthWord = ""
'  End If
'
'  If XPosition > XPlayer Then
'    EastWord = " East"
'  ElseIf XPosition < XPlayer Then
'    EastWord = " West"
'  Else
'    EastWord = ""
'  End If
'
'  If Trim(NorthWord & EastWord) = "" Then
'    'Chr(10) results in a <return>  ie New Line
'    Message = "Well done, you found him!" & "(in " & Counter & " attempts)."
'    HurkleFound = True
'  Else
'    If Counter < NumGuessAllowed Then
'      Message = "[Gues No " & Counter & " ] Go" & NorthWord & EastWord
'    Else
'      Message = "Hurkles was hiding at " & XPosition & "/" & YPosition
'    End If
'  End If
'
'End Sub
