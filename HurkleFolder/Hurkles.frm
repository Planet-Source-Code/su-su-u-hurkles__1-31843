VERSION 5.00
Begin VB.Form Hurkles 
   BackColor       =   &H80000005&
   Caption         =   "Hurkles"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Buttons 
      Caption         =   "Grid"
      Default         =   -1  'True
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Buttons 
      Caption         =   "Quit"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox TextBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox TextBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Buttons 
      Caption         =   "New Game"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Buttons 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Labels 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Labels 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Labels 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Labels 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "Position  ("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Hurkles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private NewGame As Boolean
'
''----------------------------------------------------------------
'Private Sub Form_Load()
'  NumGuessAllowed = 5
'  XMax = 9: YMax = 9
'  StartSize = XMax + 1
'  Hurkles.Caption = "Hurkles is hiding in a " & XMax & " by " & YMax & " grid."
'  NewGame = True
'  Buttons_Click (1)
'
'End Sub
'
''----------------------------------------------------------------
'Private Sub Messages(Job As Integer)
'Dim Message As String
'  Select Case Job
'  Case 0
'
'    Buttons(0).Enabled = True
'    Counter = Counter + 1
'
'    Select Case Counter
'    Case 0 To NumGuessAllowed - 2
'      Labels(0).Caption = "Guess No. " & Counter
'
'    Case NumGuessAllowed - 1
'      Labels(1).Caption = "Last Guess!"
'
'    Case NumGuessAllowed
'      If HurkleFound = False Then
'        Labels(1).Caption = "Game Over!!"
'        Labels(0).Caption = "Hurkles was hiding at " & XPosition & "/" & YPosition
'      End If
'
'    Case Else
'    End Select
'
'  Case 1
'    TextBox(0).Text = ""
'    TextBox(1).Text = ""
'    Labels(0).Caption = "Where is Hurkles hiding?"
'    HurkleFound = False
'    Call Messages(0)
'
'  Case Else
'  End Select
'
'  If HurkleFound = True Then
'    Buttons(0).Enabled = False
'    Labels(1).Caption = ""
'  End If
'
'Labels(3).Caption = HurkleFound
'End Sub
'
''----------------------------------------------------------------
'Private Sub CheckPosition()
'  Dim Message As String
'
'  If (Trim(TextBox(0).Text) = "") Or (Trim(TextBox(1).Text) = "") Then
'    Labels(0).Caption = "Re-enter co-ordinates."
'    Exit Sub
'  End If
'
'  If Counter > NumGuessAllowed Then
'    Exit Sub
'  End If
'
'  XPlayer = Int(Val(TextBox(0).Text))
'  YPlayer = Int(Val(TextBox(1).Text))
'
'  Call CheckThePosition(Message)
'
'  lblMsg.Caption = Message
'
'  If Trim(Message) = "" Then
'    Labels(0).Caption = ""    'Counter
'    HurkleFound = True
'    NewGame = True
'  End If
'
'  Call Messages(0)
'
'
'End Sub
'
'
''----------------------------------------------------------------
'Private Sub Buttons_Click(Index As Integer)
'  Select Case Index
'  Case 0    'Enter
'    CheckPosition
'
'  Case 1    'New Game
'
'    XMax = XMax + 1
'    YMax = YMax + 1
'    Hurkles.Caption = "Hurkles is hiding in a " & XMax & " by " & YMax & " grid."
'    ' Generate random value between 0 and XMax & YMax.
'    XPosition = Int(XMax * Rnd) + 1
'    YPosition = Int(YMax * Rnd) + 1
'    Labels(2).Caption = "Position " & XPosition & "/" & YPosition
'    Buttons(0).Enabled = True
'    Counter = 0
'    Call Messages(1)
'
'    ' Increase NumGuessAllowed every 2 new games
'    If (XMax - StartSize) Mod 2 = 0 Then
'        NumGuessAllowed = NumGuessAllowed + 1
'    End If
'
'
'  Case 2    'Quit
'    Unload Me
'
'  Case 3
'    graphics.Show 1
'
'  Case Else
'  End Select
'End Sub
'
'Private Sub Label2_Click()
'    Label2.Caption = "hello"
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Label2.Caption = "X = " & X & " Y = " & Y & " Shift = " & Shift & " Button = " & Button
'End Sub
