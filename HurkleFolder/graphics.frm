VERSION 5.00
Begin VB.Form graphics 
   Caption         =   "Grid Display"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Other Info."
      Height          =   1695
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
      Begin VB.Label lblGridRef 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblComments 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblColWdth 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   6615
   End
   Begin VB.Label cntLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NumCols As Integer
Dim ColWdth As Integer
Dim TheXPos As Integer
Dim TheYPos As Integer

Const AllowedSpaceSize = 4000
Const XStart = 100
Const YStart = 100

Private Sub CalculateXYValue(X As Single, Y As Single)
    TheXPos = Int((X - XStart) / ColWdth)
    TheYPos = Int((Y - YStart) / ColWdth)
    
    'Display grid reference
    TheXPos = TheXPos + 1
    TheYPos = NumCols - TheYPos
    
    lblGridRef.Caption = "Coordinates (X/Y): " & TheXPos & "/" & TheYPos
    
End Sub
Private Sub ColorGrid(Job As Integer)
    Dim Xtop As Long
    Dim Ytop As Long
    
    If Job = 1 Then 'Color in square clicked
        If TheXPos < 1 Then Exit Sub
        If TheXPos > NumCols Then Exit Sub
        If TheYPos < 1 Then Exit Sub
        If TheYPos > NumCols Then Exit Sub
    
        'Calculate top corner of clicked square
        Xtop = XStart + ((TheXPos - 1) * ColWdth)
        Ytop = YStart + ((NumCols - TheYPos) * ColWdth)
    
        'Color the square
        Line (Xtop + 20, Ytop + 25)-(Xtop + ColWdth - 20, Ytop + ColWdth - 29), vbYellow, BF
    
        'Draw a cross
        Line (Xtop, Ytop)-(Xtop + ColWdth, Ytop + ColWdth), vbRed
        Line (Xtop + ColWdth, Ytop)-(Xtop, Ytop + ColWdth), vbRed
        
    ElseIf Job = 2 Then 'Color in square in which hurkle was hiding
        Xtop = XStart + ((XPosition - 1) * ColWdth)
        Ytop = YStart + ((NumCols - YPosition) * ColWdth)
        
        'Color square in red, indicating where hurkle was hiding
        Line (Xtop + 20, Ytop + 25)-(Xtop + ColWdth - 20, Ytop + ColWdth - 29), vbRed, BF
    ElseIf Job = 3 Then
        Xtop = XStart + ((TheXPos - 1) * ColWdth)
        Ytop = YStart + ((NumCols - TheYPos) * ColWdth)
        
        'Color in square in green, indicating that the hurkle was found
        Line (Xtop + 20, Ytop + 25)-(Xtop + ColWdth - 20, Ytop + ColWdth - 29), vbGreen, BF
    End If
    
    
End Sub
Private Sub DrawGrid()

    Cls

    ColWdth = AllowedSpaceSize / NumCols
    
    For i = 0 To NumCols
        Line (XStart, YStart + (i * ColWdth))-(XStart + (NumCols * ColWdth), YStart + (i * ColWdth)), vbBlue
        
        Line (XStart + (i * ColWdth), YStart)-(XStart + (i * ColWdth), YStart + (NumCols * ColWdth)), vbBlue
    Next i
    
    
End Sub

Private Sub cmdNewGame_Click(Index As Integer)

    NumCols = NumCols + 1
    graphics.Caption = "Hurkles is hiding in a " & NumCols & " by " & NumCols & " grid."
    
    
    Randomize
    ' Generate random value between 0 and XMax & YMax.
    XPosition = Int(NumCols * Rnd) + 1
    YPosition = Int(NumCols * Rnd) + 1
    
    Counter = 0
    cntLabel.Caption = "Guess # " & Counter + 1
    
    ColWdth = AllowedSpaceSize / NumCols
    
    If NumCols <> StartSize Then
        If (NumCols - StartSize) Mod 4 = 0 Then
            NumGuessAllowed = NumGuessAllowed + 1
        End If
    End If
    
    HurkleFound = False
    
    Call DrawGrid
    
End Sub
Private Sub CheckThePosition()

    Dim NorthWord, EastWord
    
    If YPosition > TheYPos Then
        NorthWord = " North"
    ElseIf YPosition < TheYPos Then
        NorthWord = " South"
    Else
        NorthWord = ""
    End If
    
    If XPosition > TheXPos Then
        EastWord = " East"
    ElseIf XPosition < TheXPos Then
        EastWord = " West"
    Else
        EastWord = ""
    End If
    
    If Trim(NorthWord & EastWord) = "" Then
        lblMessage.Caption = "Well done, you found him! " & "(in " & Counter & " attempts)."
        HurkleFound = True
        ColorGrid (3)
    Else
        If Counter < NumGuessAllowed Then
            lblMessage.Caption = "[Guess No. " & Counter & " ] Go" & NorthWord & EastWord
        Else
            lblMessage.Caption = "Hurkles was hiding at " & XPosition & "/" & YPosition
            ColorGrid (2)
        End If
    End If
    
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub



Private Sub cmdReset_Click()
    graphics.Caption = "Hurkles is hiding in a " & NumCols & " by " & NumCols & " grid."
    
    Randomize
    ' Generate random value between 0 and XMax & YMax.
    XPosition = Int(NumCols * Rnd) + 1
    YPosition = Int(NumCols * Rnd) + 1
    
    Counter = 0
    cntLabel.Caption = "Guess # " & Counter + 1
    
    ColWdth = AllowedSpaceSize / NumCols
    
    HurkleFound = False
    
    Call DrawGrid
End Sub

Private Sub Form_Load()
    NumCols = 9
    NumGuessAllowed = 5
    StartSize = NumCols + 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblComments.Caption = "X = " & X & " Y = " & Y

    'Boundary checking
    If X < XStart Then Exit Sub
    If X > XStart + AllowedSpaceSize Then Exit Sub
    If Y < YStart Then Exit Sub
    If Y > YStart + AllowedSpaceSize Then Exit Sub
    
    'Increment counter
    Counter = Counter + 1
    
    If HurkleFound = True Then
        Counter = 0
        cntLabel.Caption = "Guess # " & Counter + 1
        Call DrawGrid
        HurkleFound = False
        Randomize
        XPosition = Int(NumCols * Rnd) + 1
        YPosition = Int(NumCols * Rnd) + 1
        Exit Sub
    End If
    
    If Counter > NumGuessAllowed Then
        Counter = 0
        cntLabel.Caption = "Guess # " & Counter + 1
        Call DrawGrid
        XPosition = Int(NumCols * Rnd) + 1
        YPosition = Int(NumCols * Rnd) + 1
        Exit Sub
    ElseIf Counter = NumGuessAllowed Then
        cntLabel.Caption = "Game over"
        
    ElseIf Counter + 1 = NumGuessAllowed Then
        cntLabel.Caption = "Last Guess"
    Else
        cntLabel.Caption = "Guess # " & Counter
    End If
    
    lblColWdth.Caption = "ColWdth = " & ColWdth
    
    
    Call CalculateXYValue(X, Y)
    Call ColorGrid(1)
    Call CheckThePosition
       
End Sub


