VERSION 5.00
Begin VB.Form FrmHighScores 
   BackColor       =   &H00BBCCB3&
   Caption         =   "High Scores"
   ClientHeight    =   7485
   ClientLeft      =   2475
   ClientTop       =   2085
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BBCCB3&
      Height          =   4935
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   3800
         TabIndex        =   32
         Top             =   3637
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   9
         Left            =   3800
         TabIndex        =   29
         Top             =   4020
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   7
         Left            =   3800
         TabIndex        =   28
         Top             =   3248
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   6
         Left            =   3800
         TabIndex        =   27
         Top             =   2859
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   5
         Left            =   3800
         TabIndex        =   26
         Top             =   2470
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   4
         Left            =   3800
         TabIndex        =   25
         Top             =   2081
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   3
         Left            =   3800
         TabIndex        =   24
         Top             =   1680
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   2
         Left            =   3800
         TabIndex        =   23
         Top             =   1303
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   1
         Left            =   3800
         TabIndex        =   22
         Top             =   914
         Width           =   1600
      End
      Begin VB.Label lblHighScoreDate 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Index           =   0
         Left            =   3800
         TabIndex        =   21
         Top             =   525
         Width           =   1600
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   2280
         TabIndex        =   20
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   2280
         TabIndex        =   19
         Top             =   3637
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   2280
         TabIndex        =   18
         Top             =   3248
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2280
         TabIndex        =   17
         Top             =   2859
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2280
         TabIndex        =   16
         Top             =   2470
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2280
         TabIndex        =   15
         Top             =   2081
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Top             =   1692
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Top             =   1303
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   914
         Width           =   1575
      End
      Begin VB.Label lblHighScore 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   10
         Top             =   4020
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   360
         TabIndex        =   9
         Top             =   3637
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   3248
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2859
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2470
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2081
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1692
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1303
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   914
         Width           =   1700
      End
      Begin VB.Label lblPlayerName 
         BackColor       =   &H00BBCCB3&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   525
         Width           =   1700
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00BBCCB3&
      Caption         =   "TOP TEN PLAYER SCORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   30
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "FrmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdExit_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
Dim i As Integer
For i = 0 To 9
    With HighScoreRecord(i + 1)
        lblPlayerName(i) = .PlayerName
        lblHighScore(i) = .PlayerScore
        lblHighScoreDate(i) = Format(.ScoreDate, "mmm dd, yyyy")
    End With
    Next i

End Sub


Private Sub Form_Load()
Dim i As Integer

'fraHighScore.BackColor = Backcolor2nd
'fraHighScore.ForeColor = ForeColor2nd
Frame1.BackColor = BackColor
Frame1.ForeColor = ForeColor
Label1.ForeColor = ForeColor
Label1.BackColor = BackColor
cmdExit.BackColor = BackColor


End Sub

Private Sub Frame1_Click()
cmdExit.value = True                'cause cmdexit click event

End Sub


