VERSION 5.00
Begin VB.Form FrmYahtzeeForm 
   BackColor       =   &H00A3BAC5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Yahtzee by Dennis Stroh"
   ClientHeight    =   6600
   ClientLeft      =   3330
   ClientTop       =   3075
   ClientWidth     =   8490
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00AF0000&
   Icon            =   "FrmYatzeeForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSound 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Sound"
      Height          =   615
      Left            =   3600
      TabIndex        =   77
      Top             =   5880
      Width           =   1695
      Begin VB.OptionButton optSoundOn 
         BackColor       =   &H00A3BAC5&
         Caption         =   "O&n"
         Height          =   315
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSoundOff 
         BackColor       =   &H00A3BAC5&
         Caption         =   "O&ff"
         Height          =   315
         Left            =   960
         TabIndex        =   78
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraHighScore 
      BackColor       =   &H00A3BAC5&
      Caption         =   "All Time High Score"
      Height          =   1215
      Left            =   5520
      TabIndex        =   66
      Top             =   120
      Width           =   2895
      Begin VB.Label lblHighScoreAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00A3BAC5&
         Caption         =   "high score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblHighScoreName 
         Alignment       =   2  'Center
         BackColor       =   &H00A3BAC5&
         Caption         =   "player's name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraPlayerScore 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Player's Score"
      Height          =   3410
      Left            =   5520
      TabIndex        =   65
      ToolTipText     =   "Move mouse over each player's score to show game status."
      Top             =   2325
      Width           =   2895
      Begin VB.Frame FraPlayer 
         BackColor       =   &H00A3BAC5&
         Caption         =   "Player 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   450
         TabIndex        =   73
         ToolTipText     =   "Click on Player Name to Change or Remove"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2000
         Begin VB.Label lblPlayerScore 
            Alignment       =   2  'Center
            BackColor       =   &H00A3BAC5&
            Caption         =   "score 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Index           =   3
            Left            =   120
            TabIndex        =   74
            Top             =   285
            Width           =   1695
         End
      End
      Begin VB.Frame FraPlayer 
         BackColor       =   &H00A3BAC5&
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   450
         TabIndex        =   71
         ToolTipText     =   "Click on Player Name to Change or Remove"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2000
         Begin VB.Label lblPlayerScore 
            Alignment       =   2  'Center
            BackColor       =   &H00A3BAC5&
            Caption         =   "score 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   285
            Width           =   1680
         End
      End
      Begin VB.Frame FraPlayer 
         BackColor       =   &H00A3BAC5&
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   450
         TabIndex        =   69
         ToolTipText     =   "Click on Player Name to Change or Remove"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2000
         Begin VB.Label lblPlayerScore 
            Alignment       =   2  'Center
            BackColor       =   &H00A3BAC5&
            Caption         =   "score 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   285
            Width           =   1695
         End
      End
      Begin VB.Frame FraPlayer 
         BackColor       =   &H00A3BAC5&
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   450
         TabIndex        =   67
         ToolTipText     =   "Click on Player Name to Change or Remove"
         Top             =   360
         Visible         =   0   'False
         Width           =   2000
         Begin VB.Label lblPlayerScore 
            Alignment       =   2  'Center
            BackColor       =   &H00A3BAC5&
            Caption         =   "score 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   68
            Top             =   285
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraCurrentPlayer 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Current Player"
      Height          =   735
      Left            =   5520
      TabIndex        =   63
      Top             =   1440
      Width           =   2895
      Begin VB.Label lblCurrentPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H00A3BAC5&
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AF0000&
         Height          =   435
         Left            =   120
         TabIndex        =   64
         ToolTipText     =   "Click to change players"
         Top             =   240
         Width           =   2700
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00BBCCB3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0083FAAC&
      Height          =   3375
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   62
      Text            =   "FrmYatzeeForm.frx":5A4A
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Fours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   915
      TabIndex        =   42
      Top             =   3855
      Width           =   950
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      Top             =   2812
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   3337
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   3862
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   36
      Top             =   4380
      Width           =   495
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   35
      Top             =   4905
      Width           =   495
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   2287
      Width           =   480
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Ones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   915
      TabIndex        =   33
      Top             =   2280
      Width           =   950
   End
   Begin VB.Frame fraScore 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Score"
      Height          =   615
      Left            =   4200
      TabIndex        =   31
      Top             =   960
      Width           =   1095
      Begin VB.Label lblTotalScore 
         Alignment       =   2  'Center
         BackColor       =   &H00A3BAC5&
         Caption         =   "177"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   32
         Top             =   165
         Width           =   1000
      End
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Yahtzee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3210
      TabIndex        =   30
      Top             =   4440
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Chance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3210
      TabIndex        =   29
      Top             =   4080
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "F&ull House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   3210
      TabIndex        =   28
      Top             =   3720
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Large Straight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3210
      TabIndex        =   27
      Top             =   3360
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "S&mall Straight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3210
      TabIndex        =   26
      Top             =   3000
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&4 of a Kind"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   3210
      TabIndex        =   25
      Top             =   2640
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&3 of a Kind"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3210
      TabIndex        =   24
      Top             =   2280
      Width           =   1600
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Sixes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   915
      TabIndex        =   23
      Top             =   4920
      Width           =   950
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Fi&ves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   915
      TabIndex        =   22
      Top             =   4380
      Width           =   950
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Thr&ees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   915
      TabIndex        =   21
      Top             =   3330
      Width           =   950
   End
   Begin VB.CheckBox ChkDice 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Twos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   915
      TabIndex        =   20
      Top             =   2800
      Width           =   950
   End
   Begin VB.OptionButton optRedDie 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Re&d"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   6120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optGreenDie 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Green"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin VB.OptionButton OptWhiteDie 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&White"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00A3BAC5&
      Caption         =   "E&xit"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdRoll 
      BackColor       =   &H00A3BAC5&
      Caption         =   "&Roll"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox NumPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   8520
      Picture         =   "FrmYatzeeForm.frx":5A65
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2040
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2640
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox Die 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox SourceDicePic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00A3BAC5&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00A3BAC5&
      Height          =   2895
      Left            =   8520
      Picture         =   "FrmYatzeeForm.frx":6D3B
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame FraHoldButtons 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Hold Buttons"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   3375
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00809BA8&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   500
      End
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00809BA8&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   500
      End
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00809BA8&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   500
      End
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00809BA8&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   500
      End
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00809BA8&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   500
      End
   End
   Begin VB.Frame fraDiceColor 
      BackColor       =   &H00A3BAC5&
      Caption         =   "Dice Color"
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BackStyle       =   0  'Transparent
      Caption         =   "djs Windows"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00527383&
      Height          =   255
      Left            =   960
      TabIndex        =   80
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblBonus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "+35 Bonus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblYahtzeeBonus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00A3BAC5&
      Caption         =   "Extra Yahtzees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   59
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   4680
      TabIndex        =   58
      Top             =   4920
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00A3BAC5&
      Caption         =   "Total"
      Height          =   255
      Left            =   3720
      TabIndex        =   57
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblSubTotal2 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   56
      Top             =   5400
      Width           =   585
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00A3BAC5&
      Caption         =   "Total"
      Height          =   255
      Left            =   840
      TabIndex        =   55
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   4800
      TabIndex        =   54
      Top             =   4200
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   4800
      TabIndex        =   53
      Top             =   3840
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   4800
      TabIndex        =   52
      Top             =   3480
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   4800
      TabIndex        =   51
      Top             =   3120
      Width           =   465
   End
   Begin VB.Label LblSubTotal1 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   49
      Top             =   5400
      Width           =   585
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1920
      TabIndex        =   48
      Top             =   3975
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1920
      TabIndex        =   47
      Top             =   5025
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1920
      TabIndex        =   46
      Top             =   4500
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1920
      TabIndex        =   45
      Top             =   3450
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1920
      TabIndex        =   44
      Top             =   2925
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1920
      TabIndex        =   43
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   4800
      TabIndex        =   41
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   4800
      TabIndex        =   40
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      Caption         =   "YAHTZEE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AF0000&
      Height          =   1215
      Left            =   360
      TabIndex        =   18
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00A3BAC5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   4800
      TabIndex        =   50
      Top             =   4560
      Width           =   465
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "&Save Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "&Player"
      Begin VB.Menu mnuHighScore 
         Caption         =   "&Show High Scores"
      End
      Begin VB.Menu mnuAddPlayer 
         Caption         =   "&Add Player"
      End
      Begin VB.Menu mnuEditPlayer 
         Caption         =   "&Edit Player"
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit 1"
            Index           =   1
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit 2"
            Index           =   2
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit 3"
            Index           =   3
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit 4"
            Index           =   4
         End
      End
      Begin VB.Menu mnuRemovePlayer 
         Caption         =   "&Remove Player"
         Index           =   1
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove 1"
            Index           =   1
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove 2"
            Index           =   2
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove 3"
            Index           =   3
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove 4"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmYahtzeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Throws As Integer    '3 throws per turn, remember between calls
Dim TurnValid As Boolean
Dim TurnOver As Boolean
Dim YahtzeeBonus As Integer
Dim debug1 As Integer       'used for debug & test
Dim SignOnSound As String, RollSound As String
Dim YahtzeeSound As String, SelectDiceSound As String
Dim GameOverSound As String, MakeSelectionsound As String
Dim errorSound As String, UnselectDiceSound As String
Dim BonusSound As String, ChangePlayerSound As String
Dim CurrentPath As String
Dim CurrentPlayer As Integer
Dim GameOver As Boolean
Dim DoneOnce As Boolean
Dim ignore As Boolean


Dim BackColorMain As Long
Dim ForeColorMain As Long
Dim TitleColor As Long
Dim Backcolor2nd As Long
Dim ForeColor2nd As Long
Dim GameOverBackColor As Long
Dim GameOverForeColor As Long
Dim ChkBoxBackcolor As Long
Dim ChkBoxForeColor As Long


Public Enum DiceRollResults
    eOnes = 1            'force enumeration to begin at one
    eTwos
    eThrees
    eFours
    eFives
    eSixes
    e3ofakind
    e4ofaKind
    eSmStraight
    eLgStraight
    eFullHouse
    eChance
    eYahtzee
End Enum
Public Enum PlayerRoster
    eNew
    eRemove
    eAdd
    eEdit
End Enum

Dim RollResults(eOnes To eYahtzee) As Integer                'holds the results of the throw

Private Sub ChkDice_Click(Index As Integer)
     'Player has checked one of 13 check boxes to
     '  score the dice in that category.
 
    Dim i As Integer
    Dim j As Integer
    Dim st As Integer
    Dim matches As Integer
    Dim match(1 To 6) As Integer
    Dim twoofakind As Boolean           'need for full house
    Dim threeofakind As Boolean
    Dim fourofakind As Boolean
    Dim fiveofakind As Boolean
If Not ignore Then
With ChkDice(Index)
    'Check to see if checkbox is enabled, not grayed,
    '   and valid turn
    ' CheckBox.value = 0 is Unchecked (default),
    '                  1 is Checked, and
    '                  2 is Grayed (dimmed).
If (.Enabled = True Or Index = eYahtzee) And TurnValid = True Then
    
    TurnValid = False                     'no more turns this go around
    PlayerControls DISABLE                'disable dice & checkboxes

  If Index = eYahtzee Then                'Yahtzee = 50 & yahtzees are automatic
        If RollResults(eYahtzee) = 50 And ChkDice(eYahtzee).value <> 2 Then
            PlaySound YahtzeeSound          'it's a Yahtzee
            delay 2                         'allow sound to play
            If Val(lblTotal(13).Caption) < 50 Then
                st = 50                     '50 for 1st yahtzee
                lblTotal(13).Caption = Val(st)
            Else
                st = 100                    '100 for each extra yahtzee
                lblTotal(14).Caption = Str(Val(lblTotal(14).Caption) + st)
                YahtzeeBonus = YahtzeeBonus + 1
                lblYahtzeeBonus.Caption = Str(YahtzeeBonus) + " Extra Yahtzee"
            End If
            ChkDice(eYahtzee).value = 2         'change checked value to 2
            RollResults(eYahtzee) = 0           'reset roll results
        End If
    End If
        
'  lblTotal(Index).Caption = Str(0)
  
  '**************************************************************************
  If Index <> eYahtzee Then             'Yahtzees are different
      lblTotal(Index).Caption = Str(RollResults(Index))      'show score for selection
  End If
  
  Select Case Index
    Case eOnes To eSixes
        UpdateScore True         'update lower part of score sheet
    Case e3ofakind To eYahtzee
        UpdateScore False         'update lower part of score sheet
  End Select
    
  'UpdatePlayerRecord CurrentPlayer     'do this at scoring time
  
  Resetfor1stRoll
 End If          '.enabled = true
End With
TurnOver = True
End If          'if not ignore
End Sub

Private Sub chkHold_Click(Index As Integer)
    If chkHold(Index).value = 1 Then
      chkHold(Index).Caption = "Hold"
      PlaySound SelectDiceSound
    Else
      chkHold(Index).Caption = ""
      PlaySound UnselectDiceSound
     delay 0.1
    End If
    PlayerRecord(CurrentPlayer).Holdstatus(Index) = chkHold(Index).value
End Sub

Private Sub cmdExit_Click()
Dim i As Integer
Dim result As Integer

If Not GameOver Then
    result = MsgBox("Save this Game?", vbYesNoCancel)
    Select Case result
    Case vbCancel
        Exit Sub
    Case vbYes
        SaveGame
    Case vbNo
        GameOver = True         'set gameover flag so player stats are reset
    End Select
End If
  
If GameOver Then              'game is over reset player stats
    For i = 1 To 4
        ResetPlayerValues (i)
        Put #1, i, PlayerRecord(i)  'save blanks
    Next i
End If
ExitGame

End Sub

Private Sub cmdNew_Click()          'new game
Dim i As Integer
    For i = 1 To NumberOfPlayers
        ResetPlayerValues (i)
        PlayerRecord(i).PlayerName = FraPlayer(i - 1).Caption
    Next i
    Initialize
End Sub

Private Sub cmdRoll_Click()
    Dim i As Integer
    Dim j As Integer
    Dim TimeDelay As Single
    Dim diecolor As Integer

    If Throws < 4 Then
      For i = 1 To 13
        ChkDice(i).ForeColor = vbBlack  'unhighlight all categories for the new roll
      Next i
      DoEvents      'needed to update the chkdice(n).foreColor
      TurnValid = True
      CheckDiceColor diecolor
      PlaySound RollSound
                        
      TimeDelay = Timer + 2             'delay 2 seconds
      Do While Timer < TimeDelay
         For i = 1 To 5                 'roll 5 dice
           RollDice SourceDicePic, Die(i), i, diecolor
         Next i
      Loop
      
      If debug1 Then                    'used for debug
        For i = 1 To 5
            DiceValue(i) = debug1
        Next i
        debug1 = 0                      'reset debug flag
      End If
      
      '*** here check to see what has been thrown
      PlayerControls ENABLE             'enable holds
      GetRollResults
      Throws = Throws + 1
      
      'Yahtzees are automatic
      If RollResults(eYahtzee) > 0 Then
            ChkDice(eYahtzee).value = 1   'force the yahtzee checkbox
      End If
                 
      If Throws < 4 Then
        cmdRoll.Caption = "&Roll" + Str(Throws)
        For i = 1 To 5
            chkHold(i).Enabled = True
        Next i
      Else
        cmdRoll.Enabled = False
        For i = 1 To 5
            chkHold(i).Enabled = False
        Next i
      End If
  End If
  UpdatePlayerRecord (CurrentPlayer)
End Sub

Private Sub Die_Click(Index As Integer)
  Select Case Index
    Case 1 To 5                  'horizontal dice
        chkHold(Index).value = 1 Xor (chkHold(Index).value) 'toggle value
    Case 6 To 11                'vertical dice
        ChkDice(Index - 5).value = 1    'invoke ChkDice click event
        'Call ChkDice_Click(Index - 5)
    Case Else
  End Select
End Sub

Private Sub Die_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is for debug.  "shift-click" over this label forces a Yahtzee
    If Shift = 1 Then
      debug1 = Index
      If debug1 > 5 Then debug1 = Index - 6
      cmdRoll.value = True              'invoke roll button's click event
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim diecolor As Integer

  CurrentPath = App.Path + "\"          'try to get the current path

  GetSounds                             'load sound effects
  SetColors                             '*** not fully functional ***
  GetHighScores
  Load FrmHighScores
  ShowHighScorePerson
  Select Case UseOldGame        'check for old game saved
    Case False                  'start with new players
        GetNumberofPlayers eNew
        Initialize
    Case True                   'read old game
        For i = 1 To 4
            Get #1, i, PlayerRecord(i)
        Next
        UpdatePlayerNames       'get number of players, names, etc.
        For i = 1 To NumberOfPlayers
            With PlayerRecord(i)   'find who is current player to roll
                If .IsCurrentPlayer = True Then
                    CurrentPlayer = i
                    Throws = .Throws
                End If
            End With
        Next i
        ShowPlayerRecord (CurrentPlayer)
        lblCurrentPlayer.Caption = PlayerRecord(CurrentPlayer).PlayerName
        CheckDiceColor diecolor         'show vert. dice
        For i = 1 To 5
            DiceValue(i) = PlayerRecord(CurrentPlayer).diceFaceValues(i)
            ShowDie SourceDicePic, Die(i), i, DiceValue(i), diecolor    'show the last roll
            chkHold(i).value = PlayerRecord(CurrentPlayer).Holdstatus(i)
            If chkHold(i).value = 1 Then
                chkHold(i).Caption = "Hold"
            Else
                chkHold(i).Caption = ""
            End If
        Next
        cmdRoll.Caption = "&Roll " + Str(Throws)
        For i = 1 To NumberOfPlayers
            GetRollResults
            ShowPlayerRecord (i)
        Next

  End Select
 
End Sub

Private Sub CheckDiceColor(diecolor As Integer)
  Dim i As Integer
      
    If optRedDie.value = True Then
        diecolor = 3
    ElseIf optGreenDie.value = True Then
        diecolor = 2
    Else
        diecolor = 1
    End If
    For i = 6 To 11          'show six dice for scoring
       ShowDie SourceDicePic, Die(i), i, (i - 5), diecolor
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer

For i = 0 To 3                  'make sure colors are reset if fast mouse move
  If FraPlayer(i).BackColor <> fraPlayerScore.BackColor Then
        lblTotalScore.BackColor = fraPlayerScore.BackColor
        For j = 0 To 3
            FraPlayer(j).BackColor = fraPlayerScore.BackColor
            lblPlayerScore(j).BackColor = fraPlayerScore.BackColor
        Next j
        ShowPlayerRecord (CurrentPlayer)
        For j = 1 To 14
            lblTotal(j).BackColor = Me.BackColor
        Next j
        Exit For    'colors changed already
    End If
Next i

End Sub



Private Sub fraPlayer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RemovePlayer Index + 1
End Sub

Private Sub fraPlayer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowPlayerStatus Index
End Sub

Private Sub fraPlayerScore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Not DoneOnce Then            'already displayed once, don't do again
    ShowPlayerRecord (CurrentPlayer)
    lblTotalScore.BackColor = fraPlayerScore.BackColor
    For i = 0 To 3                          'return backgrounds to normal color
        FraPlayer(i).BackColor = fraPlayerScore.BackColor
        lblPlayerScore(i).BackColor = fraPlayerScore.BackColor
    Next i
    For i = 1 To 14
        lblTotal(i).BackColor = fraPlayerScore.BackColor
    Next i
    
    'make sure chkdice boxes are disabled for CurrentPlayer if player had already made selection
    'If TurnValid = False Or
    If Throws = 1 Then PlayerControls DISABLE
End If
DoneOnce = True
End Sub

Private Sub fraHighScore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmHighScores.Hide
End Sub



Private Sub lblHighScoreAmount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmHighScores.Show
End Sub

Private Sub lblHighScoreName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmHighScores.Show
End Sub

Private Sub lblPlayerScore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowPlayerStatus (Index)
End Sub

Private Sub lblYahtzeeBonus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'this is for debug.  "shift-click" over this label forces a Yahtzee
    If Shift = 1 Then
      debug1 = 5
      cmdRoll.value = True              'invoke roll button's click event
    End If
End Sub

Private Sub mnuAbout_Click()
    frmSplash.Show
End Sub

Private Sub mnuAddPlayer_Click()
Dim result As Integer
Dim i As Integer

If NumberOfPlayers = 4 Then
    MsgBox "Cannot add more than 4 players", "Maximum Player Error", vbOKOnly
    Exit Sub
End If

'need to restart new game if game in progress
result = MsgBox("Adding a player will start a new game", vbOKCancel)
If result = vbCancel Then Exit Sub        'don't restart the game
For i = 1 To 4
    ResetPlayerValues (i)
Next
Initialize

GetNumberofPlayers eAdd
End Sub

Private Sub mnuEdit_Click(Index As Integer)
RemovePlayer (Index)                    'this procedure can either remove or edit
End Sub



Private Sub mnuExit_Click()
    cmdExit.value = True
End Sub

Private Sub mnuHighScore_Click()
    FrmHighScores.Show
End Sub

Private Sub MnuInstructions_Click()
    frmYahtzeeHS_Help.Show
    'frmHelp.Show
End Sub

Private Sub mnuNewGame_Click()
    cmdNew.value = True
End Sub

Private Sub mnuRemove_Click(Index As Integer)
    RemovePlayer Index          'remove player n
End Sub

Private Sub mnuSaveGame_Click()
    Dim result As Integer
    SaveGame
    result = MsgBox("Game Saved. Do You Want to Quit?", vbYesNo)
    If result = vbYes Then ExitGame
End Sub

Private Sub optGreenDie_Click()
    CheckDiceColor 2
End Sub

Private Sub optRedDie_Click()
    CheckDiceColor 3
End Sub

Private Sub OptWhiteDie_Click()
    CheckDiceColor 1
End Sub

Private Sub Resetfor1stRoll()
Dim i As Integer
For i = 1 To 5
  chkHold(i).value = 0
  chkHold(i).Enabled = False
Next i
cmdRoll.Enabled = True
Throws = 1
cmdRoll.Caption = "&Roll 1"
For i = eOnes To eYahtzee
  ChkDice(i).ToolTipText = ""
  ChkDice(i).Enabled = False
Next i
lblCurrentPlayer.Caption = PlayerRecord(CurrentPlayer).PlayerName
For i = 1 To 4
    If i = CurrentPlayer Then
     PlayerRecord(i).IsCurrentPlayer = True
    Else
    PlayerRecord(i).IsCurrentPlayer = False
    End If
Next i
    
End Sub

Private Sub UpdateScore(Upper As Boolean)
Dim BonusGranted As Boolean

Dim i As Integer
Dim st1 As Integer
Dim st2 As Integer
st1 = 0: st2 = 0
BonusGranted = Val(LblSubTotal1) > 63
For i = e3ofakind To eYahtzee + 1           'include Yahtzee bonus
    st2 = st2 + Val(lblTotal(i).Caption)
Next i
For i = eOnes To eSixes                     'total upper score area
    st1 = st1 + Val(lblTotal(i).Caption)
Next i

If Upper = True Then                        'This is the upper score on the score sheet
    If st1 <= 63 Then BonusGranted = False

    If st1 > 62 And BonusGranted = False Then
        'BonusGranted = True
        st1 = st1 + 35           'bonus of 35 if over => 63
        lblBonus.Visible = True
        'lblBonus.BackColor = vbBlack
        'lblBonus.BackStyle = 1               'opaque, 0 = transparent
        'lblBonus.Caption = "+35 BONUS"
        'lblBonus.Enabled = True
        'lblBonus.ForeColor = vbRed
        PlaySound BonusSound
        delay 2                              'delay to allow sound to play
    End If
    LblSubTotal1.Caption = Str(st1)         'show upper score
Else                                        'update lower score
    lblSubTotal2.Caption = Str(st2)
End If

PlaySound MakeSelectionsound
lblTotalScore.Caption = Str(st1 + st2)

'update current player info & record
lblPlayerScore(CurrentPlayer - 1).Caption = Str(st1 + st2)
Throws = 1
UpdatePlayerRecord CurrentPlayer
delay 1

'if last player, check to see if game is over
If CurrentPlayer = NumberOfPlayers Then     'not over till last player plays
    GameOver = True                         'set GameOver flag
    For i = 1 To 13
      'check to see of any boxes still not yet checked
        If ChkDice(i).value = 0 Then GameOver = False
    Next
End If
If GameOver Then
    Text1.Visible = True
    PlaySound GameOverSound
    delay 2
    SaveHighScores
    ShowHighScorePerson
Else
   PlaySound ChangePlayerSound
    delay 1.5                'allow time to see score
   ChangePlayerFocus
End If

End Sub

Public Sub GetSounds()
Dim WavSounds As String
Dim i As Integer
'Yahtzee.ini is a textfile with names of sound files, can include path
'This file can be edited but order must stay the same.
'Sounds are kept in "Sounds" directory
  SoundDirectory = App.Path + "\Sounds\" 'This is where the sounds are
Open CurrentPath + "Yahtzee.ini" For Input As #1
  i = 0
  Do While Not EOF(1)
    Input #1, WavSounds
    'if ' then comment field
   If Left$(WavSounds, 1) <> "'" And WavSounds <> "" Then
      Select Case i                     'assign sounds to global variables
       Case 0                               'This is where the sounds are
         If StrComp(WavSounds, "Sounds", vbTextCompare) = 0 Then      'Use the default directory
            SoundDirectory = App.Path + "\Sounds\"
         Else
            SoundDirectory = WavSounds
         End If
       Case 1
         SignOnSound = WavSounds
       Case 2
         RollSound = WavSounds
       Case 3
         YahtzeeSound = WavSounds
       Case 4
         SelectDiceSound = WavSounds
       Case 5
         GameOverSound = WavSounds
       Case 6
         MakeSelectionsound = WavSounds
       Case 7
         errorSound = WavSounds
       Case 8
         UnselectDiceSound = WavSounds
       Case 9
         BonusSound = WavSounds
       Case 10
         ChangePlayerSound = WavSounds
       Case Else
         '
     End Select
     i = i + 1                          'point to next element in array
   End If
  Loop
Close #1
End Sub

Public Sub Initialize()
  Dim i As Integer
  Dim diecolor As Integer
  ignore = False         'flag to not respond to chkdice click when .value changes
  DoneOnce = True
  debug1 = 0
  CurrentPlayer = 1
  PlaySound SignOnSound
  TurnValid = True
  YahtzeeBonus = 0
  CheckDiceColor diecolor
  PlayerControls DISABLE
  For i = 1 To 5
    RollDice SourceDicePic, Die(i), i, diecolor
  Next

  For i = 1 To 13
    lblTotal(i).Caption = ""
    ChkDice(i).value = 0                    'uncheck all boxes
  Next i
  YahtzeeBonus = 0
  lblTotal(14).Caption = ""
  LblSubTotal1.Caption = ""
  lblSubTotal2.Caption = ""
  lblBonus.Visible = False
  lblTotalScore.Caption = ""
  Text1.Visible = False
  lblYahtzeeBonus.Caption = "Extra Yahtzees"
  
  PlayerControls DISABLE
  Resetfor1stRoll
  
  For i = 0 To NumberOfPlayers - 1       'set players names on main form
    With PlayerRecord(i + 1)
      lblPlayerScore(i).Caption = ""
      FraPlayer(i).Caption = Trim(.PlayerName)
      mnuRemove(i + 1).Caption = .PlayerName    'set up menu items
      mnuEdit(i + 1).Caption = .PlayerName
      If .PlayerName = Space$(Len(.PlayerName)) Then
         FraPlayer(i).Visible = False
      Else
         FraPlayer(i).Visible = True
      End If
    End With
  Next i
  
  For i = NumberOfPlayers + 1 To 4
   mnuRemove(i).Visible = False
   mnuEdit(i).Visible = False
  Next i
  
  lblCurrentPlayer.Caption = PlayerRecord(CurrentPlayer).PlayerName
  ShowHighScorePerson
 
End Sub
Public Function OldGameSaved() As Boolean
Dim i, j, r As Integer

'PlayerRecord declared in yahtzee.bas module as public of type PRecord
'Dim PlayerRecord As PRecord
'Public Type PRecord                          ' Define user-defined type.
'    PlayerName As String * 10
'    PlayerScore As Integer
'    GameStatus(1 To 2, 1 To 14) As Integer  'keeps checkmark status and score of each area
'    Holdstatus(1 To 5) As Integer           'keeps 5 hold button status
'    diceFaceValues(1 To 5) As Integer       'keeps the current 5 dice face values
'    Throws As Integer                       'number of throws taken
'    IsCurrentPlayer as boolean              'flag for current player
'End Type

Open CurrentPath + "Yahtzee.dat" For Random As #1 Len = Len(PlayerRecord(1))
If LOF(1) = 0 Then                          'build new file
    OldGameSaved = False
    For r = 1 To 4
       ResetPlayerValues (r)
    Next r
    For i = 1 To 4
        Put #1, 1, PlayerRecord(i)
    Next
Else
    For i = 1 To 4
        Get #1, i, PlayerRecord(i)      'get the array of four players
        With PlayerRecord(i)            'see if there is a name
            If .PlayerName <> Space$(Len(.PlayerName)) Then OldGameSaved = True
        End With
    Next
End If
End Function

Public Sub GetHighScores()

'Public Type HSRecord
'    PlayerName As String * 10
'    PlayerScore As Integer
'    ScoreDate As Date
'End Type
'Dim HighScoreRecord(1 To 10) As HSRecord

Dim i As Integer, j As Integer, k As Integer
Dim temp As HSRecord

Open CurrentPath + "YahScore.dat" For Random As #2 Len = Len(HighScoreRecord(i))

If LOF(2) = 0 Then           'build new file

    For i = 1 To NumberOfPlayers
        With HighScoreRecord(i)
        .PlayerName = PlayerRecord(i).PlayerName
        .PlayerScore = PlayerRecord(i).PlayerScore
        .ScoreDate = Date
        End With
        Put #2, i, HighScoreRecord(i)
    Next i
    temp.PlayerName = ""                'set rest to null
    temp.PlayerScore = 0
    temp.ScoreDate = Date
    For i = NumberOfPlayers + 1 To 10
        HighScoreRecord(i) = temp
        Put #2, i, HighScoreRecord(i)
    Next i

  i = 1                             'sort the scores
  Do While i < 10
      With HighScoreRecord(i)
      If .PlayerScore < HighScoreRecord(i + 1).PlayerScore Then
          temp = HighScoreRecord(i)
          HighScoreRecord(i) = HighScoreRecord(i + 1)
          HighScoreRecord(i + 1) = temp
          i = 0
      End If
      End With
      i = i + 1
  Loop
 
 For i = 1 To 10                    'rewrite the sorted array
    Put #2, i, HighScoreRecord(i)
 Next
 End If        'lof 2 = 0
 
'read the file
  For i = 1 To 10
    Get #2, i, HighScoreRecord(i)
  Next i
End Sub
  
Public Sub SaveHighScores()
'Public Type HSRecord
'    PlayerName As String * 10
'    PlayerScore As Integer
'    ScoreDate As Date
'End Type
'Dim HighScoreRecord(1 To 10) As HSRecord

Dim i As Integer, j As Integer, k As Integer
Dim temp As HSRecord
   
'check the current game scores to see if are among highest
For j = 0 To NumberOfPlayers - 1
  For k = 10 To 1 Step -1
    If Val(lblPlayerScore(j)) > HighScoreRecord(k).PlayerScore Then
       HighScoreRecord(k).PlayerScore = Val(lblPlayerScore(j))
       HighScoreRecord(k).PlayerName = FraPlayer(j).Caption
       HighScoreRecord(k).ScoreDate = Date
       'bubblesort the ten top records
       i = 1
       Do While i < 10                     'sort the scores
          With HighScoreRecord(i)
              If .PlayerScore < HighScoreRecord(i + 1).PlayerScore Then
                  temp = HighScoreRecord(i)
                  HighScoreRecord(i) = HighScoreRecord(i + 1)
                  HighScoreRecord(i + 1) = temp
                  i = 0
              End If
          i = i + 1
          End With
       Loop
       Exit For       'for k = 1 to 10
    End If
  Next k
Next j              'next player score
For i = 1 To 10
    Put #2, i, HighScoreRecord(i)
Next i
    
End Sub

Private Sub GetNumberofPlayers(value)

Select Case value
  Case eNew, eAdd
    frmPlayerNames.Show vbModal             'show the get player's names form
    Initialize
    
  Case eRemove, eEdit
        frmPlayerChanges.Show vbModal
        'do not initialize just remove the player
  Case Else
End Select
    
End Sub

Private Sub ResetPlayerValues(PlayerNum As Integer)
Dim i, j
    With PlayerRecord(PlayerNum)
        .PlayerName = ""
        .PlayerScore = 0
        For i = 1 To 2
            For j = 1 To 14
            .GameStatus(i, j) = 0
        Next j, i
        For i = 1 To 5
            .Holdstatus(i) = 0
            .diceFaceValues(i) = 1
        Next
        .Throws = 0
        .IsCurrentPlayer = False
    End With
End Sub

Private Function UseOldGame()
    NewGame = True
    If OldGameSaved Then
        frmOldOrNewGame.Show (1)             'use saved game?
    End If
    UseOldGame = Not NewGame
End Function

Private Sub UpdatePlayerRecord(Player As Integer)
Dim i As Integer, j As Integer
With PlayerRecord(Player)
    .PlayerScore = Val(lblTotalScore.Caption)           'get total score
    For j = 1 To 13
        .GameStatus(1, j) = ChkDice(j).value            'get checked value
        .GameStatus(2, j) = Val(lblTotal(j).Caption)    'getindividual scores
    Next j
        .GameStatus(2, 14) = Val(lblTotal(j).Caption)   'extra Yahtzee total
    
    For i = 1 To 5
        .Holdstatus(i) = chkHold(i).value           '.Enabled
        .diceFaceValues(i) = DiceValue(i)
    Next i
    .Throws = Throws
    .IsCurrentPlayer = (CurrentPlayer = Player)
End With
End Sub

Private Sub ShowPlayerRecord(Player As Integer)
Dim i As Integer, j As Integer
Dim LowerScore As Integer, UpperScore As Integer
ignore = True               'no click events when .value changes
With PlayerRecord(Player)
    lblTotalScore.Caption = Str(.PlayerScore)          'get total score
    lblPlayerScore(Player - 1).Caption = Str(.PlayerScore)
    For j = eOnes To eYahtzee
       ChkDice(j).Enabled = (.GameStatus(1, j) = 0)
       ChkDice(j).value = .GameStatus(1, j)            'checked value
       If ChkDice(j).value = 0 Then
        lblTotal(j).Caption = ""
       Else
        lblTotal(j).Caption = Str(.GameStatus(2, j))    'getindividual scores
       End If
       Select Case j
       Case eOnes To eSixes
            LowerScore = LowerScore + .GameStatus(2, j)
        Case e3ofakind To eYahtzee
            UpperScore = UpperScore + .GameStatus(2, j)
       End Select
    Next j
    lblTotal(14).Caption = Str(.GameStatus(2, 14))       'extra Yahtzee total
    LblSubTotal1.Caption = Str(LowerScore)
    lblSubTotal2.Caption = Str(UpperScore + .GameStatus(2, 14))
    lblTotalScore.Caption = Str(LowerScore + UpperScore + .GameStatus(2, 14))
    If LowerScore > 63 Then
        lblBonus.Visible = True
    Else
        lblBonus.Visible = False
    End If
    'For i = 1 To 5
    '    chkHold(i).Enabled = .Holdstatus(i)
    '    DiceValue(i) = .diceFaceValues(i)
    'Next
    'Throws = .Throws
End With
ignore = False
End Sub

Private Sub PlayerControls(State As Boolean)
 Dim i As Integer
 For i = 1 To 13                    'disable all checkboxes
    If ChkDice(i).value <> 1 Then   'if not checked already, disable
      ChkDice(i).Enabled = State
    End If
 Next i
 For i = 1 To 11                    'disable all dice
    Die(i).Enabled = State
 Next i
 
End Sub

Private Sub GetRollResults()
'places value in global array RollResults(1..13)
    Dim i As Integer, j As Integer
    Dim TotalDiceValue As Integer
    Dim matches As Integer
    Dim match(1 To 6) As Integer
    Dim twoofakind As Boolean           'need for full house
    Dim threeofakind As Boolean
    Dim fourofakind As Boolean
    Dim fiveofakind As Boolean
    Dim SmallStraight As Boolean
    Dim LargeStraight As Boolean
    Dim FullHouse As Boolean
    Dim Yahtzee As Boolean
 
    '*********************************************************************
    For i = eOnes To eYahtzee         'clear any previous roll result, use enum values
      RollResults(i) = 0
    Next i
    For j = 1 To 6          'there are 6 numbers on five dice
        match(j) = 0        'clear matches counters
        For i = 1 To 5      'count the number of ones, twos, threes, etc
            If DiceValue(i) = j Then match(j) = match(j) + 1
    Next i, j
    
    'check for ones, twos, three, fours, fives and sixes
    For i = 1 To 5
      Select Case DiceValue(i)
        Case 1
            RollResults(eOnes) = RollResults(eOnes) + 1
        Case 2
            RollResults(eTwos) = RollResults(eTwos) + 2
        Case 3
            RollResults(eThrees) = RollResults(eThrees) + 3
        Case 4
            RollResults(eFours) = RollResults(eFours) + 4
        Case 5
            RollResults(eFives) = RollResults(eFives) + 5
        Case 6
           RollResults(eSixes) = RollResults(eSixes) + 6
      End Select
    Next i
    TotalDiceValue = 0
    For i = 1 To 5          'get total of all the dice
        TotalDiceValue = TotalDiceValue + DiceValue(i)
    Next i
    
    For i = 1 To 6                      'find how many of a kind
        Select Case match(i)
          Case 2
            twoofakind = True           'pair only used for full house
          Case 3
            threeofakind = True
            RollResults(e3ofakind) = TotalDiceValue
          Case 4
            threeofakind = True         '*** is this needed?
            RollResults(e3ofakind) = TotalDiceValue
            fourofakind = True
            RollResults(e4ofaKind) = TotalDiceValue
          Case 5                        'No more Yahtzees if checked and 0
            'threeofakind = True         '*** is this needed?
            'fourofakind = True          '*** is this needed?
            fiveofakind = True          'could take five of a kind as 3 or 4 of a kind
            If Not (ChkDice(eYahtzee).value = 1 And Val(ChkDice(eYahtzee).Caption) = 0) Then
                RollResults(eYahtzee) = 50
            End If
        End Select
    Next i
    
    'check for small straight = 30, will be 1,2,3,4 or 2,3,4,5 or 3,4,5,6
    If match(1) > 0 And match(2) > 0 And match(3) > 0 And match(4) > 0 Then SmallStraight = True
    If match(2) > 0 And match(3) > 0 And match(4) > 0 And match(5) > 0 Then SmallStraight = True
    If match(3) > 0 And match(4) > 0 And match(5) > 0 And match(6) > 0 Then SmallStraight = True
    If SmallStraight Then RollResults(eSmStraight) = 30
    
    'check for large straight = 40 will be 1,2,3,4,5 or 2,3,4,5,6
    If match(1) = 1 And match(2) = 1 And match(3) = 1 And match(4) = 1 And match(5) = 1 Then LargeStraight = True
    If match(2) = 1 And match(3) = 1 And match(4) = 1 And match(5) = 1 And match(6) = 1 Then LargeStraight = True
    If LargeStraight Then RollResults(eLgStraight) = 40
    
    'check for FullHouse = 25
       If twoofakind And threeofakind Then RollResults(eFullHouse) = 25
    
    'check for chance = value of dice
    RollResults(eChance) = TotalDiceValue
  '**************************************************************************
  For i = eOnes To eYahtzee
    ChkDice(i).ToolTipText = "Score Value = " + Str(RollResults(i))
    If RollResults(i) > 0 And ChkDice(i).Enabled = True Then
        ChkDice(i).ForeColor = vbBlue
    End If
  Next i
  ChkDice(eChance).Caption = "Chance =" + Str(RollResults(eChance))
End Sub

Private Sub ChangePlayerFocus()
Dim i As Integer

'point to next player
 If CurrentPlayer < NumberOfPlayers Then
        CurrentPlayer = CurrentPlayer + 1
    Else
        CurrentPlayer = 1
    End If
For i = 1 To 4
    PlayerRecord(i).IsCurrentPlayer = (i = CurrentPlayer)
Next i
'update screen with new player
 ShowPlayerRecord CurrentPlayer
'
lblCurrentPlayer.Caption = PlayerRecord(CurrentPlayer).PlayerName
End Sub


Private Sub SetColors()
'**********************************************************************
' This sub is not fully used
'   must be better way of changing form colors
'**********************************************************************
Dim i As Integer

Dim StdMainBackColor As Long
Dim StdMainForeColor As Long
Dim StdMainTitleColor As Long
Dim StdBackcolor2nd As Long
Dim StdForeColor2nd As Long
Dim StdGameOverBackColor As Long
Dim stdGameOverForeColor As Long
Dim stdChkBoxBackcolor As Long
Dim stdChkBoxForeColor As Long

Dim AltMainBackColor As Long
Dim AltMainForeColor As Long
Dim AltMainTitleColor As Long
Dim Alt2ndBackcolor As Long
Dim Alt2ndForeColor As Long
Dim AltGameOverBackColor As Long
Dim AltGameOverForeColor As Long
Dim AltChkBoxBackcolor As Long
Dim AltChkBoxForeColor As Long

'These vars are dim'ed in the declaration section
 BackColorMain = RGB(197, 186, 163)      'tan
 ForeColorMain = RGB(0, 0, 0)            'black
 'BackColorMain = RGB(255, 255, 255)    'tan
 'ForeColorMain = RGB(0, 0, 0)          'black
 
 TitleColor = RGB(0, 0, 175)             'dark blue
 Backcolor2nd = RGB(179, 204, 187)       'light green
 ForeColor2nd = RGB(172, 250, 131)       'pale green
 GameOverBackColor = RGB(179, 204, 187)  'light green
 GameOverForeColor = RGB(172, 250, 131)  'light green
 ChkBoxBackcolor = RGB(168, 155, 128)    'dark brown
 ChkBoxForeColor = RGB(255, 0, 0)        'red

'Dim BackColorMain As Long
'Dim ForeColorMain As Long
'Dim TitleColor As Long
'Dim Backcolor2nd As Long
'Dim ForeColor2nd As Long
'Dim GameOverBackColor As Long
'Dim GameOverForeColor As Long
'Dim ChkBoxBackcolor As Long
'Dim ChkBoxForeColor As Long
 
For i = 1 To 5
    chkHold(i).BackColor = ChkBoxBackcolor
    chkHold(i).ForeColor = ChkBoxForeColor
Next i
For i = 0 To 3
    lblPlayerScore(i).BackColor = BackColorMain
    lblPlayerScore(i).ForeColor = ForeColorMain
    FraPlayer(i).ForeColor = ForeColorMain
    FraPlayer(i).BackColor = BackColorMain
Next i
For i = 1 To 13
    ChkDice(i).ForeColor = ForeColorMain
    ChkDice(i).BackColor = BackColorMain
    lblTotal(i).ForeColor = ForeColorMain
    lblTotal(i).BackColor = BackColorMain
Next i
 
 lblTotal(14).ForeColor = ForeColorMain
 lblTotal(14).BackColor = BackColorMain
 FrmYahtzeeForm.BackColor = BackColorMain
 FrmYahtzeeForm.ForeColor = ForeColorMain
 FraHoldButtons.ForeColor = ForeColorMain
 FraHoldButtons.BackColor = BackColorMain
 lblTotalScore.BackColor = BackColorMain
 lblTotalScore.ForeColor = ForeColorMain
 cmdRoll.BackColor = BackColorMain
 Label15.BackColor = BackColorMain
 Label15.ForeColor = ForeColorMain
 Label2.BackColor = BackColorMain
 Label2.ForeColor = ForeColorMain
 LblSubTotal1.BackColor = BackColorMain
 LblSubTotal1.ForeColor = ForeColorMain
 lblSubTotal2.BackColor = BackColorMain
 lblSubTotal2.ForeColor = ForeColorMain
 lblYahtzeeBonus.ForeColor = ForeColorMain
 lblYahtzeeBonus.BackColor = BackColorMain
 fraDiceColor.ForeColor = ForeColorMain
 fraDiceColor.BackColor = BackColorMain
 fraScore.ForeColor = ForeColorMain
 fraScore.BackColor = BackColorMain
 
 fraSound.ForeColor = ForeColorMain
 fraSound.BackColor = BackColorMain
 fraHighScore.ForeColor = ForeColorMain
 fraHighScore.BackColor = BackColorMain
 lblHighScoreAmount.ForeColor = ForeColorMain
 lblHighScoreAmount.BackColor = BackColorMain
 lblHighScoreName.ForeColor = ForeColorMain
 lblHighScoreName.BackColor = BackColorMain
 fraCurrentPlayer.ForeColor = TitleColor
 fraCurrentPlayer.BackColor = BackColorMain
 lblCurrentPlayer.ForeColor = TitleColor
 lblCurrentPlayer.BackColor = BackColorMain
 fraPlayerScore.ForeColor = ForeColorMain
 fraPlayerScore.BackColor = BackColorMain
 Label1.ForeColor = TitleColor
 Label1.BackColor = BackColorMain
 cmdNew.BackColor = BackColorMain
 cmdExit.BackColor = BackColorMain
 Text1.BackColor = GameOverBackColor
 Text1.ForeColor = GameOverForeColor
 OptWhiteDie.ForeColor = ForeColorMain
 OptWhiteDie.BackColor = BackColorMain
 optGreenDie.ForeColor = ForeColorMain
 optGreenDie.BackColor = BackColorMain
 optRedDie.ForeColor = ForeColorMain
 optRedDie.BackColor = BackColorMain
 optSoundOn.ForeColor = ForeColorMain
 optSoundOn.BackColor = BackColorMain
 optSoundOff.ForeColor = ForeColorMain
 optSoundOff.BackColor = BackColorMain
 End Sub

Private Sub ShowHighScorePerson()
With HighScoreRecord(1)
 lblHighScoreName = .PlayerName
 lblHighScoreAmount = Str(.PlayerScore) + " on " + Format(.ScoreDate, "mmm d, yyyy")
End With
End Sub

Public Sub RemovePlayer(Index As Integer)

frmPlayerChanges.txtPlayer = FraPlayer(Index - 1)     'prepare form
frmPlayerChanges.Show vbModal
PlayerRecord(Index).PlayerName = frmPlayerChanges.txtPlayer   'change record
UpdatePlayerNames

End Sub

Private Sub ShowPlayerStatus(Player As Integer)
Dim i As Integer    'change color of background to show player scores
DoneOnce = False    'to prevent multiple redisplays
    FraPlayer(Player).BackColor = Backcolor2nd            '&HBBCCB3 'light greenn
    lblPlayerScore(Player).BackColor = FraPlayer(Player).BackColor
    lblTotalScore.BackColor = FraPlayer(Player).BackColor
    ShowPlayerRecord (Player + 1)
    For i = 1 To 14
        lblTotal(i).BackColor = FraPlayer(Player).BackColor
    Next i
    For i = 0 To 3
        If i <> Player Then
            FraPlayer(i).BackColor = fraPlayerScore.BackColor
            lblPlayerScore(i).BackColor = fraPlayerScore.BackColor
        End If
    Next i
End Sub

Private Sub SaveGame()
Dim i As Integer
 For i = 1 To 4
           If i = CurrentPlayer Then
             PlayerRecord(CurrentPlayer).IsCurrentPlayer = True
           Else
             PlayerRecord(CurrentPlayer).IsCurrentPlayer = False
           End If
           Put #1, i, PlayerRecord(i)  'save file
       Next i
End Sub


Private Sub ExitGame()
Dim i As Integer
For i = 1 To 10                     'save high scores
     Put #2, i, HighScoreRecord(i)
Next i
Close #1
Close #2
Unload frmYahtzeeHS_Help
Unload FrmHighScores
Unload frmOldOrNewGame
Unload frmPlayerChanges
Unload frmPlayerNames
Unload frmSplash
Unload FrmYahtzeeForm

End Sub
