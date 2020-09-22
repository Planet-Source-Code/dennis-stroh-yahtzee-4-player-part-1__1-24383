VERSION 5.00
Begin VB.Form frmPlayerNames 
   BackColor       =   &H00BBCCB3&
   BorderStyle     =   0  'None
   Caption         =   "Enter Player Names"
   ClientHeight    =   5310
   ClientLeft      =   3270
   ClientTop       =   3210
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00BBCCB3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   1
      ToolTipText     =   "Enter player's name"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C000&
      Caption         =   "&Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00BBCCB3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "Enter player's name"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00BBCCB3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Enter player's name"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00BBCCB3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Enter player's name"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00BBCCB3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter 1 to 4 Player Names"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00BBCCB3&
      Caption         =   "4th Player's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00BBCCB3&
      Caption         =   "3rd Player's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00BBCCB3&
      Caption         =   "2nd Player's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00BBCCB3&
      Caption         =   "1st Player's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "frmPlayerNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Dim i As Integer

For i = 1 To 4
   PlayerRecord(i).PlayerName = txtPlayerName(i - 1)
Next i

UpdatePlayerNames

If NumberOfPlayers = 0 Then
    MsgBox "No player names have been entered.", vbOKOnly
    Exit Sub
End If
frmPlayerNames.Hide     'hide & return to caller
End Sub

