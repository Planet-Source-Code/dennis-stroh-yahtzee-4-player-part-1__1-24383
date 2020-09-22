VERSION 5.00
Begin VB.Form frmOldOrNewGame 
   BackColor       =   &H00A3BAC5&
   BorderStyle     =   0  'None
   ClientHeight    =   6630
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H00809BA8&
      Caption         =   "Start A New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton cmdRestoreOldGame 
      BackColor       =   &H00809BA8&
      Caption         =   "Restore Old Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A3BAC5&
      Height          =   3855
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00BBCCB3&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmOldOrNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNewGame_Click()
NewGame = True
Me.Hide
End Sub

Private Sub cmdRestoreOldGame_Click()
NewGame = False
Me.Hide
End Sub
