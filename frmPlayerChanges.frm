VERSION 5.00
Begin VB.Form frmPlayerChanges 
   BackColor       =   &H00BBCCB3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Player Changes"
   ClientHeight    =   2370
   ClientLeft      =   4155
   ClientTop       =   3990
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5190
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   3368
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2048
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdChanges 
      Caption         =   "C&hange"
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
      Left            =   728
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1328
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmPlayerChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dummy As String

Private Sub cmdCancel_Click()
txtPlayer = dummy
Me.Hide
End Sub

Private Sub cmdChanges_Click()
Me.Hide
End Sub

Private Sub cmdDelete_Click()
Dim result As Integer
result = MsgBox("Remove Player " + txtPlayer + "?", vbYesNoCancel)
Select Case result
Case vbCancel
    Exit Sub
Case vbNo
    Exit Sub
Case vbYes
    'remove the player
    txtPlayer = ""
    NumberOfPlayers = 0                     'force recount
  
End Select
txtPlayer = ""
Me.Hide
End Sub

Private Sub Form_Activate()
dummy = txtPlayer
End Sub

