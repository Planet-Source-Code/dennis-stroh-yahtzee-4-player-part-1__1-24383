VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00A3BAC5&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   3075
   ClientTop       =   3435
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00A3BAC5&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   3000
         Top             =   3240
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3BAC5&
         Caption         =   "Copyright 2001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3BAC5&
         Caption         =   "Stroh Electronic Systems"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00A3BAC5&
         Caption         =   "Warning:  This program may not be sold in any form."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   4215
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00A3BAC5&
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5415
         TabIndex        =   5
         Top             =   2760
         Width           =   1470
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00A3BAC5&
         Caption         =   "Windows 95 && 98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4260
         TabIndex        =   6
         Top             =   2340
         Width           =   2595
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   1080
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   6840
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3BAC5&
         Caption         =   "LicenseTo: Freeware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00A3BAC5&
         Caption         =   "djsWindows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    ExitSplash
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ExitSplash
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    Show
    Timer1.Interval = 5000          'set timer interval
    DoEvents
End Sub

Private Sub Frame1_Click()
    ExitSplash
End Sub

Public Sub ExitSplash()
    Unload Me
    'Hide
    Load FrmYahtzeeForm
    'FrmYahtzeeForm.StartUpPosition = vbStartUpScreen  'center on the screen
    FrmYahtzeeForm.Show   ' Show form.
End Sub

Private Sub lblProductName_Click()
    ExitSplash
   
End Sub

Private Sub Timer1_Timer()
    ExitSplash            'after timer interval exit
End Sub
