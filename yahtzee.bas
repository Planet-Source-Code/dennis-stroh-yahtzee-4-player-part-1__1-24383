Attribute VB_Name = "Yahtzee"
' ******************************************************************
'*  Name:    djs Windows Yahtzee                                    *
'*  Author:  Dennis Stroh                                           *
'*  Email:   djs314@hotmail.com                                     *
'*  Date:    March 30, 2001                                         *
'*                                                                  *
'*  Written in Visual Basic 5.0 SP3                                 *
'*                                                                  *
'*  Source Copyright - Dennis J. Stroh                              *
'*  Yahtzee is owned by Milton Bradley                              *
'*  Yahtzee is a Registered Trademark of the Milton Bradley Company *
'*                                                                  *
'*  Notice:  This software was written as a learning experience     *
'*           for educational purposes.  The source code is made     *
'*           available only for personal (non commercial) use.      *
'*           The source code in whole or in part may not be sold    *
'*           or used in any commercial way.                         *
'*                                                                  *
'*  Please email me any updates you make                            *
'*  Please keep this header with any source you distribute          *
'*                                                                  *
' ******************************************************************


'********************************************************************
'The  API bitblt call for the dice was written by
'Timothy J. Mitchell, CompuServe 71461,303
'Which was inspired by the VISUAL BASIC WORKSHOP column in the May 1994
'issue of Britain's PC PLUS magazine (pg 296)
'
'To use the Dice, place the dice bitmap (dice.bmp) in a PictureBox on the
'form. You can hide it by pulling the side of the form over it (see
'Roller1.FRM) in the example. Place more PictureBoxes on the form, one for
'each of the die you want to show.
'
'Display the dice by calling the following subroutine:
'
'ShowDie (Source, Target, Digit, Color)
'
'   where   Source is the PictureBox containing the Dice Bitmap (dice.bmp)
'           Target is the PictureBox that will contain the displayed die
'           Digit is an integer between 1 and 6 referring to the die face
'           Color is an integer between 1 and 3 referring to the die color
'
' Note: Make sure you set the AUTOREDRAW property of the hidden
' PictureBox(es) to TRUE or the graphics repaints won't work correctly.
'*************************************************************************
'
'Besides the bitblt call, I also used Timothy Mitchell's ShowDie routine
'from his Roller11 program.
'All other coding was done by me.
'**************************************************************************
Option Explicit
Public Const ENABLE As Boolean = True
Public Const DISABLE As Boolean = False

Public NumberOfPlayers As Integer                  'number of players
Public NewGame As Boolean
Public DiceValue(1 To 5) As Integer
Public GameTotal As Integer

Global Const SRCCOPY = &HCC0020     'const for the bitblt API function
Global res As Integer               'Holder for Funtion BitBlt result

'Declare API call BitBlt in gdi32.dll
'
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Declare API call sndPlaySound in winmm.dll
'
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public soundflag As Integer
Global SoundDirectory As String
'
'Set up for Game File
'
Public Type PRecord                          ' Define user-defined type.
    PlayerName As String * 10
    PlayerScore As Integer
    GameStatus(1 To 2, 1 To 14) As Integer  'keeps checkmark status and score of each area
    Holdstatus(1 To 5) As Integer           'keeps 5 hold button status, "chkhold.value"
    diceFaceValues(1 To 5) As Integer       'keeps the current dice face values
    Throws As Integer                       'number of throws taken
    IsCurrentPlayer As Boolean              'flag for current player
End Type
Public Type HSRecord
    PlayerName As String * 10
    PlayerScore As Integer
    ScoreDate As Date
End Type
Public PlayerRecord(1 To 4) As PRecord         'holds stats on up to four players
Public HighScoreRecord(1 To 10) As HSRecord    'holds 10 top scores

Public Sub PlaySound(strFileName As String)
If FrmYahtzeeForm.optSoundOn Then
    soundflag = 1   'seems to be 0,1,2
                    '0 plays sound, controls are inactive, remembers mouse clicks
                    '1 plays sound while form is active, can be interrupted
                    '2 plays sound, controls are inactive, remembers mouse clicks
    sndPlaySound SoundDirectory + strFileName, soundflag
End If
End Sub

Sub ShowDie(Source As PictureBox, Target As PictureBox, DieNumber As Integer, Digit As Integer, Color As Integer)
    '
    'this subroutine clips a die face from the bitmap in SOURCE and puts it in TARGET

    '       Target.hdc = destination object, hdc = windows device context handel
    '                0 = x destination coordinates
    '                0 = y destination coordinates
    '               32 = width
    '               32 = height
    '       Source.hDC = source object
    '       32 * color = source x
    '       32 * digit = source y
    '          SRCCOPY = dwrop???  public const &hCC0020 '(dword) dest=source
   
    res = BitBlt(Target.hDC, 0, 0, 32, 32, Source.hDC, 32 * (Color - 1), 32 * (Digit - 1), SRCCOPY)
    Target.Refresh
   
End Sub

Sub RollDice(SourceDicePic As PictureBox, Die As PictureBox, DieNumber As Integer, Color As Integer)

Dim i As Integer

'Rolls a set of 5 dice by randomly determining the result and calling Sub ShowDie
'If the die is not being held then roll.
For i = 1 To 5
  If FrmYahtzeeForm.chkHold(DieNumber).value = 0 Then   'button not checked
    DiceValue(DieNumber) = GetRandom
    ShowDie SourceDicePic, Die, i, DiceValue(DieNumber), Color 'update the contents of each
  End If
Next i
End Sub
Private Function GetRandom() As Integer
    Randomize               'get a random number between 1 and 6 for each die
    GetRandom = Int(6 * Rnd + 1)
End Function
Public Sub delay(s As Single)
Dim TimeDelay As Single
    TimeDelay = Timer + s               'delay s seconods
    Do While TimeDelay > Timer
    DoEvents
    Loop
End Sub

Public Sub UpdatePlayerNames()
Dim i As Integer, j As Integer
NumberOfPlayers = 0
For i = 1 To 4                          'count number of players
    With PlayerRecord(i)
        If .PlayerName <> Space$(Len(.PlayerName)) Then
            Inc NumberOfPlayers
        End If
    End With
Next i
If NumberOfPlayers = 0 Then Exit Sub        'no players to work with

'move any blank names to the last positions
i = 1
Do While i < 4 And i <= NumberOfPlayers
 With PlayerRecord(i)
   If .PlayerName = Space$(Len(.PlayerName)) Then
      For j = i To 3
           PlayerRecord(j) = PlayerRecord(j + 1)   'move next record down
           PlayerRecord(j + 1).PlayerName = ""      'blank record up
      Next j
    End If
    If .PlayerName <> Space$(Len(.PlayerName)) Then Inc i
 End With
Loop
   
For i = 1 To 4                  'adjust main form & names form
    With PlayerRecord(i)
        FrmYahtzeeForm.FraPlayer(i - 1).Caption = Trim(.PlayerName)
        FrmYahtzeeForm.mnuRemove(i).Caption = Trim(.PlayerName)
        FrmYahtzeeForm.mnuEdit(i).Caption = Trim(.PlayerName)
        frmPlayerNames.txtPlayerName(i - 1) = Trim(.PlayerName)
        If FrmYahtzeeForm.FraPlayer(i - 1).Caption = "" Then
            FrmYahtzeeForm.FraPlayer(i - 1).Visible = False
            FrmYahtzeeForm.mnuRemove(i).Visible = False
            FrmYahtzeeForm.mnuEdit(i).Visible = False
        Else
            FrmYahtzeeForm.mnuRemove(i).Visible = True
            FrmYahtzeeForm.mnuEdit(i).Visible = True
            FrmYahtzeeForm.FraPlayer(i - 1).Visible = True
        End If
    End With
Next
End Sub

Public Sub Inc(value)
    value = value + 1
End Sub
