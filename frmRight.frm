VERSION 5.00
Begin VB.Form frmRight 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   Icon            =   "frmRight.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmRight.frx":00EA
   ScaleHeight     =   1530
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdColor 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   240
   End
   Begin VB.CommandButton cmdRota 
      BackColor       =   &H00FF8080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   585
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   495
      Width           =   225
   End
   Begin VB.CommandButton cmdHRRight 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1125
      TabIndex        =   5
      Top             =   735
      Width           =   300
   End
   Begin VB.CommandButton cmdHRLeft 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1125
      TabIndex        =   4
      Top             =   495
      Width           =   300
   End
   Begin VB.CommandButton cmdHLRight 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   705
      Width           =   270
   End
   Begin VB.CommandButton cmdHLLeft 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   510
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   330
      TabIndex        =   1
      Top             =   720
      Width           =   225
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   225
      X2              =   225
      Y1              =   885
      Y2              =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   210
      X2              =   210
      Y1              =   0
      Y2              =   465
   End
End
Attribute VB_Name = "frmRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'**                              Screen Calipers
'**                               Version 4.4.0
'**                               By Ken Foster
'**                               January  2007
'**                     Freeware--- no copyrights claimed
'*******************************************************************
'Screen Zoom by Rocky Clark (Kath-Rock Software)

'Right mouse click on Right Calipier to show/hide the zoom window
'C = color picker for calipier edges
'P/T selects between pixels and twips
'^ or > selects vertical or horizontal
'X is the exit button
'Left side arrow buttons move left side of the calipiers
'Right side arrow buttons move the right side of the calipiers
'=============================================

Option Explicit

Dim OldX As Integer
Dim value As Long

Private Sub Form_Load()
    MakeTransparent Me
    FormOnTop Me
End Sub

Private Sub Form_Activate()
    cValue                           'show display on startup
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    If Button = 2 Then frmMain.Visible = Not frmMain.Visible
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmRight.Left = frmRight.Left + (X - OldX)
        frmRight.Top = frmLeft.Top + 47
        If frmRight.Left < frmLeft.Left + 74 Then
            LockWindowUpdate frmRight.hwnd
            frmRight.Left = frmLeft.Left + 74
            LockWindowUpdate 0&
        End If
        If frmRight.Left < frmCenter.Left Then
            frmCenter.Width = 20
            Exit Sub
        End If
        frmCenter.Width = frmRight.Left - (frmLeft.Left + frmLeft.Width)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'just to make sure cals are still on top
    FormOnTop frmLeft
    FormOnTop frmCenter
    FormOnTop Me
    cValue                               'make sure value display is still visible
End Sub

Private Sub cmdColor_Click()
   frmColor.Top = frmRight.Top + frmRight.Height
   frmColor.Left = frmRight.Left + frmRight.Width / 2
   frmColor.Visible = Not frmColor.Visible
   cmdRota.SetFocus
End Sub

Private Sub cmdRota_Click()
    frmLeft.Visible = False
    frmCenter.Visible = False
    frmRight.Visible = False
    frmTop.Visible = True
    frmMid.Visible = True
    frmBot.Visible = True
    frmBot.cmdRot.SetFocus
    frmColor.Visible = False
End Sub

Private Sub Command1_Click()
    Smode = Not Smode          'show pixels or twips, see cValue
    cValue                     'update display
    frmBot.cValue
    cmdRota.SetFocus
End Sub

Private Sub cmdHLLeft_Click()            'horizontal left,left button
    frmLeft.Left = frmLeft.Left - 15     '15 twips = 1 pixel
    frmCenter.Left = frmCenter.Left - 15
    frmRight.Left = frmRight.Left - 15
End Sub

Private Sub cmdHLRight_Click()           'horizontal left, right button
    frmLeft.Left = frmLeft.Left + 15
    frmCenter.Left = frmCenter.Left + 15
    frmRight.Left = frmRight.Left + 15
End Sub

Private Sub cmdHRLeft_Click()            'horizontal right, left button
    frmRight.Left = frmRight.Left - 15
    If frmRight.Left <= frmLeft.Left + 74 Then
        frmRight.Left = frmLeft.Left + 74
        cValue
        Exit Sub
    Else
        If frmRight.Left <= frmLeft.Left + frmLeft.Width Then GoTo here
        frmCenter.Width = frmRight.Left - (frmLeft.Left + frmLeft.Width)
here:
        cValue                           'update value display
    End If
End Sub

Private Sub cmdHRRight_Click()           'horizontal right, right button
    frmRight.Left = frmRight.Left + 15
    cValue
    If frmRight.Left < frmLeft.Left + frmLeft.Width Then Exit Sub
    frmCenter.Width = frmRight.Left - (frmLeft.Left + frmLeft.Width)
End Sub

Private Sub cmdExit_Click()
    Unload frmCenter
    Unload frmMid
    Unload frmLeft
    Unload frmTop
    Unload frmBot
    Unload frmColor
    Unload frmMain
    Unload Me
End Sub

Public Sub cValue()
    If Smode = False Then                    '---pixels---
    value = ((frmRight.Left - (Int(frmLeft.Left + 74))) \ 15)  '15 twips = 1 pixel
    Command1.Caption = "P"
    frmRight.Cls
    frmRight.ForeColor = vbWhite
    frmRight.CurrentX = 650
    frmRight.CurrentY = 500
    frmRight.Print value                  'update pixel display
Else                                     '---twips---
    value = (frmRight.Left - (Int(frmLeft.Left + 74)))
    Command1.Caption = "T"
    frmRight.Cls
    frmRight.ForeColor = vbWhite
    frmRight.CurrentX = 550
    frmRight.CurrentY = 500
    frmRight.Print value                  'update pixel display
End If
End Sub
