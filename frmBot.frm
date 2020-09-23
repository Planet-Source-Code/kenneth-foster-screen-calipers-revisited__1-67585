VERSION 5.00
Begin VB.Form frmBot 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmBot.frx":00EA
   ScaleHeight     =   1185
   ScaleWidth      =   1545
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
      Left            =   810
      TabIndex        =   7
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton cmdVBDown 
      Caption         =   "6"
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
      Left            =   825
      TabIndex        =   6
      Top             =   960
      Width           =   180
   End
   Begin VB.CommandButton cmdVBUp 
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
      Height          =   195
      Left            =   585
      TabIndex        =   5
      Top             =   960
      Width           =   180
   End
   Begin VB.CommandButton cmdVTUp 
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
      Left            =   615
      TabIndex        =   4
      Top             =   15
      Width           =   180
   End
   Begin VB.CommandButton cmdVTDown 
      Caption         =   "6"
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
      Left            =   825
      TabIndex        =   3
      Top             =   15
      Width           =   180
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
      Height          =   255
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   450
      Width           =   225
   End
   Begin VB.CommandButton cmdRot 
      BackColor       =   &H00FF8080&
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
      Height          =   255
      Left            =   795
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   450
      Width           =   225
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
      Left            =   570
      TabIndex        =   1
      Top             =   720
      Width           =   225
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   585
      X2              =   -15
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1020
      X2              =   1500
      Y1              =   210
      Y2              =   210
   End
End
Attribute VB_Name = "frmBot"
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

'=============================================
Option Explicit

Dim OldY As Integer
Dim value As Long

Private Sub Form_Load()
    MakeTransparent Me
    FormOnTop Me
End Sub

Private Sub Form_Activate()
    cValue
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmBot.Top = frmBot.Top + (Y - OldY)
        frmBot.Left = frmTop.Left + 74
        If frmBot.Top < frmTop.Top + 74 Then
            LockWindowUpdate frmBot.hwnd
            frmBot.Top = frmTop.Top + 74
            LockWindowUpdate 0&
        End If
        If frmBot.Top < frmMid.Top Then
            frmMid.Height = 20
            Exit Sub
        End If
        frmMid.Height = frmBot.Top - (frmTop.Top + frmTop.Height)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'just to make sure cals are still on top
    FormOnTop frmTop
    FormOnTop frmMid
    FormOnTop Me
    cValue
End Sub

Private Sub Command1_Click()
   Smode = Not Smode
   cValue                   'update display
   cmdRot.SetFocus
End Sub

Private Sub cmdColor_Click()
   frmColor.Top = frmBot.Top + frmBot.Height
   frmColor.Left = frmBot.Left + frmBot.Width
   frmColor.Visible = Not frmColor.Visible
   cmdRot.SetFocus
End Sub

Private Sub cmdVBDown_Click()
    frmBot.Top = frmBot.Top + 15
    cValue
    If frmBot.Top < frmTop.Top + frmTop.Height Then Exit Sub  'no need to draw center bar so exit
    frmMid.Height = frmBot.Top - (frmTop.Top + frmTop.Height)
End Sub

Private Sub cmdVBUp_Click()
    frmBot.Top = frmBot.Top - 15
    If frmBot.Top <= frmTop.Top + 74 Then  'if less than home position exit
       frmBot.Top = frmTop.Top + 74
       cValue
    Exit Sub
Else
    If frmBot.Top <= frmTop.Top + frmTop.Height Then GoTo here
    frmMid.Height = frmBot.Top - (frmTop.Top + frmTop.Height)
here:
    cValue                      'update display
End If
End Sub

Private Sub cmdVTDown_Click()
    frmTop.Top = frmTop.Top + 15    '15 twips = 1 pixel
    frmMid.Top = frmMid.Top + 15
    frmBot.Top = frmBot.Top + 15
End Sub

Private Sub cmdVTUp_Click()
    frmTop.Top = frmTop.Top - 15    '15 twips = 1 pixel
    frmMid.Top = frmMid.Top - 15
    frmBot.Top = frmBot.Top - 15
End Sub

Private Sub cmdRot_Click()
    frmTop.Visible = False
    frmMid.Visible = False
    frmBot.Visible = False
    frmLeft.Visible = True
    frmCenter.Visible = True
    frmRight.Visible = True
    frmColor.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload frmCenter
    Unload frmMid
    Unload frmLeft
    Unload frmTop
    Unload frmRight
    Unload frmColor
    Unload frmMain
    Unload Me
End Sub

Public Sub cValue()
    If Smode = False Then
        value = ((frmBot.Top - (Int(frmTop.Top + 74))) \ 15)   '15 twips = 1 pixel
        Command1.Caption = "P"
        frmBot.Cls
        frmBot.ForeColor = vbWhite
        frmBot.CurrentX = 580
        frmBot.CurrentY = 230
        frmBot.Print value                   'update display
    Else
        value = (frmBot.Top - (Int(frmTop.Top + 74)))
        Command1.Caption = "T"
        frmBot.Cls
        frmBot.ForeColor = vbWhite
        frmBot.CurrentX = 480
        frmBot.CurrentY = 230
        frmBot.Print value                   'update display
    End If
End Sub
