VERSION 5.00
Begin VB.Form frmColor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   1005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   8
      Left            =   420
      TabIndex        =   9
      Top             =   750
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   7
      Left            =   705
      TabIndex        =   8
      Top             =   420
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   705
      TabIndex        =   7
      Top             =   105
      Width           =   210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right Mouse click on Right Cals to Show/Hide Zoom."
      Height          =   990
      Left            =   30
      TabIndex        =   6
      Top             =   1125
      Width           =   945
   End
   Begin VB.Shape Shape1 
      Height          =   1080
      Left            =   0
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   705
      TabIndex        =   5
      Top             =   750
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   420
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   105
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   750
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   420
      Width           =   210
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   210
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Integer
Dim OldY As Integer

Private Sub Form_DblClick()
    frmColor.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmColor.Left = frmColor.Left + (X - OldX)
        frmColor.Top = frmColor.Top + (Y - OldY)
     End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblColor_Click(Index As Integer)
   frmTop.Line1.BorderColor = lblColor(Index).BackColor
   frmTop.Line2.BorderColor = lblColor(Index).BackColor
   frmBot.Line1.BorderColor = lblColor(Index).BackColor
   frmBot.Line2.BorderColor = lblColor(Index).BackColor
   frmLeft.Line1.BorderColor = lblColor(Index).BackColor
   frmLeft.Line2.BorderColor = lblColor(Index).BackColor
   frmRight.Line1.BorderColor = lblColor(Index).BackColor
   frmRight.Line2.BorderColor = lblColor(Index).BackColor
   frmColor.Visible = False
End Sub

