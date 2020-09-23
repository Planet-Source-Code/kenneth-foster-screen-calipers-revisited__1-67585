VERSION 5.00
Begin VB.Form frmLeft 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   Icon            =   "frmLeft.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLeft.frx":00EA
   ScaleHeight     =   1575
   ScaleWidth      =   900
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   285
      X2              =   285
      Y1              =   945
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   300
      X2              =   300
      Y1              =   45
      Y2              =   555
   End
End
Attribute VB_Name = "frmLeft"
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
' Screen zoom was written by Rocky Clark (Kath-Rock Software)
' see frmRight for more information
'=============================================
Option Explicit

Dim OldX As Integer
Dim OldY As Integer

Private Sub Form_Load()
    'values and positions for all forms are setup here at startup
    FormOnTop frmLeft
    MakeTransparent Me
    frmLeft.Show
    frmCenter.Show
    frmRight.Show
    'horizontal
    frmLeft.Top = 5000
    frmLeft.Left = 7000
    frmCenter.Left = frmLeft.Left + frmLeft.Width
    frmCenter.Top = frmLeft.Top + 535
    frmCenter.Width = 200
    frmRight.Left = frmLeft.Left + 74     'set to home position (0) for horz
    frmRight.Top = frmLeft.Top + 47
    'vertical
    frmTop.Left = 5000
    frmTop.Top = 5000
    frmMid.Top = frmTop.Top + frmTop.Height
    frmMid.Left = frmTop.Left + 654
    frmMid.Height = 200
    frmBot.Top = frmTop.Top + 74          'set home position (0)for vert
    frmBot.Left = frmTop.Left + 58
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
    FormOnTop frmCenter
    FormOnTop frmRight
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmLeft.Left = frmLeft.Left + (X - OldX)
        frmLeft.Top = frmLeft.Top + (Y - OldY)
        frmRight.Left = frmRight.Left + (X - OldX)
        frmRight.Top = frmLeft.Top + 47
        frmCenter.Left = frmLeft.Left + frmLeft.Width
        frmCenter.Top = frmLeft.Top + 535
    End If
End Sub
