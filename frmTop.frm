VERSION 5.00
Begin VB.Form frmTop 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   Icon            =   "frmTop.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmTop.frx":00EA
   ScaleHeight     =   915
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   660
      X2              =   0
      Y1              =   285
      Y2              =   285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1050
      X2              =   1560
      Y1              =   300
      Y2              =   300
   End
End
Attribute VB_Name = "frmTop"
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

'see frmLeft for all form startup properties
Dim OldX As Integer
Dim OldY As Integer

Private Sub Form_Load()
    FormOnTop frmTop
    MakeTransparent Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
    FormOnTop frmMid          'just to make sure everything is still there
    FormOnTop frmBot
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmTop.Top = frmTop.Top + (Y - OldY)
        frmTop.Left = frmTop.Left + (X - OldX)
        frmBot.Top = frmBot.Top + (Y - OldY)
        frmBot.Left = frmTop.Left + 60
        frmMid.Top = frmTop.Top + frmTop.Height
        frmMid.Left = frmTop.Left + 654
    End If
End Sub
