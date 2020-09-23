VERSION 5.00
Begin VB.Form frmMid 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   ControlBox      =   0   'False
   Icon            =   "frmMid.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMid.frx":00EA
   ScaleHeight     =   1065
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmMid"
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

Dim OldX As Integer
Dim OldY As Integer

Private Sub Form_Load()
    FormOnTop frmMid
    frmMid.AutoRedraw = True
    StretchBlt frmMid.hDC, 0, 0, frmMid.ScaleWidth, frmMid.ScaleHeight, frmMid.hDC, 0, 0, frmMid.ScaleWidth, 1, vbSrcCopy
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmTop.Top = frmTop.Top + (Y - OldY)
        frmTop.Left = frmTop.Left + (X - OldX)
        frmMid.Top = frmTop.Top + frmTop.Height
        frmMid.Left = frmTop.Left + 654
        frmBot.Top = frmTop.Top + frmTop.Height + frmMid.Height
        frmBot.Left = frmTop.Left + 60
    End If
End Sub
