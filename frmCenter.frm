VERSION 5.00
Begin VB.Form frmCenter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   Icon            =   "frmCenter.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmCenter.frx":00EA
   ScaleHeight     =   405
   ScaleWidth      =   1185
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmCenter"
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
    FormOnTop frmCenter
    frmCenter.AutoRedraw = True
    StretchBlt frmCenter.hDC, 0, 0, frmCenter.ScaleWidth, frmCenter.ScaleHeight, frmCenter.hDC, 0, 0, 1, frmCenter.ScaleHeight, vbSrcCopy
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frmLeft.Left = frmLeft.Left + (X - OldX)
        frmLeft.Top = frmLeft.Top + (Y - OldY)
        frmCenter.Left = frmLeft.Left + frmLeft.Width
        frmCenter.Top = frmLeft.Top + 535
        frmRight.Left = frmLeft.Left + frmLeft.Width + frmCenter.Width
        frmRight.Top = frmLeft.Top + 47
    End If
End Sub
