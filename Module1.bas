Attribute VB_Name = "Module1"
Option Explicit
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    Public Declare Function StretchBlt Lib "GDI32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    
    Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
    Public Const RGN_XOR = 3
    Public Smode As Boolean

Public Sub FormOnTop(Frm As Form)
    Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Function MakeTransparent(ByRef Frm As Form) As Long
    Dim rgnMain As Long, rgnPixel As Long, bmpMain As Long, dcMain As Long
    Dim Width As Long, Height As Long, X As Long, Y As Long
    Dim ScaleSize As Long, RGBColor As Long
    Dim TransparentColor As Long
    
    ScaleSize& = Frm.ScaleMode
    Frm.ScaleMode = 3
    Frm.BorderStyle = 0
    Width& = Frm.ScaleX(Frm.Picture.Width, vbHimetric, vbPixels)
    Height& = Frm.ScaleY(Frm.Picture.Height, vbHimetric, vbPixels)
    Frm.Width = Width& * Screen.TwipsPerPixelX
    Frm.Height = Height& * Screen.TwipsPerPixelY
    rgnMain& = CreateRectRgn(0&, 0&, Width&, Height&)
    dcMain& = CreateCompatibleDC(Frm.hDC)
    bmpMain& = SelectObject(dcMain&, Frm.Picture.Handle)
    TransparentColor = GetPixel(dcMain&, 0, 0)     'set transparent color to upper left pixel
    For Y& = 0& To Height&
        For X& = 0& To Width&
            RGBColor& = GetPixel(dcMain&, X&, Y&)
            If RGBColor& = TransparentColor& Then
                rgnPixel& = CreateRectRgn(X&, Y&, X& + 1&, Y& + 1&)
                CombineRgn rgnMain&, rgnMain&, rgnPixel&, RGN_XOR
                DeleteObject rgnPixel&
            End If
        Next X&
    Next Y&
    SelectObject dcMain&, bmpMain&
    DeleteDC dcMain&
    DeleteObject bmpMain&
    If rgnMain& <> 0& Then
        SetWindowRgn Frm.hwnd, rgnMain&, True
        MakeTransparent = rgnMain&
    End If
    Frm.ScaleMode = ScaleSize&
End Function
