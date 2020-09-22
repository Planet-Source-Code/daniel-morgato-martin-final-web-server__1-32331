Attribute VB_Name = "mdlSkin"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_FRAMECHANGED = &H20
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const BackGround As Long = &HFFFFFF
Public Const World As Long = &H0&
Public Const ResetCounter As Long = &H808080
Public Const Minimize As Long = &HC0C0C0
Public Const CloseServer As Long = &H80&
Public Const EnableGestBook As Long = &HFF&
Public Const ActivateOnStart As Long = &H8080&
Public Const WriteLogFile As Long = &HFFFF&
Public Const EnableCounter As Long = &H8000&
Public Const SendIPTo As Long = &HFF00&
Public Const Activate As Long = &H808000
Public Const Deactivate As Long = &HFFFF00
Public Const ViewLog As Long = &H800000
Public Const GreenBall As Long = &HFF0000
Public Const RedBall As Long = &H800080
Public Const BlueBall As Long = &HFF00FF

Function GetButton(ByVal myColor As Long) As Long
    GetButton = 0
    Select Case myColor
        Case World
            GetButton = 1
        Case ResetCounter
            GetButton = 2
        Case Minimize
            GetButton = 3
        Case CloseServer
            GetButton = 4
        Case EnableGestBook
            GetButton = 5
        Case ActivateOnStart
            GetButton = 6
        Case WriteLogFile
            GetButton = 7
        Case EnableCounter
            GetButton = 8
        Case SendIPTo
            GetButton = 9
        Case Activate
            GetButton = 10
        Case Deactivate
            GetButton = 11
        Case ViewLog
            GetButton = 12
        Case GreenBall
            GetButton = 13
        Case RedBall
            GetButton = 14
        Case BlueBall
            GetButton = 15
    End Select
End Function


Public Sub MoveNow(FrmDest As Form)
    Call ReleaseCapture
    Call SendMessage(FrmDest.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim x As Long, y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hdc, 0, 0)
    
    For y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            
            If GetPixel(hdc, x, y) = TransparentColor Or x = PicWidth - 1 Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function
