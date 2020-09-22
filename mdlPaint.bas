Attribute VB_Name = "mdlPaint"
    
Option Explicit

'14 bytes
Type Bitmap
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type


'Chamadas ao API
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'Constantes
Public Const SRCCOPY = &HCC0020
Public Const NOTSRCCOPY = &H330008
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046
Public Const RC_BITBLT = 1
Public Const RASTERCAPS = 38


Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
        X As Long
        Y As Long
End Type

Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_LBUTTONUP = &H202

Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long

Public Const R2_XORPEN = 7
Public Const PS_DOT = 2                     '  .......

Public Sub TransparentBlt(Dest As Object, ByVal SrcBmp As Long, ByVal DestX As Integer, ByVal DestY As Integer, ByVal TransColor As Long)
    Const PIXEL = 3
    
    Dim DestScale As Integer
    Dim SrcDC As Long           'Bitmap fonte (colorido)
    Dim SaveDC As Long          'BackUp do bitmap fonte
    Dim MaskDC As Long          'Bitmap mask (1 cor)
    Dim InvDC As Long           'Inverso do bitmap mask (1 cor)
    Dim ResultDC As Long        'Combinação do do bitmap fonte com o fundo
    Dim bmp As Bitmap           'Descrição do bitmap fonte
    Dim hResultBmp As Long      'Combinação do do bitmap fonte com o fundo
    
    Dim hSaveBmp As Long        'BackUp do bitmapfonte
    Dim hMaskBmp As Long        'Bitmap mask (1 cor)
    Dim hInvBmp As Long         'Inverso do bitmpa mask
    Dim hPrevBmp As Long        'Bitmap anteriormente selecionado no DC
    
    Dim hSrcPrevBmp As Long      'Bitmap anterior no DC fonte
    Dim hSavePrevBmp As Long     'Bitmap anterior no DC salvo
    Dim hDestPrevBmp As Long     'Bitmap anterior no DC de destino
    Dim hMaskPrevBmp As Long     'Bitmap anterior no DC mask
    Dim hInvPrevBmp As Long      'Bitmap anterior no DC mask invertido
    
    Dim OrigColor As Long       'Cor original do fundo do DC fonte
    Dim Success As Integer      'Resultado da chamada ao API
    
    DestScale = Dest.ScaleMode      'Guarda ScaleMode para ser reativado depois
    Dest.ScaleMode = PIXEL          'Muda ScaleMode para pixel para que o GDI possa trabalhar
    
    Success = GetObject(SrcBmp, Len(bmp), bmp)      'Retorna dados do bitmap (altura, largura, etc...)
    
    SrcDC = CreateCompatibleDC(Dest.hdc)        'Cria DC para conter uma etapa da mudança
    SaveDC = CreateCompatibleDC(Dest.hdc)       'Cria DC para conter uma etapa da mudança
    MaskDC = CreateCompatibleDC(Dest.hdc)       'Cria DC para conter uma etapa da mudança
    InvDC = CreateCompatibleDC(Dest.hdc)        'Cria DC para conter uma etapa da mudança
    ResultDC = CreateCompatibleDC(Dest.hdc)     'Cria DC para conter uma etapa da mudança
    
    'Cria bitmap monocromático para os bitmaps mask
    hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    
    'Cria bitmaps coloridos para o resultado final e para beckups do fonte
    hResultBmp = CreateCompatibleBitmap(Dest.hdc, bmp.bmWidth, bmp.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(Dest.hdc, bmp.bmWidth, bmp.bmHeight)
    
    hSrcPrevBmp = SelectObject(SrcDC, SrcBmp)
    hSavePrevBmp = SelectObject(SaveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(MaskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(InvDC, hInvBmp)
    hDestPrevBmp = SelectObject(ResultDC, hResultBmp)
    
    'Faz cópia do bitmap fonte para restaurá-lo depois
    Success = BitBlt(SaveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCCOPY)
    
    'Cria mask: muda a cor do fundo do bitmap fonte para transparente.
    OrigColor = SetBkColor(SrcDC, TransColor)
    Success = BitBlt(MaskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCCOPY)
    TransColor = SetBkColor(SrcDC, OrigColor)
    
    'Cria inverso do bitmap mask para combinar (AND) com o bitmap fonte e combinar com o fundo
    Success = BitBlt(InvDC, 0, 0, bmp.bmWidth, bmp.bmHeight, MaskDC, 0, 0, NOTSRCCOPY)
    
    'Copia o bitmap de fundo para o bitmap resultado e cria bitmap final transparente
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, Dest.hdc, DestX, DestY, SRCCOPY)
    
    'Combina (AND) bitmap mask com bitmap resultado, pondo um buraco no fundo (pintando de preto)
    'a área não transparente do bitmap fonte.
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, MaskDC, 0, 0, SRCAND)
    
    'Combina (AND) o bitmap mask invertido com o bitmap fonte para tirar os bits associados
    'com a área transparente do bitmap fonte tornando-a preta.
    Success = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, InvDC, 0, 0, SRCAND)
    
    'Combina (XOR) bitmap resultado com bitmap fonte para fazer o fundo ?throught?
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCPAINT)
    
    'Mostra bitmap transparente no fundo.
    Success = BitBlt(Dest.hdc, DestX, DestY, bmp.bmWidth, bmp.bmHeight, ResultDC, 0, 0, SRCCOPY)
    
    'Restaura beckup do bitmap
    Success = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SaveDC, 0, 0, SRCCOPY)
    
    hPrevBmp = SelectObject(SrcDC, hSrcPrevBmp)         'Seleciona objeto original
    hPrevBmp = SelectObject(SaveDC, hSavePrevBmp)       'Seleciona objeto original
    hPrevBmp = SelectObject(ResultDC, hDestPrevBmp)     'Seleciona objeto original
    hPrevBmp = SelectObject(MaskDC, hMaskPrevBmp)       'Seleciona objeto original
    hPrevBmp = SelectObject(InvDC, hInvPrevBmp)         'Seleciona objeto original
    
    Success = DeleteObject(hSaveBmp)        'Libera recursos de sistema
    Success = DeleteObject(hMaskBmp)        'Libera recursos de sistema
    Success = DeleteObject(hInvBmp)         'Libera recursos de sistema
    Success = DeleteObject(hResultBmp)      'Libera recursos de sistema
    
    Success = DeleteDC(SrcDC)       'Libera recursos de sistema
    Success = DeleteDC(SaveDC)      'Libera recursos de sistema
    Success = DeleteDC(InvDC)       'Libera recursos de sistema
    Success = DeleteDC(MaskDC)      'Libera recursos de sistema
    Success = DeleteDC(ResultDC)    'Libera recursos de sistema
    
    'Restaura ScaleMode
    Dest.ScaleMode = DestScale
End Sub



Public Sub TransparentBltA(Dest As Object, ByVal SrcBmp As Long, _
    ByVal UseMask As Long, ByVal DestX As Integer, _
    ByVal DestY As Integer, ByVal nTransColor As Long)
    Const PIXEL = 3
    
    Dim DestScale As Integer
    Dim SrcDC As Long           'Bitmap fonte (colorido)
    Dim SaveDC As Long          'BackUp do bitmap fonte
    Dim MaskDC As Long          'Bitmap mask (1 cor)
    Dim myMaskDC As Long
    Dim InvDC As Long           'Inverso do bitmap mask (1 cor)
    Dim ResultDC As Long        'Combinação do do bitmap fonte com o fundo
    Dim bmp As Bitmap           'Descrição do bitmap fonte
    Dim hResultBmp As Long      'Combinação do do bitmap fonte com o fundo
    Dim Temp As Long
        
    Dim hSaveBmp As Long        'BackUp do bitmapfonte
    Dim hMaskBmp As Long        'Bitmap mask (1 cor)
    Dim hInvBmp As Long         'Inverso do bitmpa mask
    Dim hPrevBmp As Long        'Bitmap anteriormente selecionado no DC
    
    Dim hSrcPrevBmp As Long      'Bitmap anterior no DC fonte
    Dim hSavePrevBmp As Long     'Bitmap anterior no DC salvo
    Dim hDestPrevBmp As Long     'Bitmap anterior no DC de destino
    Dim hMaskPrevBmp As Long     'Bitmap anterior no DC mask
    Dim myhMaskPrevBmp As Long
    Dim hInvPrevBmp As Long      'Bitmap anterior no DC mask invertido
    
    Dim OrigColor As Long       'Cor original do fundo do DC fonte
    Dim Success As Integer      'Resultado da chamada ao API
    
    DestScale = Dest.ScaleMode      'Guarda ScaleMode para ser reativado depois
    Dest.ScaleMode = PIXEL          'Muda ScaleMode para pixel para que o GDI possa trabalhar
    
    Success = GetObject(SrcBmp, Len(bmp), bmp)      'Retorna dados do bitmap (altura, largura, etc...)
    
    SrcDC = CreateCompatibleDC(Dest.hdc)
    ' Add...
    myMaskDC = CreateCompatibleDC(Dest.hdc)
    SaveDC = CreateCompatibleDC(Dest.hdc)       'Cria DC para conter uma etapa da mudança
    MaskDC = CreateCompatibleDC(Dest.hdc)       'Cria DC para conter uma etapa da mudança
    InvDC = CreateCompatibleDC(Dest.hdc)        'Cria DC para conter uma etapa da mudança
    ResultDC = CreateCompatibleDC(Dest.hdc)     'Cria DC para conter uma etapa da mudança
    
    'Cria bitmap monocromático para os bitmaps mask
    hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    
    'Cria bitmaps coloridos para o resultado final e para beckups do fonte
    hResultBmp = CreateCompatibleBitmap(Dest.hdc, bmp.bmWidth, bmp.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(Dest.hdc, bmp.bmWidth, bmp.bmHeight)
    
    hSrcPrevBmp = SelectObject(SrcDC, SrcBmp)
    hSavePrevBmp = SelectObject(SaveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(MaskDC, hMaskBmp)
    ' Add...
    myhMaskPrevBmp = SelectObject(myMaskDC, UseMask)
    hInvPrevBmp = SelectObject(InvDC, hInvBmp)
    hDestPrevBmp = SelectObject(ResultDC, hResultBmp)
    
    'Faz cópia do bitmap fonte para restaurá-lo depois
    Success = BitBlt(SaveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCCOPY)
    
    'Cria mask: muda a cor do fundo do bitmap fonte para transparente.
    OrigColor = SetBkColor(myMaskDC, nTransColor)
    Success = BitBlt(InvDC, 0, 0, bmp.bmWidth, bmp.bmHeight, myMaskDC, 0, 0, SRCCOPY)
    nTransColor = SetBkColor(myMaskDC, OrigColor)
    
    'Cria inverso do bitmap mask para combinar (AND) com o bitmap fonte e combinar com o fundo
    Success = BitBlt(MaskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, InvDC, 0, 0, NOTSRCCOPY)
    '******************************************************************************
'    Temp = InvDC
'    InvDC = MaskDC
'    MaskDC = Temp
    'Success = BitBlt(InvDC, 0, 0, bmp.bmWidth, bmp.bmHeight, MaskDC, 0, 0, SRCCOPY)
    '******************************************************************************
    
    
    'Copia o bitmap de fundo para o bitmap resultado e cria bitmap final transparente
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, Dest.hdc, DestX, DestY, SRCCOPY)
    
    'Combina (AND) bitmap mask com bitmap resultado, pondo um buraco no fundo (pintando de preto)
    'a área não transparente do bitmap fonte.
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, MaskDC, 0, 0, SRCAND)
    
    'Combina (AND) o bitmap mask invertido com o bitmap fonte para tirar os bits associados
    'com a área transparente do bitmap fonte tornando-a preta.
    Success = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, InvDC, 0, 0, SRCAND)
    
    'Combina (XOR) bitmap resultado com bitmap fonte para fazer o fundo ?throught?
    Success = BitBlt(ResultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCPAINT)
    
    'Mostra bitmap transparente no fundo.
    Success = BitBlt(Dest.hdc, DestX, DestY, bmp.bmWidth, bmp.bmHeight, ResultDC, 0, 0, SRCCOPY)
    
    'Restaura beckup do bitmap
    Success = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SaveDC, 0, 0, SRCCOPY)
    
    hPrevBmp = SelectObject(SrcDC, hSrcPrevBmp)         'Seleciona objeto original
    hPrevBmp = SelectObject(SaveDC, hSavePrevBmp)       'Seleciona objeto original
    hPrevBmp = SelectObject(ResultDC, hDestPrevBmp)     'Seleciona objeto original
    hPrevBmp = SelectObject(MaskDC, hMaskPrevBmp)       'Seleciona objeto original
    hPrevBmp = SelectObject(InvDC, hInvPrevBmp)         'Seleciona objeto original
    hPrevBmp = SelectObject(myMaskDC, myhMaskPrevBmp)
    
    Success = DeleteObject(hSaveBmp)        'Libera recursos de sistema
    Success = DeleteObject(hMaskBmp)        'Libera recursos de sistema
    Success = DeleteObject(hInvBmp)         'Libera recursos de sistema
    Success = DeleteObject(hResultBmp)      'Libera recursos de sistema
    
    Success = DeleteDC(SrcDC)       'Libera recursos de sistema
    Success = DeleteDC(SaveDC)      'Libera recursos de sistema
    Success = DeleteDC(InvDC)       'Libera recursos de sistema
    Success = DeleteDC(MaskDC)      'Libera recursos de sistema
    Success = DeleteDC(myMaskDC)
    Success = DeleteDC(ResultDC)    'Libera recursos de sistema
    
    'Restaura ScaleMode
    Dest.ScaleMode = DestScale
End Sub

