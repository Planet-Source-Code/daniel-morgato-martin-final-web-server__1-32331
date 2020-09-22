VERSION 5.00
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Message"
   ClientHeight    =   4575
   ClientLeft      =   6135
   ClientTop       =   2400
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4575
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picButtonUnpressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   480
      Picture         =   "frmMessage.frx":0000
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   4
      Top             =   3720
      Width           =   1380
   End
   Begin VB.PictureBox picButtonMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1920
      Picture         =   "frmMessage.frx":1AE6
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   3
      Top             =   3720
      Width           =   1380
   End
   Begin VB.PictureBox picButtonPressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   480
      Picture         =   "frmMessage.frx":35CC
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   2
      Top             =   4080
      Width           =   1380
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   4680
      Picture         =   "frmMessage.frx":50B2
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   299
      TabIndex        =   1
      Top             =   240
      Width           =   4485
   End
   Begin VB.PictureBox picBackGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   4560
      Picture         =   "frmMessage.frx":348D0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   299
      TabIndex        =   0
      Top             =   120
      Width           =   4485
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Put message here..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DownX As Long
Private DownY As Long


Private Sub Form_Load()
    Dim WindowRegion As Long
    
    picMask.ScaleMode = vbPixels
    picMask.AutoRedraw = True
    picMask.AutoSize = True
    picMask.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = 4495
    Me.Height = 3235
    
    Me.Picture = picBackGround.Picture
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
        
    WindowRegion = MakeRegion(picMask)
    SetWindowRgn Me.hwnd, WindowRegion, True
    
    Call TransparentBlt(Me, picButtonUnpressed, CInt(1570 / _
        Screen.TwipsPerPixelX), CInt(2650 / Screen.TwipsPerPixelY), _
        RGB(255, 255, 255))
        
    Me.Refresh
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        DownX = CInt((X - 1570) / Screen.TwipsPerPixelX)
        DownY = CInt((Y - 2650) / Screen.TwipsPerPixelY)
        If picButtonMask.Point(DownX, DownY) = 0 Then
            Call TransparentBlt(Me, picButtonPressed, CInt(1570 / _
                Screen.TwipsPerPixelX), CInt(2650 / Screen.TwipsPerPixelY), _
                RGB(255, 255, 255))
            Me.Refresh
        Else
            Call MoveNow(Me)
        End If
    End If
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picButtonMask.Point(DownX, DownY) = 0 Then
        Call TransparentBlt(Me, picButtonUnpressed, CInt(1570 / _
            Screen.TwipsPerPixelX), CInt(2650 / _
            Screen.TwipsPerPixelY), RGB(255, 255, 255))
        Me.Refresh

        If picButtonMask.Point((X - 1570) / Screen.TwipsPerPixelX, _
            (Y - 2650) / Screen.TwipsPerPixelY) = 0 Then
            DoEvents
            Unload Me
        End If
    End If
End Sub


Private Sub lblMessage_Change()
    Call SetWindowPos(Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE _
        Or SWP_NOSIZE Or SWP_SHOWWINDOW)
End Sub

Private Sub lblMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, lblMessage.Left + X, lblMessage.Top + Y)
End Sub


