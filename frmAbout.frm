VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2280
      Tag             =   "0"
      Top             =   2280
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   4440
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   3
      Top             =   600
      Width           =   4035
   End
   Begin VB.PictureBox picMainSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   4320
      Picture         =   "frmAbout.frx":6792
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   2
      Top             =   480
      Width           =   4035
   End
   Begin VB.PictureBox picPressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   4200
      Picture         =   "frmAbout.frx":CF24
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   1
      Top             =   360
      Width           =   4035
   End
   Begin VB.PictureBox picUnpressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2910
      Left            =   4080
      Picture         =   "frmAbout.frx":337B6
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   240
      Width           =   4035
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ColorDown As Long

Private Sub Form_Load()
    Dim WindowRegion As Long
    
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = 4035
    Me.Height = 2910
    
    Me.Picture = picUnpressed.Picture
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
        
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempButton As Long
    If Button = vbLeftButton Then
        ColorDown = picMask.Point(X, Y)
        tempButton = GetButton(ColorDown)
        If ColorDown = BackGround Then
            Call MoveNow(Me)
        Else
            If tempButton < 3 Then
                Call TransparentBltA(Me, picPressed, _
                    picMask, 0, 0, ColorDown)
            End If
        End If
    End If
    Me.Refresh
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempButton As Long
    Dim OldButton As Long
    
    If Button = vbLeftButton Then
        tempButton = GetButton(picMask.Point(X, Y))
        OldButton = GetButton(ColorDown)
        Call TransparentBltA(Me, picUnpressed, _
            picMask, 0, 0, ColorDown)
        If OldButton = tempButton Then
            If tempButton = 1 Then
                Call ShellExecute(GetDesktopWindow(), _
                    vbNullString, progURL, vbNullString, _
                    vbNullString, vbNormalFocus)
            ElseIf tempButton = 2 Then
                Unload Me
                End
            End If
        End If
    End If
    Me.Refresh
End Sub


Private Sub Timer1_Timer()
    Unload Me
    End
End Sub


