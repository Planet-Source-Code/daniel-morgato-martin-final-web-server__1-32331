VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   5070
   ClientLeft      =   5880
   ClientTop       =   1905
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   3240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   1035
      TabIndex        =   4
      Top             =   2955
      Width           =   1830
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   1035
      TabIndex        =   3
      Top             =   2235
      Width           =   1845
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   5040
      Picture         =   "frmOptions.frx":0000
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   2
      Top             =   480
      Width           =   3285
   End
   Begin VB.PictureBox picMainSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4920
      Picture         =   "frmOptions.frx":8432
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   1
      Top             =   360
      Width           =   3270
   End
   Begin VB.PictureBox picPressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4800
      Picture         =   "frmOptions.frx":38934
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   0
      Top             =   240
      Width           =   3270
   End
   Begin VB.PictureBox picUnpressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4680
      Picture         =   "frmOptions.frx":68E36
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   218
      TabIndex        =   5
      Top             =   120
      Width           =   3270
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ColorDown As Long
Private ButtonState(1 To 2) As Boolean

Private Sub Form_Load()
    Dim WindowRegion As Long
        
    Me.Width = 3285
    Me.Height = 4500
    
    Me.Picture = picUnpressed.Picture
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
        
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
    
    If Intro = "1" Then
        ButtonState(GetButton(0)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, 0)
    End If
    
    If AlwaysOnTop = "1" Then
        ButtonState(GetButton(&H808080)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, &H808080)
    End If
    
    txtDir = indexDir
    txtPage = indexPage
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempButton As Long
    If Button = vbLeftButton Then
        ColorDown = picMask.Point(X, Y)
        tempButton = GetButton(ColorDown)
        Debug.Print "MouseDown em " + CStr(tempButton) + ". Color = " + Hex(ColorDown) + "."
        If ColorDown = BackGround Then
            Call MoveNow(Me)
        Else
            If tempButton <= 6 Then
                If tempButton < 3 And tempButton >= 1 Then
                    If ButtonState(tempButton) = True Then
                        Call TransparentBltA(Me, picUnpressed, _
                            picMask, 0, 0, ColorDown)
                    Else
                        Call TransparentBltA(Me, picPressed, _
                            picMask, 0, 0, ColorDown)
                    End If
                Else
                    Call TransparentBltA(Me, picPressed, _
                        picMask, 0, 0, ColorDown)
                End If
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
        If OldButton <> tempButton Then
            Debug.Print "Fora: ";
            If OldButton < 3 And OldButton >= 1 Then
                Debug.Print "botão nº" + CStr(OldButton) + " -> ";
                If ButtonState(OldButton) = True Then
                    Debug.Print "On"
                    Call TransparentBltA(Me, picPressed, _
                        picMask, 0, 0, ColorDown)
                Else
                    Debug.Print "Off"
                    Call TransparentBltA(Me, picUnpressed, _
                        picMask, 0, 0, ColorDown)
                End If
            Else
                Debug.Print "não botão -> Off"
                Call TransparentBltA(Me, picUnpressed, _
                    picMask, 0, 0, ColorDown)
            End If
        Else
            Debug.Print "Dentro: ";
            If tempButton >= 3 Or tempButton < 1 Then
                Debug.Print "não botão -> Off"
                Call TransparentBltA(Me, picUnpressed, _
                    picMask, 0, 0, ColorDown)
                If tempButton = 6 Then
                    Dim r As VbMsgBoxResult
                    r = MsgBox("Changes will be lost. Exit anyway?", vbYesNo _
                        Or vbQuestion, "Discard changes?")
                    If r = vbYes Then
                        Unload Me
                        Me.Refresh
                        Exit Sub
                    End If
                End If
                If tempButton = 5 Then
                    If Dir(txtDir, vbDirectory) = "" Then
                        MsgBox "Directory doesn't exist. Go back and select a valid one.", _
                            vbInformation, "Invalid directory..."
                        Me.Refresh
                        Exit Sub
                    End If
                    If Dir(AddASlash(txtDir) + txtPage) = "" Then
                        MsgBox "Index page doesn't exist. Go back and select a valid one.", _
                            vbInformation, "Invalid page..."
                        Me.Refresh
                        Exit Sub
                    End If
                    indexDir = txtDir
                    indexPage = txtPage
                    Call SaveSetting(AppName, "Settings", "IndexDir", _
                        AddASlash(indexDir))
                    Call SaveSetting(AppName, "Settings", "IndexPage", indexPage)
                    If ButtonState(1) = False Then
                        Intro = "0"
                    Else
                        Intro = "1"
                    End If
                    Call SaveSetting(AppName, "Settings", "Intro", Intro)
                    If ButtonState(2) = False Then
                        AlwaysOnTop = "0"
                        Call SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, _
                            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
                    Else
                        AlwaysOnTop = "1"
                        Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, _
                            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
                    End If
                    Call SaveSetting(AppName, "Settings", "AlwaysOnTop", _
                        AlwaysOnTop)
                    Me.Refresh
                    Unload Me
                    Exit Sub
                End If
                If tempButton = 4 Then
                    On Error Resume Next
                    CMD1.InitDir = AddASlash(txtDir)
                    CMD1.flags = CMD1.flags Or cdlOFNFileMustExist Or _
                        cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn
                    CMD1.ShowOpen
                    If Err = cdlCancel Then
                        Err = 0
                        Me.Refresh
                        Exit Sub
                    End If
                    If Not Exists(CMD1.FileName) Then
                        MsgBox "Invalid file name. Choose another!", _
                            vbInformation, "Bad file name."
                        Me.Refresh
                        Exit Sub
                    End If
                    If InStr(1, UCase(CMD1.FileName), UCase(AddASlash(txtDir))) = 0 Then
                        MsgBox "File must be in the server's choosen directory." _
                            , vbInformation, "File in wrong directory!"
                        Me.Refresh
                        Exit Sub
                    End If
                    txtPage = GetShortName(CMD1.FileName)
                End If
             Else
                Debug.Print "botão nº" + CStr(tempButton) + " -> Keep"
                ButtonState(tempButton) = Not ButtonState(tempButton)
            End If
        End If
    End If
    Me.Refresh
End Sub


