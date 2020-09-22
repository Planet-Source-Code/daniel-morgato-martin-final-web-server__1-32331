VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Web Server"
   ClientHeight    =   5070
   ClientLeft      =   990
   ClientTop       =   1905
   ClientWidth     =   8100
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   Begin VB.PictureBox Aux 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   5040
      Picture         =   "frmMain.frx":628A
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   360
      Width           =   4500
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3960
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox picPressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4920
      Picture         =   "frmMain.frx":1152C
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   240
      Width           =   4500
   End
   Begin VB.PictureBox picUnpressed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4800
      Picture         =   "frmMain.frx":5381E
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   120
      Width           =   4500
   End
   Begin VB.PictureBox picMainSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4680
      Picture         =   "frmMain.frx":95B10
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   0
      Width           =   4500
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3960
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin FinalWebServer.SMTPSender SMTPSender1 
      Left            =   3960
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   926
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   405
      Stretch         =   -1  'True
      Top             =   225
      Width           =   390
   End
   Begin VB.Label lblCounter 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   1395
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ColorDown As Long
Private ButtonState(5 To 9) As Boolean
Private Sub Form_Load()
    Dim WindowRegion As Long
    
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = 4500
    Me.Height = 4500
    
    Me.Picture = picUnpressed.Picture
    Me.MouseIcon = LoadResPicture(101, vbResCursor)
        
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
    
    If StartOn = "1" Then
        ButtonState(GetButton(ActivateOnStart)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, ActivateOnStart)
        sckServer(0).LocalPort = 80
        sckServer(0).Listen
        Debug.Print "Start listenning on port 80..."
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, GreenBall)
        If SendIP = "1" Then DeliverIP
    End If
    
    If UseGestBook = "1" Then
        ButtonState(GetButton(EnableGestBook)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, EnableGestBook)
    End If
    
    If WriteLog = "1" Then
        ButtonState(GetButton(WriteLogFile)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, WriteLogFile)
    End If
    
    If UseCounter = "1" Then
        ButtonState(GetButton(EnableCounter)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, EnableCounter)
        lblCounter = Counter
    Else
        lblCounter = "0000000"
    End If
    
    If SendIP = "1" Then
        ButtonState(GetButton(SendIPTo)) = True
        Call TransparentBltA(Me, picPressed, _
            picMask, 0, 0, SendIPTo)
    End If
    If Intro = "1" Then
        playIntro
        frmMessage.Show
        frmMessage.lblMessage = "Welcome to the Personal Web Server" + _
            " Environment! Many thanks to: Joox " + _
            ", Gregg Housh, Arkadiy Olovyannikov, and many othe" + _
            "rs who did contribute to this code."
        Call SetWindowPos(frmMessage.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE _
            Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End If
    If AlwaysOnTop = "1" Then
        Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, _
            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tempButton As Long
    If Button = vbLeftButton Then
        ColorDown = picMask.Point(x, y)
        tempButton = GetButton(ColorDown)
        'Debug.Print "MouseDown em " + CStr(tempButton) + ". Color = " + Hex(ColorDown) + "."
        If ColorDown = BackGround Then
            Call MoveNow(Me)
        Else
            If tempButton < 13 Then
                If tempButton < 10 And tempButton >= 5 Then
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


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tempButton As Long
    Dim OldButton As Long
    
    If Button = vbLeftButton Then
        tempButton = GetButton(picMask.Point(x, y))
        OldButton = GetButton(ColorDown)
        If OldButton <> tempButton Then
            'Debug.Print "Fora: ";
            If OldButton < 10 And OldButton >= 5 Then
                'Debug.Print "botão nº" + CStr(OldButton) + " -> ";
                If ButtonState(OldButton) = True Then
                    'Debug.Print "On"
                    Call TransparentBltA(Me, picPressed, _
                        picMask, 0, 0, ColorDown)
                Else
                    'Debug.Print "Off"
                    Call TransparentBltA(Me, picUnpressed, _
                        picMask, 0, 0, ColorDown)
                End If
            Else
                'Debug.Print "não botão -> Off"
                Call TransparentBltA(Me, picUnpressed, _
                    picMask, 0, 0, ColorDown)
            End If
        Else
            'Debug.Print "Dentro: ";
            If tempButton >= 10 Or tempButton < 5 Then
                'Debug.Print "não botão -> Off"
                Call TransparentBltA(Me, picUnpressed, _
                    picMask, 0, 0, ColorDown)
                If ColorDown = CloseServer Then Unload Me ': End
                If ColorDown = Minimize Then Me.WindowState = vbMinimized
                If ColorDown = Activate Then
                    If sckServer(0).State = sckClosed Then
                        sckServer(0).LocalPort = 80
                        sckServer(0).Listen
                        Debug.Print "Start listenning on port 80..."
                        Call TransparentBltA(Me, picPressed, _
                            picMask, 0, 0, GreenBall)
                        If SendIP = "1" Then DeliverIP
                    End If
                End If
                If ColorDown = Deactivate Then
                    sckServer(0).Close
                    Call TransparentBltA(Me, picUnpressed, _
                        picMask, 0, 0, GreenBall)
                End If
                If ColorDown = ViewLog Then _
                    If Exists(LogFile) Then _
                        Call Shell("c:\windows\notepad " + LogFile, vbNormalFocus)
                If ColorDown = ResetCounter Then
                    If UseCounter = "1" Then
                        Counter = "0"
                        Call SaveSetting(AppName, "Settings", "Counter", Counter)
                        lblCounter = "0"
                        If Exists(LogFile) Then Kill LogFile
                    End If
                End If
            Else
                'Debug.Print "botão nº" + CStr(tempButton) + " -> Keep"
                ButtonState(tempButton) = Not ButtonState(tempButton)
                Select Case tempButton
                    Case 6
                        If ButtonState(tempButton) = False Then
                            StartOn = "0"
                        Else
                            StartOn = "1"
                        End If
                        Call SaveSetting(AppName, "Settings", "StartOn", _
                            StartOn)
                    Case 5
                        If ButtonState(tempButton) = False Then
                            UseGestBook = "0"
                        Else
                            UseGestBook = "1"
                        End If
                        Call SaveSetting(AppName, "Settings", "UseGestBook", _
                            UseGestBook)
                    Case 7
                        If ButtonState(tempButton) = False Then
                            WriteLog = "0"
                        Else
                            On Error Resume Next
                            WriteLog = "1"
                            CMD1.DialogTitle = "Logs..."
                            CMD1.InitDir = AddASlash(indexDir)
                            CMD1.FileName = LogFile
                            CMD1.ShowSave
                            If Err = cdlCancel Then
                                WriteLog = "0"
                                ButtonState(tempButton) = Not ButtonState(tempButton)
                                Call TransparentBltA(Me, picUnpressed, _
                                    picMask, 0, 0, WriteLogFile)
                                Me.Refresh
                                Exit Sub
                            End If
                            If Not Exists(CMD1.FileName) Then
                                Err = 0
                                Open CMD1.FileName For Append As #1
                                If Err Then
                                    Err = 0
                                    MsgBox "Unknow error message. Try another filename.", _
                                        vbInformation, "Error:"
                                    WriteLog = "0"
                                    ButtonState(tempButton) = Not ButtonState(tempButton)
                                    Call TransparentBltA(Me, picUnpressed, _
                                        picMask, 0, 0, WriteLogFile)
                                    Me.Refresh
                                    Close #1
                                    Exit Sub
                                End If
                                Print #1, "Created " + Format(Now, "dd/mm/yyyy hh:mm:ss AM/PM")
                                Close #1
                            End If
                            LogFile = CMD1.FileName
                            Call SaveSetting(AppName, "Settings", "LogFile", _
                                LogFile)
                        End If
                        Call SaveSetting(AppName, "Settings", "WriteLog", _
                            WriteLog)
                    Case 8
                        If ButtonState(tempButton) = False Then
                            UseCounter = "0"
                            lblCounter = "0000000"
                        Else
                            UseCounter = "1"
                            lblCounter = Counter
                        End If
                        Call SaveSetting(AppName, "Settings", "UseCounter", _
                            UseCounter)
                    Case 9
                        Dim OutgoingMailRcpt As String
                        
                        If ButtonState(tempButton) = False Then
                            SendIP = "0"
                        Else
                            OutgoingMailRcpt = InputBox("Enter email recipient:", _
                                "Send IP to...", IPMailRcpt)
                            If OutgoingMailRcpt = "" Then
                                ButtonState(tempButton) = Not ButtonState(tempButton)
                                Call TransparentBltA(Me, picUnpressed, _
                                    picMask, 0, 0, SendIPTo)
                                Me.Refresh
                                Exit Sub
                            End If
                            IPMailRcpt = OutgoingMailRcpt
                            SendIP = "1"
                            Call SaveSetting(AppName, "Settings", "IPMailRcpt", _
                                IPMailRcpt)
                        End If
                        Call SaveSetting(AppName, "Settings", "SendIP", _
                            SendIP)
                        'MsgBox "Feature not implemnted yet!", vbInformation, "Info..."
                End Select
            End If
        End If
    End If
    Me.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "Closing..."
    sckServer(0).Close
    Me.Hide
    Unload frmMessage
    stopAll
    frmAbout.Show 1
    End
End Sub


Private Sub Image1_DblClick()
    frmOptions.Show 1
End Sub


Private Sub sckServer_Close(index As Integer)
    Debug.Print "Connection closed. Remote IP: " + _
        CStr(sckServer(index).RemoteHostIP)
End Sub

Private Sub sckServer_Connect(index As Integer)
    Debug.Print "Connection stabilished. Remote IP: " + _
        CStr(sckServer(index).RemoteHostIP)
End Sub

Private Sub sckServer_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Static requestNumber As Long
    If index = 0 Then
        requestNumber = requestNumber + 1
        'Debug.Print "Connection request number: " + CStr(requestNumber)
        Load sckServer(requestNumber)
        sckServer(requestNumber).LocalPort = 0
        sckServer(requestNumber).Accept requestID
    End If
End Sub

Private Sub sckServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Static colData As Collection
    Dim requestedPage As String, strData As String, postedData As String
    Dim secondSpace As Integer, findGet As Integer, findPostedData As Integer
    Dim Status As Long, findPost As Integer, useMethod As TypeMethod
    Dim requestRemoteHostIP As String, Pos As Long
    
    Debug.Print "Incomming data: "
    Debug.Print "    Connection #" + CStr(index) + "."
    Debug.Print "***************************"
    Debug.Print sckServer(index).BytesReceived; " Recebidos;";
    sckServer(index).GetData strData
    Debug.Print bytesTotal; " Totais;" + vbCrLf + "--"
    Debug.Print strData
    If colData.Item(CStr(index)) Is Nothing Then
        Err = 0
    Else
        strData = colData.Item(CStr(index)) + strData
        Call colData.Remove(CStr(index))
    End If

    Debug.Print "***************************"
    Pos = InStr(1, strData, "Host:")
    requestRemoteHostIP = Mid(strData, Pos + 6, InStr(Pos, strData, vbCrLf) - Pos - 6)
    Debug.Print "    Remote host: " + LCase(requestRemoteHostIP);
            
    If Mid(strData, 1, 3) = "GET" Then
        findGet = InStr(strData, "GET ")
        secondSpace = InStr(findGet + 5, strData, " ")
        requestedPage = Mid$(strData, findGet + 4, secondSpace - (findGet + 4))
        useMethod = MethodGet
        Debug.Print "    Requested object: " + requestedPage
        Debug.Print "    Using method GET."
        '********************************************
        If Mid(strData, Len(strData) - 3, 4) <> vbCrLf + vbCrLf Then
            colData.Add strData, CStr(index)
            Exit Sub
        End If
        '********************************************
    ElseIf Mid(strData, 1, 4) = "POST" Then
        findPost = InStr(strData, "POST ")
        secondSpace = InStr(findPost + 6, strData, " ")
        requestedPage = Mid$(strData, findPost + 5, secondSpace - (findPost + 5))
        findPostedData = InStr(1, strData, "Connection:")
        findPostedData = InStr(findPostedData, strData, vbCrLf)
        findPostedData = InStr(findPostedData + 1, strData, vbCrLf)
        postedData = Mid(strData, findPostedData + 2, _
            Len(strData) - findPostedData - 1)
            '********************************************
            If InStr(1, postedData, "&btnSubmit=") = 0 Then
                colData.Add strData, CStr(index)
                Exit Sub
            End If
            '********************************************
        useMethod = MethodPost
        Debug.Print "    Requested object: " + requestedPage
        Debug.Print "    Using method POST."
        Debug.Print "        Data to post: " + postedData
        Debug.Print "        Unformatted data: " + UnformatString(postedData)
    End If
    
    If WriteLog = "1" Then
        If Not Exists(LogFile) Then
            Open LogFile For Append As #1
            Print #1, "Created " + Format(Now, "dd/mm/yyyy hh:mm:ss AM/PM")
            Close #1
        End If
        Open LogFile For Append As #1
        Print #1, requestedPage; " -- ";
        Print #1, sckServer(index).RemoteHostIP; " ("; _
            sckServer(index).RemoteHost; ") -- ";
        Print #1, Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM");
        If useMethod = MethodPost Then
            Print #1, " -- "; postedData
        Else
            Print #1, ""
        End If
        Close #1
    End If
    
    If UseCounter = "1" Then
        Counter = CStr(CLng(Counter) + 1)
        Call SaveSetting(AppName, "Settings", "Counter", Counter)
        lblCounter = Counter
    End If
    
    If requestedPage = "/" Then
        sckServer(index).SendData FormatPage(GetFile(AddASlash(indexDir) + indexPage))
        Exit Sub
    Else
        requestedPage = Mid(requestedPage, 2, Len(requestedPage) - 1)
    End If
    
    If Mid(requestedPage, 1, 5) = "list;" Then
        Call SendRequested(CLng(index), _
            Mid(requestedPage, 6, Len(requestedPage) - 5))
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 5) = "show;" Then
        Call SendInfo(CLng(index), _
            Mid(requestedPage, 6, Len(requestedPage) - 5))
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 5) = "exec;" Then
        requestedPage = UnformatString(Mid(requestedPage, _
            6, Len(requestedPage) - 6))
        If Not Exists(requestedPage) Then
            sckServer(index).SendData "Program " + requestedPage + " doesn't exist."
        Else
            Shell requestedPage, vbNormalFocus
            sckServer(index).SendData "Program " + requestedPage + " started!"
        End If
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 7) = "msg.cgi" Then
        frmMessage.Show
        Pos = InStr(1, postedData, "body=") + 4
        postedData = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, postedData, "&")
        postedData = Mid(postedData, 1, Pos - 1)
        frmMessage.lblMessage = UnformatString(postedData)
        sckServer(index).SendData "Message sent successfuly!"
        Call sndPlaySound(AddASlash(App.Path) + "notify.wav", SND_ASYNC)
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 9) = "shutdown;" Then
        Call ExitWindowsEx(EWX_SHUTDOWN, 0)
        Exit Sub
    End If
        
    If Mid(requestedPage, 1, 8) = "restart;" Then
        Call ExitWindowsEx(EWX_REBOOT, 0)
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 10) = "guestbook;" Then
        Dim SenderName As String, SenderMail As String, SenderComment As String
        Pos = InStr(1, postedData, "Name=") + 4
        SenderName = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SenderName, "&")
        SenderName = Mid(SenderName, 1, Pos - 1)
        SenderName = UnformatString(SenderName)
        Pos = InStr(1, postedData, "&E-Mail=") + 7
        SenderMail = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SenderMail, "&")
        SenderMail = Mid(SenderMail, 1, Pos - 1)
        SenderMail = UnformatString(SenderMail)
        Pos = InStr(1, postedData, "&Comment=") + 8
        SenderComment = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SenderComment, "&")
        SenderComment = Mid(SenderComment, 1, Pos - 1)
        SenderComment = UnformatString(SenderComment)
        If UseGestBook = "1" Then
            Open AddASlash(indexDir) + "guestbook.txt" For Append As #1
            Print #1, "Name: " + SenderName
            Print #1, "Mail: " + SenderMail
            Print #1, "Comment: " + SenderComment
            Print #1, "--"
            Close #1
            sckServer(index).SendData "Your comment was added to the guestbook!"
        Else
            sckServer(index).SendData "Guestbook off! Send later..."
        End If
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 5) = "mail;" Then
        Pos = InStr(1, postedData, "MAILFROM=") + 8
        SMTPSender1.FromAdress = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SMTPSender1.FromAdress, "&")
        SMTPSender1.FromAdress = Mid(SMTPSender1.FromAdress, 1, Pos - 1)
        SMTPSender1.FromAdress = UnformatString(SMTPSender1.FromAdress)
        Pos = InStr(1, postedData, "&RCPTTO=") + 7
        SMTPSender1.ToAdress = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SMTPSender1.ToAdress, "&")
        SMTPSender1.ToAdress = Mid(SMTPSender1.ToAdress, 1, Pos - 1)
        SMTPSender1.ToAdress = UnformatString(SMTPSender1.ToAdress)
        Pos = InStr(1, postedData, "&SUBJECT=") + 8
        SMTPSender1.Subject = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SMTPSender1.Subject, "&")
        SMTPSender1.Subject = Mid(SMTPSender1.Subject, 1, Pos - 1)
        SMTPSender1.Subject = UnformatString(SMTPSender1.Subject)
        Pos = InStr(1, postedData, "&MAILDATA=") + 9
        SMTPSender1.MessageText = Mid(postedData, Pos + 1, Len(postedData) - Pos)
        Pos = InStr(1, SMTPSender1.MessageText, "&")
        SMTPSender1.MessageText = Mid(SMTPSender1.MessageText, 1, Pos - 1)
        SMTPSender1.MessageText = UnformatString(SMTPSender1.MessageText)
  
        Pos = InStr(1, SMTPSender1.ToAdress, "@")
        MX_Query Trim$(Mid(SMTPSender1.ToAdress, Pos + 1, Len(SMTPSender1.ToAdress) - Pos))
        If Err Then
            Debug.Print Err.Description
            Err = 0
            sckServer(index).SendData "Mail not sent. Probably incorrect data."
            Exit Sub
        End If
        
        If MX.count Then
            SMTPSender1.RemoteHost = MX.Best
            DoEvents
            Debug.Print "Trying... " + frmMain.SMTPSender1.RemoteHost
            SMTPSender1.Send
        End If
        sckServer(index).SendData "Mail sent successfuly!"
        Exit Sub
    End If
    

    If Mid(requestedPage, 1, 5) = "play;" Then
        requestedPage = Mid(requestedPage, 6, Len(requestedPage) - 5)
        requestedPage = UnformatString(requestedPage)
        requestedPage = Trim(Mid$(requestedPage, InStr(1, _
            requestedPage, "=") + 1, InStr(InStr(1, requestedPage, _
            "=") + 1, requestedPage, ";") - 2))
        requestedPage = Mid(requestedPage, 1, Len(requestedPage) - 1)
        Debug.Print "    Playing sound: " + requestedPage
        Call sndPlaySound(requestedPage, SND_ASYNC)
        sckServer(index).SendData "Sound has been played."
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 6) = "eject;" Then
        Status = mciSendString("Set CDAudio Door Open Wait", 0&, 0, 0)
        sckServer(index).SendData "Trying eject...done."
        Exit Sub
    End If
    
    If Mid(requestedPage, 1, 6) = "close;" Then
        Status = mciSendString("Set CDAudio Door Closed Wait", 0&, 0, 0)
        sckServer(index).SendData "Trying close...done."
        Exit Sub
    End If
    
    If Exists(AddASlash(indexDir) + requestedPage) Then
        Dim Exte As String, Pre As String
        Exte = UCase(GetExtension(AddASlash(indexDir) + requestedPage))
        If Exte = "TXT" Or Exte = "INI" Or Exte = "C" Or Exte = "CPP" Or _
            Exte = "H" Or Exte = "HXX" Or Exte = "INC" Or Exte = "BAS" Or _
            Exte = "FRM" Or Exte = "VBP" Or Exte = "LIC" Or Exte = "BAT" Then
            Pre = "HTTP/1.1 200 OK" + vbCrLf
            Pre = Pre + "Server: Final Web Server" + vbCrLf
            Pre = Pre + "Content-Type: text/plain" + vbCrLf
            Pre = Pre + "Accept-Ranges: bytes" + vbCrLf
            Pre = Pre + vbCrLf + vbCrLf
            sckServer(index).SendData Pre + FormatPage(GetFile(AddASlash(indexDir) + requestedPage))
        Else
            sckServer(index).SendData _
                FormatPage(GetFile(AddASlash(indexDir) + requestedPage))
        End If

        Debug.Print "Sending file: " + AddASlash(indexDir) + requestedPage
        'sckServer(index).SendData FormatPage(GetFile(AddASlash(indexDir) + requestedPage))
    Else
        Debug.Print "File not found: " + AddASlash(indexDir) + requestedPage
        sckServer(index).SendData "404 - File not found."
    End If
End Sub

Private Sub sckServer_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Winsock error: " + Description
End Sub

Private Sub sckServer_SendComplete(index As Integer)
    Debug.Print "Send completed. Closing connection #" + CStr(index) + "."
    sckServer(index).Close
End Sub


Private Sub sckServer_SendProgress(index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Debug.Print "Send progress. Sent: " + CStr(bytesSent) + _
        ". Remaining: " + CStr(bytesRemaining)
End Sub


Private Sub SMTPSender1_Connect(RemoteHostIP As String)
    Debug.Print "Connected to the SMTP server."
End Sub

Private Sub SMTPSender1_ConnectionClosed()
    Debug.Print "Connection closed."
End Sub


Private Sub SMTPSender1_Error(code As Integer, Description As String)
    Debug.Print "Error: " + Description
End Sub

Private Sub SMTPSender1_SendComplete()
    Debug.Print "Message has been sent successfuly!"
End Sub


