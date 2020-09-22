VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl SMTPSender 
   CanGetFocus     =   0   'False
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "SMTP.ctx":0000
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   109
   ToolboxBitmap   =   "SMTP.ctx":09BA
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   25
   End
End
Attribute VB_Name = "SMTPSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum StateConstants
    State_Closed = 0
    State_Connecting = 1
    State_Connected = 2
    State_Helo = 3
    State_MailFrom = 4
    State_RcptTo = 5
    State_Data = 6
    State_Dot = 7
    State_Quit = 8
End Enum

'Default property values:
Const m_def_State = 0
Const m_def_ToAdress = ""
Const m_def_FromAdress = ""
Const m_def_Subject = ""
Const m_def_HeloName = ""
Const m_def_MessageText = ""

'Property variables:
Dim m_State As StateConstants
Dim m_ToAdress As String
Dim m_FromAdress As String
Dim m_Subject As String
Dim m_HeloName As String
Dim m_MessageText As String

'Event declarations:
Event SendComplete()
Attribute SendComplete.VB_Description = "Occurs after a send operation has completed"
Event Connect(RemoteHostIP As String)
Event Error(Code As Integer, Description As String)
Event StateChanged(OldState As StateConstants)
Event SendCancel()
Event ConnectionClosed()

Private Function GetCode(Source As String) As Integer
    GetCode = CInt(Mid(Source, 1, 3))
End Function

Private Function GetText(Source As String) As String
    GetText = Mid(Source, 5, Len(Source) - 4)
End Function


Private Sub sckMail_Close()
    Dim OldState As StateConstants
    OldState = m_State
    If (OldState = State_Quit) Then
        RaiseEvent SendComplete
    End If
    RaiseEvent ConnectionClosed
    m_State = State_Closed
    RaiseEvent StateChanged(OldState)
End Sub

Private Sub sckMail_Connect()
    Dim OldState As StateConstants
    OldState = m_State
    If m_State = State_Connecting Then
        m_State = State_Connected
        RaiseEvent Connect(sckMail.RemoteHostIP)
        RaiseEvent StateChanged(OldState)
    Else
        sckMail.Close
    End If
End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)
Dim IncommingData As String

    If sckMail.State <> sckConnected Then
        Exit Sub
    End If
    sckMail.GetData IncommingData
    Select Case m_State
    
        Case State_Connected
            If GetCode(IncommingData) = 421 Then
                RaiseEvent Error(421, GetText(IncommingData))
                Exit Sub
            ElseIf GetCode(IncommingData) = 220 Then
                sckMail.SendData "HELO " + sckMail.LocalIP + vbCrLf
                m_State = State_Helo
                RaiseEvent StateChanged(State_Connected)
            Else
                'SMTP receiver don't sent a valid code...
                'We supose, for more flexibility, that 250 is also valid...
                If GetCode(IncommingData) = 250 Then
                    sckMail.SendData "HELO " + HeloName + vbCrLf
                    m_State = State_Helo
                    RaiseEvent StateChanged(State_Connected)
                Else
                    RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
                End If
            End If
            
        Case State_Helo
            'Error when code = 421, 500, 501, 504...
            If GetCode(IncommingData) > 400 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            ElseIf GetCode(IncommingData) = 250 Then
                sckMail.SendData "MAIL FROM: " + "<" + FromAdress + ">" + vbCrLf
                m_State = State_MailFrom
                RaiseEvent StateChanged(State_Helo)
            Else
                RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
            End If
            
        Case State_MailFrom
            'Error when code = 421, 552, 452, 500, 501, 451...
            If GetCode(IncommingData) > 400 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            ElseIf GetCode(IncommingData) = 250 Then
                sckMail.SendData "RCPT TO: " + "<" + ToAdress + ">" + vbCrLf
                m_State = State_RcptTo
                RaiseEvent StateChanged(State_MailFrom)
            Else
                RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
            End If
            
        Case State_RcptTo
            'Error when code = 421, 450, 451, 452,
            '500, 501, 503, 550, 551, 552, 553...
            If GetCode(IncommingData) > 400 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            ElseIf GetCode(IncommingData) = 250 Or _
                GetCode(IncommingData) = 251 Then
                sckMail.SendData "DATA" + vbCrLf
                m_State = State_Data
                RaiseEvent StateChanged(State_RcptTo)
            Else
                RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
            End If
            
        Case State_Data
            'Error when code = 421, 451, 554, 500, 501, 503...
            If GetCode(IncommingData) > 400 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            ElseIf GetCode(IncommingData) = 354 Then
                sckMail.SendData "Date: " + Format(Now, "Long Time") + vbCrLf
                sckMail.SendData "From: " + FromAdress + vbCrLf
                sckMail.SendData "Subject: " + Subject + vbCrLf
                sckMail.SendData "To: " + ToAdress + vbCrLf + vbCrLf
                sckMail.SendData MessageText
                sckMail.SendData vbCrLf + "." + vbCrLf
                m_State = State_Dot
                RaiseEvent StateChanged(State_Data)
            Else
                RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
            End If
            
        Case State_Dot
            'Error when code = 552, 554, 451, 452...
            If GetCode(IncommingData) > 400 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            ElseIf GetCode(IncommingData) = 250 Then
                sckMail.SendData "QUIT" + vbCrLf
                m_State = State_Quit
                RaiseEvent StateChanged(State_Dot)
            Else
                RaiseEvent Error(516, "SMTP Receiver don't answered correctly")
            End If
            
        Case Sate_Quit
            'Erro: 500...
            If GetCode(IncommingData) = 500 Then
                RaiseEvent Error(GetCode(IncommingData), GetText(IncommingData))
            End If
    End Select
End Sub

Private Sub sckMail_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, Description)
    sckMail.Close
    m_State = State_Closed
End Sub

Private Sub UserControl_Resize()
    'Keep alwais the same size...
    UserControl.Width = 570
    UserControl.Height = 525
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckMail,sckMail,-1,LocalIP
Public Property Get LocalIP() As String
Attribute LocalIP.VB_Description = "Returns the local machine IP address"
    LocalIP = sckMail.LocalIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckMail,sckMail,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = sckMail.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    sckMail.RemoteHost = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,2,0
Public Property Get State() As Integer
Attribute State.VB_Description = "Returns the state of the socket connection"
Attribute State.VB_MemberFlags = "400"
    State = m_State
End Property

Public Property Let State(ByVal New_State As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub CancelSend()
Attribute CancelSend.VB_Description = "Stops current mail transfer proccess."
Dim OldState As StateConstants
    If m_State = State_Data Then
        sckMail.SendData vbCrLf + "." + vbCrLf
    End If
    If m_State > State_Closed Then
        sckMail.SendData "RSET" + vbCrLf
        sckMail.SendData "QUIT" + vbCrLf
        sckMail.Close
        OldState = m_State
        m_State = State_Closed
        RaiseEvent SendCancel
        RaiseEvent StateChanged(OldState)
        RaiseEvent ConnectionClosed
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub Send(Optional Host As String, Optional Helo As String, _
Optional MailFrom As String, Optional MailTo As String, _
Optional TSubject As String, Optional Message As String)

    If m_State > State_Closed Then
        Err.Raise 513, "SMTPSender.Send", "Cannot send mail at this time."
        Exit Sub
    End If
    
    If Not (IsMissing(Host) = True Or Host = "") Then
        RemoteHost = Host
    ElseIf RemoteHost = "" Then
        Err.Raise 514, "SMTPSender.Send", "Remote host missing."
    End If
    
    If Not (IsMissing(Helo) = True Or Helo = "") Then
        HeloName = Helo
    ElseIf HeloName = "" Then
        HeloName = sckMail.LocalIP
    End If
    
    If Not (IsMissing(MailFrom) = True Or MailFrom = "") Then
        FromAdress = MailFrom
    ElseIf FromAdress = "" Then
        Err.Raise 515, "SMTPSender.Send", "From adress missing."
    End If
    
    If Not (IsMissing(MailTo) = True Or MailTo = "") Then
        ToAdress = MailTo
    ElseIf ToAdress = "" Then
        Err.Raise 516, "SMTPSender.Send", "To adress missing."
    End If
 
    If Not (IsMissing(TSubject) = True Or TSubject = "") Then
        Subject = TSubject
    ElseIf Subject = "" Then
        Subject = "(No subject)"
    End If
            
    If Not IsMissing(Message) And Message <> "" Then
        FromAdress = MailFrom
    End If

    sckMail.Connect RemoteHost, 25
    m_State = State_Connecting
    RaiseEvent StateChanged(State_Closed)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToAdress() As String
    ToAdress = m_ToAdress
End Property

Public Property Let ToAdress(ByVal New_ToAdress As String)
    m_ToAdress = New_ToAdress
    PropertyChanged "ToAdress"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FromAdress() As String
    FromAdress = m_FromAdress
End Property

Public Property Let FromAdress(ByVal New_FromAdress As String)
    m_FromAdress = New_FromAdress
    PropertyChanged "FromAdress"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Subject() As String
    Subject = m_Subject
End Property

Public Property Let Subject(ByVal New_Subject As String)
    m_Subject = New_Subject
    PropertyChanged "Subject"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get HeloName() As String
    HeloName = m_HeloName
End Property

Public Property Let HeloName(ByVal New_HeloName As String)
    m_HeloName = New_HeloName
    PropertyChanged "HeloName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get MessageText() As String
    MessageText = m_MessageText
End Property

Public Property Let MessageText(ByVal New_MessageText As String)
    m_MessageText = New_MessageText
    PropertyChanged "MessageText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub Reset()
    If m_State = State_Closed Then
        m_ToAdress = ""
        m_FromAdress = ""
        m_Subject = ""
        m_HeloName = ""
        m_MessageText = ""
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_State = m_def_State
    m_ToAdress = m_def_ToAdress
    m_FromAdress = m_def_FromAdress
    m_Subject = m_def_Subject
    m_HeloName = m_def_HeloName
    m_MessageText = m_def_MessageText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    sckMail.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    m_State = PropBag.ReadProperty("State", m_def_State)
    m_ToAdress = PropBag.ReadProperty("ToAdress", m_def_ToAdress)
    m_FromAdress = PropBag.ReadProperty("FromAdress", m_def_FromAdress)
    m_Subject = PropBag.ReadProperty("Subject", m_def_Subject)
    m_HeloName = PropBag.ReadProperty("HeloName", m_def_HeloName)
    m_MessageText = PropBag.ReadProperty("MessageText", m_def_MessageText)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("RemoteHost", sckMail.RemoteHost, "")
    Call PropBag.WriteProperty("State", m_State, m_def_State)
    Call PropBag.WriteProperty("ToAdress", m_ToAdress, m_def_ToAdress)
    Call PropBag.WriteProperty("FromAdress", m_FromAdress, m_def_FromAdress)
    Call PropBag.WriteProperty("Subject", m_Subject, m_def_Subject)
    Call PropBag.WriteProperty("HeloName", m_HeloName, m_def_HeloName)
    Call PropBag.WriteProperty("MessageText", m_MessageText, m_def_MessageText)
End Sub

