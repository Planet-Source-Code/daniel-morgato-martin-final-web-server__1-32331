Attribute VB_Name = "mdlMain"
Global indexPage As String
Global indexDir As String
Global StartOn As String
Global SendIP As String
Global IPMailRcpt As String
Global LogFile As String
Global WriteLog As String
Global Counter As String
Global UseCounter As String
Global UseGestBook As String
Global AppName As String
Global AlwaysOnTop As String
Global StartUpLaunch As String
Global Intro As String

Global Const progURL = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=32331&lngWId=1"

Public Declare Function DrawIcon Lib "user32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(256) As Byte
End Type

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Const MB_OK = &H0&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const OF_EXIST = &H4000

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2

Enum TypeMethod
    MethodGet
    MethodPost
End Enum

Sub DeliverIP()
On Error GoTo Err_MXQuery
Dim Pos As Integer
  
    Pos = InStr(1, IPMailRcpt, "@")
    MX_Query Trim$(Mid(IPMailRcpt, Pos + 1, Len(IPMailRcpt) - Pos))

    If MX.count Then
        frmMain.SMTPSender1.FromAdress = "webserver@home.sp.br"
        frmMain.SMTPSender1.ToAdress = IPMailRcpt
        frmMain.SMTPSender1.Subject = CStr(frmMain.sckServer(0).LocalIP)
        frmMain.SMTPSender1.MessageText = "On line!"
        frmMain.SMTPSender1.RemoteHost = MX.Best
       
        Debug.Print "Trying... " + frmMain.SMTPSender1.RemoteHost
        frmMain.SMTPSender1.Send
    End If
    
    Exit Sub

Err_MXQuery:

    Debug.Print Err.Description
End Sub

Function FormatPage(ByVal pageToFormat As String) As String
    Dim Pos As Long, Temp As String
    Pos = InStr(1, pageToFormat, "$")
    Do While Pos <> 0
        If UCase(Mid(pageToFormat, Pos, 3)) = "$IP" Then
            Temp = Mid(pageToFormat, 1, Pos - 1)
            Temp = Temp + CStr(frmMain.sckServer(0).LocalIP)
            Temp = Temp + Mid(pageToFormat, Pos + 3, Len(pageToFormat) - Pos - 2)
            pageToFormat = Temp
        End If
        If UCase(Mid(pageToFormat, Pos, 8)) = "$COUNTER" Then
            If UseCounter = "1" Then
                Temp = Mid(pageToFormat, 1, Pos - 1)
                Temp = Temp + Counter
                Temp = Temp + Mid(pageToFormat, Pos + 8, Len(pageToFormat) - Pos - 7)
                pageToFormat = Temp
            Else
                Temp = Mid(pageToFormat, 1, Pos - 1)
                Temp = Temp + "- Service not ready -"
                Temp = Temp + Mid(pageToFormat, Pos + 8, Len(pageToFormat) - Pos - 7)
                pageToFormat = Temp
            End If
        End If
        Pos = InStr(Pos + 1, pageToFormat, "$")
    Loop
    FormatPage = pageToFormat
End Function

Function GetExtension(FileName As String) As String
    Dim Pos As Integer

    If FileName = "" Then
        GetExtension = ""
        Exit Function
    End If
    
    Pos = InStr(1, FileName, ".")
    
    Do While Pos > 0 And Pos < Len(FileName) - 3
        Pos = InStr(Pos + 1, FileName, ".")
    Loop
    
    If Pos = 0 Then
        GetExtension = ""
        Exit Function
    End If
    
    GetExtension = Mid(FileName, Pos + 1, Len(FileName) - Pos)
End Function


Public Function GetPath(ByVal FileName As String) As String
    On Error Resume Next
    If Not (GetAttr(FileName) And vbDirectory) Then
        Do While Mid(FileName, Len(FileName), 1) <> "\"
            FileName = Mid(FileName, 1, Len(FileName) - 1)
        Loop
        If FileName = "" Then
            GetPath = indexDir
        Else
            GetPath = FileName
        End If
        Exit Function
    End If
    GetPath = FileName
End Function

Public Function GetShortName(ByVal nFileName As String) As String
Dim FileName As String
    FileName = nFileName
    On Error Resume Next
    Do While Mid(FileName, Len(FileName), 1) <> "\"
        FileName = Mid(FileName, 1, Len(FileName) - 1)
    Loop
    If FileName = "" Then
        GetShortName = nFileName
    Else
        GetShortName = Mid(nFileName, Len(FileName) + 1, Len(nFileName) - Len(FileName))
    End If
    Exit Function
End Function


Sub Main()
    AppName = "WebServerApp " + CStr(App.Major) + CStr(App.Minor)
    
    indexPage = GetSetting(AppName, "Settings", "IndexPage", "index.html")
    indexDir = GetSetting(AppName, "Settings", "IndexDir", _
        AddASlash(App.Path) + "www\")
    StartOn = GetSetting(AppName, "Settings", "StartOn", "1")
    SendIP = GetSetting(AppName, "Settings", "SendIP", "0")
    IPMailRcpt = GetSetting(AppName, "Settings", "IPMailRcpt", _
        "somebody@somewhere.com")
    LogFile = GetSetting(AppName, "Settings", "LogFile", _
        AddASlash(App.Path) + "www\serverlog.txt")
    WriteLog = GetSetting(AppName, "Settings", "WriteLog", "1")
    Counter = GetSetting(AppName, "Settings", "Counter", "0")
    UseCounter = GetSetting(AppName, "Settings", "UseCounter", "1")
    UseGestBook = GetSetting(AppName, "Settings", "UseGestBook", "1")
    Intro = GetSetting(AppName, "Settings", "Intro", "1")
    'StartUpLaunch = GetSetting(AppName, "Settings", "StartUpLaunch", "0")
    AlwaysOnTop = GetSetting(AppName, "Settings", "AlwaysOnTop", "0")
    
    frmMain.Show
End Sub

Sub SendRequested(sckIndex As Long, NameToSend As String)
    On Error Resume Next
    Static NumOfPic As Long
    Dim pageTitle As String, DownloadFile As Boolean, FileAttributes As Integer
    Dim Temp As String, OrigPath As String, Img As String, DFlag As String
    Dim IconID As Long, Word As Long, DataToSend As String, Extens As String
    Dim PicName As String, GifName As String, cGif As New GIF
    
    NameToSend = UnformatString(NameToSend)
    DownloadFile = False
    Word = 1
    If InStr(1, NameToSend, "o0o-download-o0o") > 0 Then DownloadFile = True
    NameToSend = Trim(Mid$(NameToSend, InStr(1, NameToSend, "=") + 1, InStr(InStr(1, NameToSend, "=") + 1, NameToSend, ";") - 2))
    NameToSend = Mid(NameToSend, 1, Len(NameToSend) - 1)
        
    Temp = Dir(NameToSend, vbDirectory Or vbHidden Or vbNormal Or _
        vbReadOnly Or vbSystem Or vbVolume)
    
    '   Por aqui (no lugar de "Exit sub") uma mensagem de arquivo inexistente...
    If Temp = "" Then
        Exit Sub
    End If
    
    If DownloadFile = True Then
        Debug.Print "Sending file: " + NameToSend
        Dim Exte As String
        Exte = UCase(GetExtension(NameToSend))
        If Exte = "TXT" Or Exte = "INI" Or Exte = "C" Or Exte = "CPP" Or _
            Exte = "H" Or Exte = "HXX" Or Exte = "INC" Or Exte = "BAS" Or _
            Exte = "FRM" Or Exte = "VBP" Or Exte = "LIC" Or Exte = "BAT" Then
            Pre = "HTTP/1.1 200 OK" + vbCrLf
            Pre = Pre + "Server: Final Web Server" + vbCrLf
            Pre = Pre + "Content-Type: text/plain" + vbCrLf
            Pre = Pre + "Accept-Ranges: bytes" + vbCrLf
            Pre = Pre + vbCrLf + vbCrLf

            frmMain.sckServer(sckIndex).SendData Pre + FormatPage(GetFile(NameToSend))
        Else
            frmMain.sckServer(sckIndex).SendData FormatPage(GetFile(NameToSend))
        End If
        Exit Sub
    End If
    
    DataToSend = "<html>" + vbCrLf
    DataToSend = DataToSend + "<head>" + vbCrLf
    DataToSend = DataToSend + "<title>Sending contents of " + NameToSend + " </title>" + vbCrLf
    DataToSend = DataToSend + "</head>" + vbCrLf
    DataToSend = DataToSend + "<body bgcolor=""#FFFFFF"">" + vbCrLf
    
    Img = "&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
            " src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
            "/open.gif" + Chr$(34) + " width=" + Chr$(34) + "32" + _
            Chr$(34) + " height=" + Chr$(34) + "32" + Chr$(34) + ">"
    
    DataToSend = DataToSend + "<h1>" + Img + "&nbsp;Sending contents of " + NameToSend + " ...</h1>" + vbCrLf
    DataToSend = DataToSend + "<hr>" + vbCrLf
    
    OrigPath = AddASlash(GetPath(NameToSend))
    'OrigPath = NameToSend
    Debug.Print NameToSend
    NameToSend = Temp
    
    DataToSend = DataToSend + "<font face=" + Chr$(34) + "Courier New" + Chr$(34) + ">" + vbCrLf
    DataToSend = DataToSend + "<table border=""0"" width=""100%"" cellspacing=""1"">" + vbCrLf
    
    If Not Exists(AddASlash(indexDir) + "extensions\") Then _
        MkDir (AddASlash(indexDir) + "extensions\")
    Kill AddASlash(indexDir) + "extensions\pic*.gif"
    If Err = 53 Then
        Err = 0
    End If
    Do While NameToSend <> ""
    
        DataToSend = DataToSend + "<tr>" + vbCrLf
        If NameToSend = "." Then
            Temp = Mid(OrigPath, 1, Len(OrigPath) - 1)
        ElseIf NameToSend = ".." Then
            Temp = Mid(OrigPath, 1, Len(OrigPath) - 1)
            Do While Mid(Temp, Len(Temp), 1) <> "\"
                Temp = Mid(Temp, 1, Len(Temp) - 1)
            Loop
            Temp = Mid(Temp, 1, Len(Temp) - 1)
        Else
            Temp = OrigPath + NameToSend
        End If
    
        Debug.Print "Creating link to: " + NameToSend
        If Not GetAttr(OrigPath + NameToSend) And vbDirectory Then
            Extens = GetExtension(Trim(OrigPath + NameToSend))
            If Extens = "" Or UCase(Extens) = "EXE" Or UCase(Extens) = "LNK" Or _
                UCase(Extens) = "CPL" Or UCase(Extens) = "ICO" Or _
                UCase(Extens) = "ANI" Or UCase(Extens) = "CUR" Or _
                UCase(Extens) = "SCR" Or UCase(Extens) = "DRV" Or _
                Not Exists(AddASlash(indexDir) + _
                "extensions\" + Extens + ".gif") Then
                NumOfPic = NumOfPic + 1
                'frmMain.Pic.Picture = Nothing
                frmMain.Aux.Cls
                'frmMain.Pic.Picture = frmMain.Pic.Image
                'frmMain.Pic.Refresh
                Call DrawIcon(frmMain.Aux.hDc, 0, 0, _
                    ExtractAssociatedIcon(App.hInstance, _
                    OrigPath + NameToSend, Word))
                'MsgBox OrigPath + NameToSend
                frmMain.Aux.Refresh
                DoEvents
                If Extens = "" Or UCase(Extens) = "EXE" Or _
                UCase(Extens) = "CPL" Or UCase(Extens) = "LNK" Or _
                UCase(Extens) = "ANI" Or UCase(Extens) = "CUR" Or _
                UCase(Extens) = "ICO" Or UCase(Extens) = "SCR" Or _
                UCase(Extens) = "DRV" Then
                    PicName = "pic" + CStr(NumOfPic) + ".bmp"
                    GifName = "pic" + CStr(NumOfPic) + ".gif"
                Else
                    PicName = Extens + ".bmp"
                    GifName = Extens + ".gif"
                End If
                
                frmMain.pic.Picture = frmMain.Aux.Image
                'frmMain.Pic.Refresh
                Call cGif.SaveGIF(frmMain.pic.Image, AddASlash(indexDir) + _
                    "extensions\" + GifName, frmMain.pic.hDc)
                frmMain.pic.Picture = Nothing
                PicName = GifName
                'SavePicture frmMain.Pic.Image, AddASlash(indexDir) + "extensions\" + PicName
                DoEvents
                'If Not Exists(AddASlash(indexDir) + "extension\" + PicName) Then
                '    PicName = "unknow.gif"
                'Else
                    PicName = "extensions/" + PicName
                'End If
            
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/" + PicName + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            Else
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/extensions/" + Extens + ".gif" + Chr$(34) + "width=" + Chr$(34) _
                    + "32" + Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            End If
            
            DFlag = "o0o-download-o0o;"
        Else
            If NameToSend = ".." Then
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/back.gif" + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            Else
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/folder.gif" + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            End If

            Temp = Temp + "\"
            DFlag = ""
        End If
            
        'Img = "<p>&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
            "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
            "/document.gif" + Chr$(34) + " width=" + Chr$(34) + "32" + _
            Chr$ (34) + " height=" + Chr$(34) + "32" + Chr$(34) + ">"

        Temp = ReplaceStr(Temp, "\", "%5c")
        Temp = ReplaceStr(Temp, " ", "+")
        
        If Not GetAttr(OrigPath + NameToSend) And vbDirectory Then
            DataToSend = DataToSend + "<td width=""55%"">" + vbCrLf
            DataToSend = DataToSend + Img
            DataToSend = DataToSend + "&nbsp;<a href=" + Chr$(34) + _
                "http://" + CStr(frmMain.sckServer(0).LocalIP) + "/" + _
                "list;" + DFlag + "file=" + Temp + ";" + _
                Chr$(34) + ">" + NameToSend + "</a>"
            DataToSend = DataToSend + "</td>"
            DataToSend = DataToSend + "<td width=""17%"">" + _
                CStr(FileLen(OrigPath + NameToSend)) + " bytes"
            DataToSend = DataToSend + "</td>" + vbCrLf
            DataToSend = DataToSend + "<td width=""28%"">" + _
                CStr(FileDateTime(OrigPath + NameToSend)) + "</td>" + vbCrLf
        Else
            DataToSend = DataToSend + "<td width=""100%"" colspan=""3"">"
            DataToSend = DataToSend + Img
            DataToSend = DataToSend + "&nbsp;<a href=" + Chr$(34) + _
                "http://" + CStr(frmMain.sckServer(0).LocalIP) + "/" + _
                "list;" + DFlag + "file=" + Temp + ";" + _
                Chr$(34) + ">" + NameToSend + "</a>"
            DataToSend = DataToSend + "</td>" + vbCrLf
        End If
        DataToSend = DataToSend + "</tr>" + vbCrLf
        NameToSend = Dir
    Loop
    
    DataToSend = DataToSend + "</table>" + vbCrLf
    DataToSend = DataToSend + "</font>" + vbCrLf
    DataToSend = DataToSend + "</body>" + vbCrLf
    DataToSend = DataToSend + "</html>" + vbCrLf
    
    Debug.Print "Sending formatted page...";
    frmMain.sckServer(sckIndex).SendData DataToSend
    Debug.Print "Done."
End Sub


Sub SendInfo(sckIndex As Long, NameToSend As String)
    On Error Resume Next
    Static NumOfPic As Long
    Dim pageTitle As String, DownloadFile As Boolean, FileAttributes As Integer
    Dim Temp As String, OrigPath As String, Img As String, DFlag As String
    Dim IconID As Long, Word As Long, DataToSend As String, Extens As String
    Dim PicName As String
    
    NameToSend = UnformatString(NameToSend)
    DownloadFile = False
    Word = 1
    If InStr(1, NameToSend, "o0o-download-o0o") > 0 Then DownloadFile = True
    NameToSend = Trim(Mid$(NameToSend, InStr(1, NameToSend, "=") + 1, InStr(InStr(1, NameToSend, "=") + 1, NameToSend, ";") - 2))
    NameToSend = Mid(NameToSend, 1, Len(NameToSend) - 1)
        
    Temp = Dir(NameToSend, vbDirectory Or vbHidden Or vbNormal Or _
        vbReadOnly Or vbSystem Or vbVolume)
    
    '   Por aqui (no lugar de "Exit sub") uma mensagem de arquivo inexistente...
    If Temp = "" Then
        Exit Sub
    End If
    
    If DownloadFile = True Then
        Debug.Print "Sending file: " + NameToSend
        frmMain.sckServer(sckIndex).SendData FormatPage(GetFile(NameToSend))
        Exit Sub
    End If
    
    DataToSend = "<html>" + vbCrLf
    DataToSend = DataToSend + "<head>" + vbCrLf
    DataToSend = DataToSend + "<title>Sending contents of " + NameToSend + " </title>" + vbCrLf
    DataToSend = DataToSend + "</head>" + vbCrLf
    DataToSend = DataToSend + "<body>" + vbCrLf
    
    Img = "&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
            " src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
            "/open.gif" + Chr$(34) + " width=" + Chr$(34) + "32" + _
            Chr$(34) + " height=" + Chr$(34) + "32" + Chr$(34) + ">"
    
    DataToSend = DataToSend + "<h1>" + Img + "&nbsp;Sending contents of " + NameToSend + " ...</h1>" + vbCrLf
    DataToSend = DataToSend + "<hr>" + vbCrLf
    
    OrigPath = AddASlash(GetPath(NameToSend))
    'OrigPath = NameToSend
    Debug.Print NameToSend
    NameToSend = Temp
    
    DataToSend = DataToSend + "<font face=" + Chr$(34) + "Courier New" + Chr$(34) + ">" + vbCrLf
    DataToSend = DataToSend + "<table border=""0"" width=""100%"" cellspacing=""1"">" + vbCrLf
    
    If Not Exists(AddASlash(indexDir) + "extensions\") Then _
        MkDir (AddASlash(indexDir) + "extensions\")
    Kill AddASlash(indexDir) + "extensions\pic*.bmp"
    If Err = 53 Then
        Err = 0
    End If
    Do While NameToSend <> ""
    
        DataToSend = DataToSend + "<tr>" + vbCrLf
        If NameToSend = "." Then
            Temp = Mid(OrigPath, 1, Len(OrigPath) - 1)
        ElseIf NameToSend = ".." Then
            Temp = Mid(OrigPath, 1, Len(OrigPath) - 1)
            Do While Mid(Temp, Len(Temp), 1) <> "\"
                Temp = Mid(Temp, 1, Len(Temp) - 1)
            Loop
            Temp = Mid(Temp, 1, Len(Temp) - 1)
        Else
            Temp = OrigPath + NameToSend
        End If
    
        Debug.Print "Creating link to: " + NameToSend
        If Not GetAttr(OrigPath + NameToSend) And vbDirectory Then
            Extens = GetExtension(Trim(OrigPath + NameToSend))
            If Extens = "" Or UCase(Extens) = "EXE" Or UCase(Extens) = "LNK" Or _
                UCase(Extens) = "CPL" Or UCase(Extens) = "ICO" Or _
                UCase(Extens) = "ANI" Or UCase(Extens) = "CUR" Or _
                UCase(Extens) = "SCR" Or UCase(Extens) = "DRV" Or _
                Not Exists(AddASlash(indexDir) + _
                "extensions\" + Extens + ".bmp") Then
                NumOfPic = NumOfPic + 1
                frmMain.pic.Cls
                Call DrawIcon(frmMain.pic.hDc, 0, 0, _
                    ExtractAssociatedIcon(App.hInstance, _
                    OrigPath + NameToSend, Word))
                'MsgBox OrigPath + NameToSend
                frmMain.pic.Refresh
                DoEvents
                If Extens = "" Or UCase(Extens) = "EXE" Or _
                UCase(Extens) = "CPL" Or UCase(Extens) = "LNK" Or _
                UCase(Extens) = "ANI" Or UCase(Extens) = "CUR" Or _
                UCase(Extens) = "ICO" Or UCase(Extens) = "SCR" Or _
                UCase(Extens) = "DRV" Then
                    PicName = "pic" + CStr(NumOfPic) + ".bmp"
                Else
                    PicName = Extens + ".bmp"
                End If
                
                SavePicture frmMain.pic.Image, AddASlash(indexDir) + "extensions\" + PicName
                DoEvents
                'If Not Exists(AddASlash(indexDir) + "extension\" + PicName) Then
                '    PicName = "unknow.gif"
                'Else
                    PicName = "extensions/" + PicName
                'End If
            
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/" + PicName + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            Else
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/extensions/" + Extens + ".bmp" + Chr$(34) + "width=" + Chr$(34) _
                    + "32" + Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            End If
            
            DFlag = "o0o-download-o0o;"
        Else
            If NameToSend = ".." Then
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/back.gif" + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            Else
                Img = "&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
                    "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
                    "/folder.gif" + Chr$(34) + "width=" + Chr$(34) + "32" + _
                    Chr$(34) + "height=" + Chr$(34) + "32" + Chr$(34) + ">"
            End If

            Temp = Temp + "\"
            DFlag = ""
        End If
            
        'Img = "<p>&nbsp;&nbsp;&nbsp;<img border=" + Chr$(34) + "0" + Chr$(34) + _
            "src=" + Chr$(34) + "http://" + CStr(frmMain.sckServer(0).LocalIP) + _
            "/document.gif" + Chr$(34) + " width=" + Chr$(34) + "32" + _
            Chr$ (34) + " height=" + Chr$(34) + "32" + Chr$(34) + ">"

        Temp = ReplaceStr(Temp, "\", "%5c")
        Temp = ReplaceStr(Temp, " ", "+")
        
        If Not GetAttr(OrigPath + NameToSend) And vbDirectory Then
            DataToSend = DataToSend + "<td width=""55%"">" + vbCrLf
            DataToSend = DataToSend + Img
            DataToSend = DataToSend + "&nbsp;<a href=" + Chr$(34) + _
                "http://" + CStr(frmMain.sckServer(0).LocalIP) + "/" + _
                "list;" + DFlag + "file=" + Temp + ";" + _
                Chr$(34) + ">" + NameToSend + "</a>"
            DataToSend = DataToSend + "</td>"
            DataToSend = DataToSend + "<td width=""17%"">" + _
                CStr(FileLen(OrigPath + NameToSend)) + " bytes"
            DataToSend = DataToSend + "</td>" + vbCrLf
            DataToSend = DataToSend + "<td width=""28%"">" + _
                CStr(FileDateTime(OrigPath + NameToSend)) + "</td>" + vbCrLf
        Else
            DataToSend = DataToSend + "<td width=""100%"" colspan=""3"">"
            DataToSend = DataToSend + Img
            DataToSend = DataToSend + "&nbsp;<a href=" + Chr$(34) + _
                "http://" + CStr(frmMain.sckServer(0).LocalIP) + "/" + _
                "list;" + DFlag + "file=" + Temp + ";" + _
                Chr$(34) + ">" + NameToSend + "</a>"
            DataToSend = DataToSend + "</td>" + vbCrLf
        End If
        DataToSend = DataToSend + "</tr>" + vbCrLf
        NameToSend = Dir
    Loop
    
    DataToSend = DataToSend + "</table>" + vbCrLf
    DataToSend = DataToSend + "</font>" + vbCrLf
    DataToSend = DataToSend + "</body>" + vbCrLf
    DataToSend = DataToSend + "</html>" + vbCrLf
    
    Debug.Print "Sending formatted page...";
    frmMain.sckServer(sckIndex).SendData DataToSend
    Debug.Print "Done."
End Sub


Public Function AddASlash(InString As String) As String
    If Mid(InString, Len(InString), 1) <> "\" Then
        AddASlash = InString & "\"
    Else
        AddASlash = InString
    End If
End Function

Public Function ReplaceStr(ByVal strMain As String, strFind As String, strReplace As String) As String
'Thsi is the same thing as the Replace function in vb6.  I added this
'for those of you using vb5.  This was NOT written by me, it was written by
' someone named 'dos'.  He's a great programmer, visit his webpage @
' http://hider.com/dos

    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    ReplaceStr$ = strNew$
End Function

Function FormatString(Name As String) As String
    Name = ReplaceStr(Name, " ", "+")
    Name = ReplaceStr(Name, "<br>", "%0D%0A")
    Name = ReplaceStr(Name, "!", "%21")
    Name = ReplaceStr(Name, "&quot;", "%22")
    Name = ReplaceStr(Name, "§", "%A7")
    Name = ReplaceStr(Name, "$", "%24")
    Name = ReplaceStr(Name, "%", "%25")
    Name = ReplaceStr(Name, "&", "%26")
    Name = ReplaceStr(Name, "/", "%2F")
    Name = ReplaceStr(Name, "(", "%28")
    Name = ReplaceStr(Name, ")", "%29")
    Name = ReplaceStr(Name, "=", "%3D")
    Name = ReplaceStr(Name, "?", "%3F")
    Name = ReplaceStr(Name, "²", "%B2")
    Name = ReplaceStr(Name, "³", "%B3")
    Name = ReplaceStr(Name, "{", "%7B")
    Name = ReplaceStr(Name, "[", "%5B")
    Name = ReplaceStr(Name, "]", "%5D")
    Name = ReplaceStr(Name, "}", "%7D")
    Name = ReplaceStr(Name, "\", "%5C")
    Name = ReplaceStr(Name, "ß", "%DF")
    Name = ReplaceStr(Name, "#", "%23")
    Name = ReplaceStr(Name, "'", "%27")
    Name = ReplaceStr(Name, ":", "%3A")
    Name = ReplaceStr(Name, ",", "%2C")
    Name = ReplaceStr(Name, ";", "%3B")
    Name = ReplaceStr(Name, "`", "%60")
    Name = ReplaceStr(Name, "~", "%7E")
    Name = ReplaceStr(Name, "+", "%2B")
    Name = ReplaceStr(Name, "´", "%B4")
    FormatString = Name
End Function

Public Function Exists(strFName As String) As Boolean
    On Error Resume Next
    Dim IsThere As Long
    Dim Buffer As OFSTRUCT

    IsThere = OpenFile(strFName, Buffer, OF_EXIST)
    If IsThere < 0 Then
        GoTo CheckForError
        Else
        Open strFName For Input As #1
        If LOF(1) = 0 Then
            Exists = False
            Close #1
            Exit Function
        End If
        Close #1
        Exists = True
    End If
CheckForError:
    IsThere = Buffer.nErrCode
    If IsThere = 3 Then
        Exists = False
    End If
    Open strFName For Input As #1
        If LOF(1) = 0 Then
            Exists = False
            Close #1
            Exit Function
        End If
    Exists = True
    Close #1
End Function


Public Function GetFile(FileName As String) As String
On Error Resume Next
Dim FileNumber
Dim TextData
    FileNumber = FreeFile
    TextData = ""
    If Exists(FileName) Then
        If Len(FileName) Then
            Open FileName For Binary As #FileNumber
                TextData = Input(LOF(FileNumber), #FileNumber)
                DoEvents
            Close #FileNumber
        End If
        GetFile = TextData
    Else
        GetFile = ""
    End If
End Function

Function UnformatString(Name As String) As String
    Dim Pos As Long, Temp As String
    Name = ReplaceStr(Name, "+", " ")
    Name = ReplaceStr(Name, "%0D%0A", vbCrLf)
    Pos = InStr(1, Name, "%")
    Do While Pos <> 0
        Temp = Mid(Name, 1, Pos - 1)
        Temp = Temp + CStr(Chr("&H" + Mid(Name, Pos + 1, 2)))
        Temp = Temp + Mid(Name, Pos + 3, Len(Name) - Pos - 2)
        Name = Temp
        Pos = InStr(Pos + 1, Name, "%")
    Loop
    UnformatString = Name
End Function


