VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2160
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
      RequestTimeout  =   5
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3612
   End
   Begin VB.Menu mnuBot 
      Caption         =   "&Bot"
      Begin VB.Menu mnuBotClearOutput 
         Caption         =   "&Clear output"
      End
      Begin VB.Menu mnuBotDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBotDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuBotDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBotQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Type person
 nick As String
 username As String
 ip As String
 command As String
 to As String
 what As String
 action As String
End Type

Private Type user
 nick As String
 username As String
 ip As String
 level As Integer
 connected As Boolean
 pass As String
 mask As String
 love As Integer
End Type

Private Type lurker
 nick As String
 mode As String
End Type

Private Type channel
 name As String
 my_mode As String
 mode As String
 topic As String
 notice_message As String
 auto_greet As Boolean
 greet_msg As String
 join_msg As String
 part_msg As String
 auto_join As Boolean
 people() As lurker
End Type

Private Type self
 nick As String
 oldnick As String
 default_msg_reply As String
 version As String
End Type

Private Type server
 name As String
 port As Integer
 connected As Boolean
End Type

Dim bot As self

Dim channels() As channel
Dim servers() As server
Dim users() As user
Dim love(100) As Integer
Dim nick(100) As String
Dim currentserver As Integer
Dim inifile As String
Dim stayconnected As Boolean

Private Sub Form_Load()
 Dim ret As String, NC As Long, tempstr As String
 inifile = "c:\test.ini"
 stayconnected = True
 
 bot.nick = "sepha"
 bot.default_msg_reply = "Error: unknown command"
 bot.version = "Sephabot v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
 
 x = 0
 ret = ""
 Do Until ret = "end"
  ret = String(255, 0)
  tempstr = "server" + String(2 - Len(CStr(x)), "0") + CStr(x)
  NC = GetPrivateProfileString("servers", tempstr, "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   ReDim servers(x) As server
   x = x + 1
  End If
 Loop
 For x = 0 To UBound(servers)
  ret = String(255, 0)
  tempstr = "port" + String(2 - Len(CStr(x)), "0") + CStr(x)
  NC = GetPrivateProfileString("servers", tempstr, "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   servers(x).port = CInt(ret)
  End If
  ret = String(255, 0)
  tempstr = "server" + String(2 - Len(CStr(x)), "0") + CStr(x)
  NC = GetPrivateProfileString("servers", tempstr, "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   servers(x).name = ret
  End If
 Next x
 
 currentserver = -1
 
 x = 0
 ret = ""
 Do Until ret = "end"
  ret = String(255, 0)
  tempstr = "channel" + String(2 - Len(CStr(x)), "0") + CStr(x)
  NC = GetPrivateProfileString(tempstr, "name", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   ReDim channels(x) As channel
   x = x + 1
  End If
 Loop
 
 For x = 0 To UBound(channels)
  tempstr = "channel" + String(2 - Len(CStr(x)), "0") + CStr(x)
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "name", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).name = ret
  End If
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "topic", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).topic = ret
  End If
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "mode", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).mode = ret
  End If
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "auto_join", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   If LCase(ret) = "true" Then
    channels(x).auto_join = True
   Else
    channels(x).auto_join = False
   End If
  End If
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "auto_greet", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   If LCase(ret) = "true" Then
    channels(x).auto_greet = True
   Else
    channels(x).auto_greet = False
   End If
  End If
 
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "join_msg", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).join_msg = ret
  End If

  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "greet_msg", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).greet_msg = ret
  End If
 
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "part_msg", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).part_msg = ret
  End If
 
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "notice_message", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   channels(x).notice_message = ret
  End If
 
 Next x
 
 LoadUsers
 
 ReConnect
 Form1.Caption = "Bot: " + Chr(34) + bot.nick + Chr(34) + " not connected"
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 With Text1
  .Left = 0
  .Top = 0
  .Width = Me.Width - 6 * (Screen.TwipsPerPixelY)
  .Height = Me.Height - 44 * (Screen.TwipsPerPixelX)
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Winsock1.State = 7 Then
  SendRawData "QUIT :Offline for testing..." + vbCrLf
 End If
 Open "c:\irc.log" For Append As #1
  Print #1, Text1.Text
 Close #1
 End
End Sub

Private Sub mnuBotClearOutput_Click()
  Open "c:\irc.log" For Append As #1
   Print #1, Text1.Text
  Close #1
  Text1.Text = "Cleared by request" + vbCrLf
End Sub

Private Sub mnuBotDisconnect_Click()
 stayconnected = False
 currentserver = -1
 SendRawData "QUIT :disconnecting"
End Sub

Private Sub mnuBotQuit_Click()
 stayconnected = False
 Unload Me
End Sub

Private Sub Text1_Change()
 Text1.SelStart = Len(Text1.Text)
 If Len(Text1.Text) > 10000 Then
  Open "c:\irc.log" For Append As #1
   Print #1, Text1.Text
  Close #1
  Text1.Text = "Clearing time" + vbCrLf
 End If
End Sub

Private Sub Timer1_Timer()
 ReConnect
 Timer1.Enabled = False
End Sub

Private Sub Winsock1_Close()
 Text1.Text = Text1.Text + "LOCAL-Connection Closed" + vbCrLf
 Form1.Caption = "Bot: " + Chr(34) + bot.nick + Chr(34) + " not connected"
 ReConnect
End Sub

Private Sub Winsock1_Connect()
 Text1.Text = Text1.Text + "LOCAL-Connected" + vbCrLf
 SendRawData "USER sepha " + Chr(34) + Chr(34) + " " + Chr(34) + Winsock1.RemoteHost + Chr(34) + " Sepha" + vbCrLf + "NICK " + bot.nick
 Form1.Caption = "Bot: " + Chr(34) + bot.nick + Chr(34) + " server:" + Winsock1.RemoteHost
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Dim str As String
 Dim who As person
 Dim victim As String
 Dim whatsay As String
 Dim chan As String
 Dim shouldspeak As Boolean
 Dim currentchannum As Integer
 Dim usernum As Integer
 
 Winsock1.GetData str
 str = Trim(str)
 str = Replace(str, vbCrLf, Chr(0))
 str = Replace(str, vbCr, Chr(0))
 str = Replace(str, vbLf, Chr(0))
 comands = Split(str, Chr(0))

 For Each comand In comands
  comand = Trim(comand)
  If comand <> "" And Left(comand, Len("ping")) <> "PING" Then
   Text1.Text = Text1.Text + comand + vbCrLf
  End If
  
  If Left(comand, Len("ERROR")) = "ERROR" Then
   If Winsock1.State = 7 Then
    Winsock1.Close
   End If
   If InStr(comand, "throttled") > 0 Then
    Timer1.Interval = 60000
    Timer1.Enabled = True
   Else
   If Winsock1.State = 0 Then
    ReConnect
   End If
   End If
  End If
  
  If LCase(Left(comand, Len(":" + LCase(Winsock1.RemoteHost)))) = ":" + LCase(Winsock1.RemoteHost) Then
   ServerResponses comand
  ElseIf InStr(comand, "@") > 0 And Mid(comand, 2, Len(servers(currentserver).name)) <> servers(currentserver).name Then
   who.nick = Mid(comand, 2, InStr(comand, "!") - 2)
   who.username = Mid(comand, Len(who.nick) + 3, InStr(comand, "@") - Len(who.nick) - 3)
   who.ip = Mid(comand, Len(who.nick) + 4 + Len(who.username), InStr(comand, " ") - (Len(who.nick) + 4 + Len(who.username)))
   who.action = Mid(comand, InStr(comand, " ") + 1, InStr(InStr(comand, " ") + 1, comand, " ") - InStr(comand, " ") - 1)
   who.what = Mid(comand, InStr(comand, " :") + 2)
   If who.nick = bot.nick Then
    If who.action = "JOIN" Then
     who.to = who.what
     For channum = 1 To UBound(channels)
      If who.to = channels(channum).name Then
       If channels(channum).join_msg <> "" Then
        whatsay = channels(channum).join_msg
       End If
      End If
     Next channum
     If whatsay = "" Then
      whatsay = channels(0).join_msg
     End If
     whatsay = Replace(whatsay, "$channel", who.to)
     Speak whatsay, who.to
    End If
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
   Else
    If who.action = "PART" Then
     who.to = Mid(comand, InStr(comand, "PART") + 5)
     For channum = 1 To UBound(channels)
      If who.to = channels(channum).name Then
       If channels(channum).part_msg <> "" Then
        whatsay = channels(channum).part_msg
       End If
      End If
     Next channum
     If whatsay = "" Then
      whatsay = channels(0).part_msg
     End If
     whatsay = Replace(whatsay, "$nick", who.nick)
     whatsay = Replace(whatsay, "$channel", who.to)
     notice who.nick, whatsay
     SendRawData "NAMES " + who.to
    ElseIf who.action = "JOIN" Then
     who.to = Mid(comand, InStr(comand, "JOIN") + 6)
     whatsay = ""
     For channum = 1 To UBound(channels)
      If who.to = channels(channum).name Then
       If channels(channum).greet_msg <> "" Then
        whatsay = channels(channum).greet_msg
        currentchannum = channum
       End If
      End If
     Next channum
     If whatsay = "" Then
      whatsay = channels(0).greet_msg
     End If
     For x = 0 To UBound(users)
      If who.ip = users(x).ip And who.username = users(x).username And users(x).level < 2 And users(x).connected Then
       modechange "+o " + who.nick, who.to
       whatsay = "Greetings Master $nick"
       whatsay = Replace(whatsay, "$nick", who.nick)
       whatsay = Replace(whatsay, "$channel", who.to)
       If Not channels(currentchannum).auto_greet Then
        Speak whatsay, who.to
       End If
      End If
     Next x
     If InStr(who.ip, "aol.com") > 0 Then
      Speak "grr...", who.to
      action "doesn't like aol", who.to
     Else
      If channels(currentchannum).auto_greet Then
       Sleep 1000
       whatsay = Replace(whatsay, "$nick", who.nick)
       whatsay = Replace(whatsay, "$channel", who.to)
       Speak whatsay, who.to
      End If
     End If
            
     If channels(currentchannum).notice_message = "" Then
      notice who.nick, Replace(channels(0).notice_message, "$bot.nick", bot.nick)
     Else
      notice who.nick, Replace(channels(currentchannum).notice_message, "$bot.nick", bot.nick)
     End If
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
    ElseIf who.action = "KICK" Then
     temp = Split(comand, " ")
     who.to = temp(2)
     who.what = temp(3)
     If who.what = bot.nick And users(0).nick <> "" Then
      Speak who.nick + " kicked off of " + who.to + " because of " + Chr(34) + Mid(comand, InStr(comand, temp(4)) + 1) + Chr(34), master.nick
      joinchan who.to
     End If
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
    ElseIf who.action = "QUIT" Then
     For x = 0 To UBound(users)
      If LCase(users(x).nick) = LCase(who.nick) Then
       users(x).connected = False
       Exit For
      End If
     Next x
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
    ElseIf who.action = "NICK" Then
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
    ElseIf who.action = "MODE" Then
     temp = Split(comand, " ")
     who.to = temp(2)
     SendRawData "NAMES " + who.to
    Else
     who.to = Mid(comand, Len(":" + who.nick + "!" + who.username + "@" + who.ip + " " + who.action + "  "), InStr(Len(":" + who.nick + "!" + who.username + "@" + who.ip + " " + who.action + "  "), comand, " ") - Len(":" + who.nick + "!" + who.username + "@" + who.ip + " " + who.action + "  "))
    End If
'    Speak "Nick:" + who.nick + " Username:" + who.username + " IP:" + who.IP + " Action:" + who.action + " To:" + who.to + " WhatSaid:" + who.what, "#sepha"
    'CTCP replies
    
    If Left(who.what, 1) = Chr(1) And Right(who.what, 1) = Chr(1) Then
     who.what = Mid(who.what, 2, Len(who.what) - 3)
     If Left(who.what, Len("PING")) = "PING" Then
      notice who.nick, "Hey that tickles!"
     ElseIf who.what = "VERSION" Then
      notice who.nick, Chr(1) + "VERSION " + bot.version + Chr(1)
     ElseIf Left(who.what, Len("DCC")) = "DCC" Then
     End If
    End If
    
    'MSG replies
    If Left(who.to, Len("#")) <> "#" And who.action = "PRIVMSG" And Left(who.what, 1) <> Chr(1) Then
     tempint = 0
     For usernum = 0 To UBound(users)
      If LCase(who.nick) = LCase(users(usernum).nick) And LCase(who.ip) = LCase(users(usernum).ip) And users(usernum).connected Then
       MasterMsgResponses who, 0
       tempint = 1
       Exit For
      End If
     Next usernum
     If tempint = 0 Then
      If who.what = "hi" Then
       Speak "hi " + who.nick, who.nick
      ElseIf who.what = "hello" Then
       Speak "hello " + who.nick, who.nick
      ElseIf who.what = "bye" Then
       Speak "bye " + who.nick, who.nick
      ElseIf Left(who.what, Len("pass")) = "pass" Then
       Speak "the pass command is outdated, use " + Chr(34) + "id" + Chr(34) + " instead", who.nick
      ElseIf Left(who.what, Len("id")) = "id" Then
       temp = Split(LCase(who.what), " ")
       If UBound(temp) = 2 Then
        If temp(1) = "me" Then temp(1) = LCase(who.nick)
        For x = 0 To UBound(users)
 '       If temp(1) = users(x).pass And (who.nick + "!" + who.username + "@" + who.ip) Like users(x).mask Then
         If LCase(temp(1)) = LCase(users(x).nick) And temp(2) = users(x).pass Then
          users(x).ip = who.ip
          users(x).username = who.username
          users(x).nick = who.nick
          users(x).connected = True
          Speak "Hello master " + who.nick + ", you have level " + CStr(users(x).level) + ". I am ready to do your bidding.", who.nick
          Exit For
         End If
        Next x
       Else
        Speak "not enough prams", who.nick
        Speak "Usage: id me [userpass]", who.nick
       End If
      Else
       Speak bot.default_msg_reply, who.nick
      End If
     End If
    End If
    
    'public channel replies
    For x = 0 To 10
     who.what = Replace(who.what, Chr(x), "")
    Next x
    If (Left(who.what, Len(bot.nick)) = bot.nick) Then
     who.command = Mid(who.what, Len(bot.nick) + 2)
     temp = 0
     For usernum = 0 To UBound(users)
      If who.ip = users(usernum).ip And who.username = users(usernum).username And users(usernum).connected Then
       MasterChanCommands who, usernum
       temp = 1
       Exit For
      End If
     Next usernum
     If temp = 0 Then
      PublicCommands who, False
     End If
    ElseIf Left(who.what, Len(UCase("hi " + bot.nick))) = UCase("hi " + bot.nick) Then
     Speak "you don't have to shout", who.to
    ElseIf Left(LCase(who.what), Len(LCase("hi " + bot.nick))) = LCase("hi " + bot.nick) Then
     Speak "greetings " + who.nick, who.to
    ElseIf Left(LCase(who.what), Len(LCase("bye " + bot.nick))) = LCase("bye " + bot.nick) Then
     Speak "c-ya " + who.nick, who.to
    ElseIf Left(who.what, 1) = "!" Then
     If who.what = "!beer" Then
      action "pours " + who.nick + " a tall one", who.to
     ElseIf who.what = "!cheese" Then
      action " hurls a large block of Swiss cheese as " + who.nick + "'s head", who.to
     ElseIf Left(who.what, Len("-seen")) = "!seen" Then
      temp = Split(who.what, " ")
      If who.nick = CStr(temp(1)) Then
       Speak who.nick + ", are you haveing an identiy crisis?", who.to
      Else
       Speak "I haven't seen " + CStr(temp(1)) + ". But i'm blind so i haven't seen anyone!", who.to
      End If
     End If
    ElseIf Left(who.what, 1) = "-" Then
    ElseIf Left(who.what, 1) = "@" Then
     If who.what = "@list" Then
      notice who.nick, "! commands are:"
      notice who.nick, "!beer"
      notice who.nick, "!cheese"
      notice who.nick, "end of ! commands"
     End If
    End If
'   If InStr(who.what, "aol") Or InStr(who.what, "shit") Or InStr(who.what, "fuck") Then
'    Speak "Hey now " + who.nick + ", no cursing", who.to
'   End If
   End If
  ElseIf InStr(comand, "PING") > 0 Then
   temp = (Split(comand, " ")(1))
   SendRawData "PONG " + temp
  End If
 Next comand
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Text1.Text = Text1.Text + vbCrLf + "Winsock Error #" + CStr(Number) + ": " + Chr(34) + Description + Chr(34) + vbCrLf
 If stayconnected Then
  If Number = 11001 Or Number = 10061 Or Number = 10060 Or Number = 10054 Then
   ReConnect
  Else
   MsgBox "Winsock Error #" + CStr(Number) + ": " + Chr(34) + Description + Chr(34)
  End If
 End If
End Sub

Private Sub Winsock1_SendComplete()
'Text1.Text = Text1.Text + "LOCAL-Send complete" + vbCrLf
End Sub

Private Sub Speak(WhatToSay As String, chan As String)
 SendRawData "PRIVMSG " + chan + " :" + Trim(WhatToSay)
End Sub

Private Sub notice(nick As String, str As String)
 SendRawData "NOTICE " + nick + " :" + str
End Sub

Private Sub action(str As String, chan As String)
 Speak Chr(1) + "ACTION " + str + Chr(1), chan
End Sub

Private Sub modechange(str As String, chan As String)
 SendRawData "MODE " + chan + " " + str
End Sub

Private Sub KickEm(nick As String, chan As String, why As String)
 SendRawData "KICK " + chan + " " + nick + " :" + why
End Sub

Private Sub topic(str As String, chan As String)
 SendRawData "TOPIC " + chan + " :" + str
End Sub

Private Sub joinchan(chan As String)
 SendRawData "JOIN " + chan
End Sub

Private Sub partchan(chan As String)
 SendRawData "PART " + chan
End Sub

Private Sub changenick(nick As String)
 SendRawData "NICK " + nick
 bot.oldnick = bot.nick
 bot.nick = nick
End Sub

Private Sub invite(nick As String, chan As String)
 SendRawData "INVITE " + nick + " " + chan
End Sub

Private Sub SendRawData(str As String)
 If Winsock1.State = 8 Then
  Winsock1.SendData str + vbCrLf
  Text1.Text = Text1.Text + str + vbCrLf
  DoEvents
 End If
End Sub

Private Sub PublicCommands(who As person, isuser As Boolean, Optional usernum As Integer)
 Randomize Timer
 Dim barks(4) As String
 barks(0) = "ARF!"
 barks(1) = "WOOF!"
 barks(2) = "MOO!"
 barks(4) = "YELP!"
 
 On Error Resume Next
 If who.command = "speak" Then
  Speak barks(Int(Rnd * UBound(barks)) + 1), who.to
 ElseIf who.command = "sleep" Then
  notice who.nick, "but i'm not tired!"
 ElseIf who.command = "jump" Then
  action "jumps", who.to
 ElseIf who.command = "good boy!" And Not isuser Then
  Speak "grr..", who.to
  action "doesn't like strangers messing with his silky fur", who.to
 ElseIf who.command = "bad boy!" Then
  Speak "ARF! ARF! ARF!", who.to
  action "jumps at " + who.nick + " but get's held back by " + users(0).nick, who.to
 ElseIf Left(LCase(who.command), Len("fetch")) = "fetch" Then
  temp = Split(who.command, " ")
  If UBound(temp) = 0 Then
   who.what = "stick"
  Else
   who.what = temp(1)
  End If
  fetch who
 ElseIf who.command = "help" Then
  notice who.nick, "You have access to [speak] [sleep] [help] [jump] [buzz] [about] [fetch]"
 ElseIf who.command = "about" Then
  notice who.nick, "i'm a vb bot useing winsock, my source is almost ready for download."
 ElseIf who.command = "turn it up" Then
  action "turns up the music", who.to
 ElseIf who.command = "turn it down" Then
  Speak who.nick + ", are you crazy?", who.to
 ElseIf who.command = "rubbers" Then
  Speak who.nick + " you don't need any, you can't pick up anyone!", who.to
 ElseIf who.command = "buzz" Then
  If master.nick <> "" Then
   Speak "Buzzing my master... please wait...", who.to
   Speak "You are needed in " + who.to + " by " + who.nick, master.nick
  End If
 ElseIf who.command = "help translate" Then
  notice who.nick, "Language translation available, but still BETA"
  notice who.nick, "Format = " + Chr(34) + bot.nick + " (source language)2(destination language) (text to translate)" + Chr(34)
  notice who.nick, "example = " + Chr(34) + bot.nick + " en2de Hello" + Chr(34)
  notice who.nick, "output would be " + Chr(34) + " <" + bot.nick + "> Translated: Hallo" + Chr(34)
  notice who.nick, "Languages available:"
  notice who.nick, "[en] english"
  notice who.nick, "[fr] french"
  notice who.nick, "[de] german"
 ElseIf Left(who.command, 5) = "es2en" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=es_en")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "en2es" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=en_es")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "en2fr" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=en_fr")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "fr2en" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=fr_en")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "en2de" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=en_de")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "de2en" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=de_en")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "fr2de" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=fr_de")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 ElseIf Left(who.command, 5) = "de2fr" Then
  If Inet1.StillExecuting Then
   Speak "Still processing last command", who.to
  Else
   intext = Replace(Mid(who.command, 7), " ", "+")
   outtext = Inet1.OpenURL("http://babelfish.altavista.com/tr?doit=done&urltext=" + intext + "&lp=de_fr")
   outtext = Mid(outtext, InStr(outtext, "textarea") + Len("textarea"))
   outtext = Mid(outtext, InStr(outtext, ">") + Len(">"))
   outtext = Mid(outtext, 1, InStr(outtext, "<") - Len("<"))
   Speak "Translated: " + outtext, who.to
  End If
 Else
' notice who.nick, "I'm not sure what you want me to do..."
 End If
End Sub

Private Sub ServerResponses(ByVal comand As String)
 Dim whatsay As String, who As person
 
 If InStr(comand, server + " 376") > 0 Or InStr(comand, server + " 422") > 0 Then
  Open "c:\irc.log" For Append As #1
   Print #1, Text1.Text
  Close #1
  Text1.Text = ""
  For x = 1 To UBound(channels)
   If channels(x).auto_join Then
    joinchan channels(x).name
   End If
  Next x
 ElseIf InStr(LCase(comand), LCase(Winsock1.RemoteHost) + " 366") > 0 Then
  who.to = Mid(comand, InStr(comand, "366") + 3 + Len(bot.nick) + 2, InStr(comand, " :") - (InStr(comand, "366") + 3 + Len(bot.nick) + 2))
 ElseIf InStr(LCase(comand), LCase(Winsock1.RemoteHost) + " 482") > 0 Then
' who.to = Mid(comand, InStr(comand, "482") + 3 + Len(" " + bot.nick + " "), InStr(comand, " :") - (InStr(comand, "482") + 3 + Len(" " + bot.nick + " ")))
 ElseIf InStr(LCase(comand), LCase(Winsock1.RemoteHost) + " 353") > 0 Then
  people = Mid(comand, InStr(comand, " :") + 3)
  temp = Split(people, " ")
  who.to = Mid(comand, InStr(comand, "=") + 2, InStr(InStr(comand, "=") + 2, comand, " ") - InStr(comand, "=") - 2)
  If LCase(who.to) = LCase("#youngvb") And InStr(people, "@zerocide") > 0 And InStr(people, "@sepha") < 0 Then
   Speak "op jimmydean #youngvb", "zerocide"
  End If
  For channum = 1 To UBound(channels)
   If LCase(who.to) = LCase(channels(channum).name) Then
    ReDim Preserve channels(channum).people(UBound(temp))
    For tempint = 0 To UBound(temp)
     If Left(temp(tempint), 1) = "@" Then
     channels(channum).people(tempint).nick = Mid(temp(tempint), 2)
     channels(channum).people(tempint).mode = "@"
     ElseIf Left(temp(tempint), 1) = "+" Then
     channels(channum).people(tempint).nick = Mid(temp(tempint), 2)
     channels(channum).people(tempint).mode = "+"
     Else
     channels(channum).people(tempint).nick = temp(tempint)
     channels(channum).people(tempint).mode = ""
     End If
    Next tempint
    Exit For
   End If
  Next channum
  If people = "@" + bot.nick + "" Then
   who.to = Mid(comand, InStr(comand, "=") + 2, InStr((InStr(comand, "=") + 2), comand, " :") - (InStr(comand, "=") + 2))
   For channum = 1 To UBound(channels)
    If LCase(who.to) = LCase(channels(channum).name) Then
     If channels(channum).mode = "" Then
      SendRawData "MODE " + who.to + " +" + channels(0).mode
     Else
      SendRawData "MODE " + who.to + " +" + channels(channum).mode
     End If
     temptopic = channels(channum).topic
     temptopic = Replace(temptopic, "$bot.nick", bot.nick)
     temptopic = Replace(temptopic, "$channel", channels(channum).name)
     SendRawData "TOPIC " + who.to + " :" + temptopic
     If Len(temptopic) > 160 And users(0).nick <> "" Then
      Speak "WARNING: topic for channel " + channels(channum).name + " is too long!", master.nick
     End If
    End If
   Next channum
   If temptopic = "" Then
    SendRawData "MODE " + who.to + " +" + channels(0).mode
    temptopic = channels(0).topic
    temptopic = Replace(temptopic, "$bot.nick", bot.nick)
    temptopic = Replace(temptopic, "$channel", who.to)
    SendRawData "TOPIC " + who.to + " :" + temptopic
   End If
  End If
 ElseIf InStr(comand, "Nickname is already in use") > 0 Then
  bot.nick = bot.oldnick
  If users(0).nick <> "" Then
   Speak "NICK command failed", users(0).nick
  End If
 End If
End Sub

Private Sub MasterMsgResponses(who As person, usernum As Integer)
 Dim whatsay As String, chan As String, str As String, temp() As String, tempint As Integer
 If who.what = "clear" And users(usernum).level = 0 And users(usernum).connected Then
  Open "c:\irc.log" For Append As #1
   Print #1, Text1.Text
  Close #1
  Text1.Text = "Cleared by request" + vbCrLf
 ElseIf who.what = "sleep" Then
  stayconnected = False
  Unload Me
 ElseIf who.what = "load users" And users(usernum).level = 0 And users(usernum).connected Then
  LoadUsers
 ElseIf Left(who.what, Len("del user")) = "del user" And users(usernum).level = 0 And users(usernum).connected Then
  temp = Split(LCase(who.what), " ")
  tempint = 0
  For x = 0 To UBound(users)
   If LCase(users(x).nick) = LCase(temp(2)) Then
    For tempint = x To UBound(users) - 1
     users(x + 1) = users(x)
    Next tempint
    ReDim Preserve users(UBound(users) - 1) As user
    tempint = 1
    Exit For
   End If
  Next x
  If tempint = 0 Then
   Speak "User not found!", who.nick
  Else
   Speak "User " + temp(2) + " removed", who.nick
  End If
 ElseIf Left(who.what, Len("save users")) = "save users" And users(usernum).level = 0 And users(usernum).connected Then
  SaveUsers
  Speak "Users saved", who.nick
 ElseIf Left(who.what, Len("user pass")) = "user pass" And users(usernum).level = 0 And users(usernum).connected Then
  temp = Split(LCase(who.what), " ")
  tempint = 0
  For x = 0 To UBound(users)
   If LCase(users(x).nick) = LCase(temp(2)) Then
    users(x).pass = CStr(temp(3))
    tempint = 1
    Exit For
   End If
  Next x
  If tempint = 0 Then
   Speak "User not found!", who.nick
  Else
   Speak "User " + temp(2) + " pass set to " + temp(3), who.nick
  End If
 ElseIf Left(who.what, Len("add user")) = "add user" And users(usernum).level = 0 And users(usernum).connected Then
  temp = Split(LCase(who.what), " ")
  tempint = 0
  If UBound(temp) = 4 Then
   For x = 0 To UBound(users)
    If LCase(users(x).nick) = LCase(temp(2)) Then
     tempint = 1
     Exit For
    End If
   Next x
   If tempint = 0 Then
    tempint = UBound(users) + 1
    ReDim Preserve users(tempint)
    users(tempint).nick = temp(2)
    users(tempint).pass = temp(3)
    users(tempint).level = CInt(temp(4))
    Speak "User " + users(tempint).nick + " added.", who.nick
   Else
    Speak "User " + temp(2) + " already exists.", who.nick
   End If
  Else
   Speak "Missing params", who.nick
   Speak "add user (usernick) (userpass) (userlevel)", who.nick
  End If
 ElseIf who.what = "list users" And users(usernum).level = 0 And users(usernum).connected Then
  Speak "user|pass|level|mask|ip|username|connected|love", who.nick
  Speak "------------------------------------------", who.nick
  For x = 0 To UBound(users)
   Speak users(x).nick + "|" + users(x).pass + "|" + CStr(users(x).level) + "|" + users(x).mask + "|" + users(x).ip + "|" + users(x).username + "|" + CStr(users(x).connected) + "|" + CStr(users(x).love), who.nick
   Sleep 500
  Next x
 ElseIf Left(who.what, Len("say")) = "say" Then
  chanwhat = Mid(who.what, Len("say  "))
  chan = Mid(chanwhat, 1, InStr(chanwhat, " ") - 1)
  whatsay = Mid(chanwhat, InStr(chanwhat, " ") + 1)
  Speak whatsay, chan
 ElseIf Left(who.what, Len("action")) = "action" Then
  chanwhat = Mid(who.what, Len("action  "))
  chan = Mid(chanwhat, 1, InStr(chanwhat, " ") - 1)
  whatsay = Mid(chanwhat, InStr(chanwhat, " ") + 1)
  action whatsay, chan
 ElseIf Left(who.what, Len("join")) = "join" And users(usernum).level < 3 And users(usernum).connected Then
  chan = Mid(who.what, Len("join  "))
  joinchan chan
 ElseIf Left(who.what, Len("part")) = "part" And users(usernum).level < 3 And users(usernum).connected Then
  chan = Mid(who.what, Len("part  "))
  partchan chan
 ElseIf Left(who.what, Len("raw")) = "raw" And users(usernum).level < 2 And users(usernum).connected Then
  str = Mid(who.what, Len("raw  "))
  SendRawData str
 ElseIf Left(who.what, Len("mode")) = "mode" And users(usernum).level < 3 And users(usernum).connected Then
  temp = Split(who.what, " ")
  modechange CStr(temp(2)), CStr(temp(1))
 ElseIf Left(who.what, Len("nick")) = "nick" And users(usernum).level < 3 And users(usernum).connected Then
  str = Mid(who.what, Len("nick  "))
  changenick str
 ElseIf who.what = "remotehost" Then
  Speak Winsock1.RemoteHost, who.nick
 Else
  Speak bot.default_msg_reply, who.nick
 End If
End Sub

Private Sub ReConnect()
 If stayconnected Then
  If Winsock1.State <> 0 Then
   Winsock1.Close
  End If
  currentserver = currentserver + 1
  If currentserver > UBound(servers) Then
   currentserver = 0
  End If
  Text1.Text = Text1.Text + "Now trying server #" + CStr(currentserver) + ":" + servers(currentserver).name + ":" + CStr(servers(currentserver).port)
  Winsock1.Connect servers(currentserver).name, servers(currentserver).port
 End If
End Sub

Private Sub fetchthestick_old(who As person)
 Randomize Timer
  For i = 0 To UBound(nick)
   If nick(i) = who.to Then
    Cur = i
    Exit For
   End If
  Next i
  love(Cur) = love(Cur) + 1
  RndNb = Int(Rnd * 5)
  action "runs to get the " + who.what + " thrown by " & who.nick, who.to
  If (love(Cur) < 10) Then
   If RndNb = 0 Then
    action "wags his tail", who.to
   ElseIf RndNb = 1 Then
    action "jumps around " & who.nick, who.to
   ElseIf RndNb = 2 Then
    action "licks " & who.nick, who.to
   ElseIf RndNb = 3 Then
    action "yaps at " & who.nick, who.to
   ElseIf RndNb = 4 Then
    action "eats the " + who.what, who.to
   End If
  ElseIf (love(Cur) >= 10) And (love(Cur) < 30) Then
   If RndNb = 0 Then
    action "jumps at " & who.nick & " face and lick it", who.to
   ElseIf RndNb = 1 Then
    action "is starting to love " & who.nick, who.to
   ElseIf RndNb = 2 Then
    action "brings the wrong " + who.what + " back", who.to
   ElseIf RndNb = 3 Then
    action "runs around " & who.nick, who.to
   ElseIf RndNb = 4 Then
    action "loves when " & who.nick & " plays with him", who.to
   End If
  ElseIf (love(Cur) >= 30) Then
   If RndNb = 0 Then
    action "loves " & who.nick, who.to
   ElseIf RndNb = 1 Then
    action "licks " & who.nick & "everywhere", who.to
   ElseIf RndNb = 2 Then
    action "jumps everywhere", who.to
   ElseIf RndNb = 3 Then
    action "likes to play with " & who.nick, who.to
   ElseIf RndNb = 4 Then
    action "is very happy to play with " & who.nick, who.to
   End If
  End If
End Sub

Private Sub LoadUsers()
 Dim ret As String, tempstr As String, NC As Long
 x = 0
 ret = ""
 Do Until ret = "end"
  ret = String(255, 0)
  tempstr = "users" + String(2 - Len(CStr(x)), "0") + CStr(x)
  NC = GetPrivateProfileString(tempstr, "pass", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   ReDim users(x) As user
   x = x + 1
  End If
 Loop
 
 For x = 0 To UBound(users)
  tempstr = "users" + String(2 - Len(CStr(x)), "0") + CStr(x)
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "pass", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   users(x).pass = ret
  End If
  
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "mask", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   users(x).mask = ret
  End If
 
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "level", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   users(x).level = CInt(ret)
  End If
 
  ret = String(255, 0)
  NC = GetPrivateProfileString(tempstr, "nick", "end", ret, 255, inifile)
  If NC <> 0 Then ret = Left$(ret, NC)
  If ret <> "end" Then
   users(x).nick = ret
  End If
 Next x
End Sub

Private Sub fetch(who As person)
 Randomize Timer
 Dim love As Integer
 action "runs to fetch the " + who.what + " thrown by " + who.nick, who.to
 tempint = 0
 For x = 0 To UBound(users)
  If LCase(users(x).nick) = LCase(who.nick) Then
   users(x).love = users(x).love + 1
   If users(x).love > 30 Then
    users(x).love = 30
   End If
   love = users(x).love
   tempint = 1
   Exit For
  End If
 Next x
 If tempint = 0 Then
  ReDim Preserve users(UBound(users) + 1) As user
  users(UBound(users)).nick = who.nick
  users(UBound(users)).love = 1
  love = 1
 End If
 RndNb = Int(Rnd * 5)
 If (love < 10) Then
  If RndNb = 0 Then
   action "wags his tail", who.to
  ElseIf RndNb = 1 Then
   action "jumps around " & who.nick, who.to
  ElseIf RndNb = 2 Then
   action "licks " & who.nick, who.to
  ElseIf RndNb = 3 Then
   action "yaps at " & who.nick, who.to
  ElseIf RndNb = 4 Then
   action "eats the " + who.what, who.to
  End If
 ElseIf (love >= 10) And (love < 30) Then
  If RndNb = 0 Then
   action "jumps at " & who.nick & "'s face and lick it", who.to
  ElseIf RndNb = 1 Then
   action "is starting to love " & who.nick, who.to
  ElseIf RndNb = 2 Then
   action "brings the wrong " + who.what + " back", who.to
  ElseIf RndNb = 3 Then
   action "runs around " & who.nick, who.to
  ElseIf RndNb = 4 Then
   action "loves when " & who.nick & " plays with him", who.to
  End If
 ElseIf (love >= 30) Then
  If RndNb = 0 Then
   action "loves " & who.nick, who.to
  ElseIf RndNb = 1 Then
   action "licks " & who.nick & "everywhere", who.to
  ElseIf RndNb = 2 Then
   action "jumps everywhere", who.to
  ElseIf RndNb = 3 Then
   action "likes to play with " & who.nick, who.to
  ElseIf RndNb = 4 Then
   action "is very happy to play with " & who.nick, who.to
  End If
 End If
 notice who.nick, "your love is at " + CStr(love)
End Sub

Private Sub ByteUser(who As person, victim As String)
 Randomize Timer
 Dim love As Integer
 tempint = 0
 For x = 0 To UBound(users)
  If LCase(users(x).nick) = LCase(victim) Then
   users(x).love = users(x).love - 1
   love = users(x).love
   tempint = 1
   Exit For
  End If
 Next x
 If tempint = 0 Then
  ReDim Preserve users(UBound(users) + 1) As user
  users(UBound(users)).nick = victim
  users(UBound(users)).love = 0
  love = 0
 End If
 RndNb = Int(Rnd * 1)
 If (love < 10) Then
  If RndNb = 0 Then
   action "nips the heels of " & victim, who.to
  End If
 ElseIf (love >= 10) And (love < 30) Then
  If RndNb = 0 Then
   action "wimps at " & victim & " and backs away", who.to
  End If
 ElseIf (love >= 30) Then
  If RndNb = 0 Then
   Speak "WOOF! WOOF!", who.to
   action "barks at " & who.nick, who.to
  End If
 End If
 notice victim, "your love is at " + CStr(love)
End Sub

Private Sub SaveUsers()
 For x = 0 To UBound(users)
  temp = "users" + String(2 - Len(CStr(x)), "0") + CStr(x)
  WritePrivateProfileString temp, "nick", users(x).nick, inifile
  WritePrivateProfileString temp, "level", CStr(users(x).level), inifile
  WritePrivateProfileString temp, "love", CStr(users(x).love), inifile
  WritePrivateProfileString temp, "pass", users(x).pass, inifile
 Next x
End Sub

Private Sub MasterChanCommands(who As person, usernum As Integer)
 Dim victim As String
 For tempint = 0 To UBound(channels)
  If LCase(channels(tempint).name) = LCase(who.to) Then
   channum = tempint
   Exit For
  End If
 Next tempint
 If who.command = "sleep" Then
  Unload Me
 ElseIf who.command = "run" Then
  action "runs", who.to
 ElseIf who.command = "walk" Then
  action "doesn't like to walk", who.to
  action "runs", who.to
 ElseIf who.command = "good boy!" Then
  action "wags his tail", who.to
 ElseIf Left(who.command, Len("byte")) = "byte" Then
  temp = Split(who.command, " ")
  victim = CStr(temp(1))
  If LCase(victim) = "me" Then victim = who.nick
  ByteUser who, victim
 ElseIf who.command = "i love you" Then
  action "loves " + who.nick + " too", who.to
 ElseIf Left(who.command, Len("deop")) = "deop" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  If victim <> bot.nick Then
   modechange "-o " + victim, who.to
  Else
   Speak "i may be a bot but i'm not stupid...", who.to
  End If
 ElseIf Left(who.command, Len("op")) = "op" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  modechange "+o " + victim, who.to
 ElseIf Left(who.command, Len("voise")) = "voice" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  modechange "+v " + victim, who.to
 ElseIf Left(who.command, Len("devoice")) = "devoice" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  modechange "-v " + victim, who.to
 ElseIf Left(who.command, Len("worship")) = "worship" Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  action "bows before " + victim, who.to
 ElseIf Left(who.command, Len("beg")) = "beg" Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  action "begs for his life from the mercy of " + victim, who.to
 ElseIf Left(who.command, Len("smack")) = "smack" Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  action "smacks " + victim, who.to
 ElseIf Left(who.command, Len("beat down")) = "beat down" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(2)
  If victim = "me" Then victim = who.nick
  action "beats " + victim + "'s head until his toes bleed", who.to
  modechange "-o " + victim, who.to
  modechange "-v " + victim, who.to
  modechange "+m " + who.to, who.to
 ElseIf Left(who.command, Len("invite")) = "invite" Then
  victim = Split(who.command, " ")(1)
  invite victim, who.to
 ElseIf Left(who.command, Len("mode")) = "mode" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "+l" Then victim = Mid(who.command, InStr(who.command, "+l"))
  modechange victim, who.to
 ElseIf Left(who.command, Len("kick")) = "kick" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Split(who.command, " ")(1)
  If victim = "me" Then victim = who.nick
  If victim = bot.nick Then
   Speak "Hey!, i may be dumb, but i'm not stupid", who.to
  Else
   KickEm victim, who.to, "requested"
  End If
 ElseIf Left(who.command, Len("say")) = "say" Then
  Speak Mid(who.command, Len("say") + 2), who.to
 ElseIf Left(who.command, Len("action")) = "action" Then
  action Mid(who.command, Len("action") + 2), who.to
 ElseIf Left(who.command, Len("auto-greet")) = "auto-greet" And users(usernum).level < 3 And users(usernum).connected Then
  For channum = 1 To UBound(channels)
   If LCase(who.to) = LCase(channels(channum).name) Then
    If who.command = "auto-greet" Then
     notice who.nick, CStr(channum) + "|" + channels(channum).name + "|" + CStr(channels(channum).auto_greet)
    Else
     If Split(who.command, " ")(1) = "off" Then
      channels(channum).auto_greet = False
      action "has disabled auto-greet for " + who.to, who.to
     ElseIf Split(who.command, " ")(1) = "on" Then
      channels(channum).auto_greet = True
      action "has enabled auto-greet for " + who.to, who.to
     End If
     Exit For
    End If
   End If
  Next channum
 ElseIf Left(who.command, Len("topic")) = "topic" And users(usernum).level < 3 And users(usernum).connected Then
  victim = Mid(who.command, Len("topic ") + 1)
  If victim = "" Then
   For x = 1 To UBound(channels)
    If LCase(channels(x).name) = LCase(who.to) Then
     victim = channels(x).topic
     Exit For
    End If
   Next x
  ElseIf victim = "default" Then
   victim = channels(0).topic
   x = 0
  End If
  victim = Replace(victim, "$bot.nick", bot.nick)
  victim = Replace(victim, "$channel", channels(channum).name)
  topic victim, who.to
 ElseIf who.command = "leave" And users(usernum).level < 3 And users(usernum).connected Then
  partchan who.to
 ElseIf Left(who.command, Len("user mode")) = "user mode" Then
  victim = Split(who.command, " ")(2)
  For tempint = 0 To UBound(channels(channum).people)
   If LCase(victim) = LCase(channels(channum).people(tempint).nick) Then
    Speak victim + " has mode " + channels(channum).people(tempint).mode, who.to
    Exit For
   End If
  Next tempint
 Else
  PublicCommands who, True, usernum
 End If
End Sub
