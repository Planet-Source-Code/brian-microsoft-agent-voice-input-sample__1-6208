VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form FrmPeedy 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   225
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   225
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox LstTasks 
      Height          =   3180
      Left            =   1920
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2460
      Top             =   240
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1740
      Top             =   0
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Prepare To Speak"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   -120
      Top             =   -60
      Width           =   255
   End
End
Attribute VB_Name = "FrmPeedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Peedy As AgentObjectsCtl.IAgentCtlCharacterEx
Dim VoiceCommand As AgentObjectsCtl.IAgentCtlCommandEx
Dim NumTasks As Integer
Private Enum AskModes
    None
    Shutdown
    Reboot
End Enum
Dim AskMode As AskModes
Dim PeedyWait As AgentObjectsCtl.IAgentCtlRequest

Private Sub Agent1_Command(ByVal UserInput As Object)
Dim action As String
If UserInput.Confidence < -25 Then
    Peedy.Listen True
    Label1.Caption = "Please speak more clearly"
    Exit Sub
End If
Peedy.Stop
Label1.Caption = UserInput.Voice
Me.Refresh
Select Case UserInput.Name
    Case "IE"
        Shell "C:\Program Files\Internet Explorer\iexplore.exe", vbNormalFocus
    Case "Email"
        Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
    Case "Close"
        SendKeys "%{F4}"
    Case "File"
        SendKeys "%F"
    Case "Edit"
        SendKeys "%E"
    Case "View"
        SendKeys "%V"
    Case "Favorites"
        SendKeys "%A"
    Case "Hide"
        Peedy.Speak "Well, I see that I am not wanted here"
        Peedy.Hide
    Case "Ebay"
        Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.ebay.com", vbNormalFocus
    Case "PSC"
        Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.planetsourcecode.com", vbNormalFocus
    Case "Start"
        ShowStartMenu
    Case "Back"
        SendKeys "%{LEFT}"
    Case "Forward"
        SendKeys "%{RIGHT}"
    Case "Hide"
        SendKeys "% N"
    Case "Yes"
        Select Case AskMode
        Case Shutdown
            Peedy.Speak "Goodbye!"
            WINShutdown
        Case Reboot
            Peedy.Speak "I'll see you later!"
            WINReboot
        Case None
            SendKeys "%Y"
        End Select
    Case "No"
        Select Case AskMode
        Case Shutdown
            Peedy.Stop
            Peedy.Speak "OK. Your computer will remain on."
            AskMode = None
            Peedy.Listen True
        Case Reboot
            Peedy.Stop
            Peedy.Speak ("No reboot? But the T.V show is so cool! Oh well, maybe next time!")
            AskMode = None
            Peedy.Listen True
        Case None
            SendKeys "%N"
        End Select
    Case "OK"
        SendKeys "{ENTER}"
    Case "Cancel"
        Beep
    Case "Time"
        If Hour(Now) > 12 Then
            Peedy.Speak "The time is " & Hour(Now) - 12 & ":" & Minute(Now) & " PM"
        Else
            Peedy.Speak "The time is " & Hour(Now) & ":" & Minute(Now) & " AM"
        End If
    Case "Day"
        Peedy.Speak "Today is " & WeekdayName(Weekday(Now))
    Case "Date"
        Peedy.Speak "Today is " & WeekdayName(Weekday(Now)) & ", " & MonthName(Month(Now)) & " " & Day(Now) & ", " & Year(Now)
    Case "Thank"
        Peedy.Stop
        Peedy.Speak "Don't mention it!"
    Case "WTUp"
        Peedy.Stop
        Peedy.Speak "Not much. I sit around in the computer all day and take orders. Sometimes I fly around, hide deep in your monitor, eat crackers, sleep, or listen to music. That's pretty much my life as a parrot."
        Peedy.Listen True
    Case "Cracker"
        Peedy.Stop
        Peedy.Speak "Crackers? YUM!"
        Peedy.Play "Idle2_2"
    Case "Shutdown"
        Peedy.Stop
        Peedy.Speak "Do you really want to shut me down?"
        AskMode = Shutdown
    Case "Peedy"
        Peedy.Stop
        Peedy.Speak "Yes, mysterious user?"
        Peedy.Listen True
    Case "Reboot"
        Peedy.Stop
        Peedy.Speak "OK. I go away and then come back. Hey, a good reboot may solve my problems. Just give me the OK!"
        AskMode = Reboot
    Case Else
        Peedy.Speak "I don't know how to " & UserInput.Voice & "!"
End Select
        Debug.Print UserInput.Name
Peedy.Listen True
Peedy.Play "RestPose"
If AskMode <> None And UserInput.Name <> "Shutdown" And UserInput.Name <> "Reboot" Then
    Select Case AskMode
        Case Shutdown
            action = "shut down the computer"
        Case Reboot
            action = "reboot your computer"
    End Select
    Peedy.Stop
    Peedy.Speak "That's not an answer! No soup for you! Well, I will do what you told me to. But you'll have to ask me again if you want me to " & action & "!"
End If
End Sub

Private Sub Agent1_DeactivateInput(ByVal CharacterID As String)
'Peedy.Listen True
Peedy.Play "RestPose"
End Sub



Private Sub Agent1_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
Me.Hide
End Sub

Private Sub Agent1_ListenComplete(ByVal CharacterID As String, ByVal Cause As Integer)
Me.Shape1.BackColor = vbRed
End Sub

Private Sub Agent1_ListenStart(ByVal CharacterID As String)
Me.Shape1.BackColor = vbGreen
Peedy.Play "RestPose"
End Sub

Private Sub Agent1_RequestComplete(ByVal Request As Object)
Beep
End Sub

Private Sub Form_Load()
Agent1.Characters.Load "Peedy", "Peedy.acs"
Set Peedy = Agent1.Characters("Peedy")
Peedy.Listen True
Peedy.Top = Screen.Height / Screen.TwipsPerPixelY - Peedy.Height - 18
Peedy.Left = -40
Peedy.Speak "Hello!"
Me.Top = Screen.Height - 800
Me.Left = 1000
Peedy.Commands.Add "IE", "Launch Internet Explorer", "[(Launch | Start | Run)] [(my | the)] (web browser | [Microsoft] Internet Explorer | internet | world wide web | [the] web)"
Peedy.Commands.Add "Email", "Launch Outlook Express", "([(Launch | Start | Run)] (Outlook [express] | email | mail | electronic mail) | (read | view | check) [my] (mail | email))"

Peedy.Commands.Add "Ebay", "Go To Ebay", "([(Launch | Start | Run | Go to)] [(my | the)] [website] Ebay | [Online] Auction[s])"
Peedy.Commands.Add "PSC", "Go To Planet Source Code", "([(Launch | Start | Run | Go to)] [(my | the)] [website] Planet Source Code |(VB | Visual Basic) code)"

Peedy.Commands.Add "Close", "Close Window", "(Close [(this | the)] [(window | program | application)] | Die [you][evil]program | Kill [this] program | Get [the hell] off (my | the) screen [[you][evil]program] | DIE [DIE][DIE][DIE][DIE][DIE])"
Peedy.Commands.Add "Minimize", "Minimize Program", "(Minimize | Hide) [this] (program | application | window)"
Peedy.Commands.Add "Next", "Next Program", "Next (program | application | window)"
Peedy.Commands.Add "Hide", "Go Away", "(Go away | Don't bug me | Hide [yourself] | You're in my way | Go bug (merlin | robby | clipit))"

Peedy.Commands.Add "File", "File Menu", "(Open [the] file menu  | File [menu] | Show [the] file [menu])"
Peedy.Commands.Add "Edit", "Edit Menu", "(Open [the] edit menu  | Edit [menu] | Show [the] edit [menu])"
Peedy.Commands.Add "View", "View Menu", "(Open [the] view menu  | View [menu] | Show [the] view [menu])"
Peedy.Commands.Add "Start", "Start Menu", "(Open [the] start menu  | Start [menu] | Show [the] start [menu])"
Peedy.Commands.Add "Favorites", "Favorites Menu", "(Open [the] (favorites | bookmarks) menu  | [(favorites | bookmarks)] menu | Show [the] (favorites | bookmarks) [menu] | (view | show me | let me see | what are) [my] (bookmarks | favorites))"

Peedy.Commands.Add "Back", "Go Back", "(Go back | Back | [Go to] [the] Previous [(website | web page | internet site)])[please][peedy]"
Peedy.Commands.Add "Forward", "Go Forward", "(Go Forward | Forward | [Go to] [the] next [(website | web page | internet site)])[please][peedy]"
Peedy.Commands.Add "Stop", "Stop Loading", "((Stop | Don't) [load [ing]] [(this | the)][(website | web page | internet site)])[(please peedy | peedy please | please | peedy)]"

Peedy.Commands.Add "Yes", "Yes", "(Yes [please]| OK | Affirmitive | Go ahead | [just] Do it | Sure | Why not)"
Peedy.Commands.Add "No", "No", "([the answer is] No [way] | No thank you | Don't do it | Please not | Please don't do it | Not if I can help it)"

Peedy.Commands.Add "Time", "Current Time", "([peedy][what is the][current]time|what time[is it][peedy])"
Peedy.Commands.Add "Date", "Current Date", "([peedy][what is the][current]date|what is todays date|what is today)[peedy]"
Peedy.Commands.Add "Day", "Day of week", "[peedy]([what is the][current]day|what day of the week is it|what day is it|what day is today)[peedy]"

Peedy.Commands.Add "Thank", "Thank You", "[peedy](Thank You | Thanks) [peedy] [for *]"
Peedy.Commands.Add "WTUp", "WtUp", "(What's | What Is)(Up|Cookin|Cooking|Happenin|Happening)[peedy]"
Peedy.Commands.Add "Cracker", "Eat a cracker", "([peedy][,](eat|have) [a] cracker [peedy]|(peedy|polly)(want|wanna)cracker?)"

Peedy.Commands.Add "Shutdown", "Shut down the computer", "[peedy](shutdown | shutdown | turn off) [(the|my) computer][peedy]"
Peedy.Commands.Add "Reboot", "Restart the computer", "[peedy](restart | reboot) [(the|my) computer][peedy]"

Agent1.Characters("Peedy").Listen True
End Sub
