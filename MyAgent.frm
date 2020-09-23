VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmMain 
   Caption         =   "Accepting input"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1500
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   1080
      Width           =   3075
   End
   Begin AgentObjectsCtl.Agent Agent 
      Left            =   360
      Top             =   300
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This statement prevents the creation
                'of Varient variables when the variables
                'are referenced but not created before
                'e.g. For n = 1 to 10, if n was never Dim-ed

Dim IntroComplete As Boolean    'When the intro is complete
                                'this is set to true
                                'to activate some options
                                '(well 1 option actually)
Dim LoadRequest(4)
    'These are the LoadRequests, when one is set to a
'R  'Character command, when the command is completed the
'E  'Agent control runs extra lines of code, according to
'A  'which request was used, such as an UnLoad function.
'D  'This is needed because Agent works with Ques, so when
    'you give Agent a command, that command is put into a
'T  'que, and your program carries on to the next command
'H  'in your program, thus freeing up your program to do
'I  'other things, though it also means if you give an
'S  'agent a command followed by an UnLoad Me statement,
'!  'the Agent command will be qued, the program carries
'!  'on, and unloads the form, and the Agent control,
    'stopping everything!!
    'It is a rather strange way of programming,
    'but it works, and this is how Microsoft
    'make their scripts too :)
    'This is where many programmers go wrong, and swear at
    'their computers for a few hours!
    '(before kicking it out of a window)
    
Dim Genie As IAgentCtlCharacter
Dim Merlin As IAgentCtlCharacter
    'To be used in Set commands
    
Dim Request As IAgentCtlRequest
    'To set to an agent request, so another agent can wait
    'for that request in another Agent's que to complete
    'before continuing the commands in it's own que,
    'well, try going through this later and remove all the
    ''Set Request =' bits, see what happens when you run
    'it ;-)

Private Sub LoadGenie()
    Set LoadRequest(0) = Agent.Characters.Load("Genie", "Genie.acs")
        'This command, Sets LoadRequest(0) as the
        'Agent Load command, which is run, so when the
        'Agent control finishes loading the Genie.acs file
        'it will run through the Agent_RequestComplete
        'Subroutine (see below) and find the finished
        'request equal to LeadRequest(0), and so run the
        'commands in this part of code.  Using this method,
        'you can set commands to only run when an Agent has
        'completed an operation, such as i have used one to
        'only Unload the form when Merlin has finished
        'hiding, this is how you get around the Ques problem
        '(strange huh?!)
    
    Label1.Caption = "Loading Genie character."
End Sub

Private Sub LoadMerlin()
    Set LoadRequest(1) = Agent.Characters.Load("Merlin", "Merlin.acs")
        'Again, running code after the character is loaded
        'the "Merlin" part is the Key of the agent, or its
        'loaded name, how you can reference to it once it
        'is loaded, and the Merlin.acs part is the path to,
        'and name of the file, if the path is omitted
        '(left out), the default path is used, which is
        'the MS Agent characters directory, in this case
        'C:\WINDOWS\MSAGENT\CHARS\
        
    Label1.Caption = "Loading Merlin character."
End Sub

Private Sub AddMerlinCommands()
    'Add voice commands to Merlin
    Merlin.Commands.Add "ChPoofs", "You want some Cheesy Poofs?", "You want some Cheesy Poofs?", True, True
                        'Ref Name , Caption in the popup menu etc, What it listens for         , Whether it is enabled, whether it is visible in the popupmenu and commands window
    Merlin.Commands.Add "Internet", "Internet?", "Internet?", True, True
    Merlin.Commands.Add "OffChar", "Can I use Office characters with Agent?", "Can I use Office characters with Agent?", True, True
    Merlin.Commands.Add "Goodbye", "Goodbye", "Goodbye", True, True
    Merlin.Commands.Caption = "Merlin"
End Sub

Private Sub Agent_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    If IntroComplete And Button = vbLeftButton Then
        Merlin.Play "Surprised" 'Play Merlins's Suprised
                                'animation
        Merlin.Speak "Be careful with that pointer!|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to your programming!"
            'make Merlin speak the text, this can also be from a textbox etc.
            'the '|' character is an OR, so the Agent chooses which text to say randomly, v cool
        Merlin.Play "RestPose"
            'Return to Merlin's Rest Pose
    End If
End Sub

Private Sub Agent_Command(ByVal UserInput As Object)
    Select Case UserInput.Name
        Case "ChPoofs" 'If they said the command with the Key 'ChPoofs' then...
            If Agent.AudioOutput.Enabled Then   'If they have AudioOutput installed..
                Merlin.Speak "Yeah I want cheesy poofs!", "ChPoofs.lwv"
                'Say "Yeah I want Cheesy Poofs", and play
                'the ChPoofs Linguistically Enhanced Wave
                'Sound file at the same time,
                'they are basically Wave files with the words
                'and punctuation written into them
                'Download the program for it from
                'http://msdn.microsoft.com/workshop/imedia/agent/default.asp
                'or wherever it is now, but DON'T ASK ME!
                'I am NOT going to e-mail a 6mb file to you :)
            Else
                Merlin.Speak "You do not have the Lernout & Houspie TrueVoice engine installed!  Please download and install it from http://msdn.microsoft.com/workshop/imedia/agent/default.asp"
            End If
        Case "Internet"
            If Agent.AudioOutput.Enabled Then
                Merlin.Speak ".", "KingOfTheHill.wav"
                    'They can also speak wav files, the
                    '[TEXT] and [URL] can be omitted
                    '(left out), so you can have just text,
                    'just a sound file, or both
                    '(or neither, though that would be useless!)
                'An hour ago this worked fine without the ".",
                'though it wont now... try removing it :)
            Else
                Merlin.Speak "You do not have the Lernout & Houspie TrueVoice engine installed!  Please download and install it from http://msdn.microsoft.com/workshop/imedia/agent/default.asp"
            End If
        Case "OffChar"
            Merlin.Speak "Yes, \pau=200\you can use Office Assistant ACS character files, though they do not have the same animations as proper characters do, and they often cannot speak, only mime.. \pau=700\So I guess you will never hear Clippit's voice.."
                '\pau=200\ is a 'speech modifier'
                'see at the end for a definition
        Case "Goodbye"
            'If UserInput.Confidence < 0 Then
            '    Merlin.Speak "Are you sure you want to say goodbye?"
            '    Merlin.Commands.Add "Yes", "Yes", "Yes", True, True
            'Else 'This would be a check of the
                  'confidence level of what you said,
                  'so it doesn't go wrong so easily.
                  'If in doubt it asks you whether you
                  'wanted to say that, adds the command
                  '"YES" to the list (which you can
                  'remove after), etc etc, you'll have to
                  'write this bit yourself if you want it,
                  'though it's only neccessary if you have
                  'a LOT of commands
            Merlin.Speak "From all us Agents, and Geeza, we hope you have lots of fun with MS Agent!"
            Merlin.Play "Greet"
            Merlin.Speak "Goodbye"
            Merlin.Play "RestPose"
            Set LoadRequest(4) = Merlin.Hide
                'After Merlin's hidden, Unload the form
     End Select
End Sub

Private Sub Agent_RequestComplete(ByVal Request As Object)
    'These are the code bits which can be run after an
    'Agent completes an action, such as speech
    'And often are used to point to other Subroutines to
    'run next, though BE CAREFUL!  Overuse of this technique
    'can make spider-like programs which are very hard to
    'follow, and BAD PROGRAMMING!  (BASIC's flaw)
    'You should see the MS Agent VB Script written by
    'Microsoft I had to start with!  Not easy if you don't
    'know what you're doing :)
    
    Select Case Request
        Case LoadRequest(0)
            If Request.Status <> 0 Then  'Obvious
                Label1.Caption = "Genie Character failed to load!"
            Else
                Label1.Caption = "Genie Character loaded successfully"
                Set Genie = Agent.Characters("Genie")
            End If
        Case LoadRequest(1)
            If Request.Status <> 0 Then
                Label1.Caption = "Merlin Character failed to load!"
            Else
                Label1.Caption = "Merlin Character loaded successfully"
                Set Merlin = Agent.Characters("Merlin")
                cmdGo.Enabled = True
                
                Call AddMerlinCommands 'Add merlin's voice
                                       'commands
            End If
        
        Case LoadRequest(2)
            Agent.Characters.Unload "Genie"
        Case LoadRequest(3)
            Agent.Characters.Unload "Merlin"
        Case LoadRequest(4)
            Agent.Characters.Unload "Merlin"
            Unload Me   'This is the LoadRequest Referenced
                        'earlier
    End Select
End Sub

Private Sub cmdGo_Click()
    Genie.MoveTo 150, 240  'Move Genie to these (Pixel) coordinates
    Genie.Show 'Show Genie (plays his "Show" animation also)
    Genie.Speak "Hello!  I am Genie, \pau=100\a Microsoft Agent Character"  'Say this text
    Genie.Play "Greet"  'Play Genie's animation "Greet"
    Genie.Speak "And I am at your service"
    Genie.Play "RestPose" 'Return Genie to default position

    Genie.Speak "I hope Geeza's code will help you to create Agent scripts easily too, \pau=100\and get around the problem that we can't pronounce Geeza correctly.  \pau=200\G-eeza, \pau=200\GG eeza... \pau=70\"
    Genie.Play "Explain"
    Genie.Speak "\emp\Oh well..\pau=100\"
    Genie.Play "RestPose"
    Genie.Speak "Microsoft Agent is a very powerful tool, \pau=100\If you can find a use for it\pau=300\"
    Genie.Speak "\pau=100\Agents can give help, \pau=100\play animations, \pau=100\speak text (for example, \pau=50\you can use us to read your e-mail), \pau=100\give presentations, \pau=100\and lots more\pau=100\"
    Genie.Speak "We can even have multiple characters, \pau=100\interacting with you and each other\pau=500\"
    Set Request = Genie.Speak("And to help me explain these features, \pau=100\here's Merlin!\pau=100\")
        'Sets the Request to this, so another Character can
        'be set to wait for this instruction to complete
        'before comtinuing their que of commands
    Genie.Play "LookLeft"

    Merlin.MoveTo 320, 240
    Merlin.Wait Request 'Wait for genie to finish speaking
    Merlin.Show 'then Show Merlin
    Merlin.Play "Wave"
    Merlin.Speak "Hello everyone!"
    Merlin.Play "RestPose"
    Merlin.Speak "I am Merlin; \pau=150\another Microsoft Agent Character\pau=100\"
    Set Request = Merlin.Speak("Though there are even more characters than us, \pau=100\such as Robby the robot,")
    Merlin.Play "RestPose"

    Genie.Wait Request  'Same as before
    Genie.Play "LookLeftReturn"
    Set Request = Genie.Speak("Peedy the Parrot,\pau=100\")
    Genie.Play "LookLeft"

    Merlin.Wait Request
    Merlin.Play "Explain"
    Set Request = Merlin.Speak("and all the Microsoft Office 2000 help Assistants!\pau=100\")
    Merlin.Play "Blink"

    Genie.Wait Request
    Genie.Play "LookLeftReturn"
    Set Request = Genie.Speak("Yes, \pau=200\all the Office 2000 Assistants are MS Agents, \pau=100\though this is not what Agent was originally written for..")
    Genie.Play "LookLeft"

    Merlin.Wait Request
    Set Request = Merlin.Speak("We were originally created for use on the Internet, \pau=100\through web pages, \pau=100\and have many other features.  We can speak, \pau=100\as we are now, \pau=140\speak Wav sound files, \pau=140\speak Linguistically Enhanced Wave sound files, \pau=140\gesture, \pau=140\move, \pau=140\register when the user clicks or drags us, \pau=140\and even accept speech input!")
    Merlin.Play "LookRight"

    'Request is a MUST if you intend to use multiple
    'characters, you may even need seperate request
    'variables for each character, such as MerlinRequest
    'don't believe me? Remove 'Set Request = ' from
    'everywhere you see it, then run this

    Genie.Wait Request
    Genie.Play "LookLeftReturn"
    Set Request = Genie.Speak("Well, \pau=100\that's enough explaining our features, \pau=100\I'm off to the pub with Clippit, \pau=200\why not let Merlin show you some of our features?")
    Merlin.Wait Request
    Merlin.Play "LookRightReturn"
    Merlin.Speak "Later Genie!  \pau=260\Don't get too drunk"
    Genie.Play "Wave"
    Set LoadRequest(2) = Genie.Hide 'After he's hidden, unload him

    Call MerlinSolo 'It's a good idea to separate these
                    'routines, to avoid comfusion and make
                    'debugging easier
End Sub

Private Sub MerlinSolo()
    Merlin.Wait LoadRequest(2)  'wait until Genie is hidden
                                'and unloaded (you can set
                                'this to other things than
                                'Request, you see)
    Merlin.Speak "Well, now Genie has gone, \pau=130\why not try voice commands.\pau=1000\"
    Merlin.MoveTo ScaleX(Screen.Width, vbTwips, vbPixels) - 130, ScaleY(Screen.Height, vbTwips, vbPixels) - 170
    Merlin.Play "GestureDown"
    Merlin.Speak "All visible voice commands for an agent are listed in the Agent Commands Window"
    Merlin.Play "Restpose"
    Merlin.Speak "Just Right-Click my Magic hat icon in your system tray, \pau=130\ or right click me, \pau=130\to make my menu appear,\pau=500\"
    Merlin.Speak "Then either click 'Open Voice Commands Window' to display the commands window, or if you like you can click one of my voice commands directly from the menu to run it, say if you haven't got a microphone or do not have the Speech Input addon installed.\pau=1000\"
    Merlin.Speak "But, \pau=170\if you do have a microphone, hold down the Scroll-Lock key, wait for the tooltip to say I'm ready to listen, then say the command you wish me to execute\pau=750\"
    Merlin.Speak "The power is now in your hands!\pau=800\"
    Merlin.Speak "By the way, I like Cheesy Poofs..."
    Merlin.Play "Acknowledge"
    
    IntroComplete = True
End Sub

Private Sub Form_Load()
    Call LoadGenie
    Call LoadMerlin 'Load the characters
        'This is handled differently if the show is
        'continuous, see the other form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'This doesn't work yet, and i'm too tired to make it work
    'it'll be good programming practice for you anyhow :)
    On Error Resume Next 'needed incase one has already been
                         'unloaded
    Set LoadRequest(2) = Genie.Hide 'Hide and unload the
    Set LoadRequest(3) = Merlin.Hide 'characters
End Sub


'OK, you should be able to see how this works :)  And if
'not, you'll still be able to use it.
'Just a few more things:

'The Speech Modifiers:
'   1.  \emp\       Emphasise the next word
'   2.  \pau = m\   Pause for m milliseconds
'   3.  \pit = p\   Pitch voice to p Hertz (1 - 400)
'   4.  \spd = s\ Set speed to s Words per minute
'                   (50-250)  -Thanks Ric for that one
'These are the only three I've ever seen and know

'If you need the components for Agent, whatever they may be
'DO NOT ask me for them, i wont send them, they're way too
'big!  Get them on the microsof Agent site
'it should be at http://msdn.microsoft.com/workshop/imedia/agent/default.asp

'There is no way of listing all the animations of an Agent
'(other than going through every possible combination or
'decompiling them)  so to find them out you'll need
'AllDocs.zip from the microsoft Agent site for the Merlin
'etc. animations

'If you are offended by or if i've done something wrong by
'putting in the example sound files, i am very sorry, and
'if you want i'll remove the files,  though they are just
'for a laugh, and hey it's good advertising right? :)

'Well that's just about Everything to do with Agent! lol
'i hope this helps!

'If it did, could you just add a reference to me in your
'"About" box or whatever just to say that i did :)  thanks

'OK , I'm off to sleep!!
'Check out my other project on-line so far-
'   Floating Objects on Invisibe Forms
'   And, Killer Button!

'Goodnight All!

'GEEZA
'*Of all the things I've lost, I miss my mind the most*
'   -Ozzy Osbourne
