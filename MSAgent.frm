VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1260
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin AgentObjectsCtl.Agent Agent 
      Left            =   360
      Top             =   300
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Genie As IAgentCtlCharacter
Dim Robby As IAgentCtlCharacter
Dim Merlin As IAgentCtlCharacter
Dim LoadRequest(3)
Dim GetRequest(8)
Dim GenieRequest As IAgentCtlRequest
Dim RobbyRequest As IAgentCtlRequest
Dim MerlinRequest As IAgentCtlRequest
Dim DatapathType As String
Dim CenterX As Integer, CenterY As Integer
Dim RobbyDelayed As Boolean
Dim RobbyLoaded As Boolean
Dim MerlinLoaded As Boolean
Dim RobbyShowFailed As Boolean
Dim MerlinShowFailed As Boolean

Private Sub Agent_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    If Button = 1 Then

        'If we have already loaded Robby, but his Showing state animation failed
            If RobbyLoaded = True And RobbyShowFailed = True Then
                ' Apologize
                Genie.Play "Greet"
                Genie.Speak "Yes, O Great One. I will attempt to summon Robby again.|As you wish, my master."
            
                ' Try again
                Set GetRequest(6) = Robby.Get("state", "showing, speaking")

        ' Otherwise, if we have already loaded Merlin, but his Showing state animation failed
        ElseIf MerlinLoaded = True And MerlinShowFailed = True Then

                ' Apologize
                Genie.Speak "Yes, O Noble One. I will attempt to summon Merlin again.|Yes, Master. I will see if I can rouse the wizened one."
                
                ' Try again
                Set GetRequest(7) = Merlin.Get("state", "showing, speaking")
                            

        Else ' We haven't loaded either character yet, so just ignore the click.
            Exit Sub
        End If
    End If
End Sub

Private Sub Agent_RequestComplete(ByVal Request As Object)
    Select Case Request

        Case LoadRequest(1)
            
            If Request.Status <> 0 Then
                ' If genie's character data fails to load post a message
                StatusBar1.SimpleText = "Genie character failed to load."

                MsgBox "Unable to load the Genie character. Please try refreshing this page."

            Else
                StatusBar1.SimpleText = "Genie character successfully loaded."

                ' Create a reference to the character
                Set Genie = Agent.Characters("genie")

                ' Load the character's Showing and Speaking state animations
                Set GetRequest(1) = Genie.Get("state", "showing,speaking")
                
                StatusBar1.SimpleText = "Now loading Genie's Showing and Speaking state animations."
                
            End If

        Case LoadRequest(2)

            If Request.Status <> 0 Then
                StatusBar1.SimpleText = "Robby character failed to load."
                RobbyLoaded = False
                Genie.Speak "Uh oh, Robby failed to load."
                Genie.Speak "Try refreshing this page."
                
            Else
                StatusBar1.SimpleText = "Robby character successfully loaded."
                RobbyLoaded = True
                Set Robby = Agent.Characters("robby")
                Set GetRequest(6) = Robby.Get("state", "showing,speaking")
                
                StatusBar1.SimpleText = "Loading Robby's Showing and Speaking state animations."
                
            End If

        Case LoadRequest(3)
            
            If Request.Status <> 0 Then
                StatusBar1.SimpleText = "Merlin character failed to load."
                MerlinLoaded = False
                Genie.Play "Greet"
                Genie.Speak "Ten thousand pardons, O Merciful One. Merlin was not available."
                Genie.Speak "When we're done, you can try refreshing this page and maybe he'll show up next time."
                
                If RobbyLoaded = True And RobbyShowFailed = False Then
                    Robby.Hide

                End If

                SayGoodbye
                
            Else
                StatusBar1.SimpleText = "Merlin character successfully loaded."
                MerlinLoaded = True
                Set Merlin = Agent.Characters("merlin")
                Set GetRequest(7) = Merlin.Get("state", "showing,speaking", False)
                Merlin.Get "animation", "blink", False

                StatusBar1.SimpleText = "Loading Merlin's Showing and Speaking state animations."
            
            End If


        Case GetRequest(1) 'Request to load Genie's Showing and Speaking state animations

            If Request.Status = 0 Or DatapathType = "UNC" Then

                StatusBar1.SimpleText = "Genie's Showing and Speaking state animations loaded."
                
                Genie.Get "animation", "greet", False

                ' If the animation loaded, start the intro
                StartIntro
                
            Else
                StatusBar1.SimpleText = "Genie's Showing and Speaking state animations failed to load."
                
                MsgBox "Unable to load Genie's Showing and Speaking animations. Try refreshing the page."
                
            End If

        Case GetRequest(2) 'Request to load Genie's LookLeft, GreetReturn, Blink animations
            
            If Request.Status = 0 Or DatapathType = "UNC" Then

                StatusBar1.SimpleText = "Genie's LookLeft, GreetReturn, and Blink animations loaded."

                Genie.Play "RestPose"

            Else
                StatusBar1.SimpleText = "Genie's LookLeft, GreetReturn, and Blink animations failed to load."

            End If

            ' Now process Genie's intro of Robby
            GenieIntro

        Case GetRequest(3) 'Request to load Robby's Wave animation
            
            If Request = 0 Or DatapathType = "UNC" Then

                StatusBar1.SimpleText = "Robby's Wave animation loaded."

            Else
                StatusBar1.SimpleText = "Robby's Wave animaton failed to load."

            End If

            RobbyIntro


        Case GetRequest(4) 'Request to load Robby's LookRight animations

            If Request.Status = 0 Or DatapathType = "UNC" Then

                StatusBar1.SimpleText = "Robby's LookRight animations loaded"

            Else
                StatusBar1.SimpleText = "Robby's LookRight animation failed to load."

            End If

            RobbyExplainsFeatureOne


        Case GetRequest(5) 'Request to load Robby's Read animation

            If Request.Status = 0 Or DatapathType = "UNC" Then

                StatusBar1.SimpleText = "Robby's Read animations loaded."
                
            Else
                StatusBar1.SimpleText = "Robby's Read animations failed to load."
            
            End If

            GenieExplainsFeatureTwo


        Case GetRequest(6) 'Request to load Robby's Showing and Speaking state animations
        
            If Request.Status = 0 Or DatapathType = "UNC" Then
                RobbyShowFailed = False
                StatusBar1.SimpleText = "Robby's Showing and Speaking state animations loaded."
                ShowRobby

            Else
                RobbyShowFailed = True
                StatusBar1.SimpleText = "Robby's Showing and Speaking state animations failed to load."
                Genie.Speak "Ten thousand pardons, O Merciful One."
                Genie.Speak "I regret to tell you that I am unable to summon the metallic one."
                Genie.Speak "Left-click me if you'd like me to try again."

            End If


        Case GetRequest(7) 'Request to load Merlin's Showing and Speaking state animations

            If Request.Status = 0 Or DatapathType = "UNC" Then
                StatusBar1.SimpleText = "Merlin's Showing and Speaking state animations loaded."
                MerlinShowFailed = False
                
                ' Get Merlin's GlanceRight animation
                Merlin.Get "animation", "glanceright", False
                StatusBar1.SimpleText = "Loading Merlin's GlanceRight animation."

                ' Get the character's hiding animation states
                Merlin.Get "state", "hiding", False
                Robby.Get "state", "hiding", False
                Genie.Get "state", "hiding", False
                StatusBar1.SimpleText = "Loading the character's Hiding state animations."
                
                MerlinIntro

            Else
                StatusBar1.SimpleText = "Merlin's Showing and Speaking state animations failed to load."
                MerlinShowFailed = True
                Genie.Speak "A thousand pardons, O Masterful One,"
                Genie.Speak "but I was unable to summon the wizard."
                Genie.Speak "Left-click me if you want me to try again."

            End If


        Case GetRequest(8) ' Request to load Genie's Explain animation
            
            StatusBar1.SimpleText = "All animation requests completed."

    End Select
End Sub

Private Sub Command1_Click()
    DatapathType = "UNC" ' set to UNC or URL depending on where character data resides
    CenterX = 320
    CenterY = 240
    Agent.Connected = True  ' May be needed in some contexts
    LoadGenie
End Sub

Private Sub LoadGenie()
    Select Case DatapathType
    
        Case "UNC"
            Set LoadRequest(1) = Agent.Characters.Load("genie", "genie.acs")

        Case "URL"
            Set LoadRequest(1) = Agent.Characters.Load("genie", "http://agent.microsoft.com/characters/genie/genie.acf")
    
    End Select

    StatusBar1.SimpleText = "Loading Genie character."
End Sub

Private Sub LoadRobby()
    Select Case DatapathType

        Case "UNC"
            Set LoadRequest(2) = Agent.Characters.Load("robby", "robby.acs")
        
        Case "URL"
            Set LoadRequest(2) = Agent.Characters.Load("robby", "http://agent.microsoft.com/characters/robby/robby.acf")
    
    End Select

    StatusBar1.SimpleText = "Loading Robby character."
End Sub

Private Sub LoadMerlin()
    Select Case DatapathType

        Case "UNC"
            Set LoadRequest(3) = Agent.Characters.Load("merlin", "merlin.acs")

        Case "URL"
            Set LoadRequest(3) = Agent.Characters.Load("merlin", "http://agent.microsoft.com/characters/merlin/merlin.acf")
    
    End Select

    StatusBar1.SimpleText = "Loading Merlin character."
End Sub

Private Sub StartIntro()
    RobbyDelayed = True

    Genie.MoveTo CenterX - 150, CenterY
    Genie.Show

    Set GetRequest(2) = Genie.Get("animation", "LookLeft, LookLeftBlink, LookLeftReturn, GreetReturn, Blink")
    StatusBar1.SimpleText = "Loading Genie's LookLeft, GreetReturn, and Blink animations."

    Genie.Speak "Welcome to the release of Microsoft Agent, the new ActiveX technology that supports interactive characters"
    Genie.Speak "and enables the creation of highly interactive Web pages or stand-alone desktop applications"
    Genie.Speak "that can leverage the natural aspects of social interaction."
    Genie.Play "Greet"
    Genie.Speak "I am Genie, your most humble and loyal servant."
    Genie.Play "RestPose"
    Genie.Speak "And one of the animated Microsoft Agent actors."
    Genie.Play "Blink"
    Genie.Speak "I bring you great news of the wondrous things that you can do with this marvelous technology,"
    Genie.Speak "feats that even the greatest magicians in all of Baghdad have been unable to equal."
    Genie.Play "Blink"
End Sub

Private Sub GenieIntro()
    LoadRobby

    Genie.Speak "To help me explain the features included in Microsoft Agent,"
    Genie.Speak "let me introduce a member of the cast and my personal friend,"
    Genie.Speak "Robby the robot."
    Genie.Play "LookLeft"
    Genie.Play "LookLeftBlink"
End Sub

Private Sub OverTime()
    If LoadRequest(2).StatusBar1.SimpleText = 2 Or LoadRequest(2).StatusBar1.SimpleText = 4 Then
        Genie.Play "LookLeftBlink"
        Genie.Play "LookLeftBlink"
        Genie.Play "LookLeftReturn"
        Genie.Speak "Sorry about the delay.|Traffic must be heavy on the Internet today.|Thank you for your patience.|It should only be another moment.|A thousand pardons, O Patient One. Another moment..."
        Genie.Speak "\pau=2500\|\pau=1000\|\pau=2000\"
        Genie.Play "LookLeft"
        Genie.Play "LookLeftBlink"

        RobbyDelayed = True
        
    Else
        Exit Sub

    End If
End Sub

Private Sub ShowRobby()
    Set GetRequest(3) = Robby.Get("animation", "Wave, WaveReturn", False)
    StatusBar1.SimpleText = "Loading Robby's Wave animation."

    Robby.MoveTo CenterX, CenterY
    Genie.Play "LookLeftReturn"
    Set GenieRequest = Genie.Speak("And here he is!")
    Genie.Play "LookLeft"
    Genie.Play "LookLeftBlink"
    
    Robby.Wait GenieRequest
    Robby.Show
    
    Set RobbyRequest = Robby.Speak("Thanks for the great introduction, Genie.")
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    If RobbyDelayed = True Then
        Robby.Speak "Sorry for the delay."
        Robby.Speak "Traffic is really busy today on the Information Superhighway."
        Set RobbyRequest = Robby.Speak("All those people downloading this new Microsoft Agent technology.")
        Robby.Speak "It's incredible!"

        Genie.Wait RobbyRequest
        Genie.Play "LookLeftBlink"

        Robby.Speak "But here I am."

    End If
End Sub

Private Sub RobbyIntro()
    Set GetRequest(4) = Robby.Get("animation", "LookRight,LookRightReturn", False)
    StatusBar1.SimpleText = "Loading Robby's LookRight animations."

    Robby.Get "animation", "acknowledge", False
    StatusBar1.SimpleText = "Loading Robby's acknowledge animation."
    
    Robby.Play "Wave"
    Set RobbyRequest = Robby.Speak("Hello. I am Robby the robot.")
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Robby.Play "restpose"
    Set RobbyRequest = Robby.Speak("Like Genie said, I am one of the sample animated characters available for Microsoft Agent.")
    Robby.Play "LookRight"

    Genie.Wait RobbyRequest
    
    Genie.Play "LookLeftBlink"
    Genie.Play "LookLeftReturn"
    Set GenieRequest = Genie.Speak("My Metallic Friend, let's expound about the awesome features of Microsoft Agent.")
    Genie.Play "LookLeft"
End Sub

Private Sub RobbyExplainsFeatureOne()
    Set GetRequest(5) = Robby.Get("animation", "Read,ReadReturn", False)
    StatusBar1.SimpleText = "Loading Robby's Read animations."

    Robby.Wait GenieRequest

    Set RobbyRequest = Robby.Play("LookRightReturn")
    
    Robby.Play "acknowledge"
    Set RobbyRequest = Robby.Speak("OK, Genie.")

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Robby.Play "Explain"
    
    Set RobbyRequest = Robby.Speak("First, Microsoft Agent gives developers several options.")
    Robby.Play "LookRight"
    
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"
    Genie.Play "LookLeftReturn"
    Genie.Play "Explain"
    Genie.Speak "That's correct, my digital friend."

    Genie.Speak "For example, a developer can use just our animation services,"
    Genie.Play "RestPose"
    Genie.Speak "programming us to respond to mouse or keyboard input."
    Set GenieRequest = Genie.Speak("Or also include optional support for speech recognition and speech synthesis.")
    Genie.Play "LookLeft"

    Robby.Wait GenieRequest
    Robby.Play "LookRightReturn"
    Robby.Speak "Character animation data can be installed and loaded from the user's disk,"
    Set RobbyRequest = Robby.Play("Explain")

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Robby.Speak "or downloaded from the server as needed and stored in the browser's cache,"
    Set RobbyRequest = Robby.Play("RestPose")

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Set RobbyRequest = Robby.Speak("making it faster to access an animation if it's already been used.")
    Robby.Play "LookRight"
End Sub

Private Sub GenieExplainsFeatureTwo()
    Robby.Get "animation", "Explain, ExplainReturn", False
    StatusBar1.SimpleText = "Loading Robby's Explain animation."
    
    LoadMerlin

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"
    Genie.Play "LookLeftReturn"
    Genie.Speak "Our simple but powerful programming interface enables us to animate and talk from anywhere on the screen."

    Set GenieRequest = Genie.Play("LookLeft")
    Genie.Play "LookLeftBlink"

    Robby.Wait GenieRequest
    Robby.Play "LookRightReturn"
    Robby.Play "Read"
    Set RobbyRequest = Robby.Speak("Yes, developers can access events, such as when the user clicks or drags us.")
    Robby.Play "ReadReturn"

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Set RobbyRequest = Robby.Speak("And also detect when animation requests begin and end.")
    Robby.Play "LookRight"
        
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftReturn"

    Set GenieRequest = Genie.Speak("There's also lots of support for customizing character interaction.")
    
    Robby.Wait GenieRequest
    Robby.Play "LookRightReturn"
    Robby.Play "Acknowledge"
    Robby.Play "LookRight"
    
    Set GenieRequest = Genie.Speak("Characters can be used as guides, coaches, entertainers, or other types of assistants or specialists.")
    Genie.Play "LookLeft"
    Genie.Play "LookLeftBlink"
    
    Robby.Get "animation", "LookLeft,LookLeftReturn", False
    StatusBar1.SimpleText = "Loading Robby's LookLeft animations."

    Robby.Wait GenieRequest

    Robby.Play "LookRightReturn"
    Robby.Play "Explain"
    Set RobbyRequest = Robby.Speak("As you can see, Microsoft Agent even supports multiple characters.")
    Robby.Play "RestPose"
    
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"

    Set RobbyRequest = Robby.Speak("Characters can be programmed to animate independently or synchronized to each other's animations.")
    Robby.Play "LookRight"

    Genie.Wait RobbyRequest
    Genie.Play "LookLeftBlink"
    Genie.Play "LookLeftReturn"
    Genie.Speak "And Microsoft also provides a special tool that developers can use to compile their own characters."
    
    Genie.Speak "Like my other friend, Merlin the wizard."
    Set GenieRequest = Genie.Play("LookLeft")
    Genie.Play "LookLeftBlink"

    Robby.Wait GenieRequest
    Robby.Play "LookLeft"
End Sub

Private Sub MerlinIntro()
    Set GetRequest(8) = Genie.Get("animation", "explain, explainreturn", False)
    Merlin.Get "animation", "acknowledge", False
    
    Merlin.MoveTo CenterX + 150, CenterY
    Merlin.Wait GenieRequest
    Merlin.Show
    Merlin.Speak "Hello everyone!"
    Merlin.Play "Blink"
    Set MerlinRequest = Merlin.Speak("Microsoft Agent includes touches of my special magic that automatically keeps our mouths in sync with our voices.")
    Merlin.Play "Blink"

    Genie.Wait MerlinRequest
    Genie.Play "LookLeftBlink"

    Set MerlinRequest = Merlin.Speak("It also works with recorded voices!")
    Merlin.Play "LookRight"

    Genie.Wait MerlinRequest
    Robby.Wait MerlinRequest

    Genie.Play "LookLeftReturn"
    Robby.Play "LookLeftReturn"

    Set GenieRequest = Genie.Speak("Thanks, Merlin.")
    Genie.Play "LookLeft"
    Genie.Play "LookLeftBlink"
    
    Merlin.Wait GenieRequest
    Merlin.Play "Acknowledge"
    Set MerlinRequest = Merlin.Speak("My pleasure.")
    
    Genie.Wait MerlinRequest
    Genie.Play "LookLeftReturn"
    Set GenieRequest = Genie.Speak("And thanks also to you, Robby,")
    Genie.Play "LookLeft"
    
    Robby.Wait GenieRequest
    Robby.Play "LookRight"
    
    Genie.Play "LookLeftBlink"
    Genie.Play "LookLeftReturn"
    Set GenieRequest = Genie.Speak("for helping me point out all these great features.")
    Genie.Play "LookLeft"
    Genie.Play "LookLeftBlink"

    Robby.Wait GenieRequest
    Robby.Play "LookRightReturn"
    Robby.Play "Acknowledge"
    Set RobbyRequest = Robby.Speak("Anything for you, my blue friend.")
    Robby.Play "LookRight"
    
    Genie.Wait RobbyRequest
    Genie.Play "LookLeftReturn"
    Genie.Play "Explain"
    Genie.Speak "That's just a short summary of what's included in Microsoft Agent."

    Set GenieRequest = Genie.Speak("We hope you enjoy this new technology as much as we do.")
    Genie.Play "Blink"

    If MerlinLoaded = True Then
        Merlin.Wait GenieRequest
        Merlin.Hide

    End If

    If RobbyLoaded = True Then
        Robby.Wait GenieRequest
        Robby.Play "LookRightReturn"
        Robby.Play "LookLeft"
        Robby.Play "LookLeftReturn"
        Robby.Play "Wave"
        Robby.Hide

    End If

    Genie.Speak "It's been my humble pleasure to relay this information, O Great One."
    Genie.Play "Explain"
    Genie.Speak "The power is now in your hands."
    Genie.Play "Greet"
    Genie.Play "RestPose"
    SayGoodbye
End Sub

Private Sub SayGoodbye()
    Genie.Speak "See you around the Web."
    Genie.Hide
End Sub
