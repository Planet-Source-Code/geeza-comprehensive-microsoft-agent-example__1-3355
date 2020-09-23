VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmScript 
   Caption         =   "Running through a set script"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   1020
      Width           =   3495
   End
   Begin AgentObjectsCtl.Agent Agent 
      Left            =   360
      Top             =   300
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Genie As IAgentCtlCharacter
Dim Merlin As IAgentCtlCharacter
Dim LoadRequest(3)
Dim GenieRequest As IAgentCtlRequest
Dim MerlinRequest As IAgentCtlRequest

Private Sub Agent_RequestComplete(ByVal Request As Object)
    Select Case Request
        Case LoadRequest(1)
            If Request.Status <> 0 Then
                ' If genie's character data fails to load post a message
                Label1.Caption = "Genie character failed to load."

                MsgBox "Unable to load the Genie character. Program will terminate"
                Unload Me
            Else
                Label1.Caption = "Genie character successfully loaded."
                'Create a reference to the character
                Set Genie = Agent.Characters("Genie")
                Call GenieIntro
            End If

        Case LoadRequest(2)
            If Request.Status <> 0 Then
                Label1.Caption = "Merlin character failed to load."
                Genie.Speak "Uh oh, Merlin failed to load."
                Genie.Speak "You'll have to exit!"
                Genie.Hide
            Else
                Label1.Caption = "Merlin character successfully loaded."
                Set Merlin = Agent.Characters("Merlin")
                Call MerlinIntro
            End If
        Case LoadRequest(3)
            Unload Me
    End Select
End Sub

Private Sub LoadGenie()
    Set LoadRequest(1) = Agent.Characters.Load("Genie", "Genie.acs")
    Label1.Caption = "Loading Genie character."
End Sub

Private Sub LoadMerlin()
    Set LoadRequest(2) = Agent.Characters.Load("Merlin", "Merlin.acs")
    Label1.Caption = "Loading Merlin character."
End Sub

Private Sub GenieIntro()
    Genie.MoveTo 170, 240
    Genie.Show
    Genie.Speak "Hello!  I am Genie, a Microsoft Agent Character"
    Genie.Play "Greet"
    Genie.Speak "and I am at your service"
    Genie.Play "Blink"
    
    Genie.Speak "I hope Geeza's code will help you to create Agent scripts easily too, and get around the problem that we can't pronounce Geeza correctly.  \pau=200\G-eeza, \pau=200\GG eeza..."
    
    Call LoadMerlin

    Set GenieRequest = Genie.Speak("But for now, here's Merlin!")
    Genie.Play "LookLeft"
End Sub

Private Sub MerlinIntro()
    Merlin.MoveTo 320, 240
    Merlin.Wait GenieRequest
    Merlin.Show
    Merlin.Play "Wave"
    Merlin.Speak "Hello to you all.  I am Merlin; another Microsoft Agent Character"
    Merlin.Play "Blink"

    Call Routine1
End Sub

Private Sub Routine1()
    Genie.Play "Blink"
    
    Call Time_For_Tubbie_Bye_Byes
End Sub

Private Sub Time_For_Tubbie_Bye_Byes()
    Merlin.Wait GenieRequest
    Set MerlinRequest = Merlin.Speak("So it's goodbye from me,")
    Genie.Wait MerlinRequest
    Set GenieRequest = Genie.Speak("And it's goodbye from him")
    Merlin.Wait GenieRequest
    Set MerlinRequest = Merlin.Speak("Goodnight")
    Genie.Wait MerlinRequest  'I had to do it, sorry!
    Genie.Hide                'The Two Ronnie's rocked :)
    Set LoadRequest(3) = Merlin.Hide
End Sub

Private Sub Form_Load()
    Call LoadGenie
End Sub

'This is a very much earlier version of the other script,
'but it shows how to make a continuous script for the
'Agents, without the dreaded
'"Variable without Block variable set" or whatever it is,
'enjoy!

'GEEZA
'Off to the land of Dreams ~~~~~~~~~~~~~~
