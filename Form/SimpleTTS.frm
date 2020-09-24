VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Simple 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recording Entire Passage"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton PauseBtn 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8040
      MaskColor       =   &H00808080&
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox FormatCB 
      Height          =   315
      Left            =   9030
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   8610
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox SkipTxtBox 
      Height          =   350
      Left            =   8520
      TabIndex        =   18
      Text            =   "0"
      Top             =   8475
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.CommandButton SkipBtn 
      Caption         =   "Skip"
      Height          =   350
      Left            =   8040
      TabIndex        =   17
      Top             =   8475
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.TextBox DebugTxtBox 
      BackColor       =   &H80000000&
      Height          =   1920
      Left            =   8730
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   9480
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "Reset"
      Height          =   350
      Left            =   8040
      TabIndex        =   15
      Top             =   8910
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speak Flags"
      Height          =   1575
      Left            =   8880
      TabIndex        =   13
      Top             =   9300
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chkSpFlagIsXML 
         Caption         =   "IsXML"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CheckBox chkSpFlagPersistXML 
      Caption         =   "PersistXML"
      Height          =   255
      Left            =   8400
      TabIndex        =   12
      Top             =   10170
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkSpFlagIsFilename 
      Caption         =   "IsFilename"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   10530
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkSpFlagAync 
      Caption         =   "FlagsAsync"
      Height          =   255
      Left            =   8910
      TabIndex        =   10
      Top             =   9930
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkSpFlagPurgeBeforeSpeak 
      Caption         =   "PurgeBeforeSpeak"
      Height          =   255
      Left            =   8910
      TabIndex        =   9
      Top             =   10290
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkSpFlagNLPSpeakPunc 
      Caption         =   "NLPSpeakPunc"
      Height          =   255
      Left            =   8910
      TabIndex        =   8
      Top             =   10650
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8040
      TabIndex        =   7
      Top             =   7530
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CheckBox chkShowEvents 
      Caption         =   "Show Events"
      Height          =   195
      Left            =   8415
      TabIndex        =   6
      Top             =   9870
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TextField2 
      Height          =   855
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "SimpleTTS.frx":0000
      Top             =   3960
      Width           =   3465
   End
   Begin VB.CommandButton SpeakItBtn 
      Caption         =   "Speak It"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox SaveToWavCheckBox 
      Caption         =   "Save to .wav"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "The text is a file name"
      Top             =   1260
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox TextField 
      Height          =   705
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1244
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"SimpleTTS.frx":0006
   End
   Begin VB.Label Label5 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   8445
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recording..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   3915
   End
End
Attribute VB_Name = "Simple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'
' This SimpleTTS sample application demonstrates how to create a SpVoice object
' and how to use it to speak text and save it to a .wav file.
'
' Copyright @ 2001 Microsoft Corporation All Rights Reserved.
'=============================================================================

Option Explicit

'Declare the SpVoice object.
Dim Voice As SpVoice
Attribute Voice.VB_VarHelpID = -1

Private Sub Form_Load()
On Error Resume Next

    Set Voice = New SpVoice
End Sub
Private Sub ExitBtn_Click()
 On Error Resume Next

   Unload Simple
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set Voice = Nothing
    Unload Simple
    Set Simple = Nothing
End Sub

Public Sub SpeakItBtn_Click()

    On Error GoTo Speak_Error
    Dim msg, Style, Title, Help, Ctxt, Response, MyString
    msg = "This is a fairly long document." & vbCrLf _
        & "It may take some time to be written to a wav file" & vbCrLf _
        & "and the file may be quite large.  Continue?" ' Define message.
    Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
    Title = "Recording Text to Speech to wav"   ' Define title.
    Help = "DEMO.HLP"   ' Define Help file.
    Ctxt = 1000   ' Define topic
          ' context.
    If Len(WavText) > 2500 Then
        Response = MsgBox(msg, Style, Title, Help, Ctxt)
            If Response <> vbYes Then   ' User chose no.
                Unload Me
                Exit Sub
            End If
    End If
    DoEvents
    If LanguageV = "tts" Then
        Set Voice.Voice = Voice.GetVoices().Item(TTSApp.VoiceCB.ListIndex)
    Else
        Set Voice.Voice = Voice.GetVoices().Item(TTSAppMain.VoiceCB.ListIndex)
    End If

'   If the 'Save to wav' checkbox is checked handle this special case by
'   calling the SaveToWav() function.
    If SaveToWavCheckBox Then
        SaveToWav
    Else
'       Call the Speak method with the text from the text box. We use the
'       SVSFlagsAsync flag to speak asynchronously and return immediately
'       from this call.
        If Not TextField.Text = "" Then
            Voice.Speak TextField.Text, SVSFlagsAsync
        End If
    End If
    
'   Return focus to text box
    TextField.SetFocus
    Screen.MousePointer = 0
    Unload Me
   Exit Sub
    
Speak_Error:
'MsgBox "Moose"
    'Screen.MousePointer = 0
    Unload Me
    'MsgBox "Speak Error!", vbOKOnly
End Sub

Public Sub SaveToWav()
On Error Resume Next

'   Create a wave stream
    Dim cpFileStream As New SpFileStream
    Dim TargetPath As String
    Dim DesktopPath As String
    DesktopPath = GetShellFolderPath(&H0)
    If Trim(SavePath) <> "" Then
        TargetPath = Replace(SavePath & "\" & format(Now, "ddmmyyhhmmss") & ".wav", "\\", "\")
    Else
        TargetPath = DesktopPath & "\" & format(Now, "ddmmyyhhmmss") & ".wav"
    End If
'   Set audio format
    cpFileStream.format.Type = SAFT22kHz16BitMono
    DoEvents
        
'   Create a new .wav file for writing. False indicates that we're not
'   interested in writing events into the .wav file.
'   Note - this line of code will fail if the file exists and is currently open.
    cpFileStream.Open TargetPath, SSFMCreateForWrite, False

'   Set the .wav file stream as the output for the Voice object
    Set Voice.AudioOutputStream = cpFileStream
    
'   Calling the Speak method now will send the output to the "SimpTTS.wav" file.
'   We use the SVSFDefault flag so this call does not return until the file is
'   completely written.
    Voice.Speak TextField.Text, SVSFDefault
    
'   Close the file
    cpFileStream.Close
    Set cpFileStream = Nothing
    
'   Reset the Voice object's output to 'Nothing'. This will force it to use
'   the default audio output the next time.
    Set Voice.AudioOutputStream = Nothing
    
Cancel:
    Exit Sub
End Sub

