VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form TTSApp 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reading Window"
   ClientHeight    =   6810
   ClientLeft      =   -270
   ClientTop       =   900
   ClientWidth     =   8415
   Icon            =   "TTSApp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8415
   Begin VB.CommandButton Command8 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Dictionry"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   630
      Picture         =   "TTSApp.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Looks up any Selected English Word in the Dictionary"
      Top             =   6045
      Width           =   705
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Wav"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2785
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Global save the phrase as a wav file (all windows)"
      Top             =   5760
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2250
      Picture         =   "TTSApp.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Stop and Reset"
      Top             =   6045
      Width           =   495
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5745
      Left            =   -60
      TabIndex        =   44
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   10134
      _Version        =   393217
      BackColor       =   14215660
      ScrollBars      =   3
      TextRTF         =   $"TTSApp.frx":0D5E
   End
   Begin VB.CommandButton Command57 
      Height          =   285
      Left            =   1350
      Picture         =   "TTSApp.frx":0DE0
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Save the current page"
      Top             =   6045
      Width           =   315
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   600
      Left            =   8150
      Max             =   -8
      Min             =   -36
      TabIndex        =   41
      Top             =   6170
      Value           =   -10
      Width           =   225
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5670
      Picture         =   "TTSApp.frx":1312
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Find NEXT word"
      Top             =   5760
      Width           =   500
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5130
      Picture         =   "TTSApp.frx":1BDC
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Search for any text"
      Top             =   5760
      Width           =   500
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1710
      Picture         =   "TTSApp.frx":24A6
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Reads the entire document"
      Top             =   6045
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      Picture         =   "TTSApp.frx":2D70
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Load a Text File"
      Top             =   6045
      Width           =   645
   End
   Begin VB.CheckBox chkShowEvents 
      Caption         =   "Show Events"
      Height          =   195
      Left            =   2130
      TabIndex        =   14
      Top             =   11460
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   6525
      TabIndex        =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CheckBox chkSpFlagNLPSpeakPunc 
      Caption         =   "NLPSpeakPunc"
      Height          =   255
      Left            =   3615
      TabIndex        =   13
      Top             =   12180
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkSpFlagPurgeBeforeSpeak 
      Caption         =   "PurgeBeforeSpeak"
      Height          =   255
      Left            =   3615
      TabIndex        =   11
      Top             =   11820
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkSpFlagAync 
      Caption         =   "FlagsAsync"
      Height          =   255
      Left            =   3615
      TabIndex        =   9
      Top             =   11460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkSpFlagIsFilename 
      Caption         =   "IsFilename"
      Height          =   255
      Left            =   2115
      TabIndex        =   12
      Top             =   12120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkSpFlagPersistXML 
      Caption         =   "PersistXML"
      Height          =   255
      Left            =   2115
      TabIndex        =   10
      Top             =   11760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speak Flags"
      Height          =   1575
      Left            =   8190
      TabIndex        =   7
      Top             =   11460
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chkSpFlagIsXML 
         Caption         =   "IsXML"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "Reset"
      Height          =   350
      Left            =   6525
      TabIndex        =   4
      Top             =   10740
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox DebugTxtBox 
      BackColor       =   &H80000000&
      Height          =   1920
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   7290
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton SkipBtn 
      Caption         =   "Skip"
      Height          =   350
      Left            =   6525
      TabIndex        =   2
      Top             =   10305
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.TextBox SkipTxtBox 
      Height          =   350
      Left            =   7005
      TabIndex        =   3
      Text            =   "0"
      Top             =   10305
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.ComboBox FormatCB 
      Height          =   315
      Left            =   7515
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.CommandButton PauseBtn 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6525
      MaskColor       =   &H00808080&
      TabIndex        =   1
      Top             =   9870
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Height          =   890
      Left            =   2700
      TabIndex        =   16
      ToolTipText     =   "Resume"
      Top             =   5890
      Width           =   5355
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   1650
         Picture         =   "TTSApp.frx":31B2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Record all of the text to a wav File"
         Top             =   150
         Width           =   525
      End
      Begin VB.CommandButton Run 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   630
         Picture         =   "TTSApp.frx":35F4
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Resume"
         Top             =   150
         Width           =   525
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1140
         Picture         =   "TTSApp.frx":3B5F
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Pause"
         Top             =   150
         Width           =   525
      End
      Begin Project1.cmdopen CmDlg 
         Left            =   12090
         Top             =   840
         _ExtentX        =   661
         _ExtentY        =   635
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3060
         Picture         =   "TTSApp.frx":4429
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Picture         =   "TTSApp.frx":45B3
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8730
         Picture         =   "TTSApp.frx":473D
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   690
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1770
         Picture         =   "TTSApp.frx":4B7F
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         Picture         =   "TTSApp.frx":4D09
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   450
         Picture         =   "TTSApp.frx":5373
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   30
         Picture         =   "TTSApp.frx":54FD
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3750
         Picture         =   "TTSApp.frx":5687
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         Picture         =   "TTSApp.frx":5811
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1230
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2940
         Picture         =   "TTSApp.frx":599B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2220
         Picture         =   "TTSApp.frx":5B25
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         Picture         =   "TTSApp.frx":618F
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1020
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton SpeakBtn 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         Picture         =   "TTSApp.frx":67F9
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Pronounce Selected Text"
         Top             =   150
         Width           =   525
      End
      Begin VB.ComboBox VoiceCB 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Select the Chinese or English Voices"
         Top             =   270
         Width           =   1395
      End
      Begin VB.ComboBox AudioOutputCB 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "TTSApp.frx":70C3
         Left            =   3960
         List            =   "TTSApp.frx":70C5
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Select the default Audio Playback Device"
         Top             =   600
         Width           =   1395
      End
      Begin Project1.cpvSlider cpvSlider1 
         Height          =   240
         Left            =   2700
         ToolTipText     =   "Change the rate of pronunciation"
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         BackColor       =   14215660
         SliderIcon      =   "TTSApp.frx":70C7
         Orientation     =   0
         RailPicture     =   "TTSApp.frx":71E1
         Min             =   -10
      End
      Begin Project1.cpvSlider cpvSlider2 
         Height          =   240
         Left            =   2700
         ToolTipText     =   "Adjust the Volume"
         Top             =   615
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         BackColor       =   14215660
         SliderIcon      =   "TTSApp.frx":76A9
         Orientation     =   0
         RailPicture     =   "TTSApp.frx":77C3
         Max             =   100
         Value           =   50
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Voice"
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
         Left            =   4110
         TabIndex        =   23
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   2265
         TabIndex        =   22
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vol"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Audio Output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   20
         Top             =   930
         Width           =   585
      End
   End
   Begin RichTextLib.RichTextBox MainTxtBox 
      Height          =   1275
      Left            =   120
      TabIndex        =   45
      Top             =   4350
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   2249
      _Version        =   393217
      BackColor       =   14215660
      TextRTF         =   $"TTSApp.frx":7C8B
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8070
      TabIndex        =   49
      Top             =   5820
      Width           =   255
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8070
      TabIndex        =   48
      Top             =   5970
      Width           =   315
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1350
      TabIndex        =   43
      ToolTipText     =   "Set Global Save Path"
      Top             =   6360
      Width           =   345
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Left            =   1380
      Top             =   6480
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2430
      Left            =   10440
      Shape           =   4  'Rounded Rectangle
      Top             =   3150
      Visible         =   0   'False
      Width           =   2070
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
      Left            =   6645
      TabIndex        =   5
      Top             =   10275
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRead 
         Caption         =   "Read Selected Text"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy to Clipboard"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "TTSApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'
' This VB TTS App sample demonstrates most of the TTS functionalities
' supported in SAPI 5.1. The main object used here is SpVoice.
'
' Copyright @ 2001 Microsoft Corporation All Rights Reserved.
'=============================================================================

Option Explicit


' First, declare the main SAPI object we are using in this sample. It is
' created inside Form_Load and released inside Form_Unload.
Dim WithEvents Voice As SpVoice
Attribute Voice.VB_VarHelpID = -1
Dim yTerm As String
Dim GooseFlag As Boolean
Dim GooseFlag1 As Boolean

' Speak flags is a combination of bit flags. These individual bits correspond
' to check boxes on the UI. So m_speakFlags should always be kept in sync
' with the state of those check boxes.
Dim m_speakFlags As SpeechVoiceSpeakFlags

' This is the default format we will use.
Const DefaultFmt = "SAFT22kHz16BitMono"

' We will disable the output combo box and show this if there's no audio output.
Const NoAudioOutput = "No audio ouput object available"

' We will enable/disable menu items and buttons based on current state
' m_speaking indicates whether a speak task is in progress
' m_paused indicates whether Voice.Pause is called
Private m_bSpeaking As Boolean
Private m_bPaused As Boolean


Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
    menuFileSaveToWave_Click
    Exit Sub
    If Trim(Text1.Text) <> "" Then
        DoEvents
        Load Simple
        'Simple.Show
        WavText = Trim(Text1.Text)
        Simple.TextField.Text = WavText
        Simple.SpeakItBtn_Click
    Else
        Exit Sub
    End If
Else

End If
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Check1.Value = 0
End Sub

Private Sub Check2_Click()
On Error Resume Next

If Check2.Value = 1 Then
    FormX.Check2.Value = 1
Else
    FormX.Check2.Value = 0
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next

    ResetBtn_Click
    Text1.ZOrder 0
    MainTxtBox.ZOrder 1
    SpeakBtn.Enabled = True
    Command19.Enabled = True
    Command8.Enabled = True
End Sub


Private Sub Command12_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If MainTxtBox.SelLength <> 0 And Trim(MainTxtBox.SelText) <> "" Then
    Moose = Trim(MainTxtBox.SelText)
    If InStr(Moose, " ") > 1 Then MsgBox "Select only one word please!": Exit Sub
    If Trim(Moose) <> "" Then
        FormX.Show
        FormX.ZOrder 0
        Me.ZOrder 1
        FormX.SSTab1.Tab = 0
        FormX.Text1.Text = Moose
        FormX.Command21_Click
        FormX.WindowState = 0
    End If
End If

End Sub

Private Sub Command15_Click()
Unload Me
End Sub

Private Sub Command18_Click()
End Sub

Private Sub Command19_Click()
On Error Resume Next
AddDebugInfo ("Speak")
Dim Moose, Goose As String
Dim j As Long
    ResetBtn_Click
    SpeakBtn.Enabled = False
    Command8.Enabled = False
    MainTxtBox.Text = ""
    Moose = Trim(Text1.Text)
        If Trim(Moose) <> "" Then
            Text1.ZOrder 1
            MainTxtBox.ZOrder 0
            TTSApp.MainTxtBox.Text = Moose
        Else
            TTSApp.MainTxtBox.Text = ""
            SpeakBtn.Enabled = True
            Command19.Enabled = True
            Exit Sub
        End If
    
    If Asc(Left(Trim(MainTxtBox.Text), 1)) < 0 Then ' Chinese < 0
        TTSApp.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSApp.VoiceCB.Text = "Microsoft Mary"
    End If
    If Not (m_bPaused And m_bSpeaking) Then
        Voice.Speak MainTxtBox.Text, m_speakFlags
    End If
    If m_bPaused Then Voice.Resume
    
    SetSpeakingState True, False
    If FormX.Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If
    MainTxtBox.SelStart = 0
    MainTxtBox.SelLength = Len(Text1.Text)
    MainTxtBox.SetFocus

    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
    SetSpeakingState False, m_bPaused

End Sub

Private Sub Command2_Click()
On Error Resume Next

Dim FileExtension As String
Dim Filenamme  As String, ThbIT As String, XXYX As String
XXYX = ""
On Error GoTo SaveAsError
    Command1_Click
    With TTSApp
        .CmDlg.InitialDir = App.Path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Open" 'Set dialog title
        .CmDlg.DefaultFilename = "Animal Farm George Orwell.txt"
        .CmDlg.Filter = "txt Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "Rich Text Files (*.rtf)" & Chr$(0) & "*.rtf"
        '.CmDlg.Filter = "txt Files (*.txt)" & Chr$(0) & "*.txt"

        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen
    End With
    If TTSApp.CmDlg.cFileName(1) = "" Then Exit Sub
    Filenamme = TTSApp.CmDlg.cFileName(1)
    'Open Filenamme For Input As #1
        'Do While Not EOF(1)
            'Line Input #1, ThbIT
            'XXYX = XXYX & ThbIT & vbCrLf
        'Loop
    'Close #1
    'Text1.Text = XXYX
    Text1.LoadFile Filenamme
SaveAsError:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub Command20_Click()
On Error Resume Next

   Voice.Pause
    m_bPaused = True

End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim X As Integer
If Trim(Text1.Text) = "" Then Exit Sub
Command1_Click
Text1.SetFocus
yTerm = InputBox("Find", "Input Search Term")
MousePointer = 11
foundit = True
Call Findit(Text1, yTerm)
If foundit = False Then
        Dim f As Integer
        f = TTSApp.hWnd
        Call FloatWindow(f, SINK)
        MsgBox "Cant find it !"
        Call FloatWindow(f, FLOAT)
End If
MousePointer = 0

End Sub

Private Sub Command25_Click()
On Error Resume Next
If Trim(Text1.Text) = "" Then Exit Sub
Command1_Click
Text1.SetFocus
MousePointer = 11
foundit = True
Call FinditNext(Text1, yTerm)
        'starts the find next task in the module
MousePointer = 0
If foundit = False Then
        Dim f As Integer
        f = TTSApp.hWnd
        Call FloatWindow(f, SINK)
        MsgBox "Cant find it !"
        Call FloatWindow(f, FLOAT)
End If

End Sub


Private Sub Command57_Click()
On Error Resume Next

Dim TargetPath As String
Dim DesktopPath As String
    If Trim(Text1.Text) = "" Then Exit Sub
    Command1_Click
    DesktopPath = GetShellFolderPath(&H0)
    If Trim(SavePath) <> "" Then
        TargetPath = Replace(SavePath & "\Reading Data " & format(Now, "ddmmyyhhmmss") & ".rtf", "\\", "\")
    Else
        TargetPath = DesktopPath & "\Reading Data " & format(Now, "ddmmyyhhmmss") & ".rtf"
    End If
    Text1.SaveFile TargetPath
        
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Text1.SelLength <> 0 And Trim(Text1.SelText) <> "" Then
    Moose = Trim(Text1.SelText)
    If InStr(Moose, " ") > 1 Then MsgBox "Select only one word please!": Exit Sub
    If Trim(Moose) <> "" Then
        FormX.ZOrder 0
        FormX.SSTab1.Tab = 0
        FormX.Text1.Text = Moose
        FormX.Command9_Click
        FormX.Command51_Click
    End If
End If
End Sub

Private Sub cpvSlider1_ValueChanged()
On Error Resume Next

    Voice.Rate = cpvSlider1.Value
    FormX.cpvSlider1.Value = TTSAppMain.cpvSlider1.Value

End Sub

Private Sub cpvSlider2_ValueChanged()
On Error Resume Next

    Voice.Volume = cpvSlider2.Value
    FormX.cpvSlider2.Value = TTSAppMain.cpvSlider2.Value
End Sub

Private Sub Form_Activate()
 On Error Resume Next

   FormX.WindowState = 1
    LanguageV = "tts"
    If FormX.Check2.Value = 1 Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    
    MainTxtBox.Top = Text1.Top
    MainTxtBox.Left = Text1.Left
    MainTxtBox.Width = Text1.Width
    MainTxtBox.Height = Text1.Height

End Sub

Private Sub Form_Load()
    On Error Resume Next
    ' Creates the voice object first
    Set Voice = New SpVoice
    'Me.Height = 6285
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    'SetTopMostWindow Me.hWnd, True
    GooseFlag = True
    ' Load the voices combo box
    Dim Token As ISpeechObjectToken

    For Each Token In Voice.GetVoices
        VoiceCB.AddItem (Token.GetDescription())
    Next
    VoiceCB.ListIndex = 0
    
    'load the format combo box
    AddItemToFmtCB
    
    ' set rate and volume to the same as the Voice
    
    'set the default format
    FormatCB.Text = DefaultFmt
    
    ' Load the audio output combo box
    If Voice.GetAudioOutputs.Count > 0 Then
        For Each Token In Voice.GetAudioOutputs
            AudioOutputCB.AddItem (Token.GetDescription)
        Next
    Else
        AudioOutputCB.AddItem NoAudioOutput
        AudioOutputCB.Enabled = False
    End If
    AudioOutputCB.ListIndex = 0
    
    'load image list
    LoadMouthImages
    
    
    ' init speak flags and sync flag check boxes
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.Value = Checked
    chkSpFlagPurgeBeforeSpeak.Value = Checked
    chkSpFlagIsXML.Value = Checked
    
    SetSpeakingState False, False
    'FormX.cpvSlider1.Value = TTSAppMain.cpvSlider1.Value
    'FormX.cpvSlider2.Value = TTSAppMain.cpvSlider2.Value


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = vbLeftButton Then
  'ReleaseCapture
  'SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set TTSApp = Nothing
    Set Voice = Nothing
    FormX.WindowState = 0
    'FormX.Visible = True
End Sub

Private Sub AudioOutputCB_Click()
    On Error GoTo ErrHandler
    
    ' change the output to the selected one
    Set Voice.AudioOutput = Voice.GetAudioOutputs().Item(AudioOutputCB.ListIndex)
    
    ' changing output may have also changed the format, so call function
    ' FormatCB_Click to make sure we are using the format as selected
    FormatCB_Click
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set audio output error: ", Err.Description
End Sub

Private Sub FormatCB_Click()
    On Error GoTo ErrHandler
    
    ' Note: AllowAudioOutputFormatChangesOnNextSet is a hidden property, VB
    ' object browser doesn't show it by default. To see it, you can go to
    ' VB object viewer, right click and turn on the "show hidden members".
    Voice.AllowAudioOutputFormatChangesOnNextSet = False
    
    ' The format Type is associated with the selected list item as a long.
    Voice.AudioOutputStream.format.Type = FormatCB.ItemData(FormatCB.ListIndex)
    
    ' Currently you have to call this so that SAPI picks up the new format.
    Set Voice.AudioOutputStream = Voice.AudioOutputStream
    
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set format error: ", Err.Description
End Sub

Private Sub menuFileExit_Click()
On Error Resume Next

    Unload TTSApp
    
End Sub

Private Sub menuFileOpenText_Click()
On Error Resume Next

    Dim sLocation As String
    
    ' Set CancelError is True
    'ComDlg.CancelError = True
    On Error GoTo ErrHandler
        
    ' Set flags
    'ComDlg.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    ' Set Dialog title
    'ComDlg.DialogTitle = "Open a Text File"
    ' Set open directory
    sLocation = GetDirectory()
    If Len(sLocation) <> 0 Then
        'ComDlg.InitDir = sLocation
    End If
    
    ' Set filters
    'ComDlg.Filter = "All Files (*.*)|*.*|Text, XML Files " & "(*.txt;*.xml)|*.txt;*.xml"
 
    ' Specify default filter
    'ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    'ComDlg.ShowOpen
    
    ' Now open the text file and open it in the text box.
    ' We only support text files encoded with the system code page as the
    ' binary to unicode conversion in VB is using system code page.
    'Open ComDlg.FileName For Binary Access Read As 1
    'MainTxtBox.Text = StrConv(InputB$(LOF(1), 1), vbUnicode)
    'Close #1
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Open file: ", Err.Description
    End If
End Sub

Private Sub menuFileSaveToWave_Click()
    ' Set CancelError is True
    'ComDlg.CancelError = True
    On Error GoTo ErrHandler

    ' Set flags
    'ComDlg.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNNoReadOnlyReturn
    ' Set Dialog title
    'ComDlg.DialogTitle = "Save to a Wave File"
    ' Set filters
    'ComDlg.Filter = "All Files (*.*)|*.*|Wave Files " & "(*.wav)|*.wav"
    ' Specify default filter
    'ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    'ComDlg.ShowSave
    
    ' create a wave stream
    Dim cpFileStream As New SpFileStream
    
    ' Set output format to selected format
    cpFileStream.format.Type = FormatCB.ItemData(FormatCB.ListIndex)
    
    ' Open the file for write
    'cpFileStream.Open ComDlg.FileName, SSFMCreateForWrite, False
    
    ' Set output stream to the file stream
    Voice.AllowAudioOutputFormatChangesOnNextSet = False
    Set Voice.AudioOutputStream = cpFileStream
    
    ' show action
    AddDebugInfo "Save to .wav file"
    ' speak the given text with given flags
    Voice.Speak MainTxtBox.Text, m_speakFlags
    
    ' wait until it's done speaking with a really really long timeout.
    ' the tiemout value is in unit of millisecond. -1 means forever.
    Voice.WaitUntilDone -1
    
    ' Since the output stream was set to the file stream, we need to
    ' set back to the selected audio output by calling AudioOutputCB_Click
    ' as if user just changed it through UI
    AudioOutputCB_Click
    
    ' close the file stream
    cpFileStream.Close
    Set cpFileStream = Nothing
    
    MsgBox "WAV file successfully written!", vbOKOnly, "File Saved"
    Exit Sub

ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Save to Wave file Error: ", Err.Description
    End If
    
    If Not cpFileStream Is Nothing Then
        Set cpFileStream = Nothing
    End If
End Sub

Private Sub menuFileSpeakWave_Click()
    ' Set CancelError is True
    'ComDlg.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    'ComDlg.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    ' Set Dialog title
    'ComDlg.DialogTitle = "Speak a Wave File"
    ' Set filters
    'ComDlg.Filter = "All Files (*.*)|*.*|Wave Files " & "(*.wav)|*.wav"
    ' Specify default filter
    'ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    'ComDlg.ShowOpen

    AddDebugInfo "Speak .wav file"
    
    ' Speak the contents of the wavefile. Notice here we are passing in the
    ' file name so the filename flag is set.
    'MainTxtBox.Text = ComDlg.FileName
    chkSpFlagIsFilename.Value = Checked
    SpeakBtn_Click
    
    Exit Sub

ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Speak Wave Error: ", Err.Description
    End If
    
    SetSpeakingState False, m_bPaused
    Exit Sub
End Sub

Private Sub Label14_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
End Sub

Private Sub mnuCopy_Click()
   On Error Resume Next
   Clipboard.Clear
   Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuPaste_Click()
On Error Resume Next

Text1.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuRead_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Text1.SelLength <> 0 And Trim(Text1.SelText) <> "" Then
    Moose = Trim(Text1.SelText)
        If Trim(Moose) <> "" Then
            TTSApp.MainTxtBox.Text = Moose
        Else
            TTSApp.MainTxtBox.Text = ""
        End If
    
    If Asc(Left(Trim(Text1.SelText), 1)) < 0 Then ' Chinese < 0
        TTSApp.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSApp.VoiceCB.Text = "Microsoft Mary"
    End If
    TTSApp.SpeakBtn_Click
End If

End Sub

Private Sub mnuSelect_Click()
On Error Resume Next
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Text1.SetFocus

End Sub

Private Sub PauseBtn_Click()
On Error Resume Next

    Select Case PauseBtn.Caption
    Case "Pause"
        AddDebugInfo "Pause"
        Voice.Pause
        SetSpeakingState m_bSpeaking, True
    
    Case "Resume"
        AddDebugInfo "Resume"
        Voice.Resume
        SetSpeakingState m_bSpeaking, False
    End Select
End Sub


Private Sub ResetBtn_Click()
On Error Resume Next

    'set output to default
    AudioOutputCB.ListIndex = 0
    Set Voice.AudioOutput = Nothing
    
    'use default voice
    VoiceCB.ListIndex = 0
    
    'Format to default
    FormatCB.Text = DefaultFmt
    
    'reset main text field
    MainTxtBox.Text = "Enter text you wish spoken here."
    
    'reset volume and rate
    
    ' reset speak flags
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.Value = Checked
    chkSpFlagPurgeBeforeSpeak.Value = Checked
    chkSpFlagIsXML.Value = Checked
    chkSpFlagIsFilename.Value = Unchecked
    chkSpFlagNLPSpeakPunc.Value = Unchecked
    chkSpFlagPersistXML.Value = Unchecked
    
    'reset DebugTxtbox text
    DebugTxtBox.Text = Empty
    
    'reset skip text box
    SkipTxtBox.Text = "0"
    'Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICFULL")
    
    ' if it's paused, call Resume to reset state
    If m_bPaused Then Voice.Resume

    SetSpeakingState False, False
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub Run_Click()
On Error Resume Next

   Voice.Resume
   If GooseFlag = True Then
        MainTxtBox.SetFocus
   Else
        Text1.SetFocus
   End If
End Sub

Private Sub SkipBtn_Click()
    On Error GoTo ErrHandler
    Dim SkipType As String
    Dim SkipNum As Integer
    
    AddDebugInfo "Skip"
    
    ' skip by the number specified
    SkipNum = SkipTxtBox.Text
    SkipType = "Sentence"
    
    Voice.Skip SkipType, SkipNum
    Exit Sub
    
ErrHandler:
    'MsgBox Err.Description & ":" & Err.Number, vbOKOnly, "Skip Error"
    AddDebugInfo "Skip Error: ", Err.Description
    Exit Sub
End Sub

Public Sub SpeakBtn_Click()
On Error GoTo ErrHandler
AddDebugInfo ("Speak")
Dim Moose, Goose As String
Dim j As Long
    Command19.Enabled = False
    Command8.Enabled = False
    GooseFlag1 = True
    GooseFlag = True
    ResetBtn_Click
    MainTxtBox.Text = ""
    Moose = Trim(Text1.SelText)
        If Trim(Moose) <> "" Then
            Text1.ZOrder 1
            MainTxtBox.ZOrder 0
            TTSApp.MainTxtBox.Text = Moose
        Else
            SpeakBtn.Enabled = True
            Command19.Enabled = True
            Exit Sub
        End If
    
    If Asc(Left(Trim(Text1.SelText), 1)) < 0 Then ' Chinese < 0
        TTSApp.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSApp.VoiceCB.Text = "Microsoft Mary"
    End If
    If Not (m_bPaused And m_bSpeaking) Then
        Voice.Speak MainTxtBox.Text, m_speakFlags
        'Voice.Speak Text1.Text, m_speakFlags
    End If
    ' Resume if Voice is paused
    If m_bPaused Then Voice.Resume
    
    ' set the state of menu items and buttons
    SetSpeakingState True, False
    If FormX.Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If
    MainTxtBox.SelStart = 0
    MainTxtBox.SelLength = Len(Text1.Text)
    MainTxtBox.SetFocus

    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
    SetSpeakingState False, m_bPaused
End Sub

Private Sub StopBtn_Click()
    On Error GoTo ErrHandler
    AddDebugInfo ("Stop")
    
    ' when string to speak is NULL and dwFlags is set to SPF_PURGEBEFORESPEAK
    ' it indicates to SAPI that any remaining data to be synthesized should
    ' be discarded.
    Voice.Speak vbNullString, SVSFPurgeBeforeSpeak
    If m_bPaused Then Voice.Resume
    
    SetSpeakingState False, False
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
End Sub



Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next


If Button = 2 Then
    PopupMenu menuFile
End If
End Sub

Private Sub Voice_AudioLevel(ByVal StreamNumber As Long, _
                             ByVal StreamPosition As Variant, _
                             ByVal AudioLevel As Long)
On Error Resume Next

                             
    ShowEvent "AudioLevel", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "AudioLevel=" & AudioLevel
End Sub

Private Sub Voice_Bookmark(ByVal StreamNumber As Long, _
                           ByVal StreamPosition As Variant, _
                           ByVal Bookmark As String, _
                           ByVal BookmarkId As Long)
On Error Resume Next

                           
    ShowEvent "BookMark", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "Bookmark=" & Bookmark, _
            "BookmarkId=" & BookmarkId
End Sub

Private Sub Voice_EndStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "EndStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos
    HighLightSpokenWords 0, Len(MainTxtBox.Text)
On Error Resume Next

    
    SetSpeakingState False, m_bPaused
End Sub

Private Sub Voice_EnginePrivate(ByVal StreamNumber As Long, _
                                ByVal StreamPosition As Long, _
                                ByVal lParam As Variant)
On Error Resume Next

                                
    ShowEvent "EnginePrivate", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "lParam=" & lParam
End Sub

Private Sub Voice_Phoneme(ByVal StreamNumber As Long, _
                          ByVal StreamPosition As Variant, _
                          ByVal Duration As Long, _
                          ByVal NextPhoneId As Integer, _
                          ByVal Feature As SpeechLib.SpeechVisemeFeature, _
                          ByVal CurrentPhoneId As Integer)
On Error Resume Next

                          
    ShowEvent "Phoneme", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "NextPhoneId=" & NextPhoneId, _
            "Feature=" & Feature, "CurrentPhoneId=" & CurrentPhoneId
End Sub

Private Sub Voice_Sentence(ByVal StreamNum As Long, _
                           ByVal StreamPos As Variant, _
                           ByVal pos As Long, _
                           ByVal Length As Long)
On Error Resume Next


    ShowEvent "Sentence", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & pos, "Length=" & Length
End Sub

Private Sub Voice_StartStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "StartStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos
On Error Resume Next

    ' reset the state of buttons, checkboxes and menu items
    SetSpeakingState True, m_bPaused
End Sub

Private Sub Voice_Viseme(ByVal StreamNum As Long, _
                         ByVal StreamPos As Variant, _
                         ByVal Duration As Long, _
                         ByVal VisemeType As SpeechVisemeType, _
                         ByVal Feature As SpeechVisemeFeature, _
                         ByVal VisemeId As Long)
On Error Resume Next

    ShowEvent "Viseme", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Duration=" & Duration, "VisemeType=" & VisemeType, _
            "Feature=" & Feature, "VisemeId=" & VisemeId
    
    ' Here we are going to show different mouth positions according to the viseme.
    ' The picture we show doesn't necessarily match the real mouth position.
    ' Just trying to make it more interesting.
    If VisemeId = 0 Then
        VisemeId = VisemeId + 1
    End If
   ' Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", VisemeId)
    'If (VisemeId Mod 6 = 2) Then
        'Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICEYECLOSED")
    'Else
        'If (VisemeId Mod 6 = 5) Then
            'Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICEYENARROW")
        'End If
    'End If
End Sub

Private Sub Voice_VoiceChange(ByVal StreamNum As Long, _
                              ByVal StreamPos As Variant, _
                              ByVal Token As SpeechLib.ISpeechObjectToken)
On Error Resume Next

    ShowEvent "VoiceChange", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Token=" & Token.GetDescription
    
    ' Let's sync up the combo box with the new value
    Dim i As Long
    For i = 0 To VoiceCB.ListCount - 1
        If VoiceCB.List(i) = Token.GetDescription() Then
            VoiceCB.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub Voice_Word(ByVal StreamNum As Long, _
                       ByVal StreamPos As Variant, _
                       ByVal pos As Long, _
                       ByVal Length As Long)
On Error Resume Next

    ShowEvent "Word", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & pos, "Length=" & Length
    
    Debug.Print pos, Length, MainTxtBox.SelStart, MainTxtBox.SelLength
    HighLightSpokenWords pos, Length
End Sub

Private Sub VoiceCB_Click()
On Error Resume Next

    'MsgBox VoiceCB.ListIndex
    Set Voice.Voice = Voice.GetVoices().Item(VoiceCB.ListIndex)
End Sub


' The following functions are simply to sync up the speak flags.
' When the check box is checked, the corresponding bit is set in the flags.
Private Sub chkSpFlagAync_Click()
On Error Resume Next

   m_speakFlags = SetOrClearFlag(chkSpFlagAync.Value, m_speakFlags, SVSFlagsAsync)
End Sub

Private Sub chkSpFlagIsFilename_Click()
On Error Resume Next

    m_speakFlags = SetOrClearFlag(chkSpFlagIsFilename.Value, m_speakFlags, SVSFIsFilename)
End Sub

Private Sub chkSpFlagIsXML_Click()
    ' Note: special case here. There are two flags,SVSFIsXML and SVSFIsNotXML.
    ' When neither is set, SAPI will guess by peeking at beginning characters.
    ' In this sample, we explicitly set one of them.
On Error Resume Next

    If chkSpFlagIsXML.Value = 0 Then
        ' clear SVSFIsXML bit and set SVSFIsNotXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsXML
        m_speakFlags = m_speakFlags Or SVSFIsNotXML
    Else
        ' clear SVSFIsNotXML bit and set SVSFIsXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsNotXML
        m_speakFlags = m_speakFlags Or SVSFIsXML
    End If
End Sub

Private Sub chkSpFlagNLPSpeakPunc_Click()
On Error Resume Next

    m_speakFlags = SetOrClearFlag(chkSpFlagNLPSpeakPunc.Value, m_speakFlags, SVSFNLPSpeakPunc)
End Sub

Private Sub chkSpFlagPersistXML_Click()
 On Error Resume Next

   m_speakFlags = SetOrClearFlag(chkSpFlagPersistXML.Value, m_speakFlags, SVSFPersistXML)
End Sub

Private Sub chkSpFlagPurgeBeforeSpeak_Click()
On Error Resume Next

    m_speakFlags = SetOrClearFlag(chkSpFlagPurgeBeforeSpeak.Value, m_speakFlags, SVSFPurgeBeforeSpeak)
End Sub


Private Sub AddFmts(ByRef Name As String, ByVal fmt As SpeechAudioFormatType)
    Dim Index As String
On Error Resume Next

    ' get the count of existing list so that we are adding to the bottom of the list
    Index = FormatCB.ListCount
    
    ' add the name to the list box and associate the format type with the item
    FormatCB.AddItem Name, Index
    FormatCB.ItemData(Index) = fmt
End Sub

Private Sub AddItemToFmtCB()
On Error Resume Next

    AddFmts "SAFT8kHz8BitMono", SAFT8kHz16BitMono
    AddFmts "SAFT8kHz8BitStereo", SAFT8kHz8BitStereo
    AddFmts "SAFT8kHz16BitMono", SAFT8kHz16BitMono
    AddFmts "SAFT8kHz16BitStereo", SAFT8kHz16BitStereo
    
    AddFmts "SAFT11kHz8BitMono", SAFT11kHz8BitMono
    AddFmts "SAFT11kHz8BitStereo", SAFT11kHz8BitStereo
    AddFmts "SAFT11kHz16BitMono", SAFT11kHz16BitMono
    AddFmts "SAFT11kHz16BitStereo", SAFT11kHz16BitStereo
    
    AddFmts "SAFT12kHz8BitMono", SAFT12kHz8BitMono
    AddFmts "SAFT12kHz8BitStereo", SAFT12kHz8BitStereo
    AddFmts "SAFT12kHz16BitMono", SAFT12kHz16BitMono
    AddFmts "SAFT12kHz16BitStereo", SAFT12kHz16BitStereo
    
    AddFmts "SAFT16kHz8BitMono", SAFT16kHz8BitMono
    AddFmts "SAFT16kHz8BitStereo", SAFT16kHz8BitStereo
    AddFmts "SAFT16kHz16BitMono", SAFT16kHz16BitMono
    AddFmts "SAFT16kHz16BitStereo", SAFT16kHz16BitStereo
    
    AddFmts "SAFT22kHz8BitMono", SAFT22kHz8BitMono
    AddFmts "SAFT22kHz8BitStereo", SAFT22kHz8BitStereo
    AddFmts "SAFT22kHz16BitMono", SAFT22kHz16BitMono
    AddFmts "SAFT22kHz16BitStereo", SAFT22kHz16BitStereo
    
    AddFmts "SAFT24kHz8BitMono", SAFT24kHz8BitMono
    AddFmts "SAFT24kHz8BitStereo", SAFT24kHz8BitStereo
    AddFmts "SAFT24kHz16BitMono", SAFT24kHz16BitMono
    AddFmts "SAFT24kHz16BitStereo", SAFT24kHz16BitStereo
    
    AddFmts "SAFT32kHz8BitMono", SAFT32kHz8BitMono
    AddFmts "SAFT32kHz8BitStereo", SAFT32kHz8BitStereo
    AddFmts "SAFT32kHz16BitMono", SAFT32kHz16BitMono
    AddFmts "SAFT32kHz16BitStereo", SAFT32kHz16BitStereo
    
    AddFmts "SAFT44kHz8BitMono", SAFT44kHz8BitMono
    AddFmts "SAFT44kHz8BitStereo", SAFT44kHz8BitStereo
    AddFmts "SAFT44kHz16BitMono", SAFT44kHz16BitMono
    AddFmts "SAFT44kHz16BitStereo", SAFT44kHz16BitStereo
    
    AddFmts "SAFT48kHz8BitMono", SAFT48kHz8BitMono
    AddFmts "SAFT48kHz8BitStereo", SAFT48kHz8BitStereo
    AddFmts "SAFT48kHz16BitMono", SAFT48kHz16BitMono
    AddFmts "SAFT48kHz16BitStereo", SAFT48kHz16BitStereo
End Sub
Private Sub LoadMouthImages()
    On Error GoTo ErrHandler
    
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description & ":" & Err.Number, vbOKOnly, "Load Images Error"
End Sub

Private Sub AddDebugInfo(DebugStr As String, Optional Error As String = Empty)
    ' This function adds debug string to the info window.
On Error Resume Next

    
    ' First of all, let's delete a few charaters if the text box is about to
    ' overflow. In this sample we are using the default limit of charaters.
    If Len(DebugTxtBox.Text) > 64000 Then
        Debug.Print "Too much stuff in the debug window. Remove first 10K chars"
        DebugTxtBox.SelStart = 0
        DebugTxtBox.SelLength = 10240
        DebugTxtBox.SelText = ""
    End If
    
    ' append the string to the DebugTxtBox text box and add a newline
    DebugTxtBox.SelStart = Len(DebugTxtBox.Text)
    DebugTxtBox.SelText = DebugStr & Error & vbCrLf
End Sub

Private Sub ShowEvent(ParamArray strArray())
On Error Resume Next

    If chkShowEvents.Value = Checked Then
        Dim strText As String
        strText = Join(strArray, ", ")
        AddDebugInfo "  Event: " & strText
    End If
End Sub

Private Sub HighLightSpokenWords(ByVal pos As Long, ByVal Length As Long)
    On Error GoTo ErrHandler
    If chkSpFlagIsFilename.Value = Unchecked Then
        MainTxtBox.SelStart = pos
        MainTxtBox.SelLength = Length
    End If
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Failed to high light words. This may be caused by too many charaters in the main text box."
End Sub

' This following helper function will set or clear a bit (flag) in the given
' integer (base) according to the condition (cond). If cond is 0, the bit
' is cleared. Otherwise, the bit is set. The resulting integer is returned.
Private Function SetOrClearFlag(ByVal cond As Long, _
                                ByVal base As Long, _
                                ByVal Flag As Long) As Long
On Error Resume Next

    If cond = 0 Then
        ' the condition is false, clear the flag
        SetOrClearFlag = base And Not Flag
    Else
        ' the condition is false, set the flag
        SetOrClearFlag = base Or Flag
    End If
End Function

Private Sub SetSpeakingState(ByVal bSpeaking As Boolean, ByVal bPaused As Boolean)
    ' change state of menu items and buttons accordingly
    'menuFileOpenText.Enabled = Not bSpeaking
    'menuFileSpeakWave.Enabled = Not bSpeaking
    'menuFileSaveToWave.Enabled = Not bSpeaking
    
    'SpeakBtn.Enabled = True
On Error Resume Next

    StopBtn.Enabled = bSpeaking
    SkipBtn.Enabled = (bSpeaking And Not bPaused)
    PauseBtn.Enabled = bSpeaking
    
    If bPaused Then
        PauseBtn.Caption = "Resume"
    Else
        PauseBtn.Caption = "Pause"
    End If
    
    m_bSpeaking = bSpeaking
    m_bPaused = bPaused
End Sub

Public Function GetDirectory() As String

    Err.Clear

    On Error GoTo ErrHandler

    Dim DataKey As ISpeechDataKey
    Dim Category As New SpObjectTokenCategory
    
    'Get the sdk installation location from the registry
    'The value is under "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech". The string name is SDKPath"
    Category.SetId SpeechRegistryLocalMachineRoot
    Set DataKey = Category.GetDataKey
    GetDirectory = DataKey.GetStringValue("SDKPath")
    GetDirectory = GetDirectory + "samples\common"
    
    
    
ErrHandler:
    If Err.Number <> 0 Then
        GetDirectory = ""
    End If
End Function

Private Sub VScroll1_Change()
Text1.Font.Size = Abs(VScroll1.Value)
MainTxtBox.Font.Size = Abs(VScroll1.Value)
End Sub
