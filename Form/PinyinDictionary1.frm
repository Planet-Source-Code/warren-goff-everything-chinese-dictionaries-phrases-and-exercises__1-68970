VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Begin VB.Form FormX 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speaking Chinese English Dictionary"
   ClientHeight    =   6360
   ClientLeft      =   4935
   ClientTop       =   7575
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   Icon            =   "PinyinDictionary1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7575
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      ToolTipText     =   "You may select text to hear the Pronunciation"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Top"
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
      Height          =   240
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "Keep Program on Top"
      Top             =   60
      Width           =   500
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00E3DFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   30
      TabIndex        =   29
      Text            =   "Simplified Chinese"
      ToolTipText     =   "Paste a Simplified Chinese Search Term"
      Top             =   960
      Width           =   2700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E3DFE0&
      BorderStyle     =   0  'None
      Height          =   925
      Left            =   4485
      TabIndex        =   174
      Top             =   330
      Width           =   2775
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E3DFE0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   15
         TabIndex        =   189
         Top             =   360
         Width           =   1005
      End
      Begin VB.CheckBox Check7 
         Caption         =   "1 @ time"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   0
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox Check6 
         Caption         =   "1 @ time"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   225
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3DFE0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   1545
         Picture         =   "PinyinDictionary1.frx":08CA
         ScaleHeight     =   885
         ScaleWidth      =   975
         TabIndex        =   175
         ToolTipText     =   "Speaking Chinese English Language Tool"
         Top             =   0
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1005
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00E3DFE0&
      Caption         =   "TTS"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6465
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   "View the Text to Speech Window"
      Top             =   60
      Width           =   500
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Wav"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   152
      ToolTipText     =   "Global save the phrase as a wav file (all windows)"
      Top             =   60
      Width           =   500
   End
   Begin VB.TextBox SP 
      Alignment       =   2  'Center
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Text            =   "PinyinDictionary1.frx":1466
      ToolTipText     =   "Chinese Phrases"
      Top             =   960
      Width           =   2565
   End
   Begin VB.CommandButton Command55 
      BackColor       =   &H00E3DFE0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   101
      ToolTipText     =   "Append Simplified Chinese to the search box"
      Top             =   960
      Width           =   360
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00E3DFE0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   "Append Numbered Pinyin to the search box"
      Top             =   650
      Width           =   360
   End
   Begin VB.TextBox Pn1 
      Alignment       =   2  'Center
      BackColor       =   &H000080E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Index           =   0
      Left            =   17250
      TabIndex        =   92
      Text            =   "a1"
      Top             =   6975
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17925
      Picture         =   "PinyinDictionary1.frx":1470
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   6630
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H000080E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19290
      Picture         =   "PinyinDictionary1.frx":1D3A
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   6630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17250
      Picture         =   "PinyinDictionary1.frx":22C4
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   6630
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H000080E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   18660
      Picture         =   "PinyinDictionary1.frx":2B8E
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   6630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Convert Txt-Htm"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   20865
      TabIndex        =   87
      Top             =   6075
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox text10 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   20115
      TabIndex        =   86
      Text            =   "Text3"
      Top             =   930
      Width           =   1485
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   17670
      TabIndex        =   69
      Top             =   7410
      Width           =   5610
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2190
         TabIndex        =   75
         Text            =   "bu2bi4jian1xian3"
         Top             =   765
         Width           =   2880
      End
      Begin VB.CommandButton Command40 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Convert Pinyin to Simplified"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2235
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1845
         Width           =   1380
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1125
         Width           =   2880
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1485
         Width           =   2880
      End
      Begin VB.CommandButton Command42 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Convert File: PinyinBiblet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1845
         Width           =   1380
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Display Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   615
         TabIndex        =   70
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caution this conversion is inaccurate and beyond my ability.   Sorry!"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   1095
         TabIndex        =   80
         Top             =   2655
         Width           =   3600
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pinyin to Simplified and beyond"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   975
         TabIndex        =   79
         Top             =   255
         Width           =   3690
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Numbered Pinyin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   390
         TabIndex        =   78
         Top             =   765
         Width           =   1770
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Simplified Chinese"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   225
         TabIndex        =   77
         Top             =   1125
         Width           =   2025
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "English Definition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   330
         TabIndex        =   76
         Top             =   1485
         Width           =   2025
      End
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Both"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Both Translate and Define"
      Top             =   350
      Width           =   450
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      ItemData        =   "PinyinDictionary1.frx":2FD0
      Left            =   17460
      List            =   "PinyinDictionary1.frx":2FD7
      TabIndex        =   42
      Top             =   4800
      Width           =   3900
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   17790
      TabIndex        =   41
      Top             =   6360
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Def"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "English Definition"
      Top             =   350
      Width           =   705
   End
   Begin VB.TextBox PP 
      Alignment       =   2  'Center
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Text            =   "PinyinDictionary1.frx":3042
      ToolTipText     =   "Pinyin Phrases"
      Top             =   650
      Width           =   2565
   End
   Begin VB.TextBox Ep 
      Alignment       =   2  'Center
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Text            =   "PinyinDictionary1.frx":304C
      ToolTipText     =   "English Phrases"
      Top             =   350
      Width           =   2565
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Search for Simplified Chinese to English Translation"
      Top             =   960
      Width           =   630
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00E3DFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3750
      TabIndex        =   11
      ToolTipText     =   "Numbered Pinyin List Generates all Simplified Chinese possibilities below"
      Top             =   650
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E3DFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Text            =   "English"
      ToolTipText     =   "Input an English Search Term"
      Top             =   350
      Width           =   2700
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E3DFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   30
      TabIndex        =   5
      Text            =   "Numbered Pinyin"
      ToolTipText     =   "Input an Pinyin Search Term"
      Top             =   650
      Width           =   2700
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "English to Chinese Translation"
      Top             =   350
      Width           =   630
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E3DFE0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Search for Pinyin to English Translation"
      Top             =   650
      Width           =   630
   End
   Begin VB.CommandButton Translit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Translit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   17340
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4110
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox Combo14 
      BackColor       =   &H00E3DFE0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3750
      TabIndex        =   31
      ToolTipText     =   "Simplified Chinese possibilities for each numbered Pinyin above"
      Top             =   930
      Width           =   780
   End
   Begin VB.TextBox English 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Pronoun"
      Top             =   660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MingLiU"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3060
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   75
      Top             =   6270
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14933984
      ForeColor       =   16711680
      MouseIcon       =   "PinyinDictionary1.frx":3056
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dictionary"
      TabPicture(0)   =   "PinyinDictionary1.frx":3072
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label27"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label37"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label33"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Simp(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command52"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command47"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "VScroll1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command57"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command22"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Dictionary"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Option1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command26"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Phrases"
      TabPicture(1)   =   "PinyinDictionary1.frx":308E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(1)=   "Shape2"
      Tab(1).Control(2)=   "SSTab7"
      Tab(1).Control(3)=   "SSTab4"
      Tab(1).Control(4)=   "Command58"
      Tab(1).Control(5)=   "Command62"
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(7)=   "List4"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&Web"
      TabPicture(2)   =   "PinyinDictionary1.frx":30AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Drugs Lab"
      TabPicture(3)   =   "PinyinDictionary1.frx":30C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Combo1"
      Tab(3).Control(2)=   "Command7"
      Tab(3).Control(3)=   "Check4"
      Tab(3).Control(4)=   "Check5"
      Tab(3).Control(5)=   "WebBrowser5"
      Tab(3).Control(6)=   "Command59"
      Tab(3).Control(7)=   "Command11"
      Tab(3).Control(8)=   "Command60"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "&Medical Dictionary"
      TabPicture(4)   =   "PinyinDictionary1.frx":30E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label7"
      Tab(4).Control(1)=   "Label9"
      Tab(4).Control(2)=   "Shape6"
      Tab(4).Control(3)=   "Label40"
      Tab(4).Control(4)=   "Label6"
      Tab(4).Control(5)=   "Frame1"
      Tab(4).Control(6)=   "List5"
      Tab(4).Control(7)=   "Option11"
      Tab(4).Control(8)=   "Option12"
      Tab(4).Control(9)=   "Command13"
      Tab(4).Control(10)=   "Command15"
      Tab(4).Control(11)=   "Text3"
      Tab(4).Control(12)=   "Command19"
      Tab(4).Control(13)=   "Command20"
      Tab(4).Control(14)=   "Command61"
      Tab(4).ControlCount=   15
      Begin VB.CommandButton Command26 
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   190
         Top             =   5430
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D8E9EC&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   285
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Pronounce English"
         Top             =   5430
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D8E9EC&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   45
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Pronounce Chinese"
         Top             =   5430
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   270
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4935
         Left            =   -75000
         TabIndex        =   12
         Top             =   1290
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8705
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Browse"
         TabPicture(0)   =   "PinyinDictionary1.frx":30FE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label20"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "WebBrowser4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command34"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command35"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Command23"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Text7"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Command36"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Command37"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Combo10"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Command14"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Command39"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "WebBrowser1"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Command10"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Command1"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Combo11"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Command5"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Links"
         TabPicture(1)   =   "PinyinDictionary1.frx":311A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(2)=   "Linkz"
         Tab(1).Control(3)=   "Command6"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Chat"
         TabPicture(2)   =   "PinyinDictionary1.frx":3136
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label15"
         Tab(2).Control(1)=   "Label26"
         Tab(2).Control(2)=   "WebBrowser3"
         Tab(2).Control(3)=   "Check3"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Radio"
         TabPicture(3)   =   "PinyinDictionary1.frx":3152
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "WebBrowser2"
         Tab(3).Control(1)=   "Command67"
         Tab(3).Control(2)=   "Command68"
         Tab(3).Control(3)=   "Command69"
         Tab(3).Control(4)=   "AutoResizer51"
         Tab(3).ControlCount=   5
         Begin VB.PictureBox AutoResizer51 
            Height          =   255
            Left            =   -67920
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   186
            Top             =   4560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   750
            Picture         =   "PinyinDictionary1.frx":316E
            Style           =   1  'Graphical
            TabIndex        =   181
            ToolTipText     =   "Forward to Next Page"
            Top             =   4110
            Width           =   735
         End
         Begin VB.CheckBox Check3 
            Caption         =   "http://zhongwen.com/chat.htm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "This is a robotic site and actual Chat may not be occurring."
            Top             =   4185
            Width           =   2955
         End
         Begin VB.ComboBox Combo11 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   315
            Left            =   45
            TabIndex        =   17
            Text            =   "Url"
            ToolTipText     =   "This is a list of all open instances of Internet Explorer, the URL's"
            Top             =   3810
            Width           =   2730
         End
         Begin VB.CommandButton Command1 
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
            Height          =   465
            Left            =   1650
            Picture         =   "PinyinDictionary1.frx":3552
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Opens the present page in the default browser (IE)"
            Top             =   4110
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   15
            Picture         =   "PinyinDictionary1.frx":3894
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Back to previous page "
            Top             =   4110
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   -74895
            Style           =   1  'Graphical
            TabIndex        =   150
            ToolTipText     =   "Opens the present page in the default browser (IE)"
            Top             =   4230
            Width           =   1200
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   3735
            Left            =   0
            TabIndex        =   135
            ToolTipText     =   "Browser Display Window(right click to navigate)"
            Top             =   60
            Width           =   7365
            ExtentX         =   12991
            ExtentY         =   6588
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.CommandButton Command69 
            Caption         =   "Chinese"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -74970
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Opens the present page in the default browser (IE)"
            Top             =   4140
            Width           =   1770
         End
         Begin VB.CommandButton Command68 
            Caption         =   "English XM"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -73050
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Opens the present page in the default browser (IE)"
            Top             =   4140
            Width           =   1770
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Default      Browser"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   -69600
            Picture         =   "PinyinDictionary1.frx":415E
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Opens the present page in the default browser (IE)"
            Top             =   4110
            Width           =   1620
         End
         Begin VB.ListBox Linkz 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3570
            Left            =   -74895
            OLEDropMode     =   1  'Manual
            TabIndex        =   24
            ToolTipText     =   "A variety of relevent links (please notify if any are dead)"
            Top             =   510
            Width           =   7215
         End
         Begin VB.CommandButton Command39 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5580
            Picture         =   "PinyinDictionary1.frx":43B2
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Save the current page"
            Top             =   3840
            Width           =   315
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Url"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2670
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.ComboBox Combo10 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   315
            Left            =   2775
            TabIndex        =   16
            Text            =   "Page Title"
            ToolTipText     =   "This is a list of all open instances of Internet Explorer, the Page Titles"
            Top             =   3810
            Width           =   2730
         End
         Begin VB.CommandButton Command37 
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5940
            Picture         =   "PinyinDictionary1.frx":48E4
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Read the present Page in English"
            Top             =   3990
            Width           =   630
         End
         Begin VB.CommandButton Command36 
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6480
            Picture         =   "PinyinDictionary1.frx":51AE
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Read the present Page in Chinese"
            Top             =   3990
            Width           =   630
         End
         Begin RichTextLib.RichTextBox Text7 
            Height          =   3045
            Left            =   60
            TabIndex        =   20
            Top             =   5040
            Visible         =   0   'False
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   5371
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"PinyinDictionary1.frx":5A78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton Command23 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Chin-English"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3405
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Simplified Chinese to English (Navigate to desired page in IE and Choose Direction of translation)"
            Top             =   4110
            Width           =   1100
         End
         Begin VB.CommandButton Command35 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Chin-Pinyin"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4500
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Simplified Chinese to Numbered Pinyin (Navigate to desired page in IE and Choose Direction of translation)"
            Top             =   4110
            Width           =   990
         End
         Begin VB.CommandButton Command34 
            BackColor       =   &H00D8E9EC&
            Caption         =   "English-Chin"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2310
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "English to Simplified Chinese (Navigate to desired page in IE and Choose Direction of translation)"
            Top             =   4110
            Width           =   1110
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser3 
            Height          =   3585
            Left            =   -74850
            TabIndex        =   26
            Top             =   570
            Visible         =   0   'False
            Width           =   7185
            ExtentX         =   12674
            ExtentY         =   6324
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser4 
            Height          =   3000
            Left            =   660
            TabIndex        =   21
            Top             =   390
            Visible         =   0   'False
            Width           =   5625
            ExtentX         =   9922
            ExtentY         =   5292
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser2 
            Height          =   4005
            Left            =   -75000
            TabIndex        =   138
            ToolTipText     =   "Browser Display Window"
            Top             =   90
            Width           =   7365
            ExtentX         =   12991
            ExtentY         =   7064
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Some of these links may be inactive"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -73200
            TabIndex        =   183
            Top             =   4230
            Width           =   5070
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Check the box below to begin Chatting"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1470
            Left            =   -74520
            TabIndex        =   28
            Top             =   1200
            Width           =   5985
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "International Chat Cafe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   -74865
            TabIndex        =   27
            Top             =   120
            Width           =   6705
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Chinese Language Links"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   -74895
            TabIndex        =   25
            Top             =   90
            Width           =   6840
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Read Page"
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
            Height          =   225
            Left            =   6090
            TabIndex        =   23
            Top             =   3810
            Width           =   915
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Browser Data"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   555
            TabIndex        =   22
            Top             =   -345
            Width           =   4410
         End
      End
      Begin VB.ListBox List4 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3210
         Left            =   -74460
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   2160
         Width           =   6810
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74430
         TabIndex        =   44
         Top             =   1830
         Width           =   6765
         Begin VB.OptionButton Option10 
            BackColor       =   &H00D8E9EC&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   4335
            MaskColor       =   &H00FFC0FF&
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Pronounce Chinese"
            Top             =   -15
            Value           =   -1  'True
            Width           =   270
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00D8E9EC&
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4575
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Pronounce English"
            Top             =   -15
            Width           =   270
         End
         Begin VB.CommandButton Command54 
            Caption         =   "Pronounce Selected"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4890
            TabIndex        =   46
            ToolTipText     =   "Pronounce the Selected Item"
            Top             =   15
            Width           =   1470
         End
         Begin VB.CommandButton Command53 
            Caption         =   "Remove UnChecked"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   45
            ToolTipText     =   "Removes the unselected items.  Clicking the Tab wil refresh the items"
            Top             =   15
            Width           =   1470
         End
         Begin VB.Label Label36 
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
            Left            =   6405
            TabIndex        =   157
            ToolTipText     =   "Set Global Save Path"
            Top             =   -105
            Width           =   345
         End
         Begin VB.Shape Shape5 
            Height          =   195
            Left            =   6405
            Top             =   15
            Width           =   315
         End
      End
      Begin RichTextLib.RichTextBox Dictionary 
         Height          =   4185
         Left            =   30
         TabIndex        =   159
         ToolTipText     =   "Dictionary Display.  Select text to pronounce"
         Top             =   1230
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   7382
         _Version        =   393217
         BackColor       =   14933984
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"PinyinDictionary1.frx":5AF9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2400
         TabIndex        =   129
         Top             =   5310
         Width           =   1485
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E3DFE0&
            Caption         =   "Any 1sr"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Any word First occurance"
            Top             =   570
            Width           =   645
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E3DFE0&
            Caption         =   "1st-1st"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "First word and first occurance"
            Top             =   360
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E3DFE0&
            Caption         =   "1st ALL"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "First word and ALL occurances"
            Top             =   375
            Width           =   645
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E3DFE0&
            Caption         =   "Any-All"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Any word All occurances"
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "Search Criteria"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   30
            TabIndex        =   133
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.CommandButton Command60 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Allg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -69290
         Picture         =   "PinyinDictionary1.frx":5B86
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Adds Selected to Medical Hx Allergies"
         Top             =   5430
         Width           =   465
      End
      Begin VB.CommandButton Command11 
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
         Height          =   700
         Left            =   -68300
         Picture         =   "PinyinDictionary1.frx":6450
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Convert Simplified Chinese to Pinyin"
         Top             =   5430
         Width           =   525
      End
      Begin VB.CommandButton Command59 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Rx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   -69765
         Picture         =   "PinyinDictionary1.frx":6D1A
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Adds Selected to Medical Hx Medications"
         Top             =   5430
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E3DFE0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5850
         Picture         =   "PinyinDictionary1.frx":75E4
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Pronounce highlighted text"
         Top             =   5580
         Width           =   705
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E3DFE0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5160
         Picture         =   "PinyinDictionary1.frx":7EAE
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Convert .wav to .mp3"
         Top             =   5580
         Width           =   705
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser5 
         Height          =   4140
         Left            =   -75030
         TabIndex        =   50
         ToolTipText     =   "Browser Display Window(right click to navigate)"
         Top             =   1260
         Width           =   7440
         ExtentX         =   13123
         ExtentY         =   7302
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CheckBox Check5 
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
         Height          =   405
         Left            =   -72480
         Picture         =   "PinyinDictionary1.frx":8778
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Displays Normal Lab values"
         Top             =   5820
         Width           =   645
      End
      Begin VB.CommandButton Command62 
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
         Height          =   405
         Left            =   -74970
         Picture         =   "PinyinDictionary1.frx":8A82
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Adds Selected to Medical Hx Phrases"
         Top             =   1350
         Width           =   345
      End
      Begin VB.CheckBox Check4 
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
         Height          =   405
         Left            =   -71760
         Picture         =   "PinyinDictionary1.frx":934C
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Opens Medical History Data form"
         Top             =   5820
         Width           =   735
      End
      Begin VB.CommandButton Command61 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   505
         Left            =   -68760
         Picture         =   "PinyinDictionary1.frx":9C16
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Adds Selected to Medical Hx Conditions"
         Top             =   5570
         Width           =   690
      End
      Begin VB.CommandButton Command57 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         Picture         =   "PinyinDictionary1.frx":A4E0
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Save the current page"
         Top             =   1350
         Width           =   315
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   555
         Left            =   7125
         Max             =   -8
         Min             =   -36
         TabIndex        =   102
         Top             =   4770
         Value           =   -8
         Width           =   255
      End
      Begin VB.CommandButton Command47 
         BackColor       =   &H00E3DFE0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4470
         Picture         =   "PinyinDictionary1.frx":AA12
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Reading, Pronunciation & Recording Window"
         Top             =   5580
         Width           =   705
      End
      Begin VB.CommandButton Command20 
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
         Height          =   405
         Left            =   -74280
         Picture         =   "PinyinDictionary1.frx":AE54
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Reloads the dictionary after Removing un-checked terms"
         Top             =   5520
         Width           =   435
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Go"
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
         Left            =   -69870
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5790
         Width           =   390
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   -72690
         TabIndex        =   60
         ToolTipText     =   "Search for any term"
         Top             =   5790
         Width           =   2745
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Remove UnChecked"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74850
         TabIndex        =   59
         ToolTipText     =   "Removes the unselected items.  Clicking the Tab wil refresh the items"
         Top             =   1260
         Width           =   1470
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Pronounce Selected"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -70125
         TabIndex        =   58
         ToolTipText     =   "Pronounce the Selected Item"
         Top             =   1290
         Width           =   1470
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00D8E9EC&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -70440
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Pronounce English"
         Top             =   1260
         Width           =   270
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00D8E9EC&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -70680
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Pronounce Chinese"
         Top             =   1260
         Width           =   270
      End
      Begin VB.ListBox List5 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3885
         Left            =   -74970
         Style           =   1  'Checkbox
         TabIndex        =   55
         ToolTipText     =   "Simplified Chinese-Pinyin-English Medical Dictionary"
         Top             =   1560
         Width           =   7380
      End
      Begin VB.CommandButton Command7 
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
         Height          =   700
         Left            =   -68810
         Picture         =   "PinyinDictionary1.frx":B71E
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Copy any selected text to the Clipboard and Clicking here will pronounce it"
         Top             =   5430
         Width           =   525
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -74940
         Sorted          =   -1  'True
         TabIndex        =   51
         Text            =   "Drug Classes"
         ToolTipText     =   "A list of common Drug Classes calls up the list of drugs in English and Simplified Chinese"
         Top             =   5430
         Width           =   5175
      End
      Begin VB.CommandButton Command52 
         BackColor       =   &H00E3DFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   900
         Picture         =   "PinyinDictionary1.frx":BFE8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Clear Screen"
         Top             =   5580
         Width           =   705
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72660
         TabIndex        =   10
         Top             =   7080
         Width           =   1500
      End
      Begin VB.ComboBox Simp 
         BackColor       =   &H000080E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   660
         Index           =   0
         ItemData        =   "PinyinDictionary1.frx":C42A
         Left            =   690
         List            =   "PinyinDictionary1.frx":C434
         TabIndex        =   9
         Text            =   " °¢"
         Top             =   7830
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command58 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74970
         Picture         =   "PinyinDictionary1.frx":C440
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Save the current page"
         Top             =   1830
         Width           =   315
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5070
         Left            =   -74970
         TabIndex        =   107
         Top             =   1260
         Visible         =   0   'False
         Width           =   7410
         Begin VB.CommandButton Command64 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Clears window"
            Top             =   4485
            Width           =   645
         End
         Begin VB.CommandButton Command66 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6330
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   1890
            Width           =   645
         End
         Begin VB.CommandButton Command65 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   1260
            Width           =   645
         End
         Begin VB.CommandButton Command63 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3930
            Width           =   645
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Clear All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3090
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Clears all windows"
            Top             =   4200
            Width           =   1005
         End
         Begin VB.TextBox Text18 
            BackColor       =   &H00D8E9EC&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   3150
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   119
            Top             =   2250
            Width           =   3795
         End
         Begin VB.CommandButton Command56 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Picture         =   "PinyinDictionary1.frx":C972
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Save the current page"
            Top             =   0
            Width           =   315
         End
         Begin VB.TextBox Text17 
            BackColor       =   &H00D8E9EC&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   3150
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   113
            Top             =   300
            Width           =   3795
         End
         Begin VB.TextBox Text16 
            BackColor       =   &H00D8E9EC&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   110
            Top             =   1740
            Width           =   2745
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00D8E9EC&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   109
            Top             =   300
            Width           =   2745
         End
         Begin VB.Shape Shape3 
            Height          =   195
            Left            =   0
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label19 
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
            Left            =   5
            TabIndex        =   155
            Top             =   200
            Width           =   345
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            Caption         =   "Medical Phrases"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3930
            TabIndex        =   120
            Top             =   1980
            Width           =   2175
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Caption         =   "Medical Conditions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3870
            TabIndex        =   114
            Top             =   30
            Width           =   2175
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "Medications"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   112
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Caption         =   "Allergies"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   300
            TabIndex        =   111
            Top             =   30
            Width           =   2175
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   4305
         Left            =   -74610
         TabIndex        =   38
         Top             =   1470
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   7594
         _Version        =   393216
         Tabs            =   8
         TabsPerRow      =   8
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Medical"
         TabPicture(0)   =   "PinyinDictionary1.frx":CEA4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Chinese"
         TabPicture(1)   =   "PinyinDictionary1.frx":CEC0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Travel"
         TabPicture(2)   =   "PinyinDictionary1.frx":CEDC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Shop"
         TabPicture(3)   =   "PinyinDictionary1.frx":CEF8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "Money"
         TabPicture(4)   =   "PinyinDictionary1.frx":CF14
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "Weather"
         TabPicture(5)   =   "PinyinDictionary1.frx":CF30
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         TabCaption(6)   =   "Time #s"
         TabPicture(6)   =   "PinyinDictionary1.frx":CF4C
         Tab(6).ControlEnabled=   0   'False
         Tab(6).ControlCount=   0
         TabCaption(7)   =   "Eat"
         TabPicture(7)   =   "PinyinDictionary1.frx":CF68
         Tab(7).ControlEnabled=   0   'False
         Tab(7).ControlCount=   0
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   330
            TabIndex        =   93
            Top             =   3780
            Width           =   5775
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "ER : Heart"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   99
               ToolTipText     =   "Medical Emergency and Cardiac"
               Top             =   120
               Value           =   -1  'True
               Width           =   1080
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Head"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   1
               Left            =   1170
               Style           =   1  'Graphical
               TabIndex        =   98
               ToolTipText     =   "Medical Neuro, Eyese, Ears, Nose and Throat"
               Top             =   120
               Width           =   780
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Lungs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   2
               Left            =   1980
               Style           =   1  'Graphical
               TabIndex        =   97
               ToolTipText     =   "Medical Breathing and Chest"
               Top             =   120
               Width           =   750
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Stomach"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   3
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   96
               ToolTipText     =   "Medical Gastro and Abdominal"
               Top             =   120
               Width           =   990
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "Kidney"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   4
               Left            =   3795
               Style           =   1  'Graphical
               TabIndex        =   95
               ToolTipText     =   "Medical Kidney and Bladder"
               Top             =   120
               Width           =   915
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00D8E9EC&
               Caption         =   "General"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   5
               Left            =   4770
               Style           =   1  'Graphical
               TabIndex        =   94
               ToolTipText     =   "Medical General and Miscellaneous"
               Top             =   120
               Width           =   855
            End
         End
      End
      Begin TabDlg.SSTab SSTab7 
         Height          =   4305
         Left            =   -74595
         TabIndex        =   39
         Top             =   1785
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   7594
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   6
         TabsPerRow      =   8
         TabHeight       =   520
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Useful"
         TabPicture(0)   =   "PinyinDictionary1.frx":CF84
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Love"
         TabPicture(1)   =   "PinyinDictionary1.frx":CFA0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Bars"
         TabPicture(2)   =   "PinyinDictionary1.frx":CFBC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Insults"
         TabPicture(3)   =   "PinyinDictionary1.frx":CFD8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "Topics"
         TabPicture(4)   =   "PinyinDictionary1.frx":CFF4
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame6"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Speak"
         TabPicture(5)   =   "PinyinDictionary1.frx":D010
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "mapic(3)"
         Tab(5).Control(1)=   "mapic(1)"
         Tab(5).Control(2)=   "mapic(2)"
         Tab(5).Control(3)=   "mapic(4)"
         Tab(5).Control(4)=   "mapic(5)"
         Tab(5).Control(5)=   "mapic(0)"
         Tab(5).Control(6)=   "cpvSlider3"
         Tab(5).Control(7)=   "Option4"
         Tab(5).Control(8)=   "Option6"
         Tab(5).Control(9)=   "Text13"
         Tab(5).Control(10)=   "Command43"
         Tab(5).Control(11)=   "Mamma"
         Tab(5).Control(12)=   "Text12"
         Tab(5).Control(13)=   "Command38"
         Tab(5).Control(14)=   "Command45"
         Tab(5).Control(15)=   "Command44"
         Tab(5).Control(16)=   "Picture2"
         Tab(5).ControlCount=   17
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3645
            Left            =   -74895
            TabIndex        =   139
            Top             =   90
            Width           =   6315
            Begin VB.CommandButton cmdOK 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   3300
               Picture         =   "PinyinDictionary1.frx":D02C
               Style           =   1  'Graphical
               TabIndex        =   144
               ToolTipText     =   "Next"
               Top             =   2790
               Width           =   705
            End
            Begin VB.ComboBox Combo12 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   180
               TabIndex        =   143
               Text            =   "Conversational Topics"
               Top             =   120
               Width           =   6105
            End
            Begin VB.CommandButton Command25 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   150
               Picture         =   "PinyinDictionary1.frx":D8F6
               Style           =   1  'Graphical
               TabIndex        =   142
               ToolTipText     =   "Pronounce English"
               Top             =   2790
               Width           =   705
            End
            Begin VB.CommandButton Command24 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   5550
               Picture         =   "PinyinDictionary1.frx":E1C0
               Style           =   1  'Graphical
               TabIndex        =   141
               ToolTipText     =   "Pronounce Chinese"
               Top             =   2790
               Width           =   705
            End
            Begin VB.CommandButton Command12 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   2580
               Picture         =   "PinyinDictionary1.frx":EA8A
               Style           =   1  'Graphical
               TabIndex        =   140
               ToolTipText     =   "Back"
               Top             =   2790
               Width           =   705
            End
            Begin Project1.cpvSlider cpvSlider2 
               Height          =   240
               Left            =   3660
               Top             =   3330
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   423
               SliderIcon      =   "PinyinDictionary1.frx":F354
               Orientation     =   0
               RailPicture     =   "PinyinDictionary1.frx":F46E
               Max             =   100
               Value           =   50
            End
            Begin Project1.cpvSlider cpvSlider1 
               Height          =   240
               Left            =   540
               Top             =   3330
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   423
               SliderIcon      =   "PinyinDictionary1.frx":F936
               Orientation     =   0
               RailPicture     =   "PinyinDictionary1.frx":FA50
               Max             =   100
               Value           =   50
            End
            Begin VB.Label lblReply 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   675
               Left            =   180
               TabIndex        =   149
               Top             =   570
               Width           =   6105
            End
            Begin VB.Label Label1 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   675
               Left            =   180
               TabIndex        =   148
               Top             =   1290
               Width           =   6105
            End
            Begin VB.Label Label2 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   675
               Left            =   180
               TabIndex        =   147
               Top             =   2070
               Width           =   6105
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Volume"
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
               Left            =   4350
               TabIndex        =   146
               Top             =   3060
               Width           =   735
            End
            Begin VB.Label Label22 
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
               Left            =   1380
               TabIndex        =   145
               Top             =   3060
               Width           =   735
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   -72090
            ScaleHeight     =   540
            ScaleWidth      =   1950
            TabIndex        =   169
            Top             =   1830
            Width           =   1950
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "ma"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   525
               Left            =   840
               TabIndex        =   172
               Top             =   -30
               Width           =   795
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "Âè"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   525
               Left            =   120
               TabIndex        =   171
               Top             =   0
               Width           =   555
            End
            Begin VB.Label Label41 
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Left            =   1680
               TabIndex        =   170
               Top             =   0
               Width           =   240
            End
         End
         Begin VB.CommandButton Command44 
            BackColor       =   &H80000008&
            Caption         =   "   "
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   -72210
            Style           =   1  'Graphical
            TabIndex        =   168
            Top             =   1110
            Width           =   2130
         End
         Begin VB.CommandButton Command45 
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   -73515
            Picture         =   "PinyinDictionary1.frx":FF18
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   180
            Width           =   480
         End
         Begin VB.CommandButton Command38 
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   -70395
            Picture         =   "PinyinDictionary1.frx":107E2
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   180
            Width           =   480
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   165
            Text            =   "Âè   Âé  Âí   Âî   Âï"
            Top             =   3360
            Width           =   2655
         End
         Begin VB.TextBox Mamma 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74970
            Locked          =   -1  'True
            TabIndex        =   164
            Text            =   "ÂèÂéÂíÂîÂï"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.CommandButton Command43 
            Caption         =   "Start Speech Recognition"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   -73035
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   195
            Width           =   2640
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   -72975
            Locked          =   -1  'True
            TabIndex        =   162
            Text            =   "ma1    ma2   ma3  ma4  ma5"
            Top             =   3165
            Width           =   2670
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Chinese"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   -71580
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   675
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "English"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   -72990
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   675
            Width           =   1335
         End
         Begin Project1.cpvSlider cpvSlider3 
            Height          =   1305
            Left            =   -70020
            Top             =   1230
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            SliderIcon      =   "PinyinDictionary1.frx":110AC
            RailPicture     =   "PinyinDictionary1.frx":111A6
            Max             =   3
         End
         Begin VB.Image mapic 
            Height          =   1335
            Index           =   0
            Left            =   -73275
            Picture         =   "PinyinDictionary1.frx":1166E
            Top             =   1155
            Width           =   975
         End
         Begin VB.Image mapic 
            Height          =   480
            Index           =   5
            Left            =   -70770
            Picture         =   "PinyinDictionary1.frx":124D1
            Top             =   2550
            Width           =   480
         End
         Begin VB.Image mapic 
            Height          =   480
            Index           =   4
            Left            =   -71250
            Picture         =   "PinyinDictionary1.frx":12D9B
            Top             =   2580
            Width           =   480
         End
         Begin VB.Image mapic 
            Height          =   480
            Index           =   2
            Left            =   -72300
            Picture         =   "PinyinDictionary1.frx":13665
            Top             =   2580
            Width           =   480
         End
         Begin VB.Image mapic 
            Height          =   480
            Index           =   1
            Left            =   -72840
            Picture         =   "PinyinDictionary1.frx":13F2F
            Top             =   2580
            Width           =   480
         End
         Begin VB.Image mapic 
            Height          =   480
            Index           =   3
            Left            =   -71760
            Picture         =   "PinyinDictionary1.frx":147F9
            Top             =   2580
            Width           =   480
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1800
         Picture         =   "PinyinDictionary1.frx":150C3
         Top             =   5610
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search Dictionary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -72390
         TabIndex        =   61
         Top             =   5520
         Width           =   2205
      End
      Begin VB.Label Label40 
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
         Left            =   -68555
         TabIndex        =   158
         ToolTipText     =   "Set Global Save Path"
         Top             =   1200
         Width           =   345
      End
      Begin VB.Shape Shape6 
         Height          =   195
         Left            =   -68540
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label33 
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
         Left            =   7050
         TabIndex        =   156
         ToolTipText     =   "Set Global Save Path"
         Top             =   5730
         Width           =   315
      End
      Begin VB.Shape Shape4 
         Height          =   195
         Left            =   7065
         Top             =   5835
         Width           =   315
      End
      Begin VB.Shape Shape2 
         Height          =   195
         Left            =   -74955
         Top             =   2190
         Width           =   330
      End
      Begin VB.Label Label17 
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
         Left            =   -74970
         TabIndex        =   154
         ToolTipText     =   "Set Global Save Path"
         Top             =   2070
         Width           =   345
      End
      Begin VB.Shape Shape1 
         Height          =   195
         Left            =   7065
         Top             =   1710
         Width           =   300
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
         Left            =   7050
         TabIndex        =   153
         ToolTipText     =   "Set Global Save Path"
         Top             =   1590
         Width           =   345
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7080
         TabIndex        =   104
         ToolTipText     =   "Increase/Decrease font size"
         Top             =   4575
         Width           =   315
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7065
         TabIndex        =   103
         ToolTipText     =   "Increase/Decrease font size"
         Top             =   4410
         Width           =   315
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re-load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -74490
         TabIndex        =   65
         Top             =   5910
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Dictionary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -73290
         TabIndex        =   63
         Top             =   1290
         Width           =   2835
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "                     Caution these Insults may be offensive                       "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3465
      Left            =   495
      TabIndex        =   179
      Top             =   2100
      Visible         =   0   'False
      Width           =   6915
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "View Insults"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   1485
         Width           =   3315
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   " Caution these Insults may be offensive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   585
         TabIndex        =   182
         Top             =   315
         Width           =   5490
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   945
      Pattern         =   "*.url"
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3780
      Left            =   2025
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "PinyinDictionary1.frx":153CD
      Top             =   1440
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "http://www.moosenose.com/speakingpinyin.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1710
      TabIndex        =   85
      Top             =   30
      Width           =   5370
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I always knew that the ""Stroke Order"" was the key to LOVE..."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1755
      TabIndex        =   84
      Top             =   3225
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "http://www.moosenose.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1350
      TabIndex        =   83
      Top             =   3450
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "Online Help:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnSaveCEd 
         Caption         =   "Save CE Dictionary"
      End
      Begin VB.Menu mnuSaveCEp 
         Caption         =   "Save CE Phrases"
      End
      Begin VB.Menu mnuSaveM 
         Caption         =   "Save CE Medical"
      End
      Begin VB.Menu ddsda 
         Caption         =   "-"
      End
      Begin VB.Menu mnuw2m 
         Caption         =   "Convert wav to mp3"
      End
      Begin VB.Menu mnuuu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTP 
         Caption         =   "Desktop Path"
      End
      Begin VB.Menu mnuSTP 
         Caption         =   "Save to Path"
      End
      Begin VB.Menu mnnn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRead 
         Caption         =   "&Reading Window"
      End
      Begin VB.Menu mnn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTranslate 
      Caption         =   "&Translate Web Page"
      Begin VB.Menu mnuce 
         Caption         =   "&Simplified Chinese to English"
      End
      Begin VB.Menu mnuscp 
         Caption         =   "Simplified Chinese to &Pinyin"
      End
      Begin VB.Menu mnuEc 
         Caption         =   "&English to Chinese"
      End
   End
   Begin VB.Menu mnuClip 
      Caption         =   "&Clipboard"
      Begin VB.Menu mnyClipEng 
         Caption         =   "English"
         Begin VB.Menu mnuClipTransE 
            Caption         =   "Translate"
         End
         Begin VB.Menu mnuClipPronE 
            Caption         =   "Pronounce"
         End
         Begin VB.Menu mnuClipBothE 
            Caption         =   "Both"
         End
      End
      Begin VB.Menu mnuClipPin 
         Caption         =   "Pinyin"
         Begin VB.Menu mnuClipTransP 
            Caption         =   "Translate"
         End
         Begin VB.Menu mnuClipBothP 
            Caption         =   "Both"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuClipSimp 
         Caption         =   "Simplified"
         Begin VB.Menu mnuClipTransS 
            Caption         =   "Translate"
         End
         Begin VB.Menu mnuClipPronS 
            Caption         =   "Pronounce"
         End
         Begin VB.Menu mnuClipBothS 
            Caption         =   "Both"
         End
      End
   End
   Begin VB.Menu mnuPronno 
      Caption         =   "&Pronounce"
      Begin VB.Menu mnuSSPEAK 
         Caption         =   "Speak"
      End
      Begin VB.Menu mnuSStop 
         Caption         =   "Stop Speaking"
      End
      Begin VB.Menu mnuShowt 
         Caption         =   "Show TTS"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&Home Page"
      End
      Begin VB.Menu mnuEnabling 
         Caption         =   "&Enabling International Support"
      End
      Begin VB.Menu mnuIFEAL 
         Caption         =   "Install files for East Asian Languages"
      End
      Begin VB.Menu mnuPRC 
         Caption         =   "Select Chinese (PRC)"
      End
      Begin VB.Menu mnuSAPI 
         Caption         =   "Install SAPI5.1"
      End
      Begin VB.Menu mnuLangPk 
         Caption         =   "Install Language Pack"
      End
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuPinyin 
         Caption         =   "&Pinyin"
      End
   End
End
Attribute VB_Name = "FormX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'*** Find files in selected directory to populate listbox

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Dim Counter, back As Integer
Dim GlobalString As String
Dim UUaarreell As String
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim Lettr As String

' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, x As Single, y As Single) As String
On Error Resume Next
    Dim pt        As POINTAPI
    Dim pos       As Long
    Dim start_pos As Long
    Dim end_pos   As Long
    Dim ch        As String
    Dim txt       As String
    Dim txtlen    As Long
    Dim Testicle As String
    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function
    If Option1.Value = True Then
        Lettr = Mid$(rch.Text, pos - 2, 1)
        If Asc(Lettr) >= 0 Then Lettr = ""
    Else
        Lettr = Mid$(rch.Text, pos - 5, 1)
        If Asc(Lettr) < 0 Then Lettr = ""
    End If
    
    'Lettr = Mid$(rch.Text, pos - 5, 1)
    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If ch = " " Or ch = Chr(13) Then Exit For
    Next start_pos
    
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If ch = " " Or ch = Chr(13) Then Exit For
    Next end_pos
    
    end_pos = end_pos - 1

    If start_pos <= end_pos Then RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function


Private Sub Check1_Click()
On Error Resume Next

If Check1.Value = 1 Then
    SetTopMostWindow Me.hWnd, True
    Close #1: Open App.Path & "\Settings.ini" For Output As #1
        Print #1, "True"
    Close #1
Else
    Close #1: Open App.Path & "\Settings.ini" For Output As #1
        Print #1, "False"
    Close #1
    SetTopMostWindow Me.hWnd, False
End If
End Sub

Private Sub Check12_Click()
On Error Resume Next

If Check12.Value = 1 Then
    WebBrowser1.Navigate App.Path & "\Examplet.htm"
    SSTab1.Tab = 2
End If
End Sub

Private Sub Check13_Click()
End Sub


Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
    TTSApp.Check2.Value = 1
Else
    TTSApp.Check2.Value = 0
End If


End Sub

Private Sub Check3_Click()
On Error Resume Next

If Check3.Value = 0 Then
    WebBrowser3.Visible = False
    WebBrowser3.Navigate "about:<html><body scroll='no'><img src='" & App.Path & "\an1.gif" & "'></img></body></html>"
Else
    WebBrowser3.Visible = True
    WebBrowser3.Navigate "http://zhongwen.com/chat.htm"
End If
End Sub



Private Sub Check4_Click()
On Error Resume Next

If Check4.Value = 1 Then
    Frame7.Visible = True
    Frame7.ZOrder 0
Else
    Frame7.Visible = False
End If
Check4.ZOrder 0
End Sub

Private Sub Check5_Click()
On Error Resume Next

If Check5.Value = 1 Then
    WebBrowser5.Navigate App.Path & "\LabWorkChinese.htm"
Else
    If Combo1.Text = "Drug Classes" Then
        WebBrowser5.Navigate App.Path & "\Medications.htm"
    Else
        Combo1_Click
    End If
End If
End Sub

Private Sub cmdOK_Click()
Dim Inny, Inny1, Outy, Loudy As String
Dim Moose, Goose, Loose, ttt As String
Dim i, j, k, l As Long
On Error GoTo here
If Combo12.Text <> "Conversational Topics" Then
    i1 = i1 + 1
    Moose = Conversation(i1)
    i = InStr(Moose, "*")
    j = InStr(i + 1, Moose, "*")
    k = InStr(j + 1, Moose, "*")
    Loose = ""
    lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "")
    Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "")
    If Asc(Right(Moose, 1)) < 0 Then
        Loose = Trim(Mid(Moose, k + 1, Len(Moose) - k - 1))
    End If
    ttt = Trim(Right(Trim(Moose), Len(Trim(Moose)) - k))
    lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "") & Loose & " " & ttt
    Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "") & Loose & " " & ttt
    Label2.Caption = Replace(Trim(Mid(Moose, j + 1, k - 1)), "*", "")
    Command24_Click
End If

here:
'i1 = 0
End Sub

Private Sub Combo1_Click()
On Error Resume Next

    If Combo1.Text = "Drug Classes" Then
        WebBrowser5.Navigate App.Path & "\Medications.htm"
    Else
        WebBrowser5.Navigate App.Path & "\Medications\" & Combo1.List(Combo1.ListIndex) & ".htm"
        UUaarreell = App.Path & "\Medications\" & Combo1.List(Combo1.ListIndex) & ".htm"
        'WebBrowser4.Navigate UUaarreell
    End If
End Sub

Private Sub Combo10_Click()
On Error Resume Next

If Combo10.ListCount = 0 Then Exit Sub
Combo11.ListIndex = Combo10.ListIndex
Uarel = Trim(Combo11.Text)
Flagday = True
Unload frmMain
Load frmMain
frmMain.Show
End Sub

Private Sub Combo11_Click()
On Error Resume Next

If Combo10.ListCount = 0 Then Exit Sub
Combo10.ListIndex = Combo11.ListIndex
Uarel = Trim(Combo11.Text)
Flagday = True
Unload frmMain
Load frmMain
frmMain.Show
End Sub

Private Sub Combo12_Click()
On Error Resume Next
Dim Inny, Inny1, Outy, Loudy As String
Dim Moose, Goose, Loose, ttt As String
Dim i, j, k, l As Long

If Combo12.Text = "Conversational Topics" Then
    MsgBox "Please select a Conversational Topic."
    Exit Sub
End If
    i = 0
    i1 = 0
    Close #1: Open App.Path & "\Conversation.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Inny
        If InStr(Inny, Combo12.Text) Then
            Do
                Line Input #1, Inny
                If Trim(Inny) <> "" Then
                    Conversation(i) = Inny
                    i = i + 1
                End If
            Loop While Right(Inny, 1) <> ":"
            Moose = Conversation(0)
            'MsgBox Moose
            i = InStr(Moose, "*")
            j = InStr(i + 1, Moose, "*")
            k = InStr(j + 1, Moose, "*")
            Loose = ""
            lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "")
            Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "")
            If Asc(Right(Moose, 1)) < 0 Then
                Loose = Trim(Mid(Moose, k + 1, Len(Moose) - k - 1))
            End If
            ttt = Trim(Right(Trim(Moose), Len(Trim(Moose)) - k))
            'MsgBox TTT
            lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "") & Loose & " " & ttt
            Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "") & Loose & " " & ttt
            Label2.Caption = Replace(Trim(Mid(Moose, j + 1, k - 1)), "*", "")
            Command24_Click
            Exit Sub
        End If
    Loop
    Close #1
    Exit Sub

End Sub

Private Sub Combo14_Change()
'Text14.Text = Combo14.Text
End Sub

Private Sub Combo14_Click()
On Error Resume Next

'Text14.Text = Combo14.Text
'Command46_Click
SSTab1.Tab = 0
End Sub

Private Sub Combo6_Click()
On Error Resume Next

Dim i, a, B, c, d, e As Long
Dim Moose As String
Dim Inny, Outy, Outy1 As String
SSTab1.Tab = 0
Combo8.ListIndex = Combo6.ListIndex
Text5.Text = ""
Text5.Text = Combo8.List(Combo8.ListIndex)
If Trim(Text2.Text) = "Numbered Pinyin" Then Text2.Text = ""
If Trim(Text14.Text) = "Simplified Chinese" Then Text14.Text = ""
Combo14.Clear
For i = 1 To Len(Trim(Text5.Text))
    Combo14.AddItem Mid(Trim(Text5.Text), i, 1)
Next i
Combo14.Text = Combo14.List(0)

Pn1(PinIncr - 1).Text = Combo6.Text
Simp(PinIncr - 1).Clear
For i = 1 To Len(Text5.Text)
    Simp(PinIncr - 1).AddItem " " & Mid(Text5.Text, i, 1)
Next
Simp(PinIncr - 1).Text = " " & Left(Text5.Text, 1)
Select Case Right(Combo6.Text, 1)
    Case 1
        Image2.Picture = LoadPicture(App.Path & "\1.ico")
    Case 2
        Image2.Picture = LoadPicture(App.Path & "\2.ico")
    Case 3
        Image2.Picture = LoadPicture(App.Path & "\3.ico")
    Case 4
        Image2.Picture = LoadPicture(App.Path & "\4.ico")
    Case 5
        Image2.Picture = LoadPicture(App.Path & "\5.ico")
End Select
'Text2.Text = Trim(Combo6.Text)
'Command8_Click
End Sub

Private Sub Command1_Click()
On Error Resume Next

    OpenBrowser WebBrowser1.LocationURL, FormX.hWnd
End Sub

Private Sub Command10_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub



Private Sub Command11_Click()
On Error Resume Next

If Trim(UUaarreell) = "" Then Exit Sub
Dim Inny, Outy, Moose, Goose As String
Dim i, j, k As Long
Screen.MousePointer = 11
DoEvents
Moose = ""
k = 0
Text = Trim(Text)
MsgBox Text
    Moose = ""
    For i = 1 To Len(Texty)
        Outy = Mid(Texty, i, 1)
        'MsgBox Outy
        If Asc(Outy) < 0 Then
            Close #1: Open App.Path & "\PYnal.txt" For Input As #1
            Do While Not EOF(1)
                Line Input #1, Inny
                Inny = Trim(Inny)
                j = InStr(Inny, " ")
                If InStr(Inny, Outy) <> 0 Then
                    'MsgBox Outy & " " & Asc(Outy)
                    Moose = Moose & Trim(Left(Inny, j)) & " "
                    Exit Do
                End If
            Loop
            Close #1
        Else
            Moose = Moose & Outy
        End If
    Next
    Open App.Path & "\pinyin.htm" For Output As #1
        Print #1, Moose
    Close #1
    WebBrowser5.Navigate App.Path & "\pinyin.htm"
    'Text7.Visible = True
    'Text7.ZOrder 0
    Screen.MousePointer = 0

End Sub

Private Sub Command12_Click()
Dim Inny, Inny1, Outy, Loudy As String
Dim Moose, Goose, Loose, ttt As String
Dim i, j, k, l As Long
On Error GoTo here
If Combo12.Text <> "Conversational Topics" Then
    i1 = i1 - 1
    If i1 < 0 Then i1 = 0
    Moose = Conversation(i1)
    'MsgBox Moose
    i = InStr(Moose, "*")
    j = InStr(i + 1, Moose, "*")
    k = InStr(j + 1, Moose, "*")
    Loose = ""
    lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "")
    Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "")
    If Asc(Right(Moose, 1)) < 0 Then
        Loose = Trim(Mid(Moose, k + 1, Len(Moose) - k - 1))
    End If
    ttt = Trim(Right(Trim(Moose), Len(Trim(Moose)) - k))
    'MsgBox TTT
    lblReply.Caption = Replace(Trim(Left(Moose, i - 1)), "*", "") & Loose & " " & ttt
    Label1.Caption = Replace(Trim(Mid(Moose, i + 1, j - i)), "*", "") & Loose & " " & ttt
    Label2.Caption = Replace(Trim(Mid(Moose, j + 1, k - 1)), "*", "")
    Command24_Click
End If

here:
'i1 = 0
End Sub



Private Sub Command13_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim a, B, c, d As String
Dim i, j, k, l As Integer
a = Trim(List5.List(List5.ListIndex))
i = InStr(a, "*--")
j = InStr(i + 1, a, "*")
k = Len(a)
Ep = Left(a, i - 1)
SP = Mid(a, i + 3, j - i - 3)
PP = Right(a, k - j)


If Option11.Value = True Then
    GlobalString = Trim(SP)
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
Else
    GlobalString = Trim(Ep)
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
End If
Moose = GlobalString
If Trim(Moose) <> "" Then
    TTSAppMain.MainTxtBox.Text = Moose
Else
    TTSAppMain.MainTxtBox.Text = ""
    Exit Sub
End If

If Check9.Value = 1 Then
    TTSAppMain.Show
    'Call MakeTranslucent(TTSAppMain, tColor)
End If
    
TTSAppMain.SpeakBtn_Click
If Check2.Value = 1 Then
    Load Simple
    WavText = Trim(Moose)
    Simple.TextField = WavText
    Simple.SpeakItBtn_Click
End If

End Sub

Private Sub Command14_Click()
On Error Resume Next

If Trim(Uarel) = "" Then Exit Sub
    'Text7.Visible = True
    WebBrowser1.Navigate Uarel
End Sub


Private Sub Command15_Click()
Dim i, j As Integer
On Error Resume Next
List3.Clear
j = List5.ListCount
For i = 0 To j - 1
    If List5.Selected(i) = True Then List3.AddItem List4.List(i)
Next
For i = 0 To 5
    Option8(i).Value = False
Next
List5.Clear
For i = 0 To List3.ListCount - 1
    List5.AddItem List3.List(i)
    List5.Selected(i) = True
Next

End Sub

Private Sub Command16_Click()
On Error Resume Next

If PinIncr = 7 Then Exit Sub
Load Pn1(PinIncr)
Load Simp(PinIncr)
Pn1(PinIncr).Visible = True
Pn1(PinIncr).ZOrder 0
Pn1(PinIncr).Top = Pn1(PinIncr - 1).Top
Pn1(PinIncr).Left = Pn1(PinIncr - 1).Left + 1440
Simp(PinIncr).Visible = True
Simp(PinIncr).ZOrder 0
Simp(PinIncr).Top = Simp(PinIncr - 1).Top
Simp(PinIncr).Left = Simp(PinIncr - 1).Left + 1440
PinIncr = PinIncr + 1
End Sub

Private Sub Command17_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To PinIncr
    Unload Pn1(i)
    Unload Simp(i)
Next
PinIncr = 1

End Sub

Private Sub Command23_Click()
On Error Resume Next

mnuCe_Click
End Sub

Private Sub Command19_Click()
On Error Resume Next
Dim i As Long
Screen.MousePointer = 11
For i = 0 To List5.ListCount - 1
    If LCase(Left(List5.List(i), Len(Trim(Text3.Text)))) = LCase(Text3.Text) Then
        List5.ListIndex = i
        Exit For
    End If
Next
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Dictionary.SelLength <> 0 And Trim(Dictionary.SelText) <> "" Then
    Goose = ""
    j = 1
    Moose = Trim(Dictionary.SelText)
        GlobalString = Moose
        If Trim(Moose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Moose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    If Check9.Value = 1 Then
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
    End If
    
    If Asc(Left(Trim(Dictionary.SelText), 1)) < 0 Then ' Chinese < 0
        TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    End If
    TTSAppMain.SpeakBtn_Click
    If Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If
End If
End Sub

Private Sub Command20_Click()
On Error Resume Next

List5.Clear
Open App.Path & "\MedicalDictionary.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List5.AddItem xx
    End If
Loop
Close #1
List5Scroll
End Sub

Public Sub Command21_Click()
On Error Resume Next

Command51_Click
Command9_Click
End Sub

Private Sub Command22_Click()
On Error Resume Next

    mnuw2m_Click
End Sub

Private Sub Command24_Click()
    On Error Resume Next
    Dim Moose, Goose As String
    Dim i, j As Long
    
    'SetTopMostWindow Me.hWnd, True
    'FormX.Enabled = False
    'TTSAppMain.Show
    Goose = ""
    j = 1
    Moose = Trim(Label2.Caption)
    If Len(Moose) > 24 Then
        For i = 1 To Len(Moose)
            If j = 22 Then
                j = 1
                Goose = Goose & Mid(Moose, i, 1) & vbCrLf
            Else
                Goose = Goose & Mid(Moose, i, 1)
                j = j + 1
            End If
        Next
        If Trim(Goose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Goose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    Else
        If Trim(Moose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Moose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    End If
    
    'TTSAppMain.MainTxtBox.Text = Moose
    'TTSAppMain.Show
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    'TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    'tColor = &H8000000F
    ''Call MakeTranslucent(TTSAppMain, tColor)

End Sub

Private Sub Command25_Click()
    On Error Resume Next
    Dim Moose, Goose As String
    Dim i, j As Long
    
    'SetTopMostWindow Me.hWnd, True
    'FormX.Enabled = False
    'TTSAppMain.Show
    Goose = ""
    j = 1
    Moose = Trim(lblReply.Caption)
    'MsgBox Len(Moose)
    If Len(Moose) > 24 Then
        For i = 1 To Len(Moose)
            If j = 22 Then
                j = 1
                Goose = Goose & Mid(Moose, i, 1) & vbCrLf
            Else
                Goose = Goose & Mid(Moose, i, 1)
                j = j + 1
            End If
        Next
        If Trim(Goose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Goose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    Else
        If Trim(Moose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Moose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    End If
    'MsgBox Goose
    'TTSAppMain.MainTxtBox.Text = Moose
    'TTSAppMain.Show
    'TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    
    ''Call MakeTranslucent(TTSAppMain, tColor)


End Sub

Private Sub Command26_Click()
Me.WindowState = 1
Shell App.Path & "\Videolibrary.exe " & App.Path & "\MOV03303.MPG"
End Sub

Private Sub Command27_Click()
On Error Resume Next

WebBrowser1.GoBack

End Sub

Private Sub Command28_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Command29_Click()
On Error Resume Next

WebBrowser1.Refresh

End Sub


Private Sub Command3_Click()
On Error Resume Next

Text6.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
End Sub

Private Sub Command30_Click()
On Error Resume Next


WebBrowser1.Navigate App.Path & "\Help\SpeakingPinYin.htm"
End Sub

Private Sub Command31_Click()
On Error Resume Next

WebBrowser1.Stop

End Sub

Private Sub Command32_Click()
On Error Resume Next

Dim shellHelper As New ShellUIHelper
shellHelper.AddFavorite WebBrowser1.LocationURL, WebBrowser1.LocationName

End Sub

Private Sub Command34_Click()
On Error Resume Next

mnuEc_Click
End Sub

Private Sub Command35_Click()
On Error Resume Next

If Trim(Uarel) = "" Then Exit Sub
Dim Inny, Outy, Moose, Goose As String
Dim i, j, k As Long
Screen.MousePointer = 11
DoEvents
Moose = ""
k = 0
Text = Trim(Text)
'MsgBox Text
'GoTo Thisthere
    For i = 1 To Len(Text)
        Outy = Mid(Text, i, 1)
        'MsgBox Outy
        If Asc(Outy) < 0 Then
            Close #1: Open App.Path & "\PYnal.txt" For Input As #1
            Do While Not EOF(1)
                Line Input #1, Inny
                Inny = Trim(Inny)
                j = InStr(Inny, " ")
                If InStr(Inny, Outy) <> 0 Then
                    'MsgBox Outy & " " & Asc(Outy)
                    Moose = Moose & Trim(Left(Inny, j)) & " "
                    Exit Do
                End If
            Loop
            Close #1
        Else
            Moose = Moose & Outy
        End If
    Next
    Text7.Text = ""
    Text7.Text = Moose
    WebBrowser1.Navigate Uarel
    Screen.MousePointer = 0
    Exit Sub
'Thisthere:
    Moose = ""
    For i = 1 To Len(Texty)
        Outy = Mid(Texty, i, 1)
        'MsgBox Outy
        If Asc(Outy) < 0 Then
            Close #1: Open App.Path & "\PYnal.txt" For Input As #1
            Do While Not EOF(1)
                Line Input #1, Inny
                Inny = Trim(Inny)
                j = InStr(Inny, " ")
                If InStr(Inny, Outy) <> 0 Then
                    'MsgBox Outy & " " & Asc(Outy)
                    Moose = Moose & Trim(Left(Inny, j)) & " "
                    Exit Do
                End If
            Loop
            Close #1
        Else
            Moose = Moose & Outy
        End If
    Next
    Open App.Path & "\pinyin.htm" For Output As #1
        Print #1, Moose
    Close #1
    WebBrowser1.Navigate App.Path & "\pinyin.htm"
    'Text7.Visible = True
    'Text7.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub Command36_Click()
On Error Resume Next

Dim Moose, Goose As String
Text = WebBrowser1.Document.Body.InnerText
TTSAppMain.MainTxtBox.Text = ""
'MsgBox Text
    If Trim(Text) <> "" Then
        TTSAppMain.MainTxtBox.Text = Text
    Else
        TTSAppMain.MainTxtBox.Text = ""
        Exit Sub
    End If
    If Check9.Value = 1 Then
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
    End If
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    TTSAppMain.SpeakBtn_Click
    If Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If
End Sub

Private Sub Command37_Click()
On Error Resume Next

'If Trim(Uarel) = "" Then Exit Sub
Dim Moose, Goose As String

'If Trim(Uarel) <> "" Then
    'WebBrowser1.Navigate Uarel
'Else
    Text = WebBrowser1.Document.Body.InnerText
'End If
TTSAppMain.MainTxtBox.Text = ""
'MsgBox Text
    If Trim(Text) <> "" Then
        TTSAppMain.MainTxtBox.Text = Text
    Else
        TTSAppMain.MainTxtBox.Text = ""
        Exit Sub
    End If
    If Check9.Value = 1 Then
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
    End If
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    If Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If

End Sub

Private Sub Command38_Click()
On Error Resume Next

Command43_Click
End Sub

Private Sub Command39_Click()
On Error Resume Next

WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
End Sub





Private Sub Command4_Click()
Frame5.Visible = False
End Sub

Private Sub Command40_Click()
Dim Inny, Piin, Piin1, Simple, Tester, Pester As String
Dim a, B, c, d, e, i As Long
On Error Resume Next
DoEvents
Screen.MousePointer = 11
Text9.Text = ""
Text11.Text = ""
Tester = LCase(Trim(Text8.Text))
Pester = LCase(Trim(text10.Text))
If Pester = "" And Tester = "" Then Screen.MousePointer = 0: Exit Sub
Close #1: Open App.Path + "\adso.gb" For Input As #1
Do While Not EOF(1)
Line Input #1, Inny
Inny = LCase(Inny)
    a = InStr(Inny, "*")
    B = InStr(a + 1, Inny, "*")
    c = InStr(B + 1, Inny, "*")
    d = InStr(c + 1, Inny, "*")
    e = InStr(d + 1, Inny, "*")
    Simple = Trim(Replace(Mid(Inny, a + 1, B - a), "*", ""))
    Piin = Trim(Replace(Mid(Inny, c + 1, d - c), "*", ""))
    Piin1 = Piin
    For i = 1 To 5
        Piin1 = Replace(Piin1, i, "")
    Next i
    'If Piin = "" Then
        'If Pester = Piin1 Then
            'Text9.Text = Simple
            'Exit Do
        'End If
    'Else
        If Tester = Piin Then
            Text9.Text = Simple
            Text11.Text = Replace(Mid(Inny, a + 1, e - a), "*", "|")
            Exit Do
        End If
    'End If
Loop
Close #1
Screen.MousePointer = 0




End Sub

Private Sub Command41_Click()
On Error Resume Next

    FileToHTML App.Path & "\Exampler.txt", App.Path & "\Examplet.htm", "Moose", "blue", "yellow"
End Sub
Public Sub FileToHTML(InputFile As String, OutputFile As String, Title As String, bgcolor As String, textcolor As String)
On Error Resume Next

Dim newline$, myline$
    newline$ = Chr$(13) + Chr$(10)
    Open InputFile For Input As #1
    Open OutputFile For Output As #2
    If Title = "" Then Title = "No Document Title"
    If bgcolor = "" Then bgcolor = "white"
    If textcolor = "" Then textcolor = "black"
    Print #2, "<HTML>" + newline$
    Print #2, "<HEAD>" + newline$
    Print #2, "<TITLE>" + Title + "</TITLE>" + newline$
    Print #2, "</HEAD>" + newline$
    Print #2, "<BODY bgcolor=" + bgcolor + " text=" + textcolor + ">" + newline$


    Do Until EOF(1)
        Line Input #1, myline$
        Print #2, myline$ + "<BR>"
    Loop
    Print #2, newline$
    Print #2, "</BODY>" + newline$
    Print #2, "</HTML>"
    Close #1
    Close #2
End Sub

Private Sub Command42_Click()
On Error Resume Next
Dim Inny, Inny1, Outy, Loudy As String
Dim i, j, k, l As Long
Dim Moose, Goose, Loose As String
Dim Ontop As String
DoEvents
Moose = ""
Open App.Path & "\PinyinBiblet" For Input As #1
Do While Not EOF(1)
Line Input #1, Inny
Inny = Replace(Inny, ".", "")
Inny = Replace(Inny, "!", "")
Inny = Replace(Inny, ",", "")
Inny = Replace(Inny, "?", "")
Moose = Moose & Inny
Loop
Close #1
extract_words (Moose)
Open App.Path & "\words.txt" For Input As #4
Open App.Path & "\Pinwords.txt" For Output As #3
Do While Not EOF(4)
    Line Input #4, Inny
    Text8.Text = Trim(Inny)
    Command40_Click
    If Trim(Text9.Text) <> "" Then
        Outy = Replace(Text9.Text, vbCrLf, "")
        Print #3, Outy
    Else
        Inny = Replace(Inny, vbCrLf, "")
        Print #3, Trim(Inny)
    End If
Loop
Close #4
Close #3
End Sub
Public Sub extract_words(InputLine As String)
Dim Temp$, word$
Dim x As Long
On Error Resume Next
DoEvents
    ' ASCII values
    ' space = 32
    Temp$ = ""
    Open App.Path & "\words.txt" For Output As #2


    For x = 1 To Len(InputLine)
        Temp$ = Mid$(InputLine, x, 1)


        If Mid$(InputLine, x, 1) = Chr$(32) Then
            If Right$(word$, 1) = "." _
            Or Right$(word$, 1) = "," _
            Or Right$(word$, 1) = ":" _
            Or Right$(word$, 1) = ";" _
            Or Right$(word$, 1) = "-" Then
            'extract period
            word$ = Left$(word$, Len(word$) - 1)
        End If
        ' make lower case
        word$ = LCase$(word$)
        'MsgBox Asc(Right(Word$, 1))
        If Asc(Right(word$, 1)) < 47 Or Asc(Right(word$, 1)) > 58 Then
            If Trim(word$) <> "" Then
                word$ = word$ & "5"
            End If
        End If
        If Trim(word$) <> "" Then
            Print #2, word$
        End If
        word$ = ""
        GoTo 10
    End If
    word$ = word$ + Temp$
10     Next x
    Close #2
End Sub

Private Sub Command43_Click()
On Error Resume Next

If Command43.Caption = "Start Speech Recognition" Then
     Command43.Caption = "Stop Speech Recognition"
     Load RecogVb
     RecogVb.Show
Else
    Command43.Caption = "Start Speech Recognition"
    Unload RecogVb
    Set RecogVb = Nothing
End If
End Sub

Private Sub Command44_Click()
On Error Resume Next

If Timer1.Enabled = False Then
    Dim i, j As Long
    
    i = 1
    Do While i <= 1000
    For j = 1 To 5
        mamama(i) = Mid(Mamma.Text, j, 1)
        i = i + 1
    Next
    Loop
    Timer1.Enabled = True
Else
    Command44.Picture = mapic(0).Picture
    Timer1.Enabled = False
End If


End Sub

Private Sub Command45_Click()
On Error Resume Next

Command43_Click
End Sub

Private Sub Command46_Click()
On Error Resume Next
If Check6.Value = 1 Then GoTo OneAtTimeCh

Dim Inny, Inny1, Outy, Outyd, Loudy, y, Outy1 As String
Dim i, j, k, l, a, B, c, d, e As Long
Dim Moose, Pinnyin As String
Dim Numbered As Boolean
Screen.MousePointer = 11
SSTab1.Tab = 0
Pinyin = Trim(Text14.Text)
If Trim(Text14.Text) = "" Or Trim(Text14.Text) = "Simplified Chinese" Then
    Screen.MousePointer = 0
    Exit Sub
Else
    If Option7(0).Value = True Or Option7(1).Value = True Then
        Moose = "* " & LCase(Trim(Text14.Text)) ' first word
    End If
    If Option7(2).Value = True Or Option7(3).Value = True Then
        Moose = LCase(Trim(Text14.Text)) ' Any word
    End If
End If
Outy = ""
Dictionary.Text = ""
Close #1: Open App.Path + "\adso.gb" For Input As #1
Do While Not EOF(1)
Line Input #1, Inny
    y = LCase(Trim(Inny))
    a = InStr(y, "*")
    B = InStr(a + 1, y, "*")
    c = InStr(B + 1, y, "*")
    d = InStr(c + 1, y, "*")
    e = InStr(d + 1, y, "*")
    y = Replace(Mid(y, a + 1, d - a), "*", " ")
If InStr(Inny, Moose) <> 0 Then
    Outy = Outy & y & vbCrLf
    If Option7(0).Value = True Or Option7(2).Value = True Then Exit Do
End If
Loop
Close #1
If Trim(Outy) <> "" Then
    Dictionary.Text = Outy & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
Else
    Dictionary.Text = UCase(Moose) & "***NOT*** found in dictionaries" & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
End If
Screen.MousePointer = 0
Exit Sub


'**********************************************
OneAtTimeCh:

Dim ii As Long, Moose1 As String

Screen.MousePointer = 11
SSTab1.Tab = 0
Pinyin = Trim(Text14.Text)
If Trim(Text14.Text) = "" Or Trim(Text14.Text) = "Simplified Chinese" Then
    Screen.MousePointer = 0
    Exit Sub
Else
    If Option7(0).Value = True Or Option7(1).Value = True Then
        Moose = "* " & Trim(Text14.Text) ' first word
    End If
    If Option7(2).Value = True Or Option7(3).Value = True Then
        Moose = Trim(Text14.Text) ' Any word
    End If
End If
Outy = ""
Dictionary.Text = ""
For ii = 3 To Len(Moose)
    Moose1 = "* " & Mid(Moose, ii, 1)
    'MsgBox Moose1
    Close #1: Open App.Path + "\adso.gb" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Inny
        y = LCase(Trim(Inny))
        a = InStr(y, "*")
        B = InStr(a + 1, y, "*")
        c = InStr(B + 1, y, "*")
        d = InStr(c + 1, y, "*")
        e = InStr(d + 1, y, "*")
        y = Replace(Mid(y, a + 1, d - a), "*", " ")
        y = y
    If InStr(Inny, Moose1) <> 0 Then
        Outy = Outy & y & vbCrLf
        If Option7(0).Value = True Or Option7(2).Value = True Then Exit Do
    End If
    
    Loop
    Close #1
Next
    If Trim(Outy) <> "" Then
        Dictionary.Text = Outy & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
    Else
        Dictionary.Text = UCase(Moose) & "***NOT*** found in dictionaries" & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
    End If
Screen.MousePointer = 0


End Sub

Private Sub Command47_Click()
On Error Resume Next

Check1.Value = 0
Load TTSApp
TTSApp.Show
TTSApp.ZOrder 0
Me.ZOrder 1
End Sub

Private Sub Command48_Click()
On Error Resume Next

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub



Private Sub Command49_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Text2.Text = Text2.Text & Combo6.Text       'List(Combo6.ListIndex)
End Sub


Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.GoForward

End Sub

Public Sub Command51_Click()
On Error Resume Next
Dim WhoDo As Boolean
Dim Moose As String
Dim Inny, Inny1, Outy, Outyd, Loudy As String
Dim Def As String, Searchh As String, de1 As String, x As String, y As String, _
    i As Long, j As Long, k As Long, l As Long, m As Long
Screen.MousePointer = 11
If Trim(Text1.Text) = "" Then
    Screen.MousePointer = 0
    Exit Sub
Else
    Moose = LCase(Trim(Text1.Text))
End If
SSTab1.Tab = 0
Inny = ""
Inny1 = ""
Outy = ""
Close #1: Open App.Path + "\dict\" & Left(Moose, 1) & "\" & Left(Moose, 1) & ".txt" For Input As #1
Do While Not EOF(1)
    'DoEvents
    Line Input #1, Inny
    Moose = LCase(Trim(Moose))
    Def = Trim(Inny)
    Inny = LCase(Trim(Inny))
        If Inny = "" Then Screen.MousePointer = 0: Exit Sub
        i = InStr(Def, "(")
        j = InStr(Def, ")")
        k = InStr(i, Def, ".")
        If Left(UCase(Moose), 1) = Left(Def, 1) And i <> 0 And j <> 0 And k <> 0 _
            And k < j And Left(Inny, Len(Moose)) = Moose Then
            Dictionary.Text = Def & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
            For m = 0 To 25
                Line Input #1, x
                Def = Trim(x)
                Dictionary.Text = Def & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
            Next
            Screen.MousePointer = 0: Exit Do: Exit Sub
        End If
        'DoEvents
    Loop
    Close #1
    Screen.MousePointer = 0

End Sub

Private Sub Command52_Click()
On Error Resume Next

Dictionary.Text = ""
End Sub

Private Sub Command53_Click()
Dim i, j As Integer
On Error Resume Next
List3.Clear
j = List4.ListCount
For i = 0 To j - 1
    If List4.Selected(i) = True Then List3.AddItem List4.List(i)
Next
For i = 0 To 5
    Option8(i).Value = False
Next
List4.Clear
For i = 0 To List3.ListCount - 1
    List4.AddItem List3.List(i)
    List4.Selected(i) = True
Next
End Sub


Private Sub Command54_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim a, B, c, d As String
Dim i, j, k, l As Integer
a = Trim(List4.List(List4.ListIndex))
i = InStr(a, "*--")
j = InStr(i + 1, a, "*")
k = Len(a)
Ep = Left(a, i - 1)
SP = Mid(a, i + 3, j - i - 3)
PP = Right(a, k - j)


If Option10.Value = True Then
    GlobalString = Trim(SP)
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
Else
    GlobalString = Trim(Ep)
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
End If
Moose = GlobalString
If Trim(Moose) <> "" Then
    TTSAppMain.MainTxtBox.Text = Moose
Else
    TTSAppMain.MainTxtBox.Text = ""
    Exit Sub
End If

If Check9.Value = 1 Then
    TTSAppMain.Show
    'Call MakeTranslucent(TTSAppMain, tColor)
End If
    
TTSAppMain.SpeakBtn_Click

If Check2.Value = 1 Then
    Load Simple
    WavText = Trim(Moose)
    Simple.TextField = WavText
    Simple.SpeakItBtn_Click
End If

End Sub

Private Sub Command55_Click()
'Text14.Text = Text14.Text & Combo14.Text

End Sub

Private Sub Command55_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

'Text14.Text = Text14.Text & Combo14.List(Combo14.ListIndex)
Text14.Text = Text14.Text & Combo14.Text
End Sub

Private Sub Command56_Click()
On Error Resume Next

Dim TargetPath As String
Dim DesktopPath As String
DesktopPath = GetShellFolderPath(&H0)

If Trim(SavePath) <> "" Then
    TargetPath = Replace(SavePath & "\Medical Data " & format(Now, "ddmmyyhhmmss") & ".txt", "\\", "\")
Else
    TargetPath = DesktopPath & "\Medical Data " & format(Now, "ddmmyyhhmmss") & ".txt"
End If

Open TargetPath For Output As #1
    Print #1, "        Allergies" & vbCrLf & Text6.Text & vbCrLf & vbCrLf
    Print #1, "        Medications" & vbCrLf & Text16.Text & vbCrLf & vbCrLf
    Print #1, "        Medical Conditions" & vbCrLf & Text17.Text & vbCrLf & vbCrLf
    Print #1, "        Medical Phrases" & vbCrLf & Text18.Text & vbCrLf & vbCrLf
Close #1
End Sub

Private Sub Command57_Click()
On Error Resume Next

Dim TargetPath As String
Dim DesktopPath As String
DesktopPath = GetShellFolderPath(&H0)
If Trim(SavePath) <> "" Then
    TargetPath = Replace(SavePath & "\Dictionary Data " & format(Now, "ddmmyyhhmmss") & ".rtf", "\\", "\")
Else
    TargetPath = DesktopPath & "\Dictionary Data " & format(Now, "ddmmyyhhmmss") & ".rtf"
End If
Dictionary.SaveFile TargetPath

Exit Sub

Open TargetPath For Output As #1
    Print #1, Dictionary.Text & vbCrLf & vbCrLf
Close #1

End Sub

Private Sub Command58_Click()
On Error Resume Next

Dim i As Long
Dim TargetPath As String
Dim DesktopPath As String
DesktopPath = GetShellFolderPath(&H0)

If Trim(SavePath) <> "" Then
    TargetPath = Replace(SavePath & "\Phrase Data " & format(Now, "ddmmyyhhmmss") & ".txt", "\\", "\")
Else
    TargetPath = DesktopPath & "\Phrase Data " & format(Now, "ddmmyyhhmmss") & ".txt"
End If

Open TargetPath For Output As #1
For i = 0 To List4.ListCount - 1
    If List4.Selected(i) = True Then
        Print #1, List4.List(i)
    End If
Next
Close #1
End Sub

Private Sub Command59_Click()
On Error Resume Next

Dim Holder As String
If Trim(Clipboard.GetText) = "" Then MsgBox "Copy your medication to the clipboard please": Exit Sub
Holder = Replace(Clipboard.GetText, vbCrLf, "")
Text16 = Text16 & Holder & vbCrLf
End Sub



Private Sub Command6_Click()
On Error Resume Next

        SSTab2.Tab = 0
        WebBrowser1.Navigate App.Path & "\Help\Chinese Links.htm"
End Sub

Private Sub Command60_Click()
On Error Resume Next

Dim Holder As String
If Trim(Clipboard.GetText) = "" Then MsgBox "Copy your medication to the clipboard please": Exit Sub
Holder = Replace(Clipboard.GetText, vbCrLf, "")
Text6 = Text6 & Holder & vbCrLf
End Sub

Private Sub Command61_Click()
On Error Resume Next

Dim i As Long

For i = 0 To List5.ListCount - 1
    If List5.Selected(i) = True Then
        Text17.Text = Text17.Text & List5.List(i) & vbCrLf
    End If
Next
End Sub

Private Sub Command62_Click()
On Error Resume Next

Dim i As Long
For i = 0 To List4.ListCount - 1
    If List4.Selected(i) = True Then
        Text18.Text = Text18.Text & List4.List(i) & vbCrLf
    End If
Next
End Sub

Private Sub Command63_Click()
On Error Resume Next

Text18.Text = ""
End Sub

Private Sub Command64_Click()
On Error Resume Next

Text16.Text = ""
End Sub

Private Sub Command65_Click()
On Error Resume Next

Text6.Text = ""
End Sub

Private Sub Command66_Click()
On Error Resume Next

Text17.Text = ""
End Sub

Private Sub Command67_Click()
    WebBrowser2.Navigate "http://www.multilingualbooks.com/online-radio-chinese.html#about"
End Sub

Private Sub Command68_Click()
On Error Resume Next

    WebBrowser2.Navigate ("http://www.tas-independent-programming.com/xmradio2/index.html")
    FormX.Refresh
End Sub

Private Sub Command69_Click()
On Error Resume Next

    WebBrowser2.Navigate App.Path & "\Help\InternetRadio.htm"
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Trim(Clipboard.GetText) <> "" <> "" Then
    j = 1
    Moose = Trim(Clipboard.GetText)
        GlobalString = Moose
        If Trim(Moose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Moose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    If Check9.Value = 1 Then
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
    End If
    
    If Asc(Left(Trim(Moose), 1)) < 0 Then ' Chinese < 0
        TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    End If
    TTSAppMain.SpeakBtn_Click
    If Check2.Value = 1 Then
        Load Simple
        WavText = Trim(Moose)
        Simple.TextField = WavText
        Simple.SpeakItBtn_Click
    End If
    Clipboard.Clear
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next

Dim Inny, Inny1, Outy, Outyd, Loudy, y, Outy1 As String
Dim i, j, k, l, a, B, c, d, e As Long
Dim Moose, Pinnyin As String
Dim Numbered As Boolean
Screen.MousePointer = 11
If Trim(Text2.Text) = "" Or Trim(Text2.Text) = "Numbered Pinyin" Then
    Screen.MousePointer = 0
    Exit Sub
Else
    If Option7(0).Value = True Or Option7(1).Value = True Then
        Moose = "* " & LCase(Trim(Text2.Text)) ' first word
    End If
    If Option7(2).Value = True Or Option7(3).Value = True Then
        Moose = LCase(Trim(Text2.Text)) ' Any word
    End If
End If
SSTab1.Tab = 0
'Pinyin = Trim(LCase(Text2.Text))
'For I = 1 To Len(Pinyin)
    'If Asc(Mid(Pinyin, I, 1)) > 47 And Asc(Mid(Pinyin, I, 1)) < 58 Then
        'Numbered = True
        'Exit For
    'Else
        'MsgBox "Please input only numbered Pinyin!"
        'Screen.MousePointer = 0
        'Exit Sub
    'End If
' Next


Outy = ""
Dictionary.Text = ""
Close #1: Open App.Path + "\adso.gb" For Input As #1
Do While Not EOF(1)
Line Input #1, Inny
    y = LCase(Trim(Inny))
    a = InStr(y, "*")
    B = InStr(a + 1, y, "*")
    c = InStr(B + 1, y, "*")
    d = InStr(c + 1, y, "*")
    e = InStr(d + 1, y, "*")
    y = Replace(Mid(y, a + 1, d - a), "*", " ")
If InStr(Inny, Moose) <> 0 Then
    Outy = Outy & y & vbCrLf
    If Option7(0).Value = True Or Option7(2).Value = True Then Exit Do
End If
Loop
Close #1
If Trim(Outy) <> "" Then
    Dictionary.Text = Outy & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
Else
    Dictionary.Text = UCase(Moose) & "***NOT*** found in dictionaries" & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf
End If
Screen.MousePointer = 0
End Sub

Public Sub Command9_Click()
On Error Resume Next

Dim WhoDo As Boolean
Dim Moose As String
Dim Inny, Inny1, Outy, Outyd, Loudy, y, Outy1 As String
Dim i, j, k, l, a, B, c, d, e As Long
Screen.MousePointer = 11
If Trim(Text1.Text) = "" Then
    Screen.MousePointer = 0
    Exit Sub
Else

    If Option7(0).Value = True Or Option7(1).Value = True Then
        Moose = "* " & LCase(Trim(Text1.Text)) ' first word
    End If
    If Option7(2).Value = True Or Option7(3).Value = True Then
        Moose = LCase(Trim(Text1.Text)) ' Any word
    End If
    
End If
SSTab1.Tab = 0
Inny = ""
Inny1 = ""
Outy = ""
'Dictionary.Text = ""
Close #1: Open App.Path + "\adso.gb" For Input As #1
Do While Not EOF(1)
Line Input #1, Inny
'DoEvents
    Moose = LCase(Trim(Moose))
    Inny = LCase(Trim(Inny))
    y = Inny
    a = InStr(y, "*")
    B = InStr(a + 1, y, "*")
    c = InStr(B + 1, y, "*")
    d = InStr(c + 1, y, "*")
    e = InStr(d + 1, y, "*")
    'MsgBox a & "   " & d
    y = Replace(Mid(y, a + 1, d - a), "*", " ")
    If InStr(Inny, Moose) > 0 Then
        Outy = Outy & y & vbCrLf
        If Option7(0).Value = True Or Option7(2).Value = True Then Exit Do
    End If
Loop
Close #1

If Trim(Outy) <> "" Then
    Outer = Outy
    Dictionary.Text = Outy & vbCrLf & Dictionary.Text & vbCrLf & vbCrLf       ' & vbCrLf & Inny1
Else
   Dictionary.Text = "***" & UCase(Moose) & "***   Not found in this dictionary" & Dictionary.Text & vbCrLf & vbCrLf
End If
Screen.MousePointer = 0

End Sub

Private Sub cpvSlider1_ValueChanged()
On Error Resume Next

TTSAppMain.cpvSlider1.Value = cpvSlider1.Value
End Sub

Private Sub cpvSlider2_ValueChanged()
On Error Resume Next

TTSAppMain.cpvSlider2.Value = cpvSlider2.Value

End Sub

Private Sub cpvSlider3_ValueChanged()
On Error Resume Next
Timer1.Enabled = False
Select Case cpvSlider1.Value
Case 0
    Timer1.Interval = 500
Case 1
    Timer1.Interval = 2000
Case 2
    Timer1.Interval = 4000
Case 3
    Timer1.Interval = 6000
End Select
Timer1.Enabled = True
End Sub

Private Sub Dictionary_Change()
On Error Resume Next

Dictionary.Text = Replace(Dictionary.Text, "|", "")
Dictionary.Text = Replace(Dictionary.Text, "*", "")

End Sub

Private Sub Dictionary_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Dim txt As String

    'txt = RichWordOver(Dictionary, x, y)
    'Me.Caption = Lettr
    'If lblCurrentWord.Caption <> txt Then lblCurrentWord.Caption = txt

End Sub

Private Sub Dictionary_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 2 Then PopupMenu mnusch
    Exit Sub
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Dictionary.SelLength <> 0 And Trim(Dictionary.SelText) <> "" Then
    Goose = ""
    j = 1
    Moose = Trim(Dictionary.SelText)
    'If Len(Moose) > 24 Then
        'For I = 1 To Len(Moose)
            'If j = 22 Then
                'j = 1
                'Goose = Goose & Mid(Moose, I, 1) & vbCrLf
            'Else
                'Goose = Goose & Mid(Moose, I, 1)
                'j = j + 1
            'End If
        'Next
        'GlobalString = Goose
        'If Trim(Goose) <> "" Then
            'TTSAppMain.MainTxtBox.Text = Goose
        'Else
            'TTSAppMain.MainTxtBox.Text = ""
        'End If
    'Else
        GlobalString = Moose
        If Trim(Moose) <> "" Then
            TTSAppMain.MainTxtBox.Text = Moose
        Else
            TTSAppMain.MainTxtBox.Text = ""
        End If
    'End If
    
    'TTSAppMain.MainTxtBox.Text = Moose
    If Check9.Value = 1 Then
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
    End If
    
    If Asc(Left(Trim(Dictionary.SelText), 1)) < 0 Then ' Chinese < 0
        TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    End If
    TTSAppMain.SpeakBtn_Click
   
End If

End Sub

Private Sub English_Click()
On Error Resume Next

    English.Visible = False
End Sub

Private Sub English_GotFocus()
'Call SelText(English)

End Sub

Private Sub Ep_Change()
On Error Resume Next

Ep.ToolTipText = Ep.Text
End Sub

Private Sub File1_DblClick()
On Error Resume Next

Dim sFile, sPath As String
Dim x As Long

    For x = 0 To (File1.ListCount - 1)      '*** Listbox starts at 0 so delete 1 from count.
        If File1.Selected(x) = True Then    '*** Find selected filename.
            sPath = File1.Path
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\" '*** Root path doesn't need \
            sFile = sPath & File1.FileName
            Exit For                         '*** Stop looking after you found it...
        End If
    Next x

'*** Open file. No need to find associated program...
ShellExecute Me.hWnd, vbNullString, sFile, vbNullString, "C:\", SW_SHOWNORMAL
'File1.Visible = False

End Sub
Private Sub SelText(t As TextBox)
On Error Resume Next

With t
    .SelStart = 0
    .SelLength = Len(t.Text)
End With
End Sub
Private Sub List4Scroll()
On Error Resume Next

   Dim c As Long
   Dim rcText As RECT
   Dim newWidth As Long
   Dim ItemWidth As Long
   Dim sysScrollWidth As Long
   FormX.Font.Name = List4.Font.Name
   FormX.Font.Bold = List4.Font.Bold
   FormX.Font.Size = List4.Font.Size
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   For c = 0 To List4.ListCount - 1
      Call DrawText(FormX.hDC, (List4.List(c)), -1&, rcText, DT_CALCRECT)
            ItemWidth = rcText.Right + sysScrollWidth
         
      If ItemWidth >= newWidth Then
         newWidth = ItemWidth
      End If
   Next
   Call SendMessage(List4.hWnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)

End Sub
Private Sub List5Scroll()
On Error Resume Next

   Dim c As Long
   Dim rcText As RECT
   Dim newWidth As Long
   Dim ItemWidth As Long
   Dim sysScrollWidth As Long
   FormX.Font.Name = List5.Font.Name
   FormX.Font.Bold = List5.Font.Bold
   FormX.Font.Size = List5.Font.Size
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   For c = 0 To List4.ListCount - 1
      Call DrawText(FormX.hDC, (List5.List(c)), -1&, rcText, DT_CALCRECT)
            ItemWidth = rcText.Right + sysScrollWidth
      If ItemWidth >= newWidth Then
         newWidth = ItemWidth
      End If
   Next
   Call SendMessage(List5.hWnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)

End Sub

Private Sub Form_Initialize()
SSTab1.Tab = 0
SSTab2.Tab = 0
SSTab4.Tab = 0
End Sub

Private Sub Form_Load()
Dim Inny, Inny1, Outy, Loudy As String
Dim i, j, k, l, back As Long
Dim Moose, Goose, Loose As String
Dim Ontop As String
Dim flags As Long, Connected As String
Dim Result As Boolean
On Error Resume Next
Dim Medical, xx, yy, zz As String
Medical = "Emergency.txt"
List4.Clear
Open App.Path & "\" & Medical For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List4.AddItem xx
    End If
Loop
Close #1
List4Scroll
WebBrowser1.Navigate2 ("about:blank")
List5.Clear
Open App.Path & "\MedicalDictionary.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List5.AddItem xx
    End If
Loop
Close #1
List5Scroll

   Text4.ZOrder 0

   mnuGeneral1_Click        'Load links
   File1.ZOrder 0
   File1.Refresh
   tColor = &HE3DFE0
    
    SSTab4_Click (0)
    SSTab1_Click (0)
    SSTab4.ZOrder 0
    List4.ZOrder 0
    Frame4.ZOrder 0
    

SSTab1.Tab = 0  ' 9
SSTab4.Tab = 0
SSTab7.Tab = 0
'WebBrowser2.Navigate "about:<html><body scroll='no'><img src='" & App.Path & "\an1.gif" & "'></img></body></html>"
'WebBrowser3.Navigate "http://www.moosenose.com/Default.asp"
'WebBrowser3.Navigate "http://zhongwen.com/chat.htm"
'WebBrowser3.Navigate "about:<html><body scroll='no'><img src='" & App.Path & "\an1.gif" & "'></img></body></html>"
'WebBrowser1.Navigate "about:<html><head><title></title></head><body bgcolor=""black"" text=""black"" link=""blue"" vlink=""purple"" alink=""red""><p><span style=""font-size:14pt;""><b><font color=""red"">It appears that you do not have an internet connection. Much of this program will not function properly.</font></b></span></p></body></html>"
'<a href=" & Chr(34) & "http://moosenose.com/Pinny.zip" & Chr(34) & ">
'WebBrowser1.ToolTipText = "To translate a Webpage, Navigate to the page using Internet Explorer and Select translate menu above."
    
WebBrowser4.Left = 45
WebBrowser4.Top = 345
'WebBrowser4.Height = 3060
'WebBrowser4.Width = 5805
Text7.Left = 45
Text7.Top = 345
'Text7.Height = 3060
'Text7.Width = 5805

PinIncr = 1
i1 = 0
DoEvents
'Me.Height = 6975
'Me.Width = 7365
'SSTab1.Height = 6630
'SSTab1.Width = 7245
SSTab1.Top = 0
SSTab1.Left = 0
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
j = 0
Flagday = False

Combo12.Clear
Close #1: Open App.Path & "\Topics" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Inny
        Inny = Trim(Inny)
        Combo12.AddItem Inny
    Loop
Close #1
Combo12.Text = "Conversational Topics" 'Combo9.List(0)


Close #1: Open App.Path & "\Pynny.txt" For Input As #1 'numbered pinyin
Close #2: Open App.Path & "\PYnal.txt" For Input As #2 '  "  +   simplified
Do While Not EOF(1)
    Line Input #1, Inny
    Inny = Trim(Inny)
    Combo6.AddItem Trim(Inny)
    Line Input #2, Inny1
    i = InStr(Inny1, " ")
    Outy = Trim(Right(Inny1, Len(Inny1) - i))
    If Asc(Left(Outy, 1)) < 1 Then
        Combo8.AddItem Trim(Outy)
    Else
        Combo8.AddItem Trim(Inny1)
    End If
Loop
Close #1
Close #2
Combo6.Text = "Pinyin#"   'Combo6.List(0)
Combo8.Text = Combo8.List(0)
Text5.Text = ""   'Combo8.List(0)
Pinyin = "[a1"

'result = InternetGetConnectedState(Flags, 0)
'If result Then
'DoEvents
'File1.Path = App.Path & "\Links\\PinYinDict\Links"
'If Dir(App.Path & "\Links\Linkz.exe") <> "" And File1.ListCount = 0 Then
'    ShellExecute Me.hwnd, vbNullString, App.Path & "\Links\Linkz.exe", vbNullString, "C:\", SW_SHOWNORMAL
'End If

back = 0

File1.Path = App.Path & "\Links"
If File1.ListCount = 0 Then
    ChDir App.Path & "\Links\"
    Shell App.Path & "\Links\Links.exe"
    ChDir App.Path
End If
Close #1: Open App.Path & "\Settings.ini" For Input As #1
    Line Input #1, Ontop
Close #1
If Ontop = "True" Then
    SetTopMostWindow Me.hWnd, True
    Check1.Value = 1
Else
    Check1.Value = 0
    SetTopMostWindow Me.hWnd, False
End If
Command44.Picture = mapic(0).Picture

    WebBrowser1.Navigate App.Path & "\Help\speakingpinyin.htm"
'Else
    'WebBrowser1.Navigate "about:<html><head><title></title></head><body bgcolor=""black"" text=""black"" link=""blue"" vlink=""purple"" alink=""red""><p><span style=""font-size:14pt;""><b><font color=""red"">It appears that you do not have an internet connection. Much of this program will not function properly.</font></b></span></p></body></html>"
'End If

WebBrowser5.Navigate App.Path & "\Medications.htm"
WebBrowser2.Navigate App.Path & "\Help\InternetRadio.htm"
'WebBrowser2.Navigate "http://www.multilingualbooks.com/online-radio-chinese.html#about"
'WebBrowser1.Navigate App.Path & "\Audio\PinyinDictionary.htm"
'Anim1.AnimatedGifPath = App.Path & "\PleaseWait.gif"
    Combo1.ListIndex = 0
    Combo1.AddItem "Asthma"
    Combo1.AddItem "Blood Components and Blood Substituents"
    Combo1.AddItem "Bronchodilators"
    Combo1.AddItem "Calcium Regulating Agents"
    Combo1.AddItem "Antibiotics"
    Combo1.AddItem "Cytotoxic Agents"
    Combo1.AddItem "Digestants"
    Combo1.AddItem "Digitalis Glycoside"
    Combo1.AddItem "Diuretics"
    Combo1.AddItem "Drugs Used for Relief of Pain and Inflammation"
    Combo1.AddItem "Enteral Nutrition"
    Combo1.AddItem "Estrogens"
    Combo1.AddItem "Expectorants and Cough Preparations"
    Combo1.AddItem "General Antidote"
    Combo1.AddItem "Hematopoietic Agents"
    Combo1.AddItem "Hemostatics"
    Combo1.AddItem "Hormones and Antagonists"
    Combo1.AddItem "Immunomodulators"
    Combo1.AddItem "Laxatives"
    Combo1.AddItem "Miscellaneous"
    Combo1.AddItem "Mucolytic Agents"
    Combo1.AddItem "Neurologic Drugs"
    Combo1.AddItem "Parenteral Nutrition"
    Combo1.AddItem "Psychopharmacologic Drugs"
    Combo1.AddItem "Radiopaque Media"
    Combo1.AddItem "Replenishers and Regulators of Water and Electrolytes"
    Combo1.AddItem "Specific Antidotes"
    Combo1.AddItem "Agents for Active Immunity"
    Combo1.AddItem "Agents for Passive Immunity"
    Combo1.AddItem "Agents Related to Anterior Pituitary and Hypothalamic Function"
    Combo1.AddItem "Agents Used for Diagnosis of Physiological Function"
    Combo1.AddItem "Agents Used in Anesthesia"
    Combo1.AddItem "Agents Used in Hemophilia"
    Combo1.AddItem "Agents Used in Peptic Ulcer Disease"
    Combo1.AddItem "Agents Used to Treat Circulatory Failure"
    Combo1.AddItem "Agents Used to Treat Deficiency Anemia"
    Combo1.AddItem "Agents Used to Treat Hemorrhoid"
    Combo1.AddItem "Agents Used to Treat Hyperlipidemia"
    Combo1.AddItem "Agents Used to Treat Thyroid Disorders"
    Combo1.AddItem "Androgens"
    Combo1.AddItem "Antiallergic Agents"
    Combo1.AddItem "Antianginal Agents"
    Combo1.AddItem "Antiarrhythmics"
    Combo1.AddItem "Anticholinergic Antispasmodics"
    Combo1.AddItem "Anticoagulants"
    Combo1.AddItem "Antidiabetic Agents"
    Combo1.AddItem "Antidiarrheals"
    Combo1.AddItem "Antiemetics"
    Combo1.AddItem "Antihistamines"
    Combo1.AddItem "Antihypertensives"
    Combo1.AddItem "Antihypotensives"
    Combo1.AddItem "Antineoplastic Adjuncts"
    Combo1.AddItem "Antiplatelets and Thrombolytics"
    Combo1.AddItem "Antiretroviral Agents"
    Combo1.AddItem "Antituberculosis Agents"
    Combo1.Text = "Drug Classes"
    UUaarreell = App.Path & "\Medications.htm"
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    'ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload TTSAppMain
    Set TTSAppMain = Nothing
    Unload TTSApp
    Set TTSApp = Nothing
    Unload Webrow
    Set Webrow = Nothing
    Unload frmMain
    Set frmMain = Nothing
    Unload FormX
    Set FormX = Nothing
    Unload RecogVb
    Set RecogVb = Nothing
    Unload frmValueTip
    Set frmValueTip = Nothing
    Unload Simple
    Set Simple = Nothing
    
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Label24.FontBold = False
    Label24.FontUnderline = False
    Label24.ForeColor = vbBlue
    Label30.FontBold = False
    Label30.FontUnderline = False
    Label30.ForeColor = vbBlue
    SetCursor LoadCursor(0, IDC_ARROW)
End Sub


Private Sub Label14_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub Label17_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub Label19_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    SetCursor LoadCursor(0, IDC_HAND)

    Label24.FontBold = True
    Label24.FontUnderline = True
    Label24.ForeColor = vbRed
    'Me.MousePointer = 99

End Sub

Private Sub Label24_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

OpenBrowser Label24.Caption, FormX.hWnd
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    SetCursor LoadCursor(0, IDC_HAND)

    Label30.FontBold = True
    Label30.FontUnderline = True
    Label30.ForeColor = vbRed
    'Me.MousePointer = 99

End Sub

Private Sub Label30_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    OpenBrowser Label30.Caption, FormX.hWnd
End Sub

Private Sub Label33_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub Label36_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub Label40_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub Label41_Change()
On Error Resume Next

If Val(Label41.Caption) > 5 Then
    Label41.Caption = Label41.Caption - 5
End If
If Val(Label41.Caption) = 0 Then
    Label41.Caption = 5
End If
Command44.Picture = mapic(Label41.Caption).Picture

End Sub

Private Sub Label41_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Command44_Click
End Sub

Private Sub Label42_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Command44_Click
End Sub

Private Sub Label43_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Command44_Click
End Sub

Private Sub Linkz_DblClick()
On Error Resume Next

Dim sFile, sPath As String
Dim x As Long
    For x = 0 To (Linkz.ListCount - 1)      '*** Listbox starts at 0 so delete 1 from count.
        If Linkz.Selected(x) = True Then    '*** Find selected filename.
            sPath = File1.Path
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\" '*** Root path doesn't need \
            sFile = sPath & Linkz.List(x)
            Exit For                         '*** Stop looking after you found it...
        End If
    Next x

'*** Close #1:Open file. No need to find associated program...
ShellExecute Me.hWnd, vbNullString, sFile, vbNullString, "C:\", SW_SHOWNORMAL

'File1.Visible = False

End Sub

Private Sub Linkz_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim XY As Long
Dim Please As String
On Error Resume Next
    For XY = 1 To Data.Files.Count
        Please = Data.Files(XY)
        MsgBox Please
    Next XY
End Sub



Private Sub List2_Click()
On Error Resume Next

Dim i, i1, j As Long
Dim Inny, Outy, Outy1 As String
Dim Moose, Goose, Goose1, Loose As String
Dim m, m1, m2 As Integer
Dim a, B, c, d, e, f, G As Long
Dim lngReturnNumber, ngReturnNumber As Long
Screen.MousePointer = 11

m = InStrRev(List2.List(List2.ListIndex), " ")
Goose = Trim(Mid(List2.List(List2.ListIndex), m + 1, Len(List2.List(List2.ListIndex)) - m))
Loose = Dir(App.Path & "\Audio\am1.wma")
If Right(Goose, 4) = ".wma" And Loose = "am1.wma" Then
    ngReturnNumber = ShellExecLaunchFile(App.Path & "\Audio\" & Goose, "", App.Path)
End If
Screen.MousePointer = 0
Exit Sub

m = InStr(List2.List(List2.ListIndex), " ")

'Clipboard.Clear
'Clipboard.SetText "[" & Trim(Left(List2.List(List2.ListIndex), m))

Pinyin = "[" & Trim(Left(List2.List(List2.ListIndex), m)) ' ONLY OPTION
Goose = Trim(Left(List2.List(List2.ListIndex), m))
Combo6.Text = Trim(Goose)
Combo8.ListIndex = Combo6.ListIndex
MsgBox Combo6.ListIndex
Text5.Text = ""
Text5.Text = Combo8.List(Combo8.ListIndex)
Translit_Click


End Sub

Private Sub List3_Click()
On Error Resume Next

Dim Moose, Goose, Loose, ngReturnNumber  As String
Dim m, m1, m2 As Integer
Dim lngReturnNumber As Long
Moose = List3.List(List3.ListIndex)
m = InStr(1, Moose, "*")
m1 = InStr(m + 1, Moose, "*")
m2 = Val(Mid(Moose, m + 1, m1 - m - 1))
'MsgBox List2.List(m2)
List2.ListIndex = m2
m = InStrRev(List2.List(m2), " ")
Goose = Trim(Mid(List2.List(m2), m + 1, Len(List2.List(m2)) - m))
Loose = Dir(App.Path & "\am1.wma")
If Right(Goose, 4) = ".wma" And Loose = "am1.wma" Then
    ngReturnNumber = ShellExecLaunchFile(App.Path & "\" & Goose, "", App.Path)
End If
List2.Visible = True
List3.Visible = False
Text4.Visible = False

End Sub


Private Sub List4_Click()
On Error Resume Next
Dim a, B, c, d As String
Dim i, j, k, l As Integer
a = Trim(List4.List(List4.ListIndex))
i = InStr(a, "*--")
j = InStr(i + 1, a, "*")
k = Len(a)
Ep = Left(a, i - 1)
SP = Mid(a, i + 3, j - i - 3)
PP = Right(a, k - j)

End Sub

Private Sub List4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error Resume Next

   List4.ToolTipText = Trim(List4.List(List4.ListIndex))
End Sub

Private Sub List5_Click()
On Error Resume Next
Dim a, B, c, d As String
Dim i, j, k, l As Integer
a = Trim(List5.List(List5.ListIndex))
i = InStr(a, "*--")
j = InStr(i + 1, a, "*")
k = Len(a)
Ep = Left(a, i - 1)
SP = Mid(a, i + 3, j - i - 3)
PP = Right(a, k - j)

End Sub

Private Sub mnSaveCEd_Click()
On Error Resume Next

Command57_Click
End Sub


Private Sub mnuAbout_Click()
On Error Resume Next

WebBrowser1.Navigate "http://www.moosenose.com/v2.01/PinyinDictionary2.01.htm"
back = back + 1
SSTab1.Tab = 2

End Sub

Private Sub mnuCe_Click()
On Error Resume Next

Translatez = "zh_en"
Unload frmMain
Load frmMain
frmMain.Show
End Sub

Private Sub mnuceng_Click()
On Error Resume Next

Ep.Text = ""
End Sub

Private Sub mnuClipBothE_Click()
On Error Resume Next

mnuClipPronE_Click
mnuClipTransE_Click
End Sub

Private Sub mnuClipBothP_Click()
On Error Resume Next

mnuClipPronP_Click
mnuClipTransP_Click

End Sub

Private Sub mnuClipBothS_Click()
On Error Resume Next

mnuClipPronS_Click
mnuClipTransS_Click

End Sub

Private Sub mnuClipPronE_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
If Trim(Clipboard.GetText) = "" Or Asc(Left(Trim(Clipboard.GetText), 1)) < 0 Then Exit Sub
'Text1.Text = Clipboard.GetText
SSTab1.Tab = 0
GlobalString = Clipboard.GetText
Moose = GlobalString
If Trim(Moose) <> "" Then
    TTSAppMain.MainTxtBox.Text = Moose
Else
    TTSAppMain.MainTxtBox.Text = ""
    Exit Sub
End If

If Check9.Value = 1 Then
    TTSAppMain.Show
    'Call MakeTranslucent(TTSAppMain, tColor)
End If
TTSAppMain.VoiceCB.Text = "Microsoft Mary"
TTSAppMain.SpeakBtn_Click

End Sub

Private Sub mnuClipPronP_Click()
On Error Resume Next

Text8.Text = Clipboard.GetText
Command40_Click
End Sub

Private Sub mnuClipPronS_Click()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
Text1.Text = Clipboard.GetText
Command9_Click
SSTab1.Tab = 0
GlobalString = Clipboard.GetText
Moose = GlobalString
If Trim(Moose) <> "" Then
    TTSAppMain.MainTxtBox.Text = Moose
Else
    TTSAppMain.MainTxtBox.Text = ""
    Exit Sub
End If

If Check9.Value = 1 Then
    TTSAppMain.Show
    'Call MakeTranslucent(TTSAppMain, tColor)
End If
    
If Asc(Left(Trim(Dictionary.SelText), 1)) < 0 Then ' Chinese < 0
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
Else
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
End If
TTSAppMain.SpeakBtn_Click


End Sub

Private Sub mnuClipTransE_Click()
On Error Resume Next

Text1.Text = Clipboard.GetText
Command9_Click
SSTab1.Tab = 0
End Sub

Private Sub mnuClipTransP_Click()
On Error Resume Next

Text2.Text = Clipboard.GetText
Command8_Click
SSTab1.Tab = 0
End Sub

Private Sub mnuClipTransS_Click()
On Error Resume Next

Text14.Text = Clipboard.GetText
Command46_Click
SSTab1.Tab = 0
End Sub

Private Sub mnucpin_Click()
On Error Resume Next

PP.Text = ""
End Sub

Private Sub mnucsimp_Click()
On Error Resume Next

SP.Text = ""
End Sub

Private Sub mnuDTP_Click()
On Error Resume Next

Dim DesktopPath As String
DesktopPath = GetShellFolderPath(&H0)
SavePath = DesktopPath

End Sub

Private Sub mnuEc_Click()
On Error Resume Next

Translatez = "en_zh"
Unload frmMain
Load frmMain
frmMain.Show
End Sub


Private Sub mnuEnabling_Click()
On Error Resume Next

WebBrowser1.Navigate App.Path & "\Help\EnablingInternational.htm"
back = back + 1
SSTab1.Tab = 2
Uarel = App.Path & "\Help\EnablingInternational.htm"
End Sub

Private Sub mnuExit_Click()
On Error Resume Next

    Unload Me
End Sub

Private Sub mnuGeneral_Click()
On Error Resume Next

Dim x As Integer
Dim c As Long
File1.ZOrder 0
File1.Path = App.Path & "\Links\"
File1.Refresh

   Linkz.Clear
   For c = 0 To File1.ListCount - 1
        Linkz.AddItem File1.List(c)
   Next
SSTab1.Tab = 3
End Sub
Private Sub mnuGeneral1_Click()
On Error Resume Next

Dim x As Integer
Dim c As Long
File1.ZOrder 0

File1.Path = App.Path & "\Links\"
File1.Refresh
   Linkz.Clear
   For c = 0 To File1.ListCount - 1
        Linkz.AddItem File1.List(c)
   Next
End Sub

Private Sub mnuIFEAL_Click()
On Error Resume Next

Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,1"
End Sub

Private Sub mnuInstructions_Click()
On Error Resume Next

WebBrowser1.Navigate App.Path & "\Help\speakingpinyin.htm"

back = back + 1
SSTab1.Tab = 2

End Sub

Private Sub mnuMandarin_Click()
On Error Resume Next

OpenBrowser "http://www.mandarintools.com/", FormX.hWnd

End Sub



Private Sub mnupeng_Click()
Ep.Text = Ep.Text & GlobalString & " |"
End Sub

Private Sub mnuPinology_Click()
On Error Resume Next

OpenBrowser "http://www.pinyinology.com", FormX.hWnd

End Sub

Private Sub mnuLangPk_Click()
On Error Resume Next

WebBrowser1.Navigate "http://download.microsoft.com/download/speechSDK/SDK/5.1/WXP/EN-US/speechsdk51LangPack.exe"

End Sub


Private Sub mnuPinyin_Click()
On Error Resume Next

If Text4.Visible = False Then
    English.Visible = True
    Text4.Visible = True
Else
    English.Visible = False
    Text4.Visible = False
End If
End Sub

Private Sub mnuppin_Click()
On Error Resume Next

    PP.Text = PP.Text & GlobalString & " |"
End Sub

Private Sub mnuPRC_Click()
On Error Resume Next

Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,2"
End Sub

Private Sub mnuPronn_Click()
On Error Resume Next

Command2_Click
End Sub

Private Sub mnupsimp_Click()
On Error Resume Next

SP.Text = SP.Text & GlobalString & " |"

End Sub

Private Sub mnuRead_Click()
On Error Resume Next

Load TTSApp
TTSApp.Show
End Sub

Private Sub mnuSAPI_Click()
On Error Resume Next

WebBrowser1.Navigate "http://download.microsoft.com/download/speechSDK/SDK/5.1/WXP/EN-US/speechsdk51.exe"
End Sub

Private Sub mnuSaveCEp_Click()
On Error Resume Next

Command58_Click
End Sub

Private Sub mnuSaveM_Click()
On Error Resume Next

Command56_Click
End Sub

Private Sub mnuscp_Click()
On Error Resume Next
Command35_Click
End Sub



Private Sub mnuShowt_Click()
 On Error Resume Next

       Check9.Value = 1
        TTSAppMain.Show
        'Call MakeTranslucent(TTSAppMain, tColor)
End Sub

Private Sub mnuSSPEAK_Click()
On Error Resume Next

    TTSAppMain.SpeakBtn_Click
End Sub

Private Sub mnuSStop_Click()
On Error Resume Next

TTSAppMain.Command1_Click
End Sub

Private Sub mnuSTP_Click()
On Error Resume Next

SavePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)

End Sub

Private Sub mnuw2m_Click()
On Error Resume Next

    Shell App.Path & "\Wav2Mp3.exe"
End Sub

Private Sub Option8_Click(Index As Integer)
Dim Medical, xx, yy, zz As String
On Error Resume Next
       Select Case Index
              Case 0
                Medical = "Emergency.txt"
              Case 1
                Medical = "ENT.txt"
              Case 2
                Medical = "Respiratory.txt"
              Case 3
                Medical = "GI.txt"
              Case 4
                Medical = "GU.txt"
              Case 5
                Medical = "PMH.txt"
       End Select
List4.Clear
Open App.Path & "\" & Medical For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List4.AddItem xx
    End If
Loop
Close #1
List4Scroll

End Sub

Private Sub Picture2_Click()
On Error Resume Next

Command44_Click
End Sub

Private Sub Pinyinn_KeyPress(KeyAscii As Integer)
On Error Resume Next

  If KeyAscii = 13 Then
    'HandleReply
  End If

End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Pn1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Dim Moose, Goose As String
    Dim i, j As Long
    Moose = Trim(Simp(PinIncr - 1))
    If Trim(Moose) <> "" Then
        TTSAppMain.MainTxtBox.Text = Moose
    Else
        TTSAppMain.MainTxtBox.Text = ""
    End If
    'TTSAppMain.Show
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    'TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    'tColor = &H8000000F
    ''Call MakeTranslucent(TTSAppMain, tColor)


End Sub

Private Sub PP_Change()
On Error Resume Next

PP.ToolTipText = PP.Text

End Sub

Private Sub Simp_Change(Index As Integer)
On Error Resume Next

    Dim Moose, Goose As String
    Dim i, j As Long
    Moose = Trim(Simp(PinIncr - 1))
    If Trim(Moose) <> "" Then
        TTSAppMain.MainTxtBox.Text = Moose
    Else
        TTSAppMain.MainTxtBox.Text = ""
    End If
    'TTSAppMain.Show
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    'TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    'tColor = &H8000000F
    ''Call MakeTranslucent(TTSAppMain, tColor)

End Sub

Private Sub Simp_Click(Index As Integer)
On Error Resume Next

    Dim Moose, Goose As String
    Dim i, j As Long
    Moose = Trim(Simp(PinIncr - 1))
    If Trim(Moose) <> "" Then
        TTSAppMain.MainTxtBox.Text = Moose
    Else
        TTSAppMain.MainTxtBox.Text = ""
    End If
    'TTSAppMain.Show
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    'TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    TTSAppMain.SpeakBtn_Click
    'tColor = &H8000000F
    ''Call MakeTranslucent(TTSAppMain, tColor)

End Sub



Private Sub SP_Change()
On Error Resume Next

SP.ToolTipText = SP.Text

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim i, x As Integer, Medical As String
On Error Resume Next
Frame7.Top = 1230
Frame7.Left = 0
If PreviousTab = 7 Then
    For i = 1 To PinIncr
        Unload Pn1(i)
        Unload Simp(i)
    Next
    PinIncr = 1
End If
If SSTab1.Tab = 1 Or SSTab1.Tab = 4 Then
    Frame2.Visible = False
Else
    Frame2.Visible = True
End If
If SSTab1.Tab = 1 Then
     'SSTab4.Tab = 0
     Frame3.Visible = False
     List4.Visible = True
     List4.ZOrder 0
     List4_Click
End If
If SSTab1.Tab <> 0 Then
    Dictionary.Visible = False
Else
    Dictionary.Visible = True
End If
'If SSTab1.Tab = 2 And flags = "Connected to the Internet" Then
    'MsgBox "It appears that you do not have " & vbCrLf & _
            "an internet connection.  Much of " & vbCrLf & _
            "this program will not function properly."
'End If
'If SSTab1.Tab = 0 Then
'   Ep.Visible = False
'   SP.Visible = False
'   PP.Visible = False
'Else
'   Ep.Visible = True
'   SP.Visible = True
'   PP.Visible = True
'End If
If PreviousTab = 3 Then Check4.Value = 0
If PreviousTab = 9 Then
    Combo6.Visible = True
    Text5.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Command9.Visible = True
    Command8.Visible = True
End If
'If SSTab1.Tab = 8 Then
    'Load RecogVb
    'RecogVb.Show
'End If
If SSTab1.Tab = 9 Then
    Combo6.Visible = True
    Text5.Visible = False
    Text1.Visible = True
    Text2.Visible = True
    English.Visible = False
    Command8.Visible = True
End If
If SSTab1.Tab = 1 Then
    SSTab4.Tab = 0
    SSTab4_Click (1)
End If
If SSTab1.Tab = 2 Then
    SSTab2.Tab = 0
End If

File1.ZOrder 0
File1.Path = App.Path & "\Links\"
File1.Refresh
End Sub
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
On Error Resume Next

    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", SW_SHOWDEFAULT)
End Function

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Label24.FontBold = False
    Label24.FontUnderline = False
    Label24.ForeColor = vbBlue
    Label30.FontBold = False
    Label30.FontUnderline = False
    Label30.ForeColor = vbBlue
    SetCursor LoadCursor(0, IDC_ARROW)

End Sub

Private Sub SSTab3_DblClick()

End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
On Error Resume Next

SSTab4.ZOrder 0
End Sub

Private Sub SSTab4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Medical As String, xx As String, yy As String, zz As String
On Error Resume Next
SSTab4.ZOrder 0
List4.ZOrder 0
If SSTab4.Tab = 0 Then
    Command62.Visible = True
Else
    Command62.Visible = False
End If

Frame3.Visible = True
Frame4.ZOrder 0
       Select Case SSTab4.Tab
              Case 0
                Medical = "Emergency.txt"
              Case 1
                Medical = "Chinese.txt"
              Case 2
                Medical = "Travel.txt"
              Case 3
                Medical = "Shopping Money.txt"
              Case 4
                Medical = "Money.txt"
              Case 5
                Medical = "Time-Weather.txt"
              Case 6
                Medical = "Numbers.txt"
              Case 7
                Medical = "Dining Food.txt"
       End Select
List4.Clear
If Medical = "" Then Exit Sub
Open App.Path & "\" & Medical For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List4.AddItem xx
    End If
Loop
Close #1
List4Scroll

End Sub

Private Sub SSTab7_Click(PreviousTab As Integer)
On Error Resume Next

SSTab7.ZOrder 0
List4.ZOrder 0
End Sub

Private Sub SSTab7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Medical As String, xx As String, yy As String, zz As String
On Error Resume Next
SSTab7.ZOrder 0
List4.ZOrder 0
Frame3.Visible = False
Frame4.ZOrder 0
Command62.Visible = False

       Select Case SSTab7.Tab
              Case 0
                Frame5.Visible = False
                Medical = "Useful Phrases.txt"
                SSTab7.ZOrder 1
              Case 1
                Frame5.Visible = False
                Medical = "Love.txt"
                SSTab7.ZOrder 1
              Case 2
                Frame5.Visible = False
                Medical = "Bar.txt"
                SSTab7.ZOrder 1
              Case 3
                Frame5.Visible = True
                Frame5.ZOrder 0
                Medical = "Insults.txt"
                SSTab7.ZOrder 1
              Case 4
                Frame5.Visible = False
                Medical = ""
                i1 = 1
                SSTab7.ZOrder 0
              Case 5
                Frame5.Visible = False
                Medical = ""
                SSTab7.ZOrder 0
              Case 6
                Frame5.Visible = False
                Medical = ""
              Case 7
                Frame5.Visible = False
                Medical = ""
       End Select
List4.Clear
If Medical = "" Then Exit Sub
Open App.Path & "\" & Medical For Input As #1
Do While Not EOF(1)
    Line Input #1, xx
    xx = Trim(xx)
    If xx <> "" Then
        List4.AddItem xx
    End If
Loop
Close #1
List4Scroll
End Sub

Private Sub Text1_GotFocus()
SSTab1.Tab = 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then Command9_Click
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SelText(Text1)

End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

'Call SelText(Text1)
End Sub

Private Sub Text14_GotFocus()
SSTab1.Tab = 0
End Sub

Private Sub Text14_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Call SelText(Text14)

End Sub

Private Sub Text2_GotFocus()
SSTab1.Tab = 0
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Call SelText(Text2)

End Sub

Private Sub Text4_DblClick()
On Error Resume Next

Text4.Visible = False
English.Visible = False
End Sub

Private Sub Text5_Change()
'Screen.MousePointer = 0
End Sub

Private Sub Text5_GotFocus()
On Error Resume Next

    Call SelText(Text5)
End Sub

Private Sub Text5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Text5.SelLength <> 0 Then
    'SetTopMostWindow Me.hWnd, True
    'Form1.Enabled = False
    TTSAppMain.Show
    TTSAppMain.MainTxtBox.Text = Text5.SelText
    TTSAppMain.Show
    If Asc(Left(Trim(Text5.SelText), 1)) < 0 Then
        TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
    Else
        TTSAppMain.VoiceCB.Text = "Microsoft Mary"
    End If
    TTSAppMain.SpeakBtn_Click
End If
End Sub


Private Sub Text7_DblClick()
On Error Resume Next

    Text7.Visible = False
End Sub

Private Sub Text8_Change()
On Error Resume Next

Text8.Text = LCase(Replace(Text8.Text, " ", ""))
End Sub

Private Sub Text9_Change()
On Error Resume Next
Dim Moose, Goose As String
Dim j As Long
SSTab1.Tab = 0
GlobalString = Text9.Text
Moose = GlobalString
If Trim(Moose) <> "" Then
    TTSAppMain.MainTxtBox.Text = Moose
Else
    TTSAppMain.MainTxtBox.Text = ""
    Exit Sub
End If

If Check9.Value = 1 Then
    TTSAppMain.Show
    'Call MakeTranslucent(TTSAppMain, tColor)
End If
    
If Asc(Left(Trim(Dictionary.SelText), 1)) < 0 Then ' Chinese < 0
    TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
Else
    TTSAppMain.VoiceCB.Text = "Microsoft Mary"
End If
TTSAppMain.SpeakBtn_Click


End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Dim MyValue
Randomize
MyValue = Int((1000 * Rnd) + 1)
Label41.Caption = Right(Trim(Str(MyValue)), 1)
Label42.Caption = mamama(MyValue)
TTSAppMain.MainTxtBox.Text = mamama(MyValue)
TTSAppMain.VoiceCB.Text = "Microsoft Simplified Chinese"
TTSAppMain.SpeakBtn_Click

End Sub


Private Sub Translit_Click()
On Error Resume Next

Dim i, j, k, Rflag, TextLines As Long
Dim p2, Moose As String
Dim Spinyin(10) As String
On Error Resume Next
English.Text = ""
j = 0
Spinyin(0) = Pinyin
    For i = 1 To Len(Spinyin(j))
        'MsgBox LTrim(English.Text) & "   " & LCase(Mid(Spinyin(j), i, 1))
        If Rflag = 1 Then
            Rflag = 0
            GoTo Skipper
        End If
        If Rflag = 2 Then
            Rflag = 1
            GoTo Skipper
        End If
        
        'If LCase(Mid(Spinyin(j), i, 1)) = " " And LCase(Mid(Spinyin(j), i - 1, 1)) <> " " Then
        '    English.Text = LTrim(English.Text) & "    "
        'End If
        'MsgBox LCase(Mid(Spinyin(j), i, 1))
        
        '
        Select Case LCase(Mid(Spinyin(j), i, 1))
        Case "a"
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "i" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "y" Then
        English.Text = LTrim(English.Text) + "e-" 'a (in ian or yan) = e in ltrim(English.Text) get
        Else
        English.Text = LTrim(English.Text) + "a-" 'a (elsewhere) = a in ltrim(English.Text) father
        End If
        Case "b"
        English.Text = LTrim(English.Text) + "b-"
        Case "c"
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "tsz-" 'ci = (silent i!) as though ltrim(English.Text) had a word spelt tsz
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" Then
        English.Text = LTrim(English.Text) + "tsz-" 'ch = ch in ltrim(English.Text) church
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" And LCase(Mid(Spinyin(j), i + 8, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "chr-" 'chi = (silent i!) as though ltrim(English.Text) had a word spelt chr
        Rflag = 2
        Else
        English.Text = LTrim(English.Text) + "ts-" 'c = ts in ltrim(English.Text) hats
        End If
        End If
        End If
        
        
        Case "d"
        English.Text = LTrim(English.Text) + "d-"
        
        Case "e"
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "r" Then
        English.Text = LTrim(English.Text) + "are -" 'er = Midwestern ltrim(English.Text) are
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "n" Then
        English.Text = LTrim(English.Text) + "a -" 'e (before n or ng) = a in ltrim(English.Text) alone
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "i" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "y" Then
        English.Text = LTrim(English.Text) + "eh -" 'e (after i or y) = e in ltrim(English.Text) get
        Else
        English.Text = LTrim(English.Text) + "uhr(eh) -" 'e (alone) = usually between the u in ltrim(English.Text) lump or and the u in ltrim(English.Text) lurch  'e (alone) = occasionally in exclamations and a few particles) like e in ltrim(English.Text) get.
        End If
        End If
        End If
        
        Case "f"
        English.Text = LTrim(English.Text) + "f-"
        Case "g"
        English.Text = LTrim(English.Text) + "g-"
        Case "h"
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " Or i = 1 Then
        English.Text = LTrim(English.Text) + "y-" 'ch (initial) = ltrim(English.Text) h, only slightly harder, almost to being like the ch in Scottish Loch or German Hoch
        Else
        English.Text = LTrim(English.Text) + "h-"
        End If
        Case "i"
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " Or i = 1 Then
        English.Text = LTrim(English.Text) + "y-" 'I may never occur as an initial letter of a syllable. It always turns to y in that case.
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " And Val(LCase(Mid(Spinyin(j), i + 1, 1))) <> 0 Then
        English.Text = LTrim(English.Text) + "yih -" ' 'If i is the only sound in the syllable, then it becomes yi.
        Else
        If i = 1 And Val(LCase(Mid(Spinyin(j), i + 1, 1))) <> 0 Then
        English.Text = LTrim(English.Text) + "yih -" '
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "h" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "r" Then
        English.Text = LTrim(English.Text) + "ur -" 'i (after h or r) = Midwestern r in hurt [footnote 4]
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "z" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "c" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "s" Then
        English.Text = LTrim(English.Text) + "z -" 'i (after z, c, s) = ltrim(English.Text) z [footnote 4
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "a" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "e" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "i" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "o" Then
        English.Text = LTrim(English.Text) + "y -" 'i (before or after another vowel, except u) = ltrim(English.Text) y
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "u" Then
        English.Text = LTrim(English.Text) + "yo -" 'iu = ltrim(English.Text) yo as in "Yo ho!" or yeo as in yeoman
        Rflag = 1
        Else
        English.Text = LTrim(English.Text) + "ee -" 'i (elsewhere) = ee in ltrim(English.Text) see
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        
        Case "j"
        English.Text = LTrim(English.Text) + "j-"
        Case "k"
        English.Text = LTrim(English.Text) + "k-"
        Case "l"
        English.Text = LTrim(English.Text) + "l-"
        Case "m"
        English.Text = LTrim(English.Text) + "m-"
        Case "n"
        English.Text = LTrim(English.Text) + "n-"
        Case "o"
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "b" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "p" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "m" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "f" Then
        English.Text = LTrim(English.Text) + "uo -" 'o (after b,p,m, or f) = Italian uo
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "a" Then
        English.Text = LTrim(English.Text) + "w -" 'o (after a) = ltrim(English.Text) w
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "u" Then
        English.Text = LTrim(English.Text) + "owe  -" 'ou = ltrim(English.Text) owe
        Else
        English.Text = LTrim(English.Text) + "o -" 'o = Italian o
        End If
        End If
        End If
        
        
        
        Case "p"
        English.Text = LTrim(English.Text) + "p-"
        Case "q"
        English.Text = LTrim(English.Text) + "ch-"      'q = ltrim(English.Text) ch in cheat
        
        Case "r"
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "jr-" 'ri (silent i!) = a French j followed by a Midwestern r!
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " Or i = 1 Then
        English.Text = LTrim(English.Text) + "j-" 'r (initial) = French initial j
        Else
        If Val(LCase(Mid(Spinyin(j), i + 1, 1))) <> 0 Or i = Len(Spinyin(j)) Then
        English.Text = LTrim(English.Text) + "r-" '
        Else
        English.Text = LTrim(English.Text) + "r-" '
        End If
        End If
        End If
        
        Case "s"
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "sz-" 'si (silent i!) = as though ltrim(English.Text) had a word spelt sz
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" Then
        English.Text = LTrim(English.Text) + "sh-" 'sh = sh in ltrim(English.Text) shame
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" And LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "shri-" 'shi (silent i!) = as though ltrim(English.Text) had a word spelt shri
        Rflag = 2
        Else
        English.Text = LTrim(English.Text) + "s-"
        End If
        End If
        End If
        
        Case "t"
        English.Text = LTrim(English.Text) + "t-"
        
        Case "u"
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " And Val(LCase(Mid(Spinyin(j), i + 1, 1))) <> 0 Then
        English.Text = LTrim(English.Text) + "wu-" 'If u is the only sound in the syllable, then it becomes wu
        Else
        If i = 1 And Val(LCase(Mid(Spinyin(j), i + 1, 1))) <> 0 Then
        English.Text = LTrim(English.Text) + "wu-" '
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = " " Or i = 1 Then
        English.Text = LTrim(English.Text) + "w-" '
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "ay-" 'ui = ltrim(English.Text) way
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "j" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "q" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "x" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "i" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "y" Then
        English.Text = LTrim(English.Text) + "u-" 'u (after j,q,x,i, or y) = French u or German
        Else
        If LCase(Mid(Spinyin(j), i - 1, 1)) = "a" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "e" Or LCase(Mid(Spinyin(j), i - 1, 1)) = "o" Or LCase(Mid(Spinyin(j), i + 1, 1)) = "a" Or LCase(Mid(Spinyin(j), i + 1, 1)) = "e" Or LCase(Mid(Spinyin(j), i + 1, 1)) = "o" Then
        English.Text = LTrim(English.Text) + "w-" ' u (before or after another vowel, except i) = ltrim(English.Text) w
        Else
        English.Text = LTrim(English.Text) + "u-" 'u (elsewhere) = Italian u    ü = French u or German   ü (written like ordinary u [without the dots] after j,q,x,i, or y, since they never have a regular u-sound after them)  U may never occur as an initial letter of a syllable. It always turns to w in that case.
        End If
        End If
        End If
        End If
        End If
        End If
        
        Case "w"
        English.Text = LTrim(English.Text) + "w-"
        Case "x"
        English.Text = LTrim(English.Text) + "sh-" 'x = ltrim(English.Text) sh in sheet
        Case "y"
        English.Text = LTrim(English.Text) + "y-"
        Case "z"
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "dz-" 'zi (silent i!) = as though ltrim(English.Text) had a word spelt dz
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" Then
        English.Text = LTrim(English.Text) + "dz-" 'zh = j or g in ltrim(English.Text) judge
        Rflag = 1
        Else
        If LCase(Mid(Spinyin(j), i + 1, 1)) = "h" And LCase(Mid(Spinyin(j), i + 2, 1)) = "i" Then
        English.Text = LTrim(English.Text) + "jr-" 'zhi (silent i!) = as though ltrim(English.Text) had a word spelt jr
        Rflag = 2
        Else
        English.Text = LTrim(English.Text) + "ds-" 'z = ds in ltrim(English.Text) heads
        End If
        End If
        End If
        
        Case "1"
        English.Text = LTrim(English.Text) + "1 (!)" & vbCrLf
        TextLines = TextLines + 1
        Case "2"
        English.Text = LTrim(English.Text) + "2 (^ ?)" & vbCrLf
        TextLines = TextLines + 1
        Case "3"
        English.Text = LTrim(English.Text) + "3 (v ^ ?)" & vbCrLf
        TextLines = TextLines + 1
        Case "4"
        English.Text = LTrim(English.Text) + "4 (v)" & vbCrLf
        TextLines = TextLines + 1
        Case "5"
        English.Text = LTrim(English.Text) + "5 (-)" & vbCrLf
        TextLines = TextLines + 1
        
        End Select
Skipper:
    Next i
End Sub


Private Sub VScroll1_Change()
On Error Resume Next

Dictionary.Font.Size = Abs(VScroll1.Value)
End Sub

Private Sub WebBrowser4_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error GoTo here
    Text = WebBrowser4.Document.Body.InnerText
    'MsgBox Text
    Texty = WebBrowser4.Document.documentElement.OuterHTML
    'run simplified to pinyin conversion.   Also consider saving graphics and html
    'and reconstructing the page with pinyin
here:

End Sub

Private Sub WebBrowser5_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error GoTo here
    Text = WebBrowser5.Document.Body.InnerText
    'MsgBox Text
    Texty = WebBrowser5.Document.documentElement.OuterHTML
    'run simplified to pinyin conversion.   Also consider saving graphics and html
    'and reconstructing the page with pinyin
here:

End Sub

