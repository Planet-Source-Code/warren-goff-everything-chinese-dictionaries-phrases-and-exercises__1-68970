VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MooseNose Industries"
   ClientHeight    =   90
   ClientLeft      =   315
   ClientTop       =   315
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   5550
      TabIndex        =   3
      Top             =   810
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3750
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Window Titles ?"
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   1110
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin VB.ListBox lstUrls 
      Height          =   2010
      Left            =   105
      TabIndex        =   0
      Top             =   1905
      Width           =   8805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snobbless"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   1845
      TabIndex        =   4
      Top             =   420
      Width           =   2130
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Sub GetURL()
On Error Resume Next


lstUrls.Clear

EnumWindows AddressOf EnumWindowsProc, ByVal 0&


End Sub





Private Sub Command1_Click()
On Error Resume Next

Load Webrow
'Webrow.Show
frmMain.Visible = False
End Sub

Private Sub Command2_Click()
'    MsgBox urlz(0)

End Sub

Private Sub Form_Activate()
On Error Resume Next

Check1.Value = 1
If Flagday = True Then
    Load Webrow
Else
    Uarel = urlz(0)
    FormX.Combo10.Clear
    FormX.Combo11.Clear
    FormX.Combo10.Text = "Page Title"
    FormX.Combo11.Text = "Url"
    FormX.Combo10.Refresh
    FormX.Combo11.Refresh
    FormX.Refresh
    GetURL
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

iz = 0
Me.Top = 1
Me.Left = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Unload frmMain
Unload Webrow
Set frmMain = Nothing
Set Webrow = Nothing
End Sub

