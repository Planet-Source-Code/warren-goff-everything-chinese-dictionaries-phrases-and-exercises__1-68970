VERSION 5.00
Begin VB.Form frmValueTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "lblTip"
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   40
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblTip"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   360
   End
End
Attribute VB_Name = "frmValueTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  If you want to change Tip font:
'  Change lblTip font and select the same form font

Private Sub Form_Resize()
On Error Resume Next

    Me.Cls
    
    Me.Line (0, 0)-(ScaleWidth, 0), vb3DLight
    Me.Line (0, 0)-(0, ScaleHeight), vb3DLight
    Me.Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow
    Me.Line (0, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1), vb3DDKShadow
End Sub

