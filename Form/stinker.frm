VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Webrow 
   Caption         =   "Web Translator"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   Icon            =   "stinker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Off"
      Height          =   255
      Left            =   11535
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "On"
      Height          =   255
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fill Form and click the submit button"
      Height          =   255
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Loop all forms on the webpage for all elements"
      Height          =   255
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Html to string"
      Height          =   255
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00905F5F&
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
      Left            =   5250
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   90
      Width           =   3300
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00905F5F&
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
      Left            =   150
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   60
      Width           =   4950
   End
   Begin VB.TextBox Text1 
      Height          =   6255
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   5775
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7815
      Left            =   -30
      TabIndex        =   0
      Top             =   765
      Width           =   7455
      ExtentX         =   13150
      ExtentY         =   13785
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
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Toggle Radio Button"
      Height          =   255
      Left            =   9855
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Webrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ==============================================
' |==============================================|
' |==   Examples provided by Aspen2K. Please  ===|
' |==   forward all your questions to myself  ===|
' |==   on http://visualbasicforum.com        ===|
' |==============================================|
'  ==============================================


Option Explicit

Dim sHtml As String
Dim X As Integer 'will be used to count the form number
Dim i As Integer 'will be used to count the element number
Dim frm As Form
Dim Ctl As Control
Dim iCase As Integer


Private Sub Combo1_Click()
On Error Resume Next

Combo2.ListIndex = Combo1.ListIndex
Uarel = urlz(Combo2.ListIndex)
Command3_Click
End Sub

Private Sub Combo2_Click()
On Error Resume Next

Combo1.ListIndex = Combo2.ListIndex
Uarel = urlz(Combo1.ListIndex)
Command3_Click
End Sub

Private Sub Command1_Click()
On Error Resume Next


Text1.Text = ""
    iCase = 1
   'navigate to url that you enter in text box.
   'relevant code fires under Case 1 in WebBrowser1_DocumentComplete event
    WebBrowser1.Navigate "www.yahoo.com"
End Sub

Private Sub Command2_Click()
On Error Resume Next


Text1.Text = ""
    iCase = 2
   'http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=44242953&strAuthorName=dreamvb&txtMaxNumberOfEntriesPerPage=25
   'relevant code fires under Case 2 in WebBrowser1_DocumentComplete event
    'WebBrowser1.Navigate "http://visualbasicforum.com/search.php?"
    WebBrowser1.Navigate "http://babel.altavista.com/"
    'WebBrowser1.Navigate "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=44242953&strAuthorName=dreamvb&txtMaxNumberOfEntriesPerPage=25"
    'WebBrowser1.Navigate "http://actor.loquendo.com/actordemo/default.asp?language=en"
End Sub

Private Sub Command3_Click()
On Error Resume Next

Dim i As Long
Text1.Text = ""
    iCase = 3
   'navigate to google.
   'relevant code fires under Case 3 in WebBrowser1_DocumentComplete event
    WebBrowser1.Navigate "http://babel.altavista.com/"   '"www.google.com"
    FormX.WebBrowser1.Navigate "http://babel.altavista.com/"   '"www.google.com"
    'WebBrowser1.Navigate "http://actor.loquendo.com/actordemo/default.asp?language=en"
End Sub

Private Sub Command4_Click()
On Error Resume Next

    Text1.Text = ""
    iCase = 4
    WebBrowser1.Navigate App.Path & "\radioexample.htm"
End Sub

Private Sub Command5_Click()
On Error Resume Next

    Text1.Text = ""
    iCase = 5
    WebBrowser1.Navigate App.Path & "\radioexample.htm"
End Sub

Private Sub Command6_Click()
On Error Resume Next

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dim i As Integer
    ColorForm
    'Me.Show
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    WebBrowser1.Left = 0
    WebBrowser1.Top = 400
    WebBrowser1.Width = Me.Width - 80
    WebBrowser1.Height = Me.Height - 1000
    'Uarel = FormX.Combo11.List(0)
    Command3_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Unload frmMain
Unload Webrow
Set frmMain = Nothing
Set Webrow = Nothing
End Sub

Public Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Dim i As Long
Select Case iCase

    Case 1
        'put html into a string
        sHtml = WebBrowser1.Document.All.Item(0).innerHTML
        Text1.Text = Text1.Text & sHtml & vbCrLf
        Text1.Text = Text1.Text & "===================" & vbCrLf
        iCase = 0
    
    Case 2
        X = 0
        For X = 0 To WebBrowser1.Document.Forms.Length - 1
           'loop items in form
            For i = 0 To WebBrowser1.Document.Forms(X).Length - 1
              'print out the element number
                Text1.Text = Text1.Text & "Form Number: " & X & " Element Number: " & i & vbCrLf
              'print out the element name
                Text1.Text = Text1.Text & "Element Name: " & WebBrowser1.Document.Forms(X)(i).Name & vbCrLf
              'print out the element type
                Text1.Text = Text1.Text & "Element Type: " & WebBrowser1.Document.Forms(X)(i).Type & vbCrLf
              'print the value
                Text1.Text = Text1.Text & "Element Value: " & WebBrowser1.Document.Forms(X)(i).Value & vbCrLf
                Text1.Text = Text1.Text & vbCrLf
            Next i
            Text1.Text = Text1.Text & "===================" & vbCrLf
    
           DoEvents
        Next X
        iCase = 0
        
    Case 3
        With FormX.WebBrowser1.Document
            If Flagday = True Then
                Flagday = False
            Else
                If Trim(Uarel) <> "" Then
                    MsgBox "Translating this page: " & Uarel
                Else
                    MsgBox "No Webpage is open in Internet Explorer!!"
                    Unload Me
                End If
            End If
               
            .All("trurl").Value = Uarel
            .Forms(0)("lp").Value = Translatez '"zh_en"
            .Forms(1)("lp").Value = Translatez
            .Forms(1).submit.Click
            
                Text1.Text = Text1.Text & "Button Clicked" & vbCrLf
                Text1.Text = Text1.Text & "===================" & vbCrLf
        End With
        With WebBrowser1.Document
            'MsgBox WebBrowser1.Document.Forms(0)(5).Name
           'googles field name is q
            '.All("q").Value = "Visual basic forum"
            '.All("trtext").Value = "Fuck"
                'Text1.Text = Text1.Text & "Filled Form" & vbCrLf
            
            '.All("trurl").Value = "http://cgi.ebay.com/ws/eBayISAPI.dll?ViewItem&item=3858796988"
            .All("trurl").Value = Uarel   'urlz(0)
                'Text1.Text = Text1.Text & "Filled Form" & vbCrLf
            '.Forms(0)("lp").Value = "de_en"
            '.Forms(1)("lp").Value = "de_en"
            '.Forms(0)("lp").Value = "en_zh"
            '.Forms(1)("lp").Value = "en_zh"
            .Forms(0)("lp").Value = Translatez '"zh_en"
            .Forms(1)("lp").Value = Translatez

                'Text1.Text = Text1.Text & "Filled Form" & vbCrLf
           'one of googles button names is btnG
            '.All(WebBrowser1.Document.Forms(0)(4).Name).Click
            
            '.Forms(0).submit.Click
            .Forms(1).submit.Click
            
            '.All("btnG").Click
                Text1.Text = Text1.Text & "Button Clicked" & vbCrLf
                Text1.Text = Text1.Text & "===================" & vbCrLf
        End With
        
        FormX.WebBrowser4.Navigate Uarel
        FormX.Refresh
        FormX.Combo10.Refresh
        FormX.Combo11.Refresh
        FormX.SSTab1.Tab = 1
        FormX.SSTab1.Tab = 2
        iCase = 0
        Unload Me
    Case 4
        'When applying this to your own project, you may not always know the
        'form number, and the element number, so I suggest that you first
        'run case 2 from above on your url.  Then apply the form number
        'to the first (0) and the element in the second (0).
        WebBrowser1.Document.Forms(0)(0).Checked = True
        iCase = 0
        
    
    Case 5
        'When applying this to your own project, you may not always know the
        'form number, and the element number, so I suggest that you first
        'run case 2 from above on your url.  Then apply the form number
        'to the (0) and the element in the second (1).
        
        WebBrowser1.Document.Forms(0)(1).Checked = True
        iCase = 0

End Select
End Sub

Private Sub ColorForm()
On Error Resume Next

'Give color all the objects on the form

Dim Ctl As Control

For Each Ctl In Me
 If TypeOf Ctl Is Label Then
  With Ctl
  .BackColor = RGB(95, 95, 144)
  .ForeColor = vbWhite
  End With
 End If
Next

For Each Ctl In Me
 If TypeOf Ctl Is CommandButton Then
  With Ctl
  .BackColor = RGB(160, 160, 223)
  End With
 End If
Next

For Each Ctl In Me
 If TypeOf Ctl Is TextBox Then
  With Ctl
  .BackColor = vbWhite
  .ForeColor = RGB(95, 95, 144)
  End With
 End If
Next '


Me.BackColor = RGB(95, 95, 144)


End Sub

