Attribute VB_Name = "Find"
Dim textfound As Integer    'used in subs as success indicator
Public foundit As Boolean   'lets the form know if search was successful
Public dontsave As Boolean  'stops you screwing files while you experiment
                            'remove if you like
'API that allows us to 'Float' the frmFind on top of the main
'form during searches
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
(ByVal hwndOwner As Long, _
ByVal nFolder As Long, pidl As Long) As Long
' Ret: 0=success

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
(pidl As Long, ByVal pszPath As String) As Long
Public Declare Function GlobalFree Lib "kernel32" _
(ByVal hMem As Long) As Long

Public Const MAX_PATH = 260
Public Function GetShellFolderPath(ByVal CSIDL As Long) As String
Dim pID As Long
Dim sTmp As String

If SHGetSpecialFolderLocation(0&, CSIDL, pID) = 0& Then
sTmp = String(MAX_PATH + 2, 0)
If SHGetPathFromIDList(ByVal pID, sTmp) <> 0& Then
    GetShellFolderPath = left$(sTmp, InStr(1, sTmp, vbNullChar) - 1)
End If
End If
If pID <> 0& Then GlobalFree pID
End Function
'API that allows us to 'Float' the frmFind on top of the main
'form during searches - got this code from PlanetSourceCode
'Even though its' pretty simple stuff, I'll lay claim to
'all other code.
Sub FloatWindow(x As Integer, action As Integer)
Dim wFlags As Integer, result As Integer
wFlags = SWP_NOMOVE Or SWP_NOSIZE
If action <> 0 Then
    Call SetWindowPos(x, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
Else
    Call SetWindowPos(x, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
End If

End Sub

Public Sub Findit(RTF As RichTextBox, Findme As String)
    RTF.Find (Findme)
    textfound = RTF.Find(Findme)
    If textfound <> -1 Then
        foundit = True 'lets the form know we found it
        RTF.SetFocus 'OK we've found it so select it to show the user
        Exit Sub 'lets' get out of here
    Else
        foundit = False 'lets the form know we cant find it
        Exit Sub
    End If

End Sub

Public Sub ReplaceitAll(RTF As RichTextBox, currentfile As String, Findme As String, Replaceme As String)
    
    RTF.SetFocus
    RTF.SelStart = 0
    RTF.Find (Findme) 'find it
    Do Until RTF.SelText = "" 'keep going till we cant find it
        RTF.SelText = Replaceme 'when we find it replace it
        RTF.Find (Findme), RTF.SelStart + Len(Findme) 'look for the next one
        If RTF.SelText = "" Then 'get out of the loop when done
            RTF.SelStart = 0
            RTF.SelLength = 0
            Exit Do
        End If
    Loop
    If dontsave = False Then RTF.SaveFile currentfile
    'set to true in normal operation or remove dontsave value
    Exit Sub

End Sub

Public Sub Replaceit(RTF As RichTextBox, currentfile As String, Findme As String, Replaceme As String)
If RTF.SelText = Findme Then RTF.SelText = Replaceme
If dontsave = False Then RTF.SaveFile currentfile
'put this in for thoroughness - but didn't bother calling it -
'did the same thing on the form itself
End Sub

Public Sub FinditNext(RTF As RichTextBox, Findme As String)
On Error Resume Next
'the text has already been found once. We want the next occurrence
  textfound = RTF.Find(Findme, RTF.SelStart + Len(Findme))
    If textfound <> -1 Then 'Found it !
        foundit = True 'let the form know we succeeded
        RTF.SetFocus 'show it to the user
        Exit Sub 'lets bail out now
    Else
        foundit = False 'let the form know we cant find it
        Exit Sub 'lets bail out now
    End If

End Sub

