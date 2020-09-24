Attribute VB_Name = "Globals"
Option Explicit

'Translucent Forms...
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'Prevents function recursion...
Global iRecursion As Boolean
Global tColor As Long
Global i1 As Long
Global Outer As String
Global Uarel As String
Global WebTitle As String
Global PinIncr As Integer
Global Simplified As String
Global Text As String
Global Texty As String
Global iz As Long
Global sSave As String
Global ret As Long
Global lpClassName As String
Global urlz(20) As String
Global Titlez(20) As String
Global Flagday As Boolean
Global Conversation(100) As String
Global mamama(1000) As String
Global GlobalFIlter As String
Global Scanned As Boolean
Global WavText As String
Global SavePath As String
Global LanguageV As String

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long


Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Const IDC_HAND = 32649&
    Public Const IDC_ARROW = 32512&
Sub Main()
On Error Resume Next
If App.PrevInstance Then End
Load FormX
FormX.Show

End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    Dim GetTickCount As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
        Loop
End Function

Public Sub DragForm(Who As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Public Sub MakeTranslucent(Who As Form, Optional tColor As Long) 'Was (Who as Object) before...

On Local Error Resume Next

Dim HW As Long
Dim HA As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If IsMissing(tColor) Or tColor = 0 Then
    tColor = RGB(0, 0, 200)
End If

Who.AutoRedraw = True
Who.Hide

DoEvents

HW = GetDesktopWindow()
HA = GetDC(HW)

'Get the Left, Top, Width and Height of the Form...
iLeft = Who.left / Screen.TwipsPerPixelX
iTop = Who.top / Screen.TwipsPerPixelY '+ 25    If using a form with a titlebar (border)...
iWidth = Who.ScaleWidth
iHeight = Who.ScaleHeight

'Now, Transfer the contents of the Desktop Window to the Form...
Call BitBlt(Who.hDC, 0, 0, iWidth, iHeight, HA, iLeft, iTop, SRCCOPY) 'iLeft + 4    If using a form with a titlebar (border)...

'Show...
Who.Picture = Who.Image
Who.Show

'Release the DC...
Call ReleaseDC(HW, HA)

'Add color...
Who.DrawMode = 9
Who.ForeColor = tColor
Who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub
