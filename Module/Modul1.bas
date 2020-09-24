Attribute VB_Name = "Modul1"
Option Explicit
Global Pinyin As String
Global Translatez As String

Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const DT_CALCRECT = &H400
Public Const SM_CXVSCROLL = 2

Public Type RECT
   left As Long
   top As Long
   right As Long
   Bottom As Long
End Type

Public Declare Function DrawText Lib "User32" _
   Alias "DrawTextA" _
  (ByVal hDC As Long, _
   ByVal lpStr As String, _
   ByVal nCount As Long, _
   lpRect As RECT, ByVal _
   wFormat As Long) As Long
   
Public Declare Function GetSystemMetrics Lib "User32" _
  (ByVal nIndex As Long) As Long

Public Declare Function SendMessage Lib "User32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'declare for moving the form
Public Declare Function ReleaseCapture Lib "User32" () As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1



'Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
      
      Declare Function FindWindow _
       Lib "User32" Alias "FindWindowA" _
       (ByVal lpClassName As String, _
       ByVal lpWindowName As String) _
       As Long

      

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 If Topmost = True Then 'Make the window topmost
  SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
 Else
  SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
  SetTopMostWindow = False
 End If
End Function


