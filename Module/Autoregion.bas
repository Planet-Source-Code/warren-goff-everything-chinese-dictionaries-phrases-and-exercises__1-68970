Attribute VB_Name = "Autoregion"
'Author£ºWXJ_Lake
'Email: webmaster@archtide.com
'Homepage£ºwww.archtide.com

'This sample will show you how to automatically make a ploygon form
'accroding to the picture property of the form.

'You may use these codes in your program or redistribute it on the web freely,
'Please be sure to mention the author name who made it.
'If want to distribute it in any other media besides web, please inform me, thanks!

'Note: The cDIBSection.cls included in this project comes from
'the excellent VB site - www.vbaccelerator.com
'We should all appreciate the great works they have done.


Option Explicit

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Const RGN_OR = 2
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SE_ERR_NOASSOC As Long = 31
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOPROGRESSUI = &H2
'Const SW_SHOWDEFAULT = 10

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Public Const GMEM_MOVEABLE = &H2
'Public Const GMEM_ZEROINIT = &H40

'Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RGNDATA) As Long
'Declare Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RGNDATA) As Long
'Type RGNDATA
'        rdh As RGNDATAHEADER
'        Buffer(1) As Rect
'End Type
'Type RGNDATAHEADER
'        dwSize As Long
'        iType As Long
'        nCount As Long
'        nRgnSize As Long
'        rcBound As Rect
'End Type

'declare for moving the form
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type
'API That will not send file to recycle bin
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_DELETE = &H3
Public Const FOF_SILENT = &H30

'for translucent effect in win2k, remove this if run in win9x or NT
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1


Global ThePicture As String
Global TheNumber As String
Global MicPath As String
Global PicPath As String
Global Model As String
Global Manufacturer As String
Global DirXsts As String
Global dirFolder As String
Global FileOperation As SHFILEOPSTRUCT
Global lReturn As Long
Global OKFlag As Boolean
Global I As Integer
Global j As Integer
Global MicNumber As Long
Global iMinutes As Integer
Global Tday, DDay As String
Global T1, T2, T3 As String
Global xx As String
Global yy As String
Global TickTock As Integer


Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
      
      Declare Function FindWindow _
       Lib "user32" Alias "FindWindowA" _
       (ByVal lpClassName As String, _
       ByVal lpWindowName As String) _
       As Long
 



Public Sub xListKillDupe(listbox As ComboBox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next
Next
End Sub
Public Sub xListKillDupes(listbox As listbox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next
Next
End Sub

Public Sub SetAutoRgn(hForm As Form, Optional transColor As Long = vbNull)
  Dim x As Long, y As Long
  Dim rgn1 As Long, rgn2 As Long
  Dim sPos As Long, EPos As Long
  Dim Wid As Long, Hgt As Long
  Dim xoff As Long, yoff As Long
  Dim DIB As New cDIBSection
  Dim bDib() As Byte
  Dim tSA As SAFEARRAY2D
  
  
    'get the picture size of the form
  DIB.CreateFromPicture hForm.Picture
  Wid = DIB.Width
  Hgt = DIB.Height
  
  With hForm
    .ScaleMode = vbPixels
    'compute the title bar's offset
    xoff = (.ScaleX(.Width, vbTwips, vbPixels) - .ScaleWidth) / 2
    yoff = .ScaleY(.Height, vbTwips, vbPixels) - .ScaleHeight - xoff
    'change the form size
    .Width = (Wid + xoff * 2) * Screen.TwipsPerPixelX
    .Height = (Hgt + xoff + yoff) * Screen.TwipsPerPixelY
  End With
  
  ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.BytesPerScanLine
        .pvData = DIB.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
' if there is no transColor specified, use the first pixel as the transparent color
  If transColor = vbNull Then transColor = RGB(bDib(0, 0), bDib(1, 0), bDib(2, 0))
  rgn1 = CreateRectRgn(0, 0, 0, 0)
    For y = 0 To Hgt - 1 'line scan
    x = -3
    Do
     x = x + 3
     
     While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) = transColor And (x < Wid * 3 - 3)
       x = x + 3 'skip the transparent point
     Wend
     sPos = x / 3
     While RGB(bDib(x, y), bDib(x + 1, y), bDib(x + 2, y)) <> transColor And (x < Wid * 3 - 3)
       x = x + 3 'skip the nontransparent point
     Wend
     EPos = x / 3
     
     'combine the region
     If sPos <= EPos Then
         rgn2 = CreateRectRgn(sPos + xoff, Hgt - y + yoff, EPos + xoff, Hgt - 1 - y + yoff)
         CombineRgn rgn1, rgn1, rgn2, RGN_OR
         DeleteObject rgn2
     End If
    Loop Until x >= Wid * 3 - 3
  Next y
  
  SetWindowRgn hForm.hWnd, rgn1, True  'set the final shap region
  DeleteObject rgn1
 
End Sub
