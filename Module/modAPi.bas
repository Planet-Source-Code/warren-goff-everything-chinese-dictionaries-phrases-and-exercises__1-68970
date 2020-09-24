Attribute VB_Name = "modAPi"
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const WM_GETTEXT = &HD

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    'Dim sSave As String, Ret As Long, lpClassName As String
    lpClassName = Space(256)
    Ret = GetClassName(hWnd, lpClassName, 256)
    lpClassName = Trim(lpClassName)
    lpClassName = Left(lpClassName, Len(lpClassName) - 1)
    
    If (lpClassName <> "IEFrame") And (lpClassName <> "CabinetWClass") Then
        GoTo noie
    End If
    
    
    Ret = GetWindowTextLength(hWnd)
    sSave = Space(Ret)
    GetWindowText hWnd, sSave, Ret + 1
    
    If frmMain.Check1.Value = 1 Then
        frmMain.lstUrls.AddItem Str$(hWnd) & ":" & sSave & " URL: " & GetURLFrom(hWnd)
        If Left(LCase(Trim(GetURLFrom(hWnd))), 4) = "http" Then
            FormX.Combo11.AddItem GetURLFrom(hWnd)
            FormX.Combo10.AddItem sSave
            urlz(iz) = GetURLFrom(hWnd)
            Titlez(iz) = sSave
            iz = iz + 1
        End If
    Else
        frmMain.lstUrls.AddItem Str$(hWnd) & ":URL: " & GetURLFrom(hWnd)
        If Left(LCase(Trim(GetURLFrom(hWnd))), 4) = "http" Then
            FormX.Combo11.AddItem GetURLFrom(hWnd)
            FormX.Combo10.AddItem sSave
            urlz(iz) = GetURLFrom(hWnd)
            Titlez(iz) = sSave
            iz = iz + 1
        End If
    End If
noie:
    Uarel = urlz(0)

    EnumWindowsProc = True
    Load Webrow
    'Unload frmMain
    'Set frmMain = Nothing
'****************
    'Webrow.Show
End Function


Function GetURLFrom(hWnd As Long) As String
    Dim strCadena As String * 256

    hw1& = FindWindowEx(hWnd, 0&, "WorkerW", vbNullString)
    
    hw2& = FindWindowEx(hw1&, 0&, "ReBarWindow32", vbNullString)
    
    hw3& = FindWindowEx(hw2&, 0&, "ComboBoxEx32", vbNullString)
    
    hw4& = FindWindowEx(hw3&, 0&, "ComboBox", vbNullString)
    
    hw5& = FindWindowEx(hw4&, 0&, "Edit", vbNullString)
    
    SendMessageString hw5&, WM_GETTEXT, 256, strCadena
    
    GetURLFrom = strCadena

End Function
