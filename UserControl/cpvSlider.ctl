VERSION 5.00
Begin VB.UserControl cpvSlider 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   35
   ToolboxBitmap   =   "cpvSlider.ctx":0000
   Begin VB.Image iRailPicture 
      Height          =   300
      Left            =   60
      Top             =   315
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Slider 
      Height          =   240
      Left            =   0
      Picture         =   "cpvSlider.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "cpvSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
' cpvSlider OCX v1.1
'
' Carles P.V. - 2001
' carles_pv@terra.es
'-------------------------------------------------------------------------------------------
' Last revision: 2003.02.22
'-------------------------------------------------------------------------------------------


Option Explicit

Private Declare Function DrawEdge Lib "USER32" (ByVal hdc As Long, pRect As RECT, ByVal lEdge As Long, ByVal grfFlags As Long) As Long

Private Const BDR_SUNKEN      As Long = &HA
Private Const BDR_RAISED      As Long = &H5
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4

Private Const BF_LEFT   As Long = &H1
Private Const BF_RIGHT  As Long = &H4
Private Const BF_TOP    As Long = &H2
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
                         
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
                         
Private Const HWND_TOP       As Long = 0
Private Const HWND_TOPMOST   As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE     As Long = &H1
Private Const SWP_NOMOVE     As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
                         
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'-- UC Types and Constants:

Private Type Point
    X As Single
    Y As Single
End Type

Public Enum sOrientationConstants
    [Horizontal]
    [Vertical]
End Enum

Public Enum sRailStyleConstants
    [Sunken]
    [Raised]
    [SunkenSoft]
    [RaisedSoft]
    [ByPicture] = 99
End Enum

'-- Private Variables:
Private SliderHooked As Boolean ' Slider hooked
Private SliderOffset As Point   ' Slider anchor point
Private RailRect     As RECT    ' Rail rectangle
Private AbsCount     As Long    ' AbsCount = Max - Min
Private LastValue    As Long    ' Last slider value
Private TPPx         As Long    ' TwipsPerPixelX
Private TPPy         As Long    ' TwipsPerPixelY

'-- Default Property Values:
Private Const m_def_Enabled = True
Private Const m_def_Orientation = 1     ' Vertical
Private Const m_def_RailStyle = 0       ' Sunken
Private Const m_def_ShowValueTip = True ' Show Tip
Private Const m_def_Min = 0             ' Min = 0
Private Const m_def_Max = 10            ' Max = 10
Private Const m_def_Value = 0           ' Value = 0

'-- Property Variables:
Private m_Enabled      As Boolean
Private m_Orientation  As sOrientationConstants
Private m_RailStyle    As sRailStyleConstants
Private m_ShowValueTip As Boolean
Private m_Min          As Long
Private m_Max          As Long
Private m_Value        As Long

'-- Event Declarations:
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event ValueChanged()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)



'-- Initialize

Private Sub UserControl_Initialize()
    
    TPPx = Screen.TwipsPerPixelX
    TPPy = Screen.TwipsPerPixelY
End Sub

'-- UserControl: InitProperties/ReadProperties/WriteProperties

Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_Orientation = m_def_Orientation
    m_RailStyle = m_def_RailStyle
    m_ShowValueTip = m_def_ShowValueTip
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    
    AbsCount = 10
    LastValue = m_Value
    ResetSlider
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_RailStyle = PropBag.ReadProperty("RailStyle", m_def_RailStyle)
    m_ShowValueTip = PropBag.ReadProperty("ShowValueTip", m_def_ShowValueTip)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    Set Slider.Picture = PropBag.ReadProperty("SliderIcon", Nothing)
    Set iRailPicture = PropBag.ReadProperty("RailPicture", Nothing)
    
    '-- Get absolute count and set Slider position
    AbsCount = m_Max - m_Min
    LastValue = m_Value
    Slider.Left = (m_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
    Slider.Top = (ScaleHeight - Slider.Height) - (m_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
        .WriteProperty "Enabled", m_Enabled, m_def_Enabled
        .WriteProperty "SliderIcon", Slider.Picture, Nothing
        .WriteProperty "Orientation", m_Orientation, m_def_Orientation
        .WriteProperty "RailPicture", iRailPicture, Nothing
        .WriteProperty "RailStyle", m_RailStyle, m_def_RailStyle
        .WriteProperty "ShowValueTip", m_ShowValueTip, m_def_ShowValueTip
        .WriteProperty "Min", m_Min, m_def_Min
        .WriteProperty "Max", m_Max, m_def_Max
        .WriteProperty "Value", m_Value, m_def_Value
    End With
End Sub


'-- UserControl draw

Private Sub UserControl_Show()
    '-- Draw control
    Refresh
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    '-- Resize control
    If (m_RailStyle = 99 And iRailPicture <> 0) Then
    
        Select Case m_Orientation
            
          Case 0 '-- Horizontal
            If (Slider.Height < iRailPicture.Height) Then
                Size (iRailPicture.Width + 4) * TPPx, iRailPicture.Height * TPPx
              Else
                Size (iRailPicture.Width + 4) * TPPx, Slider.Height * TPPx
            End If
            
          Case 1 '-- Vertical
            If (Slider.Width < iRailPicture.Width) Then
                Size iRailPicture.Width * TPPy, (iRailPicture.Height + 4) * TPPy
              Else
                Size Slider.Width * TPPy, (iRailPicture.Height + 4) * TPPy
            End If
        End Select
    
      Else
    
        Select Case m_Orientation
            
          Case 0 '-- Horizontal
            If (Width = 0) Then Width = Slider.Width * TPPx
            Height = Slider.Height * TPPy
                    
          Case 1 '-- Vertical
            If (Height = 0) Then Height = Slider.Height * TPPy
            Width = (Slider.Width) * TPPx
        End Select
    
    End If
    
    '-- Update slider position
    Select Case m_Orientation
    
      Case 0 '-- Horizontal
        If (Slider.Height < iRailPicture.Height And m_RailStyle = 99 And iRailPicture <> 0) Then
            Slider.Top = (iRailPicture.Height - Slider.Height) * 0.5
          Else
            Slider.Top = 0
        End If
        Slider.Left = (m_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
        
      Case 1 '-- Vertical
        If (Slider.Width < iRailPicture.Width And m_RailStyle = 99 And iRailPicture <> 0) Then
            Slider.Left = (iRailPicture.Width - Slider.Width) * 0.5
          Else
            Slider.Left = 0
        End If
        Slider.Top = ScaleHeight - Slider.Height - (m_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
    End Select
    
    '-- Define rail rectangle
    Select Case m_Orientation
        
      Case 0 '-- Horizontal
        With RailRect
            .Top = (Slider.Height - 4) * 0.5
            .Bottom = RailRect.Top + 4
            .Left = Slider.Width * 0.5 - 2
            .Right = RailRect.Left + ScaleWidth - Slider.Width + 4
        End With
                
      Case 1 '-- Vertical
        With RailRect
            .Top = Slider.Height * 0.5 - 2
            .Bottom = RailRect.Top + ScaleHeight - Slider.Height + 4
            .Left = (Slider.Width - 4) * 0.5
            .Right = RailRect.Left + 4
        End With
    End Select
    
    '-- Refresh control
    Refresh
    
    On Error GoTo 0
End Sub

Private Sub Refresh()
    
    '-- Clear control
    Cls
    
    '-- Draw rail...
    On Error Resume Next
    
    If (m_RailStyle = 99) Then
    
        Select Case m_Orientation
        
          Case 0 '-- Horizontal
            PaintPicture iRailPicture, 2, (ScaleHeight - iRailPicture.Height) * 0.5
                 
          Case 1 '-- Vertical
            PaintPicture iRailPicture, (ScaleWidth - iRailPicture.Width) * 0.5, 2
        End Select
        
      Else
        DrawEdge hdc, RailRect, Choose(m_RailStyle + 1, &HA, &H5, &H2, &H4, 0), BF_RECT
    End If
    
    '-- ...and slider
    PaintPicture Slider, Slider.Left, Slider.Top
    
    '-- Show value tip
    If (m_ShowValueTip And SliderHooked) Then
        ShowTip
    End If
    
    On Error GoTo 0
End Sub

'-- Scrolling...

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Me.Enabled) Then
    
        With Slider
            
            '-- Hook slider, get offsets and show tip
            If (Button = vbLeftButton) Then
               
                SliderHooked = True
                
                '-- Mouse over slider
                If (X >= .Left And X < .Left + .Width And Y >= .Top And Y < .Top + .Height) Then
                   
                    SliderOffset.X = X - .Left
                    SliderOffset.Y = Y - .Top
                
                '-- Mouse over rail
                  Else
                    SliderOffset.X = .Width / 2
                    SliderOffset.Y = .Height / 2
                    UserControl_MouseMove Button, Shift, X, Y
                End If
                
                '-- Show tip
                If (m_ShowValueTip) Then
                    ShowTip
                End If
                
                RaiseEvent MouseDown(Shift)
            End If
        End With
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (SliderHooked) Then
        
        '-- Check limits
        With Slider
            Select Case m_Orientation
            
              Case 0 '-- Horizontal
                If (X - SliderOffset.X < 0) Then
                    .Left = 0
                  ElseIf (X - SliderOffset.X > ScaleWidth - .Width) Then
                    .Left = ScaleWidth - .Width
                  Else
                    .Left = X - SliderOffset.X
                End If
            
              Case 1 '-- Vertical
                If (Y - SliderOffset.Y < 0) Then
                    .Top = 0
                  ElseIf (Y - SliderOffset.Y > ScaleHeight - .Height) Then
                    .Top = ScaleHeight - .Height
                  Else
                    .Top = Y - SliderOffset.Y
                End If
            End Select
        End With
        
        '-- Get value from Slider position
        Value = GetValue
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '-- Click event (If mouse over control area)
    If (X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight And Button = vbLeftButton) Then
        RaiseEvent Click
    End If
    
    '-- MouseUp event (Slider has been hooked)
    If (SliderHooked) Then
        RaiseEvent MouseUp(Shift)
    End If
    
    '-- Unhook slider and hide value tip
    SliderHooked = False
    Unload frmValueTip
End Sub

'-- Properties:

'-- Enabled
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
End Property

'-- BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Refresh
End Property

'-- Max
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)

    If (New_Max <= m_Min) Then Err.Raise 380
    
    m_Max = New_Max
    AbsCount = m_Max - m_Min
End Property

'-- Min
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)

    If (New_Min >= m_Max) Then Err.Raise 380
    
    m_Min = New_Min
    Value = New_Min
    AbsCount = m_Max - m_Min
End Property

'-- Value
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)

    If (New_Value < m_Min Or New_Value > m_Max) Then Err.Raise 380
    
    m_Value = New_Value
        
    If (m_Value <> LastValue) Then
        
        If (Not SliderHooked) Then
                   
            Select Case m_Orientation

              Case 0 '-- Horizontal
                Slider.Left = (New_Value - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
                
              Case 1 '-- Vertical
                Slider.Top = ScaleHeight - Slider.Height - (New_Value - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
            End Select
        End If
        
        Refresh
        LastValue = m_Value
        
        RaiseEvent ValueChanged
        If (m_Value = m_Max) Then RaiseEvent ArrivedLast
        If (m_Value = m_Min) Then RaiseEvent ArrivedFirst
    End If
End Property

'-- Orientation
Public Property Get Orientation() As sOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As sOrientationConstants)
    
    m_Orientation = New_Orientation
    
    ResetSlider
    UserControl_Resize
End Property

'-- RailStyle
Public Property Get RailStyle() As sRailStyleConstants
    RailStyle = m_RailStyle
End Property

Public Property Let RailStyle(ByVal New_RailStyle As sRailStyleConstants)

    m_RailStyle = New_RailStyle
    UserControl_Resize
End Property

'-- SliderIcon
Public Property Get SliderIcon() As Picture
    Set SliderIcon = Slider.Picture
End Property

Public Property Set SliderIcon(ByVal New_SliderIcon As Picture)

    Set Slider.Picture = New_SliderIcon
    UserControl_Resize
End Property

'-- RailPicture
Public Property Get RailPicture() As Picture
    Set RailPicture = iRailPicture.Picture
End Property

Public Property Set RailPicture(ByVal New_RailPicture As Picture)

    Set iRailPicture.Picture = New_RailPicture
    UserControl_Resize
End Property

'-- ShowValueTip
Public Property Get ShowValueTip() As Boolean
    ShowValueTip = m_ShowValueTip
End Property

Public Property Let ShowValueTip(ByVal New_ShowValueTip As Boolean)
    m_ShowValueTip = New_ShowValueTip
End Property


'-- Private:

'-- Get value from Slider position
Private Function GetValue() As Long
    
    On Error Resume Next
    
    Select Case m_Orientation
    
      Case 0 '-- Horizontal
        GetValue = Slider.Left / (ScaleWidth - Slider.Width) * AbsCount + m_Min
        Slider.Left = (GetValue - m_Min) * (ScaleWidth - Slider.Width) / AbsCount
        
      Case 1 '-- Vertical
        GetValue = (ScaleHeight - Slider.Height - Slider.Top) / (ScaleHeight - Slider.Height) * AbsCount + m_Min
        Slider.Top = ScaleHeight - Slider.Height - (GetValue - m_Min) * (ScaleHeight - Slider.Height) / AbsCount
    End Select
    
    On Error GoTo 0
End Function

'-- Reset slider position
Private Sub ResetSlider()

    Select Case m_Orientation
        
      Case 0 '-- Horizontal
        Slider.Move 0, 0
             
      Case 1 '-- Vertical
        Slider.Move 0, ScaleHeight - Slider.Height
    End Select
End Sub

'-- Show value tip
Private Sub ShowTip()
    
    Dim ucRect As RECT
    Dim X      As Long
    Dim Y      As Long

    On Error Resume Next
    
    GetWindowRect hWnd, ucRect
    
    With frmValueTip
    
        .lblTip.Width = .TextWidth(m_Value)
        .lblTip.Caption = m_Value
        .lblTip.Refresh
        
        Select Case m_Orientation
            
          Case 0 '-- Horizontal
            X = ucRect.Left + Slider.Left + (Slider.Width - .lblTip.Width - 4) * 0.5
            Y = ucRect.Top + Slider.Top - .lblTip.Height - 5
                 
          Case 1 '-- Vertical
            X = ucRect.Left + Slider.Left - .lblTip.Width - 6
            Y = ucRect.Top + Slider.Top + (Slider.Height - .lblTip.Height - 4) * 0.5
                 
        End Select
        
        '-- Set Tip position...
        .Move X * TPPx, Y * TPPy, (.lblTip.Width + 4) * TPPx, (.lblTip.Height + 3) * TPPy
        
        '-- ...and show it
        SetWindowPos .hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End With
    
    On Error GoTo 0
End Sub
