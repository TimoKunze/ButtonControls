VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ButtonControls 1.10 - Color Chooser sample"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
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
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'Bildschirmmitte
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   1733
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _cx             =   2355
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   265
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":0000
      CommandLinkNote =   "frmMain.frx":0038
      Text            =   "frmMain.frx":0058
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdColor 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _cx             =   2355
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1288
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   2
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":008A
      CommandLinkNote =   "frmMain.frx":00C2
      Text            =   "frmMain.frx":00E2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set BackColor:"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1050
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook


  Private Const ARROWHEIGHT = 2
  Private Const ARROWWIDTH = 4

  Private Const CLR_DEFAULT = &HFF000000


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type POINT
    x As Long
    y As Long
  End Type


  Private m_bColorPickerIsActive As Boolean
  Private m_bEnteredControl As Boolean

  Private bComctl32Version600OrNewer As Boolean
  Private hTheme As Long
  
  Private WithEvents m_oColorPicker As clsColorPicker
Attribute m_oColorPicker.VB_VarHelpID = -1
  Private m_selectedColor As Long


  Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qrc As RECTANGLE, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, lprc As RECTANGLE) As Long
  Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hDC As Long, lprc As RECTANGLE, ByVal uType As Long, ByVal uState As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, lpRect As RECTANGLE, ByVal uFormat As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECTANGLE, ByVal pClipRect As Long) As Long
  Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECTANGLE) As Long
  Private Declare Function ExtTextOut Lib "gdi32.dll" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTANGLE, ByVal lpString As Long, ByVal cbCount As Long, ByVal lpDx As Long) As Long
  Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lprc As RECTANGLE, ByVal hbr As Long) As Long
  Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal fnObject As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECTANGLE, pContentRect As RECTANGLE) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECTANGLE) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECTANGLE, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECTANGLE, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
  Private Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As Long, ByVal lpPoints As Long, ByVal nCount As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetPolyFillMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iPolyFillMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long


Private Function IHook_KeyboardProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_KeyboardProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const KF_ALTDOWN = &H2000
  Const UIS_CLEAR = 2
  Const UISF_HIDEACCEL = &H2
  Const UISF_HIDEFOCUS = &H1
  Const VK_DOWN = &H28
  Const VK_LEFT = &H25
  Const VK_RIGHT = &H27
  Const VK_TAB = &H9
  Const VK_UP = &H26
  Const WM_CHANGEUISTATE = &H127

  If HiWord(lParam) And KF_ALTDOWN Then
    ' this will make keyboard and focus cues work
    SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEACCEL Or UISF_HIDEFOCUS), 0
  Else
    Select Case wParam
      Case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_TAB
        ' this will make focus cues work
        SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS), 0
    End Select
  End If
End Function


Private Sub cmdAbout_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  cmdAbout.About
End Sub

Private Sub cmdColor_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  m_bEnteredControl = True
  cmdColor.Refresh
End Sub

Private Sub cmdColor_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  m_bEnteredControl = False
  cmdColor.Refresh
End Sub

Private Sub cmdColor_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim buttonRectangle As RECTANGLE
  Dim cancelled As Boolean
  Dim previousColor As Long

  ' save the current color for future reference
  previousColor = m_selectedColor

  ' display the popup
  m_oColorPicker.hWndOwner = cmdColor.hWnd
  m_oColorPicker.hWndParent = GetParent(cmdColor.hWnd)
  m_oColorPicker.SelectedColor = m_selectedColor
  'm_oColorPicker.TrackSelection = True

  GetWindowRect cmdColor.hWnd, buttonRectangle

  ' mark the button as active and invalidate button to force a redraw
  m_bColorPickerIsActive = True
  cmdColor.Refresh

  cancelled = Not m_oColorPicker.Popup(buttonRectangle)
  m_bColorPickerIsActive = False
  cmdColor.Refresh
  If cancelled Then
    ' if we are tracking, restore the old selection
    If m_oColorPicker.TrackSelection Then
      If previousColor <> m_selectedColor Then
        m_selectedColor = previousColor
        OnSelectionChanged m_selectedColor, True
      End If
    End If
  Else
    If previousColor <> m_oColorPicker.SelectedColor Then
      m_selectedColor = m_oColorPicker.SelectedColor
      OnSelectionChanged m_selectedColor, True
    End If
  End If

  cmdColor.Refresh
End Sub

Private Sub cmdColor_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  Const BLACK_BRUSH = 4
  Const BDR_SUNKENOUTER = &H2
  Const BDR_RAISEDINNER = &H4
  Const EDGE_ETCHED = BDR_SUNKENOUTER Or BDR_RAISEDINNER
  Const BF_RIGHT = &H4
  Const COLOR_GRAYTEXT = 17
  Const COLOR_BTNTEXT = 18
  Const DFC_BUTTON = 4
  Const DFCS_BUTTONPUSH = &H10
  Const DFCS_INACTIVE = &H100
  Const DFCS_PUSHED = &H200
  Const DFCS_ADJUSTRECT = &H2000
  Const DT_CENTER = &H1
  Const DT_VCENTER = &H4
  Const DT_SINGLELINE = &H20
  Const ETO_OPAQUE = &H2
  Const BP_PUSHBUTTON = 1
  Const PBS_HOT = 2
  Const PBS_PRESSED = 3
  Const PBS_DISABLED = 4
  Const PBS_DEFAULTED = 5
  Const SM_CXEDGE = 45
  Const SM_CYEDGE = 46
  Const Transparent = 1
  Dim arrowRectangle As RECTANGLE
  Dim cellText As String
  Dim frameState As Long
  Dim previousBackColor As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long

  If hTheme Then
    ' draw the outer edge
    If ((controlState And odcsSelected) = odcsSelected) Or m_bColorPickerIsActive Then
      frameState = frameState Or PBS_PRESSED
    End If
    If controlState And odcsDisabled Then
      frameState = frameState Or PBS_DISABLED
    End If
    If ((controlState And odcsHot) = odcsHot) Or m_bEnteredControl Then
      frameState = frameState Or PBS_HOT
    ElseIf ((controlState And odcsFocus) = odcsFocus) Or cmdColor.Default Then
      frameState = frameState Or PBS_DEFAULTED
    End If

    DrawThemeBackground hTheme, hDC, BP_PUSHBUTTON, frameState, drawingRectangle, 0
    GetThemeBackgroundContentRect hTheme, hDC, BP_PUSHBUTTON, frameState, drawingRectangle, drawingRectangle
  Else
    ' draw the outer edge
    frameState = DFCS_BUTTONPUSH Or DFCS_ADJUSTRECT
    If ((controlState And odcsSelected) = odcsSelected) Or m_bColorPickerIsActive Then
      frameState = frameState Or DFCS_PUSHED
    End If
    If controlState And odcsDisabled Then
      frameState = frameState Or DFCS_INACTIVE
    End If

    DrawFrameControl hDC, drawingRectangle, DFC_BUTTON, frameState

    ' adjust the position if we are selected (gives a 3d look)
    If ((controlState And odcsSelected) = odcsSelected) Or m_bColorPickerIsActive Then
      OffsetRect drawingRectangle, 1, 1
    End If
  End If

  ' draw the focus rectangle
  If ((controlState And odcsFocus) = odcsFocus) Or m_bColorPickerIsActive Then
    DrawFocusRect hDC, drawingRectangle
  End If
  InflateRect drawingRectangle, -GetSystemMetrics(SM_CXEDGE), -GetSystemMetrics(SM_CYEDGE)

  ' draw the arrow
  arrowRectangle.Left = drawingRectangle.Right - ARROWWIDTH - GetSystemMetrics(SM_CXEDGE) \ 2
  arrowRectangle.Top = (drawingRectangle.Bottom + drawingRectangle.Top) \ 2 - ARROWHEIGHT \ 2
  arrowRectangle.Right = arrowRectangle.Left + ARROWWIDTH
  arrowRectangle.Bottom = (drawingRectangle.Bottom + drawingRectangle.Top) \ 2 + ARROWHEIGHT \ 2

  DrawArrow hDC, arrowRectangle, 0, IIf(controlState And odcsDisabled, GetSysColor(COLOR_GRAYTEXT), GetSysColor(COLOR_BTNTEXT))
  drawingRectangle.Right = arrowRectangle.Left - GetSystemMetrics(SM_CXEDGE) \ 2

  ' draw the separator
  DrawEdge hDC, drawingRectangle, EDGE_ETCHED, BF_RIGHT
  drawingRectangle.Right = drawingRectangle.Right - (GetSystemMetrics(SM_CXEDGE) * 2) + 1

  ' draw the color rectangle
  If (controlState And odcsDisabled) = 0 Then
    If (m_selectedColor = CLR_DEFAULT) And (m_oColorPicker.DefaultColor = CLR_DEFAULT) Then
      cellText = m_oColorPicker.DefaultColorText
      If Len(cellText) > 0 Then
        If hTheme Then
          DrawThemeText hTheme, hDC, BP_PUSHBUTTON, frameState, StrPtr(cellText), -1, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, 0, drawingRectangle
        Else
          previousTextColor = SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT))
          previousBkMode = SetBkMode(hDC, Transparent)

          DrawText hDC, StrPtr(cellText), -1, drawingRectangle, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE

          SetBkMode hDC, previousBkMode
          SetTextColor hDC, previousTextColor
        End If
      End If
    Else
      previousBackColor = SetBkColor(hDC, IIf(m_selectedColor = CLR_DEFAULT, m_oColorPicker.DefaultColor, m_selectedColor))
      ExtTextOut hDC, 0, 0, ETO_OPAQUE, drawingRectangle, 0, 0, 0
      SetBkColor hDC, previousBackColor
    End If
    FrameRect hDC, drawingRectangle, GetStockObject(BLACK_BRUSH)
  End If
End Sub

Private Sub Form_Initialize()
  Dim DLLVerData As DLLVERSIONINFO

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  m_selectedColor = CLR_DEFAULT
End Sub

Private Sub Form_Load()
  Dim cx As Long
  Dim cy As Long

  InstallKeyboardHook Me

  If bComctl32Version600OrNewer Then
    cmdAbout.GetIdealSize cx, cy
    cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cx * 2) / 2, cmdAbout.Top, cx * 2, cy * 1.2
  Else
    cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cmdAbout.Width) / 2, cmdAbout.Top
  End If

  If bComctl32Version600OrNewer Then
    If IsAppThemed Then
      hTheme = OpenThemeData(cmdColor.hWnd, StrPtr("Button"))
    End If
  End If

  Set m_oColorPicker = New clsColorPicker
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_oColorPicker = Nothing
  If hTheme Then
    CloseThemeData hTheme
    hTheme = 0
  End If
  RemoveKeyboardHook Me
End Sub

Private Sub m_oColorPicker_HotTrackingColor(ByVal hotColor As Long, ByVal colorIsValid As Boolean)
  If colorIsValid Then
    m_selectedColor = hotColor
  End If
  cmdColor.Refresh
  OnSelectionChanged m_selectedColor, colorIsValid
End Sub


Private Sub DrawArrow(ByVal hDC As Long, boundingRectangle As RECTANGLE, ByVal direction As Long, ByVal color As Long)
  Const PS_SOLID = 0
  Const WINDING = 2
  Dim corners(0 To 2) As POINT
  Dim hArrowBrush As Long
  Dim hArrowPen As Long
  Dim hPreviousBrush As Long
  Dim hPreviousPen As Long
  Dim previousFillMode As Long

  Select Case direction
    Case 0:
      ' pointing downward
      corners(0).x = boundingRectangle.Left
      corners(0).y = boundingRectangle.Top
      corners(1).x = boundingRectangle.Right
      corners(1).y = boundingRectangle.Top
      corners(2).x = (boundingRectangle.Left + boundingRectangle.Right) \ 2
      corners(2).y = boundingRectangle.Bottom
    Case 1:
      ' pointing upward
      corners(0).x = boundingRectangle.Left
      corners(0).y = boundingRectangle.Bottom
      corners(1).x = boundingRectangle.Right
      corners(1).y = boundingRectangle.Bottom
      corners(2).x = (boundingRectangle.Left + boundingRectangle.Right) \ 2
      corners(2).y = boundingRectangle.Top
    Case 2:
      ' pointing to the left
      corners(0).x = boundingRectangle.Right
      corners(0).y = boundingRectangle.Top
      corners(1).x = boundingRectangle.Right
      corners(1).y = boundingRectangle.Bottom
      corners(2).x = boundingRectangle.Left
      corners(2).y = (boundingRectangle.Top + boundingRectangle.Bottom) \ 2
    Case 3:
      ' pointing to the right
      corners(0).x = boundingRectangle.Left
      corners(0).y = boundingRectangle.Top
      corners(1).x = boundingRectangle.Left
      corners(1).y = boundingRectangle.Bottom
      corners(2).x = boundingRectangle.Right
      corners(2).y = (boundingRectangle.Top + boundingRectangle.Bottom) \ 2
  End Select

  hArrowBrush = CreateSolidBrush(color)
  hArrowPen = CreatePen(PS_SOLID, 0, color)

  hPreviousBrush = SelectObject(hDC, hArrowBrush)
  hPreviousPen = SelectObject(hDC, hArrowPen)

  previousFillMode = SetPolyFillMode(hDC, WINDING)
  Polygon hDC, VarPtr(corners(0)), 3

  SetPolyFillMode hDC, previousFillMode
  SelectObject hDC, hPreviousBrush
  SelectObject hDC, hPreviousPen
  DeleteObject hArrowBrush
  DeleteObject hArrowPen
End Sub

Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

Private Function LoWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)

  LoWord = ret
End Function

Private Function MakeDWord(ByVal Lo As Integer, ByVal Hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

  MakeDWord = ret
End Function

Private Sub OnSelectionChanged(ByVal color As Long, ByVal colorIsValid As Boolean)
  If color = CLR_DEFAULT Then
    Me.BackColor = vbButtonFace
  Else
    Me.BackColor = color
  End If
End Sub
