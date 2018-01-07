VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ButtonControls 1.10 - Menu Button sample"
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
   Begin BtnCtlsLibUCtl.CommandButton cmdMenuButton 
      Height          =   420
      Left            =   1553
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _cx             =   2990
      _cy             =   741
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1033
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
      HoverTime       =   1
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
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   1733
      TabIndex        =   0
      Top             =   1245
      Width           =   1335
      _cx             =   2355
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
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
      IconIndexes     =   "frmMain.frx":008A
      CommandLinkNote =   "frmMain.frx":00C2
      Text            =   "frmMain.frx":00E2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type LOGFONT1
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
'      lfItalic As Byte
'      lfUnderline As Byte
'      lfStrikeOut As Byte
'      lfCharSet As Byte
    lf1 As Long
'      lfOutPrecision As Byte
'      lfClipPrecision As Byte
'      lfQuality As Byte
'      lfPitchAndFamily As Byte
    lf2 As Long
    lfFaceName(1 To 64) As Byte
  End Type

  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
    hBitmap As Long
  End Type

  Private Type POINTAPI
    x As Long
    y As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private hWebDingsFont As Long


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT1) As Long
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function ExtTextOutAsLong Lib "gdi32.dll" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByRef lpRect As RECT, ByVal lpString As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, ByVal lpvObject As Long) As Long
  Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function InflateRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByRef lptpm As TPMPARAMS) As Long


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

Private Sub cmdMenuButton_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_TYPE = &H10
  Const TPM_RIGHTALIGN = &H8&
  Const TPM_RETURNCMD = &H100&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim hMenu As Long
  Dim i As Long
  Dim miiData As MENUITEMINFO
  Dim ptMenu As POINTAPI
  Dim rcWindow As RECT
  Dim strMenuItem As String
  Dim tpmexParams As TPMPARAMS

  If Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels) < cmdMenuButton.Width - 22 Then
    MsgBox "You clicked the button."
    Exit Sub
  End If

  hMenu = CreatePopupMenu
  If hMenu Then
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID
      .fType = MFT_STRING
    End With
    For i = 1 To 5
      strMenuItem = "Entry &" & CStr(i)
      With miiData
        .wID = i
        .dwTypeData = StrPtr(strMenuItem)
        .cch = Len(strMenuItem)
      End With
      InsertMenuItem hMenu, i, 1, miiData
    Next i

    GetWindowRect cmdMenuButton.hWnd, rcWindow
    ptMenu.x = rcWindow.Right
    ptMenu.y = rcWindow.Bottom

    With tpmexParams
      .cbSize = LenB(tpmexParams)
      .rcExclude = rcWindow
      InflateRect .rcExclude, -1, -1
    End With
    i = TrackPopupMenuEx(hMenu, TPM_RIGHTALIGN Or TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, Me.hWnd, tpmexParams)
    DestroyMenu hMenu

    If i > 0 Then MsgBox "You clicked entry " & CStr(i) & "."
  End If
End Sub

Private Sub cmdMenuButton_CustomDraw(ByVal drawStage As BtnCtlsLibUCtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE, furtherProcessing As BtnCtlsLibUCtl.CustomDrawReturnValuesConstants)
  Const ARROW_CHAR As Integer = &H36
  Const COLOR_BTNHIGHLIGHT = 20
  Const COLOR_BTNSHADOW = 16
  Const COLOR_BTNTEXT = 18
  Const COLOR_GRAYTEXT = 17
  Const DT_LEFT = &H0
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Dim bDisabled As Boolean
  Dim hOldFont As Long
  Dim rcArrow As RECT
  Dim rcLine As RECT

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyPostPaint

    Case CustomDrawStageConstants.cdsPostPaint
      With rcLine
        .Left = drawingRectangle.Right - 22
        .Top = drawingRectangle.Top + 6
        .Right = drawingRectangle.Right - 20
        .Bottom = drawingRectangle.Bottom - 6
      End With
      bDisabled = ((controlState And (CustomDrawControlStateConstants.cdcsDisabled Or CustomDrawControlStateConstants.cdcsGrayed)) <> 0)
      Draw3DRect hDC, rcLine, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_BTNHIGHLIGHT)
      hOldFont = SelectObject(hDC, hWebDingsFont)
      With rcArrow
        .Left = drawingRectangle.Right - 18
        .Top = drawingRectangle.Top
        .Right = drawingRectangle.Right
        .Bottom = drawingRectangle.Bottom
      End With
      SetBkMode hDC, TRANSPARENT
      SetTextColor hDC, GetSysColor(IIf(bDisabled, COLOR_GRAYTEXT, COLOR_BTNTEXT))
      DrawText hDC, VarPtr(ARROW_CHAR), 1, rcArrow, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
      SelectObject hDC, hOldFont
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
  End Select
End Sub

Private Sub Form_Initialize()
  Const DEFAULT_GUI_FONT = 17
  Const SYMBOL_CHARSET = 2
  Dim DLLVerData As DLLVERSIONINFO
  Dim lf As LOGFONT1
  Dim hGUIFont As Long

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  If Not bComctl32Version600OrNewer Then
    MsgBox "This sample requires version 6.0 or newer of comctl32.dll.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If

  hGUIFont = GetStockObject(DEFAULT_GUI_FONT)
  If hGUIFont Then
    GetObjectAPI hGUIFont, LenB(lf), VarPtr(lf)
    lstrcpy VarPtr(lf.lfFaceName(1)), StrPtr("Webdings")
    'lf.lfCharSet = SYMBOL_CHARSET
    CopyMemory VarPtr(lf) + 23, VarPtr(SYMBOL_CHARSET), 1
    lf.lfHeight = 18
    hWebDingsFont = CreateFontIndirect(lf)
  End If
End Sub

Private Sub Form_Load()
  Dim cx As Long
  Dim cy As Long

  InstallKeyboardHook Me

  cmdAbout.GetIdealSize cx, cy
  cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cx * 2) / 2, cmdAbout.Top, cx * 2, cy * 1.2
  cmdMenuButton.GetIdealSize cx, cy
  cmdMenuButton.Move Me.ScaleLeft + (Me.ScaleWidth - cx * 2) / 2, cmdMenuButton.Top, cx * 2, cy * 1.2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
  If hWebDingsFont Then DeleteObject hWebDingsFont
End Sub


Private Sub Draw3DRect(ByVal hDC As Long, rc As RECT, ByVal clrTopLeft As Long, ByVal clrBottomRight As Long)
  Const ETO_OPAQUE = 2
  Dim clrOld As Long
  Dim rc2 As RECT

  clrOld = SetBkColor(hDC, clrTopLeft)
  rc2 = rc
  rc2.Right = rc.Right - 1
  rc2.Bottom = rc.Top + 1
  ExtTextOutAsLong hDC, 0, 0, ETO_OPAQUE, rc2, 0, 0, 0
  rc2.Right = rc.Left + 1
  rc2.Bottom = rc.Bottom - 1
  ExtTextOutAsLong hDC, 0, 0, ETO_OPAQUE, rc2, 0, 0, 0
  SetBkColor hDC, clrBottomRight
  rc2.Left = rc.Right
  rc2.Right = rc.Right - 1
  rc2.Bottom = rc.Bottom
  ExtTextOutAsLong hDC, 0, 0, ETO_OPAQUE, rc2, 0, 0, 0
  rc2.Left = rc.Left
  rc2.Top = rc.Bottom
  rc2.Right = rc.Right
  rc2.Bottom = rc.Bottom - 1
  ExtTextOutAsLong hDC, 0, 0, ETO_OPAQUE, rc2, 0, 0, 0
  SetBkColor hDC, clrOld
End Sub

' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
End Function
