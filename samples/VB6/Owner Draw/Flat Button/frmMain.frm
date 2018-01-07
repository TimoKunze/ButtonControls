VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ButtonControls 1.10 - Flat Button sample"
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
   Begin BtnCtlsLibUCtl.OptionButton opt 
      Height          =   375
      Index           =   1
      Left            =   2700
      TabIndex        =   4
      Top             =   720
      Width           =   1695
      _cx             =   2990
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ContentType     =   0
      DisabledEvents  =   265
      DontRedraw      =   0   'False
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
      HAlignment      =   0
      HoverTime       =   1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      OptionMarkOnRight=   0   'False
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      Selected        =   0   'False
      Style           =   2
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":0000
      Text            =   "frmMain.frx":0038
   End
   Begin BtnCtlsLibUCtl.OptionButton opt 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      _cx             =   2990
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ContentType     =   0
      DisabledEvents  =   265
      DontRedraw      =   0   'False
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
      HAlignment      =   0
      HoverTime       =   1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      OptionMarkOnRight=   0   'False
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      Selected        =   0   'False
      Style           =   2
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":0074
      Text            =   "frmMain.frx":00AC
   End
   Begin BtnCtlsLibUCtl.CheckBox chk 
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _cx             =   2990
      _cy             =   661
      Appearance      =   0
      AutoToggleCheckMark=   0   'False
      BackColor       =   -2147483633
      BorderStyle     =   0
      CheckMarkOnRight=   0   'False
      ContentType     =   0
      DisabledEvents  =   265
      DontRedraw      =   0   'False
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
      HAlignment      =   0
      HoverTime       =   1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      PushLike        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionState  =   0
      Style           =   2
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      TriState        =   0   'False
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":00E8
      Text            =   "frmMain.frx":0120
   End
   Begin BtnCtlsLibUCtl.CommandButton cmd 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      _cx             =   2990
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
      Style           =   2
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":0154
      CommandLinkNote =   "frmMain.frx":018C
      Text            =   "frmMain.frx":01AC
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   1733
      TabIndex        =   0
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
      IconIndexes     =   "frmMain.frx":01EA
      CommandLinkNote =   "frmMain.frx":0222
      Text            =   "frmMain.frx":0242
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

  Private Type Size
    cx As Long
    cy As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
  Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


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


Private Sub chk_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Select Case chk.SelectionState
    Case SelectionStateConstants.ssUnchecked
      chk.SelectionState = SelectionStateConstants.ssChecked
    Case SelectionStateConstants.ssChecked
      chk.SelectionState = SelectionStateConstants.ssUnchecked
    Case SelectionStateConstants.ssIndeterminate
      chk.SelectionState = SelectionStateConstants.ssUnchecked
  End Select
End Sub

Private Sub chk_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  Const BDR_RAISEDINNER = &H4
  Const BDR_SUNKENINNER = &H8
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const DT_CALCRECT = &H400
  Const DT_CENTER = &H1
  Const DT_EDITCONTROL = &H2000
  Const DT_HIDEPREFIX = &H100000
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Dim cx As Long
  Dim cy As Long
  Dim darkClr As Long
  Dim flags As Long
  Dim hBMP As Long
  Dim hBMP_Old As Long
  Dim hBrush As Long
  Dim hDC_Mem As Long
  Dim lightClr As Long
  Dim raised As Boolean
  Dim rc As RECT
  Dim sunken As Boolean
  Dim txt As String

  If requiredAction And OwnerDrawActionConstants.odaDrawEntire Then
    ' draw background
    LSet rc = drawingRectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) And ((controlState And OwnerDrawControlStateConstants.odcsHot) = 0) And ((controlState And OwnerDrawControlStateConstants.odcsDisabled) = 0) Then
      darkClr = TranslateColor(SystemColorConstants.vb3DLight)
      lightClr = TranslateColor(SystemColorConstants.vb3DHighlight)

      hDC_Mem = CreateCompatibleDC(hDC)
      If hDC_Mem Then
        hBMP = CreateCompatibleBitmap(hDC, 2, 2)
        If hBMP Then
          hBMP_Old = SelectObject(hDC_Mem, hBMP)

          SetPixelV hDC_Mem, 0, 0, darkClr
          SetPixelV hDC_Mem, 1, 1, darkClr
          SetPixelV hDC_Mem, 0, 1, lightClr
          SetPixelV hDC_Mem, 1, 0, lightClr
          hBrush = CreatePatternBrush(hBMP)

          SelectObject hDC_Mem, hBMP_Old
          DeleteObject hBMP
        End If
        DeleteDC hDC_Mem
      End If
    Else
      hBrush = CreateSolidBrush(TranslateColor(chk.BackColor))
    End If
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw border
    If ((controlState And OwnerDrawControlStateConstants.odcsHot) = OwnerDrawControlStateConstants.odcsHot) Or ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) Then
      If ((controlState And OwnerDrawControlStateConstants.odcsPushed) = OwnerDrawControlStateConstants.odcsPushed) Or ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Or ((controlState And OwnerDrawControlStateConstants.odcsIndeterminate) = OwnerDrawControlStateConstants.odcsIndeterminate) Then
        sunken = True
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        raised = True
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    ElseIf ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Or ((controlState And OwnerDrawControlStateConstants.odcsIndeterminate) = OwnerDrawControlStateConstants.odcsIndeterminate) Then
      sunken = True
      DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
    End If

    ' draw text
    txt = chk.Text
    flags = DT_CENTER Or DT_VCENTER Or DT_EDITCONTROL
    If chk.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    If Not chk.MultiLine Then
      flags = flags Or DT_SINGLELINE
    End If
    If controlState And OwnerDrawControlStateConstants.odcsNoAccelerator Then
      flags = flags Or DT_HIDEPREFIX
    End If
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags Or DT_CALCRECT
    cx = rc.Right - rc.Left
    cy = rc.Bottom - rc.Top
    rc.Left = drawingRectangle.Left + (drawingRectangle.Right - drawingRectangle.Left - cx) / 2
    rc.Top = drawingRectangle.Top + (drawingRectangle.Bottom - drawingRectangle.Top - cy) / 2
    If sunken Then
      rc.Left = rc.Left + 1
      rc.Top = rc.Top + 1
    ElseIf raised Then
      'rc.Left = rc.Left - 1
      'rc.Top = rc.Top - 1
    End If
    rc.Right = rc.Left + cx
    rc.Bottom = rc.Top + cy

    SetBkMode hDC, TRANSPARENT
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      LSet rc = drawingRectangle
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If

  If requiredAction And OwnerDrawActionConstants.odaFocusChanged Then
    LSet rc = drawingRectangle
    ' ensure border is drawn
    If controlState And OwnerDrawControlStateConstants.odcsFocus Then
      If ((controlState And OwnerDrawControlStateConstants.odcsPushed) = OwnerDrawControlStateConstants.odcsPushed) Or ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Or ((controlState And OwnerDrawControlStateConstants.odcsIndeterminate) = OwnerDrawControlStateConstants.odcsIndeterminate) Then
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    End If

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If
End Sub

Private Sub cmd_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  Const BDR_RAISEDINNER = &H4
  Const BDR_SUNKENINNER = &H8
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const DT_CALCRECT = &H400
  Const DT_CENTER = &H1
  Const DT_EDITCONTROL = &H2000
  Const DT_HIDEPREFIX = &H100000
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Dim cx As Long
  Dim cy As Long
  Dim flags As Long
  Dim hBrush As Long
  Dim raised As Boolean
  Dim rc As RECT
  Dim sunken As Boolean
  Dim txt As String

  If requiredAction And OwnerDrawActionConstants.odaDrawEntire Then
    ' draw background
    LSet rc = drawingRectangle
    hBrush = CreateSolidBrush(TranslateColor(cmd.BackColor))
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw border
    If ((controlState And OwnerDrawControlStateConstants.odcsHot) = OwnerDrawControlStateConstants.odcsHot) Or ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) Then
      If controlState And OwnerDrawControlStateConstants.odcsPushed Then
        sunken = True
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        raised = True
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    End If

    ' draw text
    txt = cmd.Text
    flags = DT_CENTER Or DT_VCENTER Or DT_EDITCONTROL
    If cmd.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    If Not cmd.MultiLine Then
      flags = flags Or DT_SINGLELINE
    End If
    If controlState And OwnerDrawControlStateConstants.odcsNoAccelerator Then
      flags = flags Or DT_HIDEPREFIX
    End If
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags Or DT_CALCRECT
    cx = rc.Right - rc.Left
    cy = rc.Bottom - rc.Top
    rc.Left = drawingRectangle.Left + (drawingRectangle.Right - drawingRectangle.Left - cx) / 2
    rc.Top = drawingRectangle.Top + (drawingRectangle.Bottom - drawingRectangle.Top - cy) / 2
    If sunken Then
      rc.Left = rc.Left + 1
      rc.Top = rc.Top + 1
    ElseIf raised Then
      'rc.Left = rc.Left - 1
      'rc.Top = rc.Top - 1
    End If
    rc.Right = rc.Left + cx
    rc.Bottom = rc.Top + cy

    SetBkMode hDC, TRANSPARENT
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      LSet rc = drawingRectangle
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If

  If requiredAction And OwnerDrawActionConstants.odaFocusChanged Then
    LSet rc = drawingRectangle
    ' ensure border is drawn
    If controlState And OwnerDrawControlStateConstants.odcsFocus Then
      If ((controlState And OwnerDrawControlStateConstants.odcsPushed) = OwnerDrawControlStateConstants.odcsPushed) Or ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Then
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    End If

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If
End Sub

Private Sub cmdAbout_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  cmdAbout.About
End Sub

Private Sub Form_Initialize()
  Dim DLLVerData As DLLVERSIONINFO

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With
End Sub

Private Sub Form_Load()
  Dim combinedWidth As Long
  Dim cx As Long
  Dim cy As Long

  InstallKeyboardHook Me

  If bComctl32Version600OrNewer Then
    cmd.GetIdealSize cx, cy
    cx = cx * 1.3
  Else
    cx = cmd.Width
    cy = cmd.Height
  End If
  combinedWidth = cx * 2 + 10
  cmd.Move Me.ScaleLeft + (Me.ScaleWidth - combinedWidth) / 2, cmd.Top, cx, cy
  chk.Move cmd.Left + cx + 10, cmd.Top, cx, cy
  opt(0).Move cmd.Left, cmd.Top + cmd.Height + 8, cmd.Width, cmd.Height
  opt(1).Move chk.Left, opt(0).Top, chk.Width, chk.Height

  If bComctl32Version600OrNewer Then
    cmdAbout.GetIdealSize cx, cy
    cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cx * 2) / 2, cmdAbout.Top, cx * 2, cy * 1.2
  Else
    cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cmdAbout.Width) / 2, cmdAbout.Top
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub opt_Click(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  opt(index).Selected = True
  opt((index + 1) Mod 2).Selected = False
End Sub

Private Sub opt_OwnerDraw(index As Integer, ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  Const BDR_RAISEDINNER = &H4
  Const BDR_SUNKENINNER = &H8
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const DT_CALCRECT = &H400
  Const DT_CENTER = &H1
  Const DT_EDITCONTROL = &H2000
  Const DT_HIDEPREFIX = &H100000
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Dim cx As Long
  Dim cy As Long
  Dim darkClr As Long
  Dim flags As Long
  Dim hBMP As Long
  Dim hBMP_Old As Long
  Dim hBrush As Long
  Dim hDC_Mem As Long
  Dim lightClr As Long
  Dim raised As Boolean
  Dim rc As RECT
  Dim sunken As Boolean
  Dim txt As String

  If requiredAction And OwnerDrawActionConstants.odaDrawEntire Then
    ' draw background
    LSet rc = drawingRectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) And ((controlState And OwnerDrawControlStateConstants.odcsHot) = 0) And ((controlState And OwnerDrawControlStateConstants.odcsDisabled) = 0) Then
      darkClr = TranslateColor(SystemColorConstants.vb3DLight)
      lightClr = TranslateColor(SystemColorConstants.vb3DHighlight)

      hDC_Mem = CreateCompatibleDC(hDC)
      If hDC_Mem Then
        hBMP = CreateCompatibleBitmap(hDC, 2, 2)
        If hBMP Then
          hBMP_Old = SelectObject(hDC_Mem, hBMP)

          SetPixelV hDC_Mem, 0, 0, darkClr
          SetPixelV hDC_Mem, 1, 1, darkClr
          SetPixelV hDC_Mem, 0, 1, lightClr
          SetPixelV hDC_Mem, 1, 0, lightClr
          hBrush = CreatePatternBrush(hBMP)

          SelectObject hDC_Mem, hBMP_Old
          DeleteObject hBMP
        End If
        DeleteDC hDC_Mem
      End If
    Else
      hBrush = CreateSolidBrush(TranslateColor(opt(index).BackColor))
    End If
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw border
    If ((controlState And OwnerDrawControlStateConstants.odcsHot) = OwnerDrawControlStateConstants.odcsHot) Or ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) Then
      If ((controlState And OwnerDrawControlStateConstants.odcsPushed) = OwnerDrawControlStateConstants.odcsPushed) Or ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Then
        sunken = True
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        raised = True
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    ElseIf controlState And OwnerDrawControlStateConstants.odcsSelected Then
      sunken = True
      DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
    End If

    ' draw text
    txt = opt(index).Text
    flags = DT_CENTER Or DT_VCENTER Or DT_EDITCONTROL
    If opt(index).RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    If Not opt(index).MultiLine Then
      flags = flags Or DT_SINGLELINE
    End If
    If controlState And OwnerDrawControlStateConstants.odcsNoAccelerator Then
      flags = flags Or DT_HIDEPREFIX
    End If
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags Or DT_CALCRECT
    cx = rc.Right - rc.Left
    cy = rc.Bottom - rc.Top
    rc.Left = drawingRectangle.Left + (drawingRectangle.Right - drawingRectangle.Left - cx) / 2
    rc.Top = drawingRectangle.Top + (drawingRectangle.Bottom - drawingRectangle.Top - cy) / 2
    If sunken Then
      rc.Left = rc.Left + 1
      rc.Top = rc.Top + 1
    ElseIf raised Then
      'rc.Left = rc.Left - 1
      'rc.Top = rc.Top - 1
    End If
    rc.Right = rc.Left + cx
    rc.Bottom = rc.Top + cy

    SetBkMode hDC, TRANSPARENT
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      LSet rc = drawingRectangle
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If

  If requiredAction And OwnerDrawActionConstants.odaFocusChanged Then
    LSet rc = drawingRectangle
    ' ensure border is drawn
    If controlState And OwnerDrawControlStateConstants.odcsFocus Then
      If ((controlState And OwnerDrawControlStateConstants.odcsPushed) = OwnerDrawControlStateConstants.odcsPushed) Or ((controlState And OwnerDrawControlStateConstants.odcsSelected) = OwnerDrawControlStateConstants.odcsSelected) Then
        DrawEdge hDC, rc, BDR_SUNKENINNER, BF_RECT
      Else
        DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
      End If
    End If

    ' draw focus rectangle
    If ((controlState And OwnerDrawControlStateConstants.odcsFocus) = OwnerDrawControlStateConstants.odcsFocus) And ((controlState And OwnerDrawControlStateConstants.odcsNoFocusRectangle) = 0) Then
      InflateRect rc, -3, -3
      DrawFocusRect hDC, rc
    End If
  End If
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

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
  Const CLR_INVALID = &HFFFF&

  If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function
