VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ButtonControls 1.10 - Collapsible Frame Sample"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   6720
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
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   2415
      _cx             =   4260
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   265
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   0
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
      Text            =   "frmMain.frx":006A
   End
   Begin BtnCtlsLibUCtl.Frame frmU 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _cx             =   6800
      _cy             =   3413
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      BorderVisible   =   -1  'True
      ContentType     =   0
      DisabledEvents  =   1
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
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      MousePointer    =   0
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      IconIndexes     =   "frmMain.frx":009C
      Text            =   "frmMain.frx":00D4
      Begin BtnCtlsLibUCtl.OptionButton optU 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
         _cx             =   2990
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   267
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
         HoverTime       =   -1
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
         RegisterForOLEDragDrop=   -1  'True
         RightToLeft     =   0
         Selected        =   0   'False
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmMain.frx":0118
         Text            =   "frmMain.frx":0150
      End
      Begin BtnCtlsLibUCtl.CommandButton cmdU 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ButtonType      =   0
         ContentType     =   0
         DisabledEvents  =   1291
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
         IconMarginLeft  =   5
         IconMarginRight =   0
         IconMarginTop   =   0
         KeepDropDownArrowAspectRatio=   -1  'True
         MousePointer    =   0
         MultiLine       =   -1  'True
         ProcessContextMenuKeys=   -1  'True
         Pushed          =   0   'False
         RegisterForOLEDragDrop=   -1  'True
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
         IconIndexes     =   "frmMain.frx":0182
         CommandLinkNote =   "frmMain.frx":01BA
         Text            =   "frmMain.frx":01F8
      End
      Begin BtnCtlsLibUCtl.CheckBox chkU 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _cx             =   2566
         _cy             =   450
         Appearance      =   0
         AutoToggleCheckMark=   -1  'True
         BackColor       =   -2147483633
         BorderStyle     =   0
         CheckMarkOnRight=   0   'False
         ContentType     =   0
         DisabledEvents  =   267
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
         HoverTime       =   -1
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
         RegisterForOLEDragDrop=   -1  'True
         RightToLeft     =   0
         SelectionState  =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         TriState        =   0   'False
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmMain.frx":022C
         Text            =   "frmMain.frx":0264
      End
      Begin BtnCtlsLibUCtl.OptionButton optU 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
         _cx             =   2990
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   267
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
         HoverTime       =   -1
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
         RegisterForOLEDragDrop=   -1  'True
         RightToLeft     =   0
         Selected        =   0   'False
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmMain.frx":0294
         Text            =   "frmMain.frx":02CC
      End
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


  Private collapsed As Boolean
  Private hImgLst_Minus As Long
  Private hImgLst_Plus As Long


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


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
  frmU.About
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim bComctl32Version600OrNewer As Boolean
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim iconPath As String

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

  hImgLst_Plus = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 1, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconPath = App.Path & "\..\res\Plus.ico"
  Else
    iconPath = App.Path & "\res\Plus.ico"
  End If
  hIcon = LoadImage(0, StrPtr(iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
  If hIcon Then
    ImageList_AddIcon hImgLst_Plus, hIcon
    DestroyIcon hIcon
  End If
  hImgLst_Minus = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 1, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconPath = App.Path & "\..\res\Minus.ico"
  Else
    iconPath = App.Path & "\res\Minus.ico"
  End If
  hIcon = LoadImage(0, StrPtr(iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
  If hIcon Then
    ImageList_AddIcon hImgLst_Minus, hIcon
    DestroyIcon hIcon
  End If
End Sub

Private Sub Form_Load()
  Dim cx As Long
  Dim cy As Long

  InstallKeyboardHook Me

  cmdU.GetIdealSize cx, cy
  cmdU.Move cmdU.Left, cmdU.Top, Me.ScaleX(cx * 1.5, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
  optU(0).GetIdealSize cx, cy
  optU(0).Move optU(0).Left, optU(0).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
  optU(1).GetIdealSize cx, cy
  optU(1).Move optU(1).Left, optU(1).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
  cmdAbout.GetIdealSize cx, cy
  cmdAbout.Move Text1.Left + Text1.Width + (Me.ScaleWidth - (Text1.Left + Text1.Width) - (cx * 2)) / 2, cmdAbout.Top, cx * 2, cy * 1.2

  ExpandFrame
End Sub

Private Sub Form_Terminate()
  If hImgLst_Minus Then ImageList_Destroy hImgLst_Minus
  If hImgLst_Plus Then ImageList_Destroy hImgLst_Plus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub frmU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  If (11 < x) And (x < 22) Then
    If (5 < y) And (y < 16) Then
      If collapsed Then
        ExpandFrame
      Else
        CollapseFrame
      End If
    End If
  End If
End Sub

Private Sub frmU_RecreatedControlWindow(ByVal hWnd As Long)
  If collapsed Then
    CollapseFrame
  Else
    ExpandFrame
  End If
End Sub


Private Sub CollapseFrame()
  frmU.Height = 20
  Text1.Move Text1.Left, frmU.Top + frmU.Height + 5, Text1.Width, Me.ScaleHeight - (frmU.Top + frmU.Height + 5) - 3
  frmU.hImageList = hImgLst_Plus
  collapsed = True
End Sub

Private Sub ExpandFrame()
  frmU.Height = 129
  Text1.Move Text1.Left, frmU.Top + frmU.Height + 5, Text1.Width, Me.ScaleHeight - (frmU.Top + frmU.Height + 5) - 3
  frmU.hImageList = hImgLst_Minus
  collapsed = False
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
