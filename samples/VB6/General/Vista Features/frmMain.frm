VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ButtonControls 1.10 - Vista Features Sample"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   6270
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
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'Bildschirmmitte
   Begin BtnCtlsLibUCtl.CommandButton cmdElevation 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _cx             =   3625
      _cy             =   873
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
      ShowRightsElevationIcon=   -1  'True
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
   Begin BtnCtlsLibUCtl.CommandButton cmdSplit 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _cx             =   3625
      _cy             =   873
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   2
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
      IconIndexes     =   "frmMain.frx":00A0
      CommandLinkNote =   "frmMain.frx":00D8
      Text            =   "frmMain.frx":00F8
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdLink 
      Height          =   1455
      Left            =   128
      TabIndex        =   2
      Top             =   840
      Width           =   6015
      _cx             =   10610
      _cy             =   2566
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   1
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
      IconIndexes     =   "frmMain.frx":013A
      CommandLinkNote =   "frmMain.frx":0172
      Text            =   "frmMain.frx":023C
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   2228
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
      _cx             =   3201
      _cy             =   661
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
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
      IconIndexes     =   "frmMain.frx":0284
      CommandLinkNote =   "frmMain.frx":02BC
      Text            =   "frmMain.frx":02EE
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
    hbmpItem As Long
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


  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function InflateRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
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

Private Sub cmdElevation_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Const SW_SHOWDEFAULT = 10

  ShellExecute Me.hWnd, StrPtr("open"), StrPtr("hdwwiz.exe"), StrPtr(""), StrPtr(App.Path), SW_SHOWDEFAULT
End Sub

Private Sub cmdLink_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  MsgBox "Hello world!"
End Sub

Private Sub cmdSplit_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Const TPM_RIGHTALIGN = &H8&
  Const TPM_RETURNCMD = &H100&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim hPopupMenu As Long
  Dim menuItemID As Long
  Dim miiData As MENUITEMINFO
  Dim ptMenu As POINTAPI
  Dim rcWindow As RECT
  Dim tpmexParams As TPMPARAMS

  If Not cmdSplit.ShowSplitLine Then
    hPopupMenu = CreateDropDownMenu(cmdSplit.ShowSplitLine)
    GetWindowRect cmdSplit.hWnd, rcWindow
    ptMenu.x = rcWindow.Right
    ptMenu.y = rcWindow.Bottom

    With tpmexParams
      .cbSize = LenB(tpmexParams)
      .rcExclude = rcWindow
      InflateRect .rcExclude, -1, -1
    End With
    cmdSplit.Pushed = True
    menuItemID = TrackPopupMenuEx(hPopupMenu, TPM_RIGHTALIGN Or TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, Me.hWnd, tpmexParams)
    cmdSplit.Pushed = False
    If menuItemID = 1 Then
      cmdSplit.ShowSplitLine = Not cmdSplit.ShowSplitLine
    End If
    DestroyMenu hPopupMenu
  Else
    MsgBox "Hello world!"
  End If
End Sub

Private Sub cmdSplit_DropDown(buttonRectangle As BtnCtlsLibUCtl.RECTANGLE)
  Const TPM_RIGHTALIGN = &H8&
  Const TPM_RETURNCMD = &H100&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim hPopupMenu As Long
  Dim menuItemID As Long
  Dim miiData As MENUITEMINFO
  Dim ptMenu As POINTAPI
  Dim rcWindow As RECT
  Dim tpmexParams As TPMPARAMS

  hPopupMenu = CreateDropDownMenu(cmdSplit.ShowSplitLine)
  ptMenu.x = buttonRectangle.Right
  ptMenu.y = buttonRectangle.Bottom
  ClientToScreen cmdSplit.hWnd, ptMenu

  With tpmexParams
    .cbSize = LenB(tpmexParams)
    .rcExclude = rcWindow
    InflateRect .rcExclude, -1, -1
  End With
  cmdSplit.DropDownPushed = True
  menuItemID = TrackPopupMenuEx(hPopupMenu, TPM_RIGHTALIGN Or TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, Me.hWnd, tpmexParams)
  cmdSplit.DropDownPushed = False
  If menuItemID = 1 Then
    cmdSplit.ShowSplitLine = Not cmdSplit.ShowSplitLine
  End If
  DestroyMenu hPopupMenu
End Sub

Private Sub Form_Initialize()
  Dim bComctl32Version610OrNewer As Boolean
  Dim DLLVerData As DLLVERSIONINFO

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    If .dwMajor = 6 Then
      bComctl32Version610OrNewer = (.dwMinor >= 10)
    Else
      bComctl32Version610OrNewer = (.dwMajor > 6)
    End If
  End With

  If Not bComctl32Version610OrNewer Then
    MsgBox "This sample requires version 6.10 or newer of comctl32.dll.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If
End Sub

Private Sub Form_Load()
  Dim combinedWidth As Long
  Dim cx As Long
  Dim cy As Long
  Dim cx2 As Long

  InstallKeyboardHook Me

  cmdAbout.GetIdealSize cx, cy
  cmdAbout.Move Me.ScaleLeft + (Me.ScaleWidth - cx * 2) / 2, cmdAbout.Top, cx * 2, cy * 1.2

  cmdSplit.GetIdealSize cx, cy
  cmdElevation.GetIdealSize cx2
  cx2 = cx2 * 1.2
  combinedWidth = cx + cx2 + 10

  cmdSplit.Move Me.ScaleLeft + (Me.ScaleWidth - combinedWidth) / 2, cmdSplit.Top, cx, cy * 1.3
  cmdElevation.Move cmdSplit.Left + cx + 10, cmdSplit.Top, cx2, cmdSplit.Height

  cmdLink.GetIdealSize cx, cy
  ' doesn't really work for command links
  cx = cmdLink.Width
  cmdLink.Move cmdLink.Left, cmdLink.Top, cx, cy * 1.2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub


Private Function CreateDropDownMenu(ByVal checked As Boolean) As Long
  Const MF_CHECKED = &H8&
  Const MF_UNCHECKED = &H0&
  Const MFS_CHECKED = MF_CHECKED
  Const MFS_UNCHECKED = MF_UNCHECKED
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_STRING = &H40
  Dim hMenu As Long
  Dim miiData As MENUITEMINFO
  Dim strMenuItem As String

  hMenu = CreatePopupMenu
  If hMenu Then
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_ID Or MIIM_STATE Or MIIM_STRING
      strMenuItem = "&Split Line"
      .wID = 1
      .dwTypeData = StrPtr(strMenuItem)
      .fState = IIf(checked, MFS_CHECKED, MFS_UNCHECKED)
      .cch = Len(strMenuItem)
    End With
    InsertMenuItem hMenu, 1, 1, miiData
  End If
  CreateDropDownMenu = hMenu
End Function

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
