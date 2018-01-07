VERSION 5.00
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.10#0"; "BtnCtlsU.ocx"
Object = "{A38BC98C-3D41-44DB-B27B-CF19BAD8AE53}#1.10#0"; "BtnCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ButtonControls 1.10 - Events Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9480
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'Bildschirmmitte
   Begin BtnCtlsLibUCtl.CommandButton cmdAbout 
      Default         =   -1  'True
      Height          =   375
      Left            =   6030
      TabIndex        =   12
      Top             =   5280
      Width           =   2415
      _cx             =   4260
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
      IconIndexes     =   "frmMain.frx":0000
      CommandLinkNote =   "frmMain.frx":0038
      Text            =   "frmMain.frx":0058
   End
   Begin BtnCtlsLibUCtl.CheckBox chkLog 
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   5400
      Width           =   735
      _cx             =   1296
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
      RegisterForOLEDragDrop=   0   'False
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
      IconIndexes     =   "frmMain.frx":008A
      Text            =   "frmMain.frx":00C2
   End
   Begin BtnCtlsLibACtl.Frame frmA 
      Height          =   3015
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   4815
      _cx             =   8493
      _cy             =   5318
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      BorderVisible   =   -1  'True
      ContentType     =   0
      DisabledEvents  =   0
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
      IconIndexes     =   "frmMain.frx":00EA
      Text            =   "frmMain.frx":0122
      Begin BtnCtlsLibACtl.OptionButton optA 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
         _cx             =   2990
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":014C
         Text            =   "frmMain.frx":0184
      End
      Begin BtnCtlsLibACtl.OptionButton optA 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
         _cx             =   2990
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":01C0
         Text            =   "frmMain.frx":01F8
      End
      Begin BtnCtlsLibACtl.CommandButton cmdA 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ButtonType      =   0
         ContentType     =   0
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":0234
         CommandLinkNote =   "frmMain.frx":026C
         Text            =   "frmMain.frx":028C
      End
      Begin BtnCtlsLibACtl.CheckBox chkA 
         Height          =   255
         Left            =   240
         TabIndex        =   6
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
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":02CA
         Text            =   "frmMain.frx":0302
      End
   End
   Begin BtnCtlsLibUCtl.Frame frmU 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _cx             =   8493
      _cy             =   6165
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      BorderVisible   =   -1  'True
      ContentType     =   0
      DisabledEvents  =   0
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
      IconIndexes     =   "frmMain.frx":0336
      Text            =   "frmMain.frx":036E
      Begin BtnCtlsLibUCtl.OptionButton optU 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
         _cx             =   2990
         _cy             =   450
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ContentType     =   0
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":039E
         Text            =   "frmMain.frx":03D6
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
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":0412
         CommandLinkNote =   "frmMain.frx":044A
         Text            =   "frmMain.frx":046A
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
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":04A8
         Text            =   "frmMain.frx":04E0
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
         DisabledEvents  =   0
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
         IconIndexes     =   "frmMain.frx":0514
         Text            =   "frmMain.frx":054C
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      Top             =   120
      Width           =   4455
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


  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private hImgLst As Long
  Private objActiveCtl As Object


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


Private Sub chkA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_CustomDraw(ByVal drawStage As BtnCtlsLibACtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibACtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE, furtherProcessing As BtnCtlsLibACtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "chkA_CustomDraw: drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub chkA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "chkA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub chkA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "chkA_DragDrop"
End Sub

Private Sub chkA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "chkA_DragOver"
End Sub

Private Sub chkA_GotFocus()
  AddLogEntry "chkA_GotFocus"
  Set objActiveCtl = chkA
End Sub

Private Sub chkA_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "chkA_KeyDown: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub chkA_KeyPress(keyAscii As Integer)
  AddLogEntry "chkA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub chkA_KeyUp(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "chkA_KeyUp: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub chkA_LostFocus()
  AddLogEntry "chkA_LostFocus"
End Sub

Private Sub chkA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "chkA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_OLEDragDrop(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    chkA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub chkA_OLEDragEnter(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkA_OLEDragLeave(ByVal data As BtnCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkA_OLEDragMouseMove(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkA_OwnerDraw(ByVal requiredAction As BtnCtlsLibACtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibACtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE)
  AddLogEntry "chkA_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub chkA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "chkA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub chkA_ResizedControlWindow()
  AddLogEntry "chkA_ResizedControlWindow"
End Sub

Private Sub chkA_SelectionStateChanged(ByVal previousSelectionState As BtnCtlsLibACtl.SelectionStateConstants, ByVal newSelectionState As BtnCtlsLibACtl.SelectionStateConstants)
  AddLogEntry "chkA_SelectionStateChanged: previousSelectionState=" & previousSelectionState & ", newSelectionState=" & newSelectionState
End Sub

Private Sub chkA_Validate(Cancel As Boolean)
  AddLogEntry "chkA_Validate"
End Sub

Private Sub chkA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkLog_SelectionStateChanged(ByVal previousSelectionState As BtnCtlsLibUCtl.SelectionStateConstants, ByVal newSelectionState As BtnCtlsLibUCtl.SelectionStateConstants)
  bLog = (newSelectionState = BtnCtlsLibUCtl.SelectionStateConstants.ssChecked)
End Sub

Private Sub chkU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_CustomDraw(ByVal drawStage As BtnCtlsLibUCtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE, furtherProcessing As BtnCtlsLibUCtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "chkU_CustomDraw: drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub chkU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "chkU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub chkU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "chkU_DragDrop"
End Sub

Private Sub chkU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "chkU_DragOver"
End Sub

Private Sub chkU_GotFocus()
  AddLogEntry "chkU_GotFocus"
  Set objActiveCtl = chkU
End Sub

Private Sub chkU_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "chkU_KeyDown: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub chkU_KeyPress(keyAscii As Integer)
  AddLogEntry "chkU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub chkU_KeyUp(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "chkU_KeyUp: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub chkU_LostFocus()
  AddLogEntry "chkU_LostFocus"
End Sub

Private Sub chkU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "chkU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_OLEDragDrop(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    chkU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub chkU_OLEDragEnter(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkU_OLEDragLeave(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkU_OLEDragMouseMove(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "chkU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub chkU_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  AddLogEntry "chkU_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub chkU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "chkU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub chkU_ResizedControlWindow()
  AddLogEntry "chkU_ResizedControlWindow"
End Sub

Private Sub chkU_SelectionStateChanged(ByVal previousSelectionState As BtnCtlsLibUCtl.SelectionStateConstants, ByVal newSelectionState As BtnCtlsLibUCtl.SelectionStateConstants)
  AddLogEntry "chkU_SelectionStateChanged: previousSelectionState=" & previousSelectionState & ", newSelectionState=" & newSelectionState
End Sub

Private Sub chkU_Validate(Cancel As Boolean)
  AddLogEntry "chkU_Validate"
End Sub

Private Sub chkU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub chkU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "chkU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_CustomDraw(ByVal drawStage As BtnCtlsLibACtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibACtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE, furtherProcessing As BtnCtlsLibACtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "cmdA_CustomDraw: drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub cmdA_CustomDropDownAreaSize(clientRectangle As BtnCtlsLibACtl.RECTANGLE, commandButtonAreaRectangle As BtnCtlsLibACtl.RECTANGLE, dropDownAreaRectangle As BtnCtlsLibACtl.RECTANGLE, furtherProcessing As BtnCtlsLibACtl.CustomDropDownAreaSizeReturnValuesConstants)
  'AddLogEntry "cmdA_CustomDropDownAreaSize: clientRectangle=(" & clientRectangle.Left & ", " & clientRectangle.Top & ")-(" & clientRectangle.Right & ", " & clientRectangle.Bottom & "), commandButtonAreaRectangle=(" & commandButtonAreaRectangle.Left & ", " & commandButtonAreaRectangle.Top & ")-(" & commandButtonAreaRectangle.Right & ", " & commandButtonAreaRectangle.Bottom & "), dropDownAreaRectangle=(" & dropDownAreaRectangle.Left & ", " & dropDownAreaRectangle.Top & ")-(" & dropDownAreaRectangle.Right & ", " & dropDownAreaRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub cmdA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "cmdA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub cmdA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "cmdA_DragDrop"
End Sub

Private Sub cmdA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "cmdA_DragOver"
End Sub

Private Sub cmdA_DropDown(buttonRectangle As BtnCtlsLibACtl.RECTANGLE)
  AddLogEntry "cmdA_DropDown: buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & ")"
End Sub

Private Sub cmdA_GotFocus()
  AddLogEntry "cmdA_GotFocus"
  Set objActiveCtl = cmdA
End Sub

Private Sub cmdA_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "cmdA_KeyDown: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub cmdA_KeyPress(keyAscii As Integer)
  AddLogEntry "cmdA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub cmdA_KeyUp(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "cmdA_KeyUp: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub cmdA_LostFocus()
  AddLogEntry "cmdA_LostFocus"
End Sub

Private Sub cmdA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "cmdA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_OLEDragDrop(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    cmdA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub cmdA_OLEDragEnter(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdA_OLEDragLeave(ByVal data As BtnCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdA_OLEDragMouseMove(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdA_OwnerDraw(ByVal requiredAction As BtnCtlsLibACtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibACtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE)
  AddLogEntry "cmdA_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub cmdA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "cmdA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub cmdA_ResizedControlWindow()
  AddLogEntry "cmdA_ResizedControlWindow"
End Sub

Private Sub cmdA_Validate(Cancel As Boolean)
  AddLogEntry "cmdA_Validate"
End Sub

Private Sub cmdA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdAbout_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  objActiveCtl.About
End Sub

Private Sub cmdU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_CustomDraw(ByVal drawStage As BtnCtlsLibUCtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE, furtherProcessing As BtnCtlsLibUCtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "cmdU_CustomDraw: drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub cmdU_CustomDropDownAreaSize(clientRectangle As BtnCtlsLibUCtl.RECTANGLE, commandButtonAreaRectangle As BtnCtlsLibUCtl.RECTANGLE, dropDownAreaRectangle As BtnCtlsLibUCtl.RECTANGLE, furtherProcessing As BtnCtlsLibUCtl.CustomDropDownAreaSizeReturnValuesConstants)
  'AddLogEntry "cmdU_CustomDropDownAreaSize: clientRectangle=(" & clientRectangle.Left & ", " & clientRectangle.Top & ")-(" & clientRectangle.Right & ", " & clientRectangle.Bottom & "), commandButtonAreaRectangle=(" & commandButtonAreaRectangle.Left & ", " & commandButtonAreaRectangle.Top & ")-(" & commandButtonAreaRectangle.Right & ", " & commandButtonAreaRectangle.Bottom & "), dropDownAreaRectangle=(" & dropDownAreaRectangle.Left & ", " & dropDownAreaRectangle.Top & ")-(" & dropDownAreaRectangle.Right & ", " & dropDownAreaRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub cmdU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "cmdU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub cmdU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "cmdU_DragDrop"
End Sub

Private Sub cmdU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "cmdU_DragOver"
End Sub

Private Sub cmdU_DropDown(buttonRectangle As BtnCtlsLibUCtl.RECTANGLE)
  AddLogEntry "cmdU_DropDown: buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & ")"
End Sub

Private Sub cmdU_GotFocus()
  AddLogEntry "cmdU_GotFocus"
  Set objActiveCtl = cmdU
End Sub

Private Sub cmdU_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "cmdU_KeyDown: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub cmdU_KeyPress(keyAscii As Integer)
  AddLogEntry "cmdU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub cmdU_KeyUp(KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "cmdU_KeyUp: keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub cmdU_LostFocus()
  AddLogEntry "cmdU_LostFocus"
End Sub

Private Sub cmdU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "cmdU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_OLEDragDrop(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    cmdU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub cmdU_OLEDragEnter(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdU_OLEDragLeave(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdU_OLEDragMouseMove(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "cmdU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub cmdU_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  AddLogEntry "cmdU_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub cmdU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "cmdU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub cmdU_ResizedControlWindow()
  AddLogEntry "cmdU_ResizedControlWindow"
End Sub

Private Sub cmdU_Validate(Cancel As Boolean)
  AddLogEntry "cmdU_Validate"
End Sub

Private Sub cmdU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub cmdU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "cmdU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim iconPath As String

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 1, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconPath = App.Path & "\..\res\Recycle Bin.ico"
  Else
    iconPath = App.Path & "\res\Recycle Bin.ico"
  End If
  hIcon = LoadImage(0, StrPtr(iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
  If hIcon Then
    ImageList_AddIcon hImgLst, hIcon
    DestroyIcon hIcon
  End If
End Sub

Private Sub Form_Load()
  Dim cx As Long
  Dim cy As Long

  chkLog.SelectionState = BtnCtlsLibUCtl.SelectionStateConstants.ssChecked
  bLog = (chkLog.SelectionState = BtnCtlsLibUCtl.SelectionStateConstants.ssChecked)

  InstallKeyboardHook Me

  If bComctl32Version600OrNewer Then
    chkA.hImageList = hImgLst
    chkU.hImageList = hImgLst
    cmdA.hImageList = hImgLst
    cmdU.hImageList = hImgLst
    frmA.hImageList = hImgLst
    frmU.hImageList = hImgLst
    optA(0).hImageList = hImgLst
    optA(1).hImageList = hImgLst
    optU(0).hImageList = hImgLst
    optU(1).hImageList = hImgLst

    chkA.GetIdealSize cx, cy
    chkA.Move chkA.Left, chkA.Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    chkU.GetIdealSize cx, cy
    chkU.Move chkU.Left, chkU.Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    cmdA.GetIdealSize cx, cy
    cmdA.Move cmdA.Left, cmdA.Top, Me.ScaleX(cx * 1.1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy * 1.1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    cmdU.GetIdealSize cx, cy
    cmdU.Move cmdU.Left, cmdU.Top, Me.ScaleX(cx * 1.1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy * 1.1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    optA(0).GetIdealSize cx, cy
    optA(0).Move optA(0).Left, optA(0).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    optA(1).GetIdealSize cx, cy
    optA(1).Move optA(1).Left, optA(1).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    optU(0).GetIdealSize cx, cy
    optU(0).Move optU(0).Left, optU(0).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    optU(1).GetIdealSize cx, cy
    optU(1).Move optU(1).Left, optU(1).Top, Me.ScaleX(cx, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips), Me.ScaleY(cy, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    chkLog.GetIdealSize cx, cy
    chkLog.Move chkLog.Left, chkLog.Top, cx, cy
    cmdAbout.GetIdealSize cx, cy
    cmdAbout.Move cmdAbout.Left, cmdAbout.Top, cx * 2, cy * 1.2
  End If
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    frmU.Move 3, 3, frmU.Width, (Me.ScaleHeight - 11) / 2
    frmA.Move 3, frmU.Top + frmU.Height + 5, frmU.Width, frmU.Height
    txtLog.Move frmU.Left + frmU.Width + 5, 0, Me.ScaleWidth - frmU.Width - 8, Me.ScaleHeight - cmdAbout.Height - 10
    chkLog.Move txtLog.Left, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - chkLog.Height) / 2
    cmdAbout.Move chkLog.Left + chkLog.Width + ((txtLog.Width - chkLog.Width) - cmdAbout.Width) / 2, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - cmdAbout.Height) / 2
  End If
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub frmA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "frmA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub frmA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "frmA_DragDrop"
End Sub

Private Sub frmA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "frmA_DragOver"
End Sub

Private Sub frmA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "frmA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_OLEDragDrop(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    frmA.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub frmA_OLEDragEnter(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmA_OLEDragLeave(ByVal data As BtnCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmA_OLEDragMouseMove(ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmA_OwnerDraw(ByVal requiredAction As BtnCtlsLibACtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibACtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE)
  AddLogEntry "frmA_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub frmA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "frmA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub frmA_ResizedControlWindow()
  AddLogEntry "frmA_ResizedControlWindow"
End Sub

Private Sub frmA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "frmU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub frmU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "frmU_DragDrop"
End Sub

Private Sub frmU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "frmU_DragOver"
End Sub

Private Sub frmU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "frmU_MouseMove: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub frmU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_OLEDragDrop(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    frmU.FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub frmU_OLEDragEnter(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmU_OLEDragLeave(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmU_OLEDragMouseMove(ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "frmU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub frmU_OwnerDraw(ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  AddLogEntry "frmU_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub frmU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "frmU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub frmU_ResizedControlWindow()
  AddLogEntry "frmU_ResizedControlWindow"
End Sub

Private Sub frmU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub frmU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "frmU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_Click(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_Click: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_ContextMenu(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_ContextMenu: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_CustomDraw(index As Integer, ByVal drawStage As BtnCtlsLibACtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibACtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE, furtherProcessing As BtnCtlsLibACtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "optA_CustomDraw: Index=" & index & ", drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub optA_DblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_DblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_DestroyedControlWindow(index As Integer, ByVal hWnd As Long)
  AddLogEntry "optA_DestroyedControlWindow: Index=" & index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub optA_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
  AddLogEntry "optA_DragDrop: Index=" & index
End Sub

Private Sub optA_DragOver(index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "optA_DragOver: Index=" & index
End Sub

Private Sub optA_GotFocus(index As Integer)
  AddLogEntry "optA_GotFocus: Index=" & index
  Set objActiveCtl = optA(index)
End Sub

Private Sub optA_KeyDown(index As Integer, KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "optA_KeyDown: Index=" & index & ", keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub optA_KeyPress(index As Integer, keyAscii As Integer)
  AddLogEntry "optA_KeyPress: Index=" & index & ", keyAscii=" & keyAscii
End Sub

Private Sub optA_KeyUp(index As Integer, KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "optA_KeyUp: Index=" & index & ", keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub optA_LostFocus(index As Integer)
  AddLogEntry "optA_LostFocus: Index=" & index
End Sub

Private Sub optA_MClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseDown(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MouseDown: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseEnter(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MouseEnter: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseHover(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MouseHover: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseLeave(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MouseLeave: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseMove(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "optA_MouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_MouseUp(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_MouseUp: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_OLEDragDrop(index As Integer, ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optA_OLEDragDrop: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    optA(index).FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub optA_OLEDragEnter(index As Integer, ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optA_OLEDragEnter: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optA_OLEDragLeave(index As Integer, ByVal data As BtnCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optA_OLEDragLeave: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optA_OLEDragMouseMove(index As Integer, ByVal data As BtnCtlsLibACtl.IOLEDataObject, effect As BtnCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optA_OLEDragMouseMove: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optA_OwnerDraw(index As Integer, ByVal requiredAction As BtnCtlsLibACtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibACtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibACtl.RECTANGLE)
  AddLogEntry "optA_OwnerDraw: Index=" & index & ", requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub optA_RClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_RClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_RDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_RDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_RecreatedControlWindow(index As Integer, ByVal hWnd As Long)
  AddLogEntry "optA_RecreatedControlWindow: Index=" & index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub optA_ResizedControlWindow(index As Integer)
  AddLogEntry "optA_ResizedControlWindow: Index=" & index
End Sub

Private Sub optA_SelectionChanged(index As Integer, ByVal previousSelectionState As Boolean, ByVal newSelectionState As Boolean)
  AddLogEntry "optA_SelectionChanged: Index=" & index & ", previousSelectionState=" & previousSelectionState & ", newSelectionState=" & newSelectionState
End Sub

Private Sub optA_Validate(index As Integer, Cancel As Boolean)
  AddLogEntry "optA_Validate: Index=" & index
End Sub

Private Sub optA_XClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_XClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optA_XDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optA_XDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_Click(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_Click: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_ContextMenu(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_ContextMenu: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_CustomDraw(index As Integer, ByVal drawStage As BtnCtlsLibUCtl.CustomDrawStageConstants, ByVal controlState As BtnCtlsLibUCtl.CustomDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE, furtherProcessing As BtnCtlsLibUCtl.CustomDrawReturnValuesConstants)
  'AddLogEntry "optU_CustomDraw: Index=" & index & ", drawStage=" & drawStage & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
End Sub

Private Sub optU_DblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_DblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_DestroyedControlWindow(index As Integer, ByVal hWnd As Long)
  AddLogEntry "optU_DestroyedControlWindow: Index=" & index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub optU_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
  AddLogEntry "optU_DragDrop: Index=" & index
End Sub

Private Sub optU_DragOver(index As Integer, Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "optU_DragOver: Index=" & index
End Sub

Private Sub optU_GotFocus(index As Integer)
  AddLogEntry "optU_GotFocus: Index=" & index
  Set objActiveCtl = optU(index)
End Sub

Private Sub optU_KeyDown(index As Integer, KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "optU_KeyDown: Index=" & index & ", keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub optU_KeyPress(index As Integer, keyAscii As Integer)
  AddLogEntry "optU_KeyPress: Index=" & index & ", keyAscii=" & keyAscii
End Sub

Private Sub optU_KeyUp(index As Integer, KeyCode As Integer, ByVal shift As Integer)
  AddLogEntry "optU_KeyUp: Index=" & index & ", keyCode=" & KeyCode & ", shift=" & shift
End Sub

Private Sub optU_LostFocus(index As Integer)
  AddLogEntry "optU_LostFocus: Index=" & index
End Sub

Private Sub optU_MClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseDown(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MouseDown: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseEnter(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MouseEnter: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseHover(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MouseHover: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseLeave(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MouseLeave: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseMove(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "optU_MouseMove: Index=" & Index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_MouseUp(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_MouseUp: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_OLEDragDrop(index As Integer, ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optU_OLEDragDrop: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    optU(index).FinishOLEDragdrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub optU_OLEDragEnter(index As Integer, ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optU_OLEDragEnter: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optU_OLEDragLeave(index As Integer, ByVal data As BtnCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optU_OLEDragLeave: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optU_OLEDragMouseMove(index As Integer, ByVal data As BtnCtlsLibUCtl.IOLEDataObject, effect As BtnCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "optU_OLEDragMouseMove: index=" & index & ", data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub optU_OwnerDraw(index As Integer, ByVal requiredAction As BtnCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As BtnCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As BtnCtlsLibUCtl.RECTANGLE)
  AddLogEntry "optU_OwnerDraw: Index=" & index & ", requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub optU_RClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_RClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_RDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_RDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_RecreatedControlWindow(index As Integer, ByVal hWnd As Long)
  AddLogEntry "optU_RecreatedControlWindow: Index=" & index & ", hWnd=0x" & Hex(hWnd)
End Sub

Private Sub optU_ResizedControlWindow(index As Integer)
  AddLogEntry "optU_ResizedControlWindow: Index=" & index
End Sub

Private Sub optU_SelectionChanged(index As Integer, ByVal previousSelectionState As Boolean, ByVal newSelectionState As Boolean)
  AddLogEntry "optU_SelectionChanged: Index=" & index & ", previousSelectionState=" & previousSelectionState & ", newSelectionState=" & newSelectionState
End Sub

Private Sub optU_Validate(index As Integer, Cancel As Boolean)
  AddLogEntry "optU_Validate: Index=" & index
End Sub

Private Sub optU_XClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_XClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub optU_XDblClick(index As Integer, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "optU_XDblClick: Index=" & index & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
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
