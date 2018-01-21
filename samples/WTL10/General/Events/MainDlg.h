// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include <initguid.h>

#import <libid:2AFA7915-463D-4B61-AEB7-41B1236C143E> version("1.10") raw_dispinterfaces
#import <libid:A38BC98C-3D41-44db-B27B-CF19BAD8AE53> version("1.10") raw_dispinterfaces

DEFINE_GUID(LIBID_BtnCtlsLibU, 0x2AFA7915, 0x463D, 0x4B61, 0xAE, 0xB7, 0x41, 0xB1, 0x23, 0x6C, 0x14, 0x3E);
DEFINE_GUID(LIBID_BtnCtlsLibA, 0xA38BC98C, 0x3D41, 0x44db, 0xB2, 0x7B, 0xCF, 0x19, 0xBA, 0xD8, 0xAE, 0x53);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_FRMU, CMainDlg, &__uuidof(BtnCtlsLibU::_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_CHKU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_CMDU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_OPT1U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_OPT2U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_FRMA, CMainDlg, &__uuidof(BtnCtlsLibA::_IFrameEvents), &LIBID_BtnCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_CHKA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICheckBoxEvents), &LIBID_BtnCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_CMDA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICommandButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_OPT1A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_OPT2A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> frmUContainerWnd;
	CContainedWindowT<CWindow> frmUWnd;
	CContainedWindowT<CAxWindow> chkUWnd;
	CContainedWindowT<CAxWindow> cmdUWnd;
	CContainedWindowT<CAxWindow> opt1UWnd;
	CContainedWindowT<CAxWindow> opt2UWnd;
	CContainedWindowT<CAxWindow> frmAContainerWnd;
	CContainedWindowT<CWindow> frmAWnd;
	CContainedWindowT<CAxWindow> chkAWnd;
	CContainedWindowT<CAxWindow> cmdAWnd;
	CContainedWindowT<CAxWindow> opt1AWnd;
	CContainedWindowT<CAxWindow> opt2AWnd;

	CMainDlg() :
	    frmUContainerWnd(this, 1),
	    frmUWnd(this, 2),
	    chkUWnd(this, 3),
	    cmdUWnd(this, 4),
	    opt1UWnd(this, 5),
	    opt2UWnd(this, 6),
	    frmAContainerWnd(this, 7),
	    frmAWnd(this, 8),
	    chkAWnd(this, 9),
	    cmdAWnd(this, 10),
	    opt1AWnd(this, 11),
	    opt2AWnd(this, 12)
	{
	}

	int focusedControl;

	struct Controls
	{
		CEdit logEdit;
		CButton aboutButton;
		CComPtr<BtnCtlsLibU::IFrame> frmU;
		CComPtr<BtnCtlsLibU::ICheckBox> chkU;
		CComPtr<BtnCtlsLibU::ICommandButton> cmdU;
		CComPtr<BtnCtlsLibU::IOptionButton> opt1U;
		CComPtr<BtnCtlsLibU::IOptionButton> opt2U;
		CComPtr<BtnCtlsLibA::IFrame> frmA;
		CComPtr<BtnCtlsLibA::ICheckBox> chkA;
		CComPtr<BtnCtlsLibA::ICommandButton> cmdA;
		CComPtr<BtnCtlsLibA::IOptionButton> opt1A;
		CComPtr<BtnCtlsLibA::IOptionButton> opt2A;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusChkU)

		ALT_MSG_MAP(4)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusCmdU)

		ALT_MSG_MAP(5)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusOpt1U)

		ALT_MSG_MAP(6)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusOpt2U)

		ALT_MSG_MAP(7)
		ALT_MSG_MAP(8)
		ALT_MSG_MAP(9)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusChkA)

		ALT_MSG_MAP(10)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusCmdA)

		ALT_MSG_MAP(11)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusOpt1A)

		ALT_MSG_MAP(12)
		COMMAND_CODE_HANDLER(BN_SETFOCUS, OnSetFocusOpt2A)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 0, ClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 1, ContextMenuFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 2, DblClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 3, DestroyedControlWindowFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 4, MClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 5, MDblClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 6, MouseDownFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 7, MouseEnterFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 8, MouseHoverFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 9, MouseLeaveFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 10, MouseMoveFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 11, MouseUpFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 12, OLEDragDropFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 13, OLEDragEnterFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 14, OLEDragLeaveFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 15, OLEDragMouseMoveFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 17, RClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 18, RDblClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 19, RecreatedControlWindowFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 20, ResizedControlWindowFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 21, XClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(BtnCtlsLibU::_IFrameEvents), 22, XDblClickFrmU)

		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 1, ClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 2, ContextMenuChkU)
		//SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 3, CustomDrawChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 4, DblClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 5, DestroyedControlWindowChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 6, KeyDownChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 7, KeyPressChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 8, KeyUpChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 9, MClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 10, MDblClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 11, MouseDownChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 12, MouseEnterChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 13, MouseHoverChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 14, MouseLeaveChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 15, MouseMoveChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 16, MouseUpChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 17, OLEDragDropChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 18, OLEDragEnterChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 19, OLEDragLeaveChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 20, OLEDragMouseMoveChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 22, RClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 23, RDblClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 24, RecreatedControlWindowChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 25, ResizedControlWindowChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 0, SelectionStateChangedChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 26, XClickChkU)
		SINK_ENTRY_EX(IDC_CHKU, __uuidof(BtnCtlsLibU::_ICheckBoxEvents), 27, XDblClickChkU)

		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 0, ClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 1, ContextMenuCmdU)
		//SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 2, CustomDrawCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 3, DblClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 4, DestroyedControlWindowCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 5, KeyDownCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 6, KeyPressCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 7, KeyUpCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 8, MClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 9, MDblClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 10, MouseDownCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 11, MouseEnterCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 12, MouseHoverCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 13, MouseLeaveCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 14, MouseMoveCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 15, MouseUpCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 16, OLEDragDropCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 17, OLEDragEnterCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 18, OLEDragLeaveCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 19, OLEDragMouseMoveCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 21, RClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 22, RDblClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 23, RecreatedControlWindowCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 24, ResizedControlWindowCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 25, CustomDropDownAreaSizeCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 26, DropDownCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 27, XClickCmdU)
		SINK_ENTRY_EX(IDC_CMDU, __uuidof(BtnCtlsLibU::_ICommandButtonEvents), 28, XDblClickCmdU)

		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 1, ClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 2, ContextMenuOpt1U)
		//SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 3, CustomDrawOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 4, DblClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 5, DestroyedControlWindowOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 6, KeyDownOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 7, KeyPressOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 8, KeyUpOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 9, MClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 10, MDblClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 11, MouseDownOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 12, MouseEnterOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 13, MouseHoverOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 14, MouseLeaveOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 15, MouseMoveOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 16, MouseUpOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 17, OLEDragDropOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 18, OLEDragEnterOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 19, OLEDragLeaveOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 20, OLEDragMouseMoveOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 22, RClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 23, RDblClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 24, RecreatedControlWindowOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 25, ResizedControlWindowOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 0, SelectionChangedOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 26, XClickOpt1U)
		SINK_ENTRY_EX(IDC_OPT1U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 27, XDblClickOpt1U)

		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 1, ClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 2, ContextMenuOpt2U)
		//SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 3, CustomDrawOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 4, DblClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 5, DestroyedControlWindowOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 6, KeyDownOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 7, KeyPressOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 8, KeyUpOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 9, MClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 10, MDblClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 11, MouseDownOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 12, MouseEnterOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 13, MouseHoverOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 14, MouseLeaveOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 15, MouseMoveOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 16, MouseUpOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 17, OLEDragDropOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 18, OLEDragEnterOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 19, OLEDragLeaveOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 20, OLEDragMouseMoveOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 22, RClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 23, RDblClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 24, RecreatedControlWindowOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 25, ResizedControlWindowOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 0, SelectionChangedOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 26, XClickOpt2U)
		SINK_ENTRY_EX(IDC_OPT2U, __uuidof(BtnCtlsLibU::_IOptionButtonEvents), 27, XDblClickOpt2U)

		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 0, ClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 1, ContextMenuFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 2, DblClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 3, DestroyedControlWindowFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 4, MClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 5, MDblClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 6, MouseDownFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 7, MouseEnterFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 8, MouseHoverFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 9, MouseLeaveFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 10, MouseMoveFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 11, MouseUpFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 12, OLEDragDropFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 13, OLEDragEnterFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 14, OLEDragLeaveFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 15, OLEDragMouseMoveFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 17, RClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 18, RDblClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 19, RecreatedControlWindowFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 20, ResizedControlWindowFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 21, XClickFrmA)
		SINK_ENTRY_EX(IDC_FRMA, __uuidof(BtnCtlsLibA::_IFrameEvents), 22, XDblClickFrmA)

		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 1, ClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 2, ContextMenuChkA)
		//SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 3, CustomDrawChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 4, DblClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 5, DestroyedControlWindowChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 6, KeyDownChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 7, KeyPressChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 8, KeyUpChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 9, MClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 10, MDblClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 11, MouseDownChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 12, MouseEnterChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 13, MouseHoverChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 14, MouseLeaveChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 15, MouseMoveChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 16, MouseUpChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 17, OLEDragDropChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 18, OLEDragEnterChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 19, OLEDragLeaveChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 20, OLEDragMouseMoveChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 22, RClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 23, RDblClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 24, RecreatedControlWindowChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 25, ResizedControlWindowChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 0, SelectionStateChangedChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 26, XClickChkA)
		SINK_ENTRY_EX(IDC_CHKA, __uuidof(BtnCtlsLibA::_ICheckBoxEvents), 27, XDblClickChkA)

		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 0, ClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 1, ContextMenuCmdA)
		//SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 2, CustomDrawCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 3, DblClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 4, DestroyedControlWindowCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 5, KeyDownCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 6, KeyPressCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 7, KeyUpCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 8, MClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 9, MDblClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 10, MouseDownCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 11, MouseEnterCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 12, MouseHoverCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 13, MouseLeaveCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 14, MouseMoveCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 15, MouseUpCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 16, OLEDragDropCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 17, OLEDragEnterCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 18, OLEDragLeaveCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 19, OLEDragMouseMoveCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 21, RClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 22, RDblClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 23, RecreatedControlWindowCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 24, ResizedControlWindowCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 25, CustomDropDownAreaSizeCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 26, DropDownCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 27, XClickCmdA)
		SINK_ENTRY_EX(IDC_CMDA, __uuidof(BtnCtlsLibA::_ICommandButtonEvents), 28, XDblClickCmdA)

		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 1, ClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 2, ContextMenuOpt1A)
		//SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 3, CustomDrawOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 4, DblClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 5, DestroyedControlWindowOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 6, KeyDownOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 7, KeyPressOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 8, KeyUpOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 9, MClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 10, MDblClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 11, MouseDownOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 12, MouseEnterOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 13, MouseHoverOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 14, MouseLeaveOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 15, MouseMoveOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 16, MouseUpOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 17, OLEDragDropOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 18, OLEDragEnterOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 19, OLEDragLeaveOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 20, OLEDragMouseMoveOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 22, RClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 23, RDblClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 24, RecreatedControlWindowOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 25, ResizedControlWindowOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 0, SelectionChangedOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 26, XClickOpt1A)
		SINK_ENTRY_EX(IDC_OPT1A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 27, XDblClickOpt1A)

		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 1, ClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 2, ContextMenuOpt2A)
		//SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 3, CustomDrawOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 4, DblClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 5, DestroyedControlWindowOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 6, KeyDownOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 7, KeyPressOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 8, KeyUpOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 9, MClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 10, MDblClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 11, MouseDownOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 12, MouseEnterOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 13, MouseHoverOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 14, MouseLeaveOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 15, MouseMoveOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 16, MouseUpOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 17, OLEDragDropOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 18, OLEDragEnterOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 19, OLEDragLeaveOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 20, OLEDragMouseMoveOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 22, RClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 23, RDblClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 24, RecreatedControlWindowOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 25, ResizedControlWindowOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 0, SelectionChangedOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 26, XClickOpt2A)
		SINK_ENTRY_EX(IDC_OPT2A, __uuidof(BtnCtlsLibA::_IOptionButtonEvents), 27, XDblClickOpt2A)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusChkU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusCmdU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusOpt1U(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusOpt2U(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusChkA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusCmdA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusOpt1A(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);
	LRESULT OnSetFocusOpt2A(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);

	void __stdcall ClickFrmU(short button, short shift, long x, long y);
	void __stdcall ContextMenuFrmU(short button, short shift, long x, long y);
	void __stdcall DblClickFrmU(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowFrmU(long hWnd);
	void __stdcall MClickFrmU(short button, short shift, long x, long y);
	void __stdcall MDblClickFrmU(short button, short shift, long x, long y);
	void __stdcall MouseDownFrmU(short button, short shift, long x, long y);
	void __stdcall MouseEnterFrmU(short button, short shift, long x, long y);
	void __stdcall MouseHoverFrmU(short button, short shift, long x, long y);
	void __stdcall MouseLeaveFrmU(short button, short shift, long x, long y);
	void __stdcall MouseMoveFrmU(short button, short shift, long x, long y);
	void __stdcall MouseUpFrmU(short button, short shift, long x, long y);
	void __stdcall OLEDragDropFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveFrmU(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickFrmU(short button, short shift, long x, long y);
	void __stdcall RDblClickFrmU(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowFrmU(long hWnd);
	void __stdcall ResizedControlWindowFrmU();
	void __stdcall XClickFrmU(short button, short shift, long x, long y);
	void __stdcall XDblClickFrmU(short button, short shift, long x, long y);

	void __stdcall ClickChkU(short button, short shift, long x, long y);
	void __stdcall ContextMenuChkU(short button, short shift, long x, long y);
	void __stdcall CustomDrawChkU(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickChkU(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowChkU(long hWnd);
	void __stdcall KeyDownChkU(short* keyCode, short shift);
	void __stdcall KeyPressChkU(short* keyAscii);
	void __stdcall KeyUpChkU(short* keyCode, short shift);
	void __stdcall MClickChkU(short button, short shift, long x, long y);
	void __stdcall MDblClickChkU(short button, short shift, long x, long y);
	void __stdcall MouseDownChkU(short button, short shift, long x, long y);
	void __stdcall MouseEnterChkU(short button, short shift, long x, long y);
	void __stdcall MouseHoverChkU(short button, short shift, long x, long y);
	void __stdcall MouseLeaveChkU(short button, short shift, long x, long y);
	void __stdcall MouseMoveChkU(short button, short shift, long x, long y);
	void __stdcall MouseUpChkU(short button, short shift, long x, long y);
	void __stdcall OLEDragDropChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveChkU(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickChkU(short button, short shift, long x, long y);
	void __stdcall RDblClickChkU(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowChkU(long hWnd);
	void __stdcall ResizedControlWindowChkU();
	void __stdcall SelectionStateChangedChkU(BtnCtlsLibU::SelectionStateConstants previousSelectionState, BtnCtlsLibU::SelectionStateConstants newSelectionState);
	void __stdcall XClickChkU(short button, short shift, long x, long y);
	void __stdcall XDblClickChkU(short button, short shift, long x, long y);

	void __stdcall ClickCmdU(short button, short shift, long x, long y);
	void __stdcall ContextMenuCmdU(short button, short shift, long x, long y);
	void __stdcall CustomDrawCmdU(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall CustomDropDownAreaSizeCmdU(BtnCtlsLibU::RECTANGLE* clientRectangle, BtnCtlsLibU::RECTANGLE* commandButtonAreaRectangle, BtnCtlsLibU::RECTANGLE* dropDownAreaRectangle, BtnCtlsLibU::CustomDropDownAreaSizeReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickCmdU(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowCmdU(long hWnd);
	void __stdcall DropDownCmdU(BtnCtlsLibU::RECTANGLE* buttonRectangle);
	void __stdcall KeyDownCmdU(short* keyCode, short shift);
	void __stdcall KeyPressCmdU(short* keyAscii);
	void __stdcall KeyUpCmdU(short* keyCode, short shift);
	void __stdcall MClickCmdU(short button, short shift, long x, long y);
	void __stdcall MDblClickCmdU(short button, short shift, long x, long y);
	void __stdcall MouseDownCmdU(short button, short shift, long x, long y);
	void __stdcall MouseEnterCmdU(short button, short shift, long x, long y);
	void __stdcall MouseHoverCmdU(short button, short shift, long x, long y);
	void __stdcall MouseLeaveCmdU(short button, short shift, long x, long y);
	void __stdcall MouseMoveCmdU(short button, short shift, long x, long y);
	void __stdcall MouseUpCmdU(short button, short shift, long x, long y);
	void __stdcall OLEDragDropCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveCmdU(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickCmdU(short button, short shift, long x, long y);
	void __stdcall RDblClickCmdU(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowCmdU(long hWnd);
	void __stdcall ResizedControlWindowCmdU();
	void __stdcall XClickCmdU(short button, short shift, long x, long y);
	void __stdcall XDblClickCmdU(short button, short shift, long x, long y);

	void __stdcall ClickOpt1U(short button, short shift, long x, long y);
	void __stdcall ContextMenuOpt1U(short button, short shift, long x, long y);
	void __stdcall CustomDrawOpt1U(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickOpt1U(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowOpt1U(long hWnd);
	void __stdcall KeyDownOpt1U(short* keyCode, short shift);
	void __stdcall KeyPressOpt1U(short* keyAscii);
	void __stdcall KeyUpOpt1U(short* keyCode, short shift);
	void __stdcall MClickOpt1U(short button, short shift, long x, long y);
	void __stdcall MDblClickOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseDownOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseEnterOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseHoverOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseLeaveOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseMoveOpt1U(short button, short shift, long x, long y);
	void __stdcall MouseUpOpt1U(short button, short shift, long x, long y);
	void __stdcall OLEDragDropOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveOpt1U(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickOpt1U(short button, short shift, long x, long y);
	void __stdcall RDblClickOpt1U(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowOpt1U(long hWnd);
	void __stdcall ResizedControlWindowOpt1U();
	void __stdcall SelectionChangedOpt1U(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState);
	void __stdcall XClickOpt1U(short button, short shift, long x, long y);
	void __stdcall XDblClickOpt1U(short button, short shift, long x, long y);

	void __stdcall ClickOpt2U(short button, short shift, long x, long y);
	void __stdcall ContextMenuOpt2U(short button, short shift, long x, long y);
	void __stdcall CustomDrawOpt2U(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickOpt2U(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowOpt2U(long hWnd);
	void __stdcall KeyDownOpt2U(short* keyCode, short shift);
	void __stdcall KeyPressOpt2U(short* keyAscii);
	void __stdcall KeyUpOpt2U(short* keyCode, short shift);
	void __stdcall MClickOpt2U(short button, short shift, long x, long y);
	void __stdcall MDblClickOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseDownOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseEnterOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseHoverOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseLeaveOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseMoveOpt2U(short button, short shift, long x, long y);
	void __stdcall MouseUpOpt2U(short button, short shift, long x, long y);
	void __stdcall OLEDragDropOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveOpt2U(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickOpt2U(short button, short shift, long x, long y);
	void __stdcall RDblClickOpt2U(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowOpt2U(long hWnd);
	void __stdcall ResizedControlWindowOpt2U();
	void __stdcall SelectionChangedOpt2U(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState);
	void __stdcall XClickOpt2U(short button, short shift, long x, long y);
	void __stdcall XDblClickOpt2U(short button, short shift, long x, long y);

	void __stdcall ClickFrmA(short button, short shift, long x, long y);
	void __stdcall ContextMenuFrmA(short button, short shift, long x, long y);
	void __stdcall DblClickFrmA(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowFrmA(long hWnd);
	void __stdcall MClickFrmA(short button, short shift, long x, long y);
	void __stdcall MDblClickFrmA(short button, short shift, long x, long y);
	void __stdcall MouseDownFrmA(short button, short shift, long x, long y);
	void __stdcall MouseEnterFrmA(short button, short shift, long x, long y);
	void __stdcall MouseHoverFrmA(short button, short shift, long x, long y);
	void __stdcall MouseLeaveFrmA(short button, short shift, long x, long y);
	void __stdcall MouseMoveFrmA(short button, short shift, long x, long y);
	void __stdcall MouseUpFrmA(short button, short shift, long x, long y);
	void __stdcall OLEDragDropFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveFrmA(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickFrmA(short button, short shift, long x, long y);
	void __stdcall RDblClickFrmA(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowFrmA(long hWnd);
	void __stdcall ResizedControlWindowFrmA();
	void __stdcall XClickFrmA(short button, short shift, long x, long y);
	void __stdcall XDblClickFrmA(short button, short shift, long x, long y);

	void __stdcall ClickChkA(short button, short shift, long x, long y);
	void __stdcall ContextMenuChkA(short button, short shift, long x, long y);
	void __stdcall CustomDrawChkA(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickChkA(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowChkA(long hWnd);
	void __stdcall KeyDownChkA(short* keyCode, short shift);
	void __stdcall KeyPressChkA(short* keyAscii);
	void __stdcall KeyUpChkA(short* keyCode, short shift);
	void __stdcall MClickChkA(short button, short shift, long x, long y);
	void __stdcall MDblClickChkA(short button, short shift, long x, long y);
	void __stdcall MouseDownChkA(short button, short shift, long x, long y);
	void __stdcall MouseEnterChkA(short button, short shift, long x, long y);
	void __stdcall MouseHoverChkA(short button, short shift, long x, long y);
	void __stdcall MouseLeaveChkA(short button, short shift, long x, long y);
	void __stdcall MouseMoveChkA(short button, short shift, long x, long y);
	void __stdcall MouseUpChkA(short button, short shift, long x, long y);
	void __stdcall OLEDragDropChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveChkA(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickChkA(short button, short shift, long x, long y);
	void __stdcall RDblClickChkA(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowChkA(long hWnd);
	void __stdcall ResizedControlWindowChkA();
	void __stdcall SelectionStateChangedChkA(BtnCtlsLibA::SelectionStateConstants previousSelectionState, BtnCtlsLibA::SelectionStateConstants newSelectionState);
	void __stdcall XClickChkA(short button, short shift, long x, long y);
	void __stdcall XDblClickChkA(short button, short shift, long x, long y);

	void __stdcall ClickCmdA(short button, short shift, long x, long y);
	void __stdcall ContextMenuCmdA(short button, short shift, long x, long y);
	void __stdcall CustomDrawCmdA(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall CustomDropDownAreaSizeCmdA(BtnCtlsLibA::RECTANGLE* clientRectangle, BtnCtlsLibA::RECTANGLE* commandButtonAreaRectangle, BtnCtlsLibA::RECTANGLE* dropDownAreaRectangle, BtnCtlsLibA::CustomDropDownAreaSizeReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickCmdA(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowCmdA(long hWnd);
	void __stdcall DropDownCmdA(BtnCtlsLibA::RECTANGLE* buttonRectangle);
	void __stdcall KeyDownCmdA(short* keyCode, short shift);
	void __stdcall KeyPressCmdA(short* keyAscii);
	void __stdcall KeyUpCmdA(short* keyCode, short shift);
	void __stdcall MClickCmdA(short button, short shift, long x, long y);
	void __stdcall MDblClickCmdA(short button, short shift, long x, long y);
	void __stdcall MouseDownCmdA(short button, short shift, long x, long y);
	void __stdcall MouseEnterCmdA(short button, short shift, long x, long y);
	void __stdcall MouseHoverCmdA(short button, short shift, long x, long y);
	void __stdcall MouseLeaveCmdA(short button, short shift, long x, long y);
	void __stdcall MouseMoveCmdA(short button, short shift, long x, long y);
	void __stdcall MouseUpCmdA(short button, short shift, long x, long y);
	void __stdcall OLEDragDropCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveCmdA(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickCmdA(short button, short shift, long x, long y);
	void __stdcall RDblClickCmdA(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowCmdA(long hWnd);
	void __stdcall ResizedControlWindowCmdA();
	void __stdcall XClickCmdA(short button, short shift, long x, long y);
	void __stdcall XDblClickCmdA(short button, short shift, long x, long y);

	void __stdcall ClickOpt1A(short button, short shift, long x, long y);
	void __stdcall ContextMenuOpt1A(short button, short shift, long x, long y);
	void __stdcall CustomDrawOpt1A(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickOpt1A(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowOpt1A(long hWnd);
	void __stdcall KeyDownOpt1A(short* keyCode, short shift);
	void __stdcall KeyPressOpt1A(short* keyAscii);
	void __stdcall KeyUpOpt1A(short* keyCode, short shift);
	void __stdcall MClickOpt1A(short button, short shift, long x, long y);
	void __stdcall MDblClickOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseDownOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseEnterOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseHoverOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseLeaveOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseMoveOpt1A(short button, short shift, long x, long y);
	void __stdcall MouseUpOpt1A(short button, short shift, long x, long y);
	void __stdcall OLEDragDropOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveOpt1A(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickOpt1A(short button, short shift, long x, long y);
	void __stdcall RDblClickOpt1A(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowOpt1A(long hWnd);
	void __stdcall ResizedControlWindowOpt1A();
	void __stdcall SelectionChangedOpt1A(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState);
	void __stdcall XClickOpt1A(short button, short shift, long x, long y);
	void __stdcall XDblClickOpt1A(short button, short shift, long x, long y);

	void __stdcall ClickOpt2A(short button, short shift, long x, long y);
	void __stdcall ContextMenuOpt2A(short button, short shift, long x, long y);
	void __stdcall CustomDrawOpt2A(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickOpt2A(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowOpt2A(long hWnd);
	void __stdcall KeyDownOpt2A(short* keyCode, short shift);
	void __stdcall KeyPressOpt2A(short* keyAscii);
	void __stdcall KeyUpOpt2A(short* keyCode, short shift);
	void __stdcall MClickOpt2A(short button, short shift, long x, long y);
	void __stdcall MDblClickOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseDownOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseEnterOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseHoverOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseLeaveOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseMoveOpt2A(short button, short shift, long x, long y);
	void __stdcall MouseUpOpt2A(short button, short shift, long x, long y);
	void __stdcall OLEDragDropOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveOpt2A(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickOpt2A(short button, short shift, long x, long y);
	void __stdcall RDblClickOpt2A(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowOpt2A(long hWnd);
	void __stdcall ResizedControlWindowOpt2A();
	void __stdcall SelectionChangedOpt2A(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState);
	void __stdcall XClickOpt2A(short button, short shift, long x, long y);
	void __stdcall XDblClickOpt2A(short button, short shift, long x, long y);
};
