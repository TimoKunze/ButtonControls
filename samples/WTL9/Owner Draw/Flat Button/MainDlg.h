// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:2AFA7915-463D-4B61-AEB7-41B1236C143E> version("1.10") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_CMD, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_CHK, CMainDlg, &__uuidof(_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_OPT1, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_OPT2, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> cmdWnd;
	CContainedWindowT<CAxWindow> chkWnd;
	CContainedWindowT<CAxWindow> opt1Wnd;
	CContainedWindowT<CAxWindow> opt2Wnd;

	CMainDlg() :
	    cmdWnd(this, 1),
	    chkWnd(this, 2),
	    opt1Wnd(this, 3),
	    opt2Wnd(this, 4)
	{
	}

	struct Controls
	{
		CComPtr<ICommandButton> cmd;
		CComPtr<ICheckBox> chk;
		CComPtr<IOptionButton> opt1;
		CComPtr<IOptionButton> opt2;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		ALT_MSG_MAP(4)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_CMD, __uuidof(_ICommandButtonEvents), 20, OwnerDrawCmd)
		SINK_ENTRY_EX(IDC_CHK, __uuidof(_ICheckBoxEvents), 1, ClickChk)
		SINK_ENTRY_EX(IDC_CHK, __uuidof(_ICheckBoxEvents), 21, OwnerDrawChk)
		SINK_ENTRY_EX(IDC_OPT1, __uuidof(_IOptionButtonEvents), 21, OwnerDrawOpt1)
		SINK_ENTRY_EX(IDC_OPT1, __uuidof(_IOptionButtonEvents), 0, SelectionChanged1)
		SINK_ENTRY_EX(IDC_OPT2, __uuidof(_IOptionButtonEvents), 21, OwnerDrawOpt2)
		SINK_ENTRY_EX(IDC_OPT2, __uuidof(_IOptionButtonEvents), 0, SelectionChanged2)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);

	void __stdcall OwnerDrawCmd(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall ClickChk(short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall OwnerDrawChk(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall OwnerDrawOpt1(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall SelectionChanged1(VARIANT_BOOL /*previousSelectionState*/, VARIANT_BOOL newSelectionState);
	void __stdcall OwnerDrawOpt2(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle);
	void __stdcall SelectionChanged2(VARIANT_BOOL /*previousSelectionState*/, VARIANT_BOOL newSelectionState);
};
