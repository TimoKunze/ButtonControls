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
    public IDispEventImpl<IDC_CMDSPLIT, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_CMDELEVATION, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_CMDLNK, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> cmdSplitWnd;
	CContainedWindowT<CAxWindow> cmdElevationWnd;
	CContainedWindowT<CAxWindow> cmdLinkWnd;

	CMainDlg() :
	    cmdSplitWnd(this, 1),
	    cmdElevationWnd(this, 2),
	    cmdLinkWnd(this, 3)
	{
	}

	struct Controls
	{
		CComPtr<ICommandButton> cmdSplit;
		CComPtr<ICommandButton> cmdElevation;
		CComPtr<ICommandButton> cmdLink;
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
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_CMDSPLIT, __uuidof(_ICommandButtonEvents), 0, ClickCmdsplit)
		SINK_ENTRY_EX(IDC_CMDSPLIT, __uuidof(_ICommandButtonEvents), 26, DropDownCmdsplit)

		SINK_ENTRY_EX(IDC_CMDELEVATION, __uuidof(_ICommandButtonEvents), 0, ClickCmdelevation)

		SINK_ENTRY_EX(IDC_CMDLNK, __uuidof(_ICommandButtonEvents), 0, ClickCmdlink)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	HMENU CreateDropDownMenu(VARIANT_BOOL checked);

	void __stdcall ClickCmdsplit(short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall DropDownCmdsplit(RECTANGLE* buttonRectangle);

	void __stdcall ClickCmdelevation(short /*button*/, short /*shift*/, long /*x*/, long /*y*/);

	void __stdcall ClickCmdlink(short /*button*/, short /*shift*/, long /*x*/, long /*y*/);
};
