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
    public IDispEventImpl<IDC_CMDBTNU, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> cmdBtnUWnd;

	CMainDlg() : cmdBtnUWnd(this, 1)
	{
	}

	struct Controls
	{
		CComPtr<ICommandButton> cmdBtnU;
	} controls;

	CFont webdingsFont;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_CMDBTNU, __uuidof(_ICommandButtonEvents), 0, ClickCmdbtnu)
		SINK_ENTRY_EX(IDC_CMDBTNU, __uuidof(_ICommandButtonEvents), 2, CustomDrawCmdbtnu)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);

	void __stdcall ClickCmdbtnu(short /*button*/, short /*shift*/, long x, long /*y*/);
	void __stdcall CustomDrawCmdbtnu(CustomDrawStageConstants drawStage, CustomDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle, CustomDrawReturnValuesConstants* furtherProcessing);
};
