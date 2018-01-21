// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:2AFA7915-463D-4B61-AEB7-41B1236C143E> version("1.10") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_FRMU, CMainDlg, &__uuidof(_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> frmUContainerWnd;
	CContainedWindowT<CWindow> frmUWnd;

	CMainDlg() :
	    frmUContainerWnd(this, 1),
	    frmUWnd(this, 2)
	{
	}

	BOOL collapsed;

	struct Controls
	{
		CEdit text;
		CImageList imlMinus;
		CImageList imlPlus;
		CComPtr<IFrame> frmU;

		~Controls()
		{
			imlMinus.Destroy();
			imlPlus.Destroy();
		}
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(_IFrameEvents), 0, ClickFrmU)
		SINK_ENTRY_EX(IDC_FRMU, __uuidof(_IFrameEvents), 19, RecreatedControlWindowFrmU)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void CollapseFrame(void);
	void ExpandFrame(void);

	void __stdcall ClickFrmU(short /*button*/, short /*shift*/, long x, long y);
	void __stdcall RecreatedControlWindowFrmU(long /*hWnd*/);
};
