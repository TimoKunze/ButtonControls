// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.frmU) {
		IDispEventImpl<IDC_FRMU, CMainDlg, &__uuidof(_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.frmU);
		controls.frmU.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.imlMinus.CreateFromImage(IDB_MINUS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.imlPlus.CreateFromImage(IDB_PLUS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.text = GetDlgItem(IDC_TEXT);

	frmUContainerWnd.SubclassWindow(GetDlgItem(IDC_FRMU));
	frmUContainerWnd.QueryControl(__uuidof(IFrame), reinterpret_cast<LPVOID*>(&controls.frmU));
	if(controls.frmU) {
		IDispEventImpl< IDC_FRMU, CMainDlg, &__uuidof(_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10 >::DispEventAdvise(controls.frmU);
		frmUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.frmU->GethWnd())));
	}

	ExpandFrame();

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.frmU->About();
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::CollapseFrame(void)
{
	WTL::CRect frameRectangle;
	frmUWnd.GetWindowRect(&frameRectangle);
	ScreenToClient(&frameRectangle);
	frameRectangle.bottom = frameRectangle.top + 20;
	frmUWnd.SetWindowPos(NULL, &frameRectangle, SWP_NOREPOSITION);

	WTL::CRect textRectangle;
	controls.text.GetWindowRect(&textRectangle);
	ScreenToClient(&textRectangle);
	textRectangle.top = frameRectangle.bottom + 5;
	controls.text.SetWindowPos(NULL, &textRectangle, SWP_NOREPOSITION);
	controls.frmU->PuthImageList(HandleToLong(static_cast<HIMAGELIST>(controls.imlPlus)));
	collapsed = TRUE;
}

void CMainDlg::ExpandFrame(void)
{
	WTL::CRect frameRectangle;
	frmUWnd.GetWindowRect(&frameRectangle);
	ScreenToClient(&frameRectangle);
	frameRectangle.bottom = frameRectangle.top + 129;
	frmUWnd.SetWindowPos(NULL, &frameRectangle, SWP_NOREPOSITION);

	WTL::CRect textRectangle;
	controls.text.GetWindowRect(&textRectangle);
	ScreenToClient(&textRectangle);
	textRectangle.top = frameRectangle.bottom + 5;
	controls.text.SetWindowPos(NULL, &textRectangle, SWP_NOREPOSITION);
	controls.frmU->PuthImageList(HandleToLong(static_cast<HIMAGELIST>(controls.imlMinus)));
	collapsed = FALSE;
}

void __stdcall CMainDlg::ClickFrmU(short /*button*/, short /*shift*/, long x, long y)
{
	if((11 < x) && (x < 22)) {
		if((5 < y) && (y < 16)) {
			if(collapsed) {
				ExpandFrame();
			} else {
				CollapseFrame();
			}
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowFrmU(long /*hWnd*/)
{
	if(collapsed) {
		CollapseFrame();
	} else {
		ExpandFrame();
	}
}
