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
		IDispEventImpl<IDC_FRMU, CMainDlg, &__uuidof(BtnCtlsLibU::_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.frmU);
		controls.frmU.Release();
	}
	if(controls.chkU) {
		IDispEventImpl<IDC_CHKU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.chkU);
		controls.chkU.Release();
	}
	if(controls.cmdU) {
		IDispEventImpl<IDC_CMDU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.cmdU);
		controls.cmdU.Release();
	}
	if(controls.opt1U) {
		IDispEventImpl<IDC_OPT1U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.opt1U);
		controls.opt1U.Release();
	}
	if(controls.opt2U) {
		IDispEventImpl<IDC_OPT2U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.opt2U);
		controls.opt2U.Release();
	}
	if(controls.frmA) {
		IDispEventImpl<IDC_FRMA, CMainDlg, &__uuidof(BtnCtlsLibA::_IFrameEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventUnadvise(controls.frmA);
		controls.frmA.Release();
	}
	if(controls.chkA) {
		IDispEventImpl<IDC_CHKA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICheckBoxEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventUnadvise(controls.chkA);
		controls.chkA.Release();
	}
	if(controls.cmdA) {
		IDispEventImpl<IDC_CMDA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICommandButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventUnadvise(controls.cmdA);
		controls.cmdA.Release();
	}
	if(controls.opt1A) {
		IDispEventImpl<IDC_OPT1A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventUnadvise(controls.opt1A);
		controls.opt1A.Release();
	}
	if(controls.opt2A) {
		IDispEventImpl<IDC_OPT2A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventUnadvise(controls.opt2A);
		controls.opt2A.Release();
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

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	frmUContainerWnd.SubclassWindow(GetDlgItem(IDC_FRMU));
	frmUContainerWnd.QueryControl(__uuidof(BtnCtlsLibU::IFrame), reinterpret_cast<LPVOID*>(&controls.frmU));
	if(controls.frmU) {
		IDispEventImpl<IDC_FRMU, CMainDlg, &__uuidof(BtnCtlsLibU::_IFrameEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.frmU);
		frmUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.frmU->GethWnd())));
	}
	chkUWnd.SubclassWindow(GetDlgItem(IDC_CHKU));
	chkUWnd.QueryControl(__uuidof(BtnCtlsLibU::ICheckBox), reinterpret_cast<LPVOID*>(&controls.chkU));
	if(controls.chkU) {
		IDispEventImpl<IDC_CHKU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.chkU);
	}
	cmdUWnd.SubclassWindow(GetDlgItem(IDC_CMDU));
	cmdUWnd.QueryControl(__uuidof(BtnCtlsLibU::ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdU));
	if(controls.cmdU) {
		IDispEventImpl<IDC_CMDU, CMainDlg, &__uuidof(BtnCtlsLibU::_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.cmdU);
	}
	opt1UWnd.SubclassWindow(GetDlgItem(IDC_OPT1U));
	opt1UWnd.QueryControl(__uuidof(BtnCtlsLibU::IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt1U));
	if(controls.opt1U) {
		IDispEventImpl<IDC_OPT1U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.opt1U);
	}
	opt2UWnd.SubclassWindow(GetDlgItem(IDC_OPT2U));
	opt2UWnd.QueryControl(__uuidof(BtnCtlsLibU::IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt2U));
	if(controls.opt2U) {
		IDispEventImpl<IDC_OPT2U, CMainDlg, &__uuidof(BtnCtlsLibU::_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.opt2U);
	}

	frmAContainerWnd.SubclassWindow(GetDlgItem(IDC_FRMA));
	frmAContainerWnd.QueryControl(__uuidof(BtnCtlsLibA::IFrame), reinterpret_cast<LPVOID*>(&controls.frmA));
	if(controls.frmA) {
		IDispEventImpl<IDC_FRMA, CMainDlg, &__uuidof(BtnCtlsLibA::_IFrameEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventAdvise(controls.frmA);
		frmAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.frmA->GethWnd())));
	}
	chkAWnd.SubclassWindow(GetDlgItem(IDC_CHKA));
	chkAWnd.QueryControl(__uuidof(BtnCtlsLibA::ICheckBox), reinterpret_cast<LPVOID*>(&controls.chkA));
	if(controls.chkA) {
		IDispEventImpl<IDC_CHKA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICheckBoxEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventAdvise(controls.chkA);
	}
	cmdAWnd.SubclassWindow(GetDlgItem(IDC_CMDA));
	cmdAWnd.QueryControl(__uuidof(BtnCtlsLibA::ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdA));
	if(controls.cmdA) {
		IDispEventImpl<IDC_CMDA, CMainDlg, &__uuidof(BtnCtlsLibA::_ICommandButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventAdvise(controls.cmdA);
	}
	opt1AWnd.SubclassWindow(GetDlgItem(IDC_OPT1A));
	opt1AWnd.QueryControl(__uuidof(BtnCtlsLibA::IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt1A));
	if(controls.opt1A) {
		IDispEventImpl<IDC_OPT1A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventAdvise(controls.opt1A);
	}
	opt2AWnd.SubclassWindow(GetDlgItem(IDC_OPT2A));
	opt2AWnd.QueryControl(__uuidof(BtnCtlsLibA::IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt2A));
	if(controls.opt2A) {
		IDispEventImpl<IDC_OPT2A, CMainDlg, &__uuidof(BtnCtlsLibA::_IOptionButtonEvents), &LIBID_BtnCtlsLibA, 1, 10>::DispEventAdvise(controls.opt2A);
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.logEdit && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect rc;
			GetClientRect(&rc);

			frmUContainerWnd.SetWindowPos(NULL, 0, 0, 197, (rc.Height() - 5) / 2, SWP_NOMOVE);
			int y = 0;
			OLE_XSIZE_PIXELS cx = 0;
			OLE_YSIZE_PIXELS cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.chkU->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				chkUWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			chkUWnd.SetWindowPos(NULL, 203, y + 5, cx, cy, 0);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.cmdU->GetIdealSize(&cx, &cy);
				cx += 5;
			} else {
				CRect rect;
				cmdUWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			cmdUWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);
			controls.logEdit.SetWindowPos(NULL, 203 + cx + 20, 0, rc.Width() - (203 + cx + 20), rc.Height() - 42, 0);
			controls.aboutButton.SetWindowPos(NULL, 203 + cx + 20, rc.Height() - 37, 0, 0, SWP_NOSIZE);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.opt1U->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				opt1UWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			opt1UWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.opt2U->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				opt2UWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			opt2UWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);

			y = (rc.Height() - 5) / 2 + 5;
			frmAContainerWnd.SetWindowPos(NULL, 0, y, 197, (rc.Height() - 5) / 2, 0);
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.chkA->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				chkAWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			chkAWnd.SetWindowPos(NULL, 203, y + 5, cx, cy, 0);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.cmdA->GetIdealSize(&cx, &cy);
				cx += 5;
			} else {
				CRect rect;
				cmdAWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			cmdAWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.opt1A->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				opt1AWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			opt1AWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);
			y += cy;
			cx = cy = 0;
			if(RunTimeHelper::IsCommCtrl6()) {
				controls.opt2A->GetIdealSize(&cx, &cy);
			} else {
				CRect rect;
				opt2AWnd.GetWindowRect(&rect);
				cx = rect.Width();
				cy = rect.Height();
			}
			opt2AWnd.SetWindowPos(NULL, 203, y + 7, cx, cy, 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	switch(focusedControl) {
		case 1:
			if(controls.chkU) {
				controls.chkU->About();
			}
			break;
		case 2:
			if(controls.cmdU) {
				controls.cmdU->About();
			}
			break;
		case 3:
			if(controls.opt1U) {
				controls.opt1U->About();
			}
			break;
		case 4:
			if(controls.opt2U) {
				controls.opt2U->About();
			}
			break;
		case 5:
			if(controls.chkA) {
				controls.chkA->About();
			}
			break;
		case 6:
			if(controls.cmdA) {
				controls.cmdA->About();
			}
			break;
		case 7:
			if(controls.opt1A) {
				controls.opt1A->About();
			}
			break;
		case 8:
			if(controls.opt2A) {
				controls.opt2A->About();
			}
			break;
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusChkU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 1;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusCmdU(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 2;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusOpt1U(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 3;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusOpt2U(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 4;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusChkA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 5;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusCmdA(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 6;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusOpt1A(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 7;
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusOpt2A(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& bHandled)
{
	focusedControl = 8;
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::ClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowFrmU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("FrmU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.frmU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveFrmU(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmU_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveFrmU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowFrmU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("FrmU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowFrmU()
{
	AddLogEntry(CAtlString(TEXT("FrmU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickFrmU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawChkU(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("ChkU_CustomDraw: drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowChkU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ChkU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownChkU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ChkU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressChkU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ChkU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpChkU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ChkU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.chkU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveChkU(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkU_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveChkU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowChkU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ChkU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowChkU()
{
	AddLogEntry(CAtlString(TEXT("ChkU_ResizedControlWindow")));
}

void __stdcall CMainDlg::SelectionStateChangedChkU(BtnCtlsLibU::SelectionStateConstants previousSelectionState, BtnCtlsLibU::SelectionStateConstants newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("ChkU_SelectionStateChanged: previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickChkU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawCmdU(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("CmdU_CustomDraw: drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDropDownAreaSizeCmdU(BtnCtlsLibU::RECTANGLE* clientRectangle, BtnCtlsLibU::RECTANGLE* commandButtonAreaRectangle, BtnCtlsLibU::RECTANGLE* dropDownAreaRectangle, BtnCtlsLibU::CustomDropDownAreaSizeReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("CmdU_CustomDropDownAreaSize: clientRectangle=(%i, %i)-(%i, %i), commandButtonAreaRectangle=(%i, %i)-(%i, %i), dropDownAreaRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), clientRectangle->Left, clientRectangle->Top, clientRectangle->Right, clientRectangle->Bottom, commandButtonAreaRectangle->Left, commandButtonAreaRectangle->Top, commandButtonAreaRectangle->Right, commandButtonAreaRectangle->Bottom, dropDownAreaRectangle->Left, dropDownAreaRectangle->Top, dropDownAreaRectangle->Right, dropDownAreaRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCmdU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmdU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropDownCmdU(BtnCtlsLibU::RECTANGLE* buttonRectangle)
{
	CAtlString str;
	str.Format(TEXT("CmdU_DropDown: buttonRectangle=(%i, %i)-(%i, %i)"), buttonRectangle->Left, buttonRectangle->Top, buttonRectangle->Right, buttonRectangle->Bottom);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCmdU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CmdU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCmdU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CmdU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCmdU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CmdU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmdU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCmdU(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdU_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCmdU(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCmdU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmdU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCmdU()
{
	AddLogEntry(CAtlString(TEXT("CmdU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCmdU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_Click: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_ContextMenu: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawOpt1U(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("OptU_CustomDraw: index=1, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_DblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowOpt1U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptU_DestroyedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownOpt1U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyDown: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressOpt1U(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyPress: index=1, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpOpt1U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyUp: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseDown: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseEnter: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseHover: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseLeave: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseMove: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseUp: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragDrop: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragDrop: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.opt1U->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragEnter: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragEnter: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveOpt1U(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragLeave: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragLeave: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveOpt1U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragMouseMove: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragMouseMove: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_RClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_RDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowOpt1U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptU_RecreatedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowOpt1U()
{
	AddLogEntry(CAtlString(TEXT("OptU_ResizedControlWindow: index=1")));
}

void __stdcall CMainDlg::SelectionChangedOpt1U(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("OptU_SelectionChanged: index=1, previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);

	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt2U->Selected = VARIANT_FALSE;
		opt1UWnd.ModifyStyle(0, WS_TABSTOP);
		opt2UWnd.ModifyStyle(WS_TABSTOP, 0);
	}
}

void __stdcall CMainDlg::XClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_XClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickOpt1U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_XDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_Click: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_ContextMenu: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawOpt2U(BtnCtlsLibU::CustomDrawStageConstants drawStage, BtnCtlsLibU::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibU::RECTANGLE* drawingRectangle, BtnCtlsLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("OptU_CustomDraw: index=2, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_DblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowOpt2U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptU_DestroyedControlWindow: index=2, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownOpt2U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyDown: index=2, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressOpt2U(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyPress: index=2, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpOpt2U(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptU_KeyUp: index=2, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseDown: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseEnter: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseHover: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseLeave: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseMove: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_MouseUp: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragDrop: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragDrop: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.opt2U->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragEnter: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragEnter: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveOpt2U(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragLeave: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragLeave: index=2, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveOpt2U(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptU_OLEDragMouseMove: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptU_OLEDragMouseMove: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_RClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_RDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowOpt2U(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptU_RecreatedControlWindow: index=2, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowOpt2U()
{
	AddLogEntry(CAtlString(TEXT("OptU_ResizedControlWindow: index=2")));
}

void __stdcall CMainDlg::SelectionChangedOpt2U(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("OptU_SelectionChanged: index=2, previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);

	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt1U->Selected = VARIANT_FALSE;
		opt2UWnd.ModifyStyle(0, WS_TABSTOP);
		opt1UWnd.ModifyStyle(WS_TABSTOP, 0);
	}
}

void __stdcall CMainDlg::XClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_XClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickOpt2U(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptU_XDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowFrmA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("FrmA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.frmA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveFrmA(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmA_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveFrmA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("FrmA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("FrmA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowFrmA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("FrmA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowFrmA()
{
	AddLogEntry(CAtlString(TEXT("FrmA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickFrmA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("FrmA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawChkA(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("ChkA_CustomDraw: drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowChkA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ChkA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownChkA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ChkA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressChkA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ChkA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpChkA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ChkA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.chkA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveChkA(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkA_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveChkA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ChkA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ChkA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowChkA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ChkA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowChkA()
{
	AddLogEntry(CAtlString(TEXT("ChkA_ResizedControlWindow")));
}

void __stdcall CMainDlg::SelectionStateChangedChkA(BtnCtlsLibA::SelectionStateConstants previousSelectionState, BtnCtlsLibA::SelectionStateConstants newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("ChkA_SelectionStateChanged: previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickChkA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ChkA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawCmdA(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("CmdA_CustomDraw: drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDropDownAreaSizeCmdA(BtnCtlsLibA::RECTANGLE* clientRectangle, BtnCtlsLibA::RECTANGLE* commandButtonAreaRectangle, BtnCtlsLibA::RECTANGLE* dropDownAreaRectangle, BtnCtlsLibA::CustomDropDownAreaSizeReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("CmdA_CustomDropDownAreaSize: clientRectangle=(%i, %i)-(%i, %i), commandButtonAreaRectangle=(%i, %i)-(%i, %i), dropDownAreaRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), clientRectangle->Left, clientRectangle->Top, clientRectangle->Right, clientRectangle->Bottom, commandButtonAreaRectangle->Left, commandButtonAreaRectangle->Top, commandButtonAreaRectangle->Right, commandButtonAreaRectangle->Bottom, dropDownAreaRectangle->Left, dropDownAreaRectangle->Top, dropDownAreaRectangle->Right, dropDownAreaRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCmdA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmdA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropDownCmdA(BtnCtlsLibA::RECTANGLE* buttonRectangle)
{
	CAtlString str;
	str.Format(TEXT("CmdA_DropDown: buttonRectangle=(%i, %i)-(%i, %i)"), buttonRectangle->Left, buttonRectangle->Top, buttonRectangle->Right, buttonRectangle->Bottom);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCmdA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CmdA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCmdA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CmdA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCmdA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CmdA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.cmdA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCmdA(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdA_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCmdA(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CmdA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CmdA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCmdA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CmdA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCmdA()
{
	AddLogEntry(CAtlString(TEXT("CmdA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCmdA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("CmdA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_Click: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_ContextMenu: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawOpt1A(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("OptA_CustomDraw: index=1, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_DblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowOpt1A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptA_DestroyedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownOpt1A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyDown: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressOpt1A(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyPress: index=1, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpOpt1A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyUp: index=1, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseDown: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseEnter: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseHover: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseLeave: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseMove: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseUp: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragDrop: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragDrop: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.opt1A->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragEnter: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragEnter: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveOpt1A(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragLeave: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragLeave: index=1, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveOpt1A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragMouseMove: index=1, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragMouseMove: index=1, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_RClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_RDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowOpt1A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptA_RecreatedControlWindow: index=1, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowOpt1A()
{
	AddLogEntry(CAtlString(TEXT("OptA_ResizedControlWindow: index=1")));
}

void __stdcall CMainDlg::SelectionChangedOpt1A(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("OptA_SelectionChanged: index=1, previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);

	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt2A->Selected = VARIANT_FALSE;
		opt1AWnd.ModifyStyle(0, WS_TABSTOP);
		opt2AWnd.ModifyStyle(WS_TABSTOP, 0);
	}
}

void __stdcall CMainDlg::XClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_XClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickOpt1A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_XDblClick: index=1, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_Click: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_ContextMenu: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CustomDrawOpt2A(BtnCtlsLibA::CustomDrawStageConstants drawStage, BtnCtlsLibA::CustomDrawControlStateConstants controlState, long hDC, BtnCtlsLibA::RECTANGLE* drawingRectangle, BtnCtlsLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	str.Format(TEXT("OptA_CustomDraw: index=2, drawStage=0x%X, controlState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, controlState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_DblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowOpt2A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptA_DestroyedControlWindow: index=2, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownOpt2A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyDown: index=2, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressOpt2A(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyPress: index=2, keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpOpt2A(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("OptA_KeyUp: index=2, keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseDown: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseEnter: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseHover: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseLeave: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseMove: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_MouseUp: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragDrop: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragDrop: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.opt2A->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragEnter: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragEnter: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveOpt2A(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragLeave: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragLeave: index=2, data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveOpt2A(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<BtnCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("OptA_OLEDragMouseMove: index=2, data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("OptA_OLEDragMouseMove: index=2, data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_RClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_RDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowOpt2A(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("OptA_RecreatedControlWindow: index=2, hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowOpt2A()
{
	AddLogEntry(CAtlString(TEXT("OptA_ResizedControlWindow: index=2")));
}

void __stdcall CMainDlg::SelectionChangedOpt2A(VARIANT_BOOL previousSelectionState, VARIANT_BOOL newSelectionState)
{
	CAtlString str;
	str.Format(TEXT("OptA_SelectionChanged: index=2, previousSelectionState=%i, newSelectionState=%i"), previousSelectionState, newSelectionState);

	AddLogEntry(str);

	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt1A->Selected = VARIANT_FALSE;
		opt2AWnd.ModifyStyle(0, WS_TABSTOP);
		opt1AWnd.ModifyStyle(WS_TABSTOP, 0);
	}
}

void __stdcall CMainDlg::XClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_XClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickOpt2A(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("OptA_XDblClick: index=2, button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}
