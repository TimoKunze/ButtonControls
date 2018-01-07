// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


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
	if(controls.cmdSplit) {
		IDispEventImpl<IDC_CMDSPLIT, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.cmdSplit);
		controls.cmdSplit.Release();
	}
	if(controls.cmdElevation) {
		IDispEventImpl<IDC_CMDELEVATION, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.cmdElevation);
		controls.cmdElevation.Release();
	}
	if(controls.cmdLink) {
		IDispEventImpl<IDC_CMDLNK, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.cmdLink);
		controls.cmdLink.Release();
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
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	cmdSplitWnd.SubclassWindow(GetDlgItem(IDC_CMDSPLIT));
	cmdSplitWnd.QueryControl(__uuidof(ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdSplit));
	if(controls.cmdSplit) {
		IDispEventImpl<IDC_CMDSPLIT, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.cmdSplit);
	}
	cmdElevationWnd.SubclassWindow(GetDlgItem(IDC_CMDELEVATION));
	cmdElevationWnd.QueryControl(__uuidof(ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdElevation));
	if(controls.cmdElevation) {
		IDispEventImpl<IDC_CMDELEVATION, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.cmdElevation);
	}
	cmdLinkWnd.SubclassWindow(GetDlgItem(IDC_CMDLNK));
	cmdLinkWnd.QueryControl(__uuidof(ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdLink));
	if(controls.cmdLink) {
		IDispEventImpl<IDC_CMDLNK, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.cmdLink);
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.cmdSplit) {
		controls.cmdSplit->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

HMENU CMainDlg::CreateDropDownMenu(VARIANT_BOOL checked)
{
	CMenu menu;
	menu.CreatePopupMenu();
	CMenuItemInfo mii;
	mii.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
	mii.fState = (checked == VARIANT_FALSE ? MFS_UNCHECKED : MFS_CHECKED);
	mii.dwTypeData = TEXT("&Split Line");
	mii.cch = lstrlen(mii.dwTypeData);
	mii.wID = 1;
	menu.InsertMenuItem(1, TRUE, &mii);

	return menu.Detach();
}

void __stdcall CMainDlg::ClickCmdsplit(short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	if(controls.cmdSplit->GetShowSplitLine() == VARIANT_FALSE) {
		CMenu menu = CreateDropDownMenu(controls.cmdSplit->GetShowSplitLine());
		RECT rc;
		cmdSplitWnd.GetWindowRect(&rc);
		POINT ptMenu = {rc.right, rc.bottom};

		TPMPARAMS tpmexParams = {0};
		tpmexParams.cbSize = sizeof(TPMPARAMS);
		tpmexParams.rcExclude = rc;
		InflateRect(&tpmexParams.rcExclude, -1, -1);
		controls.cmdSplit->PutPushed(VARIANT_TRUE);
		// NOT COMPATIBLE WITH NT4: UINT i = menu.TrackPopupMenuEx(TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdSplit->GethWnd())), &tpmexParams);
		UINT i = TrackPopupMenuEx(menu, TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdSplit->GethWnd())), &tpmexParams);
		controls.cmdSplit->PutPushed(VARIANT_FALSE);
		if(i == 1) {
			controls.cmdSplit->PutShowSplitLine(controls.cmdSplit->GetShowSplitLine() == VARIANT_FALSE ? VARIANT_TRUE : VARIANT_FALSE);
		}
		menu.DestroyMenu();
	} else {
		MessageBox(TEXT("Hello World!"), TEXT("VistaFeatures"));
	}
}

void __stdcall CMainDlg::DropDownCmdsplit(RECTANGLE* buttonRectangle)
{
	CMenu menu = CreateDropDownMenu(controls.cmdSplit->GetShowSplitLine());
	POINT ptMenu = {buttonRectangle->Right, buttonRectangle->Bottom};
	::ClientToScreen(static_cast<HWND>(LongToHandle(controls.cmdSplit->GethWnd())), &ptMenu);

	RECT rc;
	cmdSplitWnd.GetWindowRect(&rc);
	TPMPARAMS tpmexParams = {0};
	tpmexParams.cbSize = sizeof(TPMPARAMS);
	tpmexParams.rcExclude = rc;
	InflateRect(&tpmexParams.rcExclude, -1, -1);
	controls.cmdSplit->PutDropDownPushed(VARIANT_TRUE);
	// NOT COMPATIBLE WITH NT4: UINT i = menu.TrackPopupMenuEx(TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdSplit->GethWnd())), &tpmexParams);
	UINT i = TrackPopupMenuEx(menu, TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdSplit->GethWnd())), &tpmexParams);
	controls.cmdSplit->PutDropDownPushed(VARIANT_FALSE);
	if(i == 1) {
		controls.cmdSplit->PutShowSplitLine(controls.cmdSplit->GetShowSplitLine() == VARIANT_FALSE ? VARIANT_TRUE : VARIANT_FALSE);
	}
	menu.DestroyMenu();
}

void __stdcall CMainDlg::ClickCmdelevation(short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	ShellExecute(*this, TEXT("open"), TEXT("hdwwiz.exe"), TEXT(""), NULL, SW_SHOWDEFAULT);
}

void __stdcall CMainDlg::ClickCmdlink(short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	MessageBox(TEXT("Hello World!"), TEXT("VistaFeatures"));
}
