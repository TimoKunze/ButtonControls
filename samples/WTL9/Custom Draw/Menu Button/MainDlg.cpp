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
	if(controls.cmdBtnU) {
		DispEventUnadvise(controls.cmdBtnU);
		controls.cmdBtnU.Release();
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

	CLogFont lf = AtlGetDefaultGuiFont();
	lstrcpy(lf.lfFaceName, TEXT("Webdings"));
	lf.lfCharSet = SYMBOL_CHARSET;
	lf.lfHeight = 18;
	webdingsFont.CreateFontIndirect(&lf);

	cmdBtnUWnd.SubclassWindow(GetDlgItem(IDC_CMDBTNU));
	cmdBtnUWnd.QueryControl(__uuidof(ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmdBtnU));
	if(controls.cmdBtnU) {
		DispEventAdvise(controls.cmdBtnU);
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.cmdBtnU) {
		controls.cmdBtnU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::ClickCmdbtnu(short /*button*/, short /*shift*/, long x, long /*y*/)
{
	RECT rc = {0};
	cmdBtnUWnd.GetClientRect(&rc);
	if(x < (rc.right - rc.left) - 22) {
		MessageBox(TEXT("You clicked the button."), TEXT("MenuButton"));
		return;
	}

	CMenu menu;
	menu.CreatePopupMenu();
	CMenuItemInfo mii;
	mii.fMask = MIIM_TYPE | MIIM_ID;
	mii.fType = MFT_STRING;
	mii.dwTypeData = reinterpret_cast<LPTSTR>(malloc(15 * sizeof(TCHAR)));
	for(UINT i = 1; i <= 5; ++i) {
		wsprintf(mii.dwTypeData, TEXT("Entry &%i"), i);
		mii.cch = lstrlen(mii.dwTypeData);
		mii.wID = i;
		menu.InsertMenuItem(i, TRUE, &mii);
	}
	free(mii.dwTypeData);

	cmdBtnUWnd.GetWindowRect(&rc);
	POINT ptMenu = {rc.right, rc.bottom};

	TPMPARAMS tpmexParams = {0};
	tpmexParams.cbSize = sizeof(TPMPARAMS);
	tpmexParams.rcExclude = rc;
	InflateRect(&tpmexParams.rcExclude, -1, -1);
	UINT i = TrackPopupMenuEx(menu, TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdBtnU->GethWnd())), &tpmexParams);
	// NOT COMPATIBLE WITH NT4: UINT i = menu.TrackPopupMenuEx(TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERTICAL | TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, static_cast<HWND>(LongToHandle(controls.cmdBtnU->GethWnd())), &tpmexParams);

	if(i > 0) {
		CAtlString str;
		str.Format(TEXT("You clicked entry %i."), i);
		MessageBox(str, TEXT("MenuButton"));
	}
}

void __stdcall CMainDlg::CustomDrawCmdbtnu(CustomDrawStageConstants drawStage, CustomDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle, CustomDrawReturnValuesConstants* furtherProcessing)
{
	#define ARROW_CHAR TEXT("\x36")

	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = cdrvNotifyPostPaint;
			break;

		case cdsPostPaint: {
			RECT rcLine = {drawingRectangle->Right - 22, drawingRectangle->Top + 6, drawingRectangle->Right - 20, drawingRectangle->Bottom - 6};
			BOOL bDisabled = ((controlState & (cdcsDisabled | cdcsGrayed)) != 0);
			CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
			dc.Draw3dRect(&rcLine, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_BTNHIGHLIGHT));
			HFONT hPreviousFont = dc.SelectFont(webdingsFont);
			RECT rcArrow = {drawingRectangle->Right - 18, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom};
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(GetSysColor(bDisabled ? COLOR_GRAYTEXT : COLOR_BTNTEXT));
			dc.DrawText(ARROW_CHAR, 1, &rcArrow, DT_SINGLELINE | DT_LEFT | DT_VCENTER);
			dc.SelectFont(hPreviousFont);
			*furtherProcessing = cdrvDoDefault;
			break;
		}
	}
}