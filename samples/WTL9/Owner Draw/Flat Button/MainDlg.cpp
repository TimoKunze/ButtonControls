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
	if(controls.cmd) {
		IDispEventImpl<IDC_CMD, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.cmd);
		controls.cmd.Release();
	}
	if(controls.chk) {
		IDispEventImpl<IDC_CHK, CMainDlg, &__uuidof(_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.chk);
		controls.chk.Release();
	}
	if(controls.opt1) {
		IDispEventImpl<IDC_OPT1, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.opt1);
		controls.opt1.Release();
	}
	if(controls.opt2) {
		IDispEventImpl<IDC_OPT2, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventUnadvise(controls.opt2);
		controls.opt2.Release();
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

	cmdWnd.SubclassWindow(GetDlgItem(IDC_CMD));
	cmdWnd.QueryControl(__uuidof(ICommandButton), reinterpret_cast<LPVOID*>(&controls.cmd));
	if(controls.cmd) {
		IDispEventImpl<IDC_CMD, CMainDlg, &__uuidof(_ICommandButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.cmd);
	}
	chkWnd.SubclassWindow(GetDlgItem(IDC_CHK));
	chkWnd.QueryControl(__uuidof(ICheckBox), reinterpret_cast<LPVOID*>(&controls.chk));
	if(controls.chk) {
		IDispEventImpl<IDC_CHK, CMainDlg, &__uuidof(_ICheckBoxEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.chk);
	}
	opt1Wnd.SubclassWindow(GetDlgItem(IDC_OPT1));
	opt1Wnd.QueryControl(__uuidof(IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt1));
	if(controls.opt1) {
		IDispEventImpl<IDC_OPT1, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.opt1);
	}
	opt2Wnd.SubclassWindow(GetDlgItem(IDC_OPT2));
	opt2Wnd.QueryControl(__uuidof(IOptionButton), reinterpret_cast<LPVOID*>(&controls.opt2));
	if(controls.opt2) {
		IDispEventImpl<IDC_OPT2, CMainDlg, &__uuidof(_IOptionButtonEvents), &LIBID_BtnCtlsLibU, 1, 10>::DispEventAdvise(controls.opt2);
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.cmd) {
		controls.cmd->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::OwnerDrawCmd(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle)
{
	CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
	if(requiredAction & odaDrawEntire) {
		// draw background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		COLORREF clr;
		OleTranslateColor(controls.cmd->BackColor, NULL, &clr);
		brush.CreateSolidBrush(clr);
		dc.FillRect(&rc, brush);

		// draw border
		BOOL sunken = FALSE;
		BOOL raised = FALSE;
		if(((controlState & odcsHot) == odcsHot) || ((controlState & odcsFocus) == odcsFocus)) {
			if(controlState & odcsPushed) {
				sunken = TRUE;
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				raised = TRUE;
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		}

		// draw text
		UINT fmt = DT_CENTER | DT_VCENTER | DT_EDITCONTROL;
		if(controls.cmd->RightToLeft & rtlText) {
			fmt |= DT_RTLREADING;
		}
		if(controls.cmd->MultiLine == VARIANT_FALSE) {
			fmt |= DT_SINGLELINE;
		}
		if(controlState & odcsNoAccelerator) {
			fmt |= DT_HIDEPREFIX;
		}
		dc.DrawText(COLE2CT(controls.cmd->Text), controls.cmd->Text.length(), &rc, fmt | DT_CALCRECT);
		SIZE sz = rc.Size();
		rc.left = drawingRectangle->Left + (drawingRectangle->Right - drawingRectangle->Left - sz.cx) / 2;
		rc.top = drawingRectangle->Top + (drawingRectangle->Bottom - drawingRectangle->Top - sz.cy) / 2;
		if(sunken) {
			rc.left++;
			rc.top++;
		} else if(raised) {
			/*rc.left--;
			rc.top--;*/
		}
		rc.right = rc.left + sz.cx;
		rc.bottom = rc.top + sz.cy;

		dc.SetBkMode(TRANSPARENT);
		dc.DrawText(COLE2CT(controls.cmd->Text), controls.cmd->Text.length(), &rc, fmt);

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.SetRect(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}

	if(requiredAction & odaFocusChanged) {
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		// ensure border is drawn
		if(controlState & odcsFocus) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected)) {
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		}

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}
}

void CMainDlg::ClickChk(short /*button*/, short /*shift*/, long /*x*/, long /*y*/)
{
	switch(controls.chk->SelectionState) {
		case ssUnchecked:
			controls.chk->SelectionState = ssChecked;
			break;
		case ssChecked:
			controls.chk->SelectionState = ssUnchecked;
			break;
		case ssIndeterminate:
			controls.chk->SelectionState = ssUnchecked;
			break;
	}
}

void CMainDlg::OwnerDrawChk(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle)
{
	CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
	if(requiredAction & odaDrawEntire) {
		// draw background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		if(((controlState & odcsSelected) == odcsSelected) && ((controlState & odcsHot) == 0) && ((controlState & odcsDisabled) == 0)) {
			CDC memDC;
			memDC.CreateCompatibleDC(dc);
			CBitmap memBMP;
			memBMP.CreateCompatibleBitmap(dc, 2, 2);
			HBITMAP hPreviousBitmap = memDC.SelectBitmap(memBMP);
			memDC.SetPixel(0, 0, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(1, 1, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(0, 1, GetSysColor(COLOR_3DHILIGHT));
			memDC.SetPixel(1, 0, GetSysColor(COLOR_3DHILIGHT));
			brush.CreatePatternBrush(memBMP);
			memDC.SelectBitmap(hPreviousBitmap);
		} else {
			COLORREF clr;
			OleTranslateColor(controls.chk->BackColor, NULL, &clr);
			brush.CreateSolidBrush(clr);
		}
		dc.FillRect(&rc, brush);

		// draw border
		BOOL sunken = FALSE;
		BOOL raised = FALSE;
		if(((controlState & odcsHot) == odcsHot) || ((controlState & odcsFocus) == odcsFocus)) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected) || ((controlState & odcsIndeterminate) == odcsIndeterminate)) {
				sunken = TRUE;
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				raised = TRUE;
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		} else if(((controlState & odcsSelected) == odcsSelected) || ((controlState & odcsIndeterminate) == odcsIndeterminate)) {
			sunken = TRUE;
			dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
		}

		// draw text
		UINT fmt = DT_CENTER | DT_VCENTER | DT_EDITCONTROL;
		if(controls.chk->RightToLeft & rtlText) {
			fmt |= DT_RTLREADING;
		}
		if(controls.chk->MultiLine == VARIANT_FALSE) {
			fmt |= DT_SINGLELINE;
		}
		if(controlState & odcsNoAccelerator) {
			fmt |= DT_HIDEPREFIX;
		}
		dc.DrawText(COLE2CT(controls.chk->Text), controls.chk->Text.length(), &rc, fmt | DT_CALCRECT);
		SIZE sz = rc.Size();
		rc.left = drawingRectangle->Left + (drawingRectangle->Right - drawingRectangle->Left - sz.cx) / 2;
		rc.top = drawingRectangle->Top + (drawingRectangle->Bottom - drawingRectangle->Top - sz.cy) / 2;
		if(sunken) {
			rc.left++;
			rc.top++;
		} else if(raised) {
			/*rc.left--;
			rc.top--;*/
		}
		rc.right = rc.left + sz.cx;
		rc.bottom = rc.top + sz.cy;

		dc.SetBkMode(TRANSPARENT);
		dc.DrawText(COLE2CT(controls.chk->Text), controls.chk->Text.length(), &rc, fmt);

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.SetRect(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}

	if(requiredAction & odaFocusChanged) {
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		// ensure border is drawn
		if(controlState & odcsFocus) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected) || ((controlState & odcsIndeterminate) == odcsIndeterminate)) {
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		}

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}
}

void CMainDlg::OwnerDrawOpt1(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle)
{
	CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
	if(requiredAction & odaDrawEntire) {
		// draw background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		if(((controlState & odcsSelected) == odcsSelected) && ((controlState & odcsHot) == 0) && ((controlState & odcsDisabled) == 0)) {
			CDC memDC;
			memDC.CreateCompatibleDC(dc);
			CBitmap memBMP;
			memBMP.CreateCompatibleBitmap(dc, 2, 2);
			HBITMAP hPreviousBitmap = memDC.SelectBitmap(memBMP);
			memDC.SetPixel(0, 0, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(1, 1, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(0, 1, GetSysColor(COLOR_3DHILIGHT));
			memDC.SetPixel(1, 0, GetSysColor(COLOR_3DHILIGHT));
			brush.CreatePatternBrush(memBMP);
			memDC.SelectBitmap(hPreviousBitmap);
		} else {
			COLORREF clr;
			OleTranslateColor(controls.opt1->BackColor, NULL, &clr);
			brush.CreateSolidBrush(clr);
		}
		dc.FillRect(&rc, brush);

		// draw border
		BOOL sunken = FALSE;
		BOOL raised = FALSE;
		if(((controlState & odcsHot) == odcsHot) || ((controlState & odcsFocus) == odcsFocus)) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected)) {
				sunken = TRUE;
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				raised = TRUE;
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		} else if(controlState & odcsSelected) {
			sunken = TRUE;
			dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
		}

		// draw text
		UINT fmt = DT_CENTER | DT_VCENTER | DT_EDITCONTROL;
		if(controls.opt1->RightToLeft & rtlText) {
			fmt |= DT_RTLREADING;
		}
		if(controls.opt1->MultiLine == VARIANT_FALSE) {
			fmt |= DT_SINGLELINE;
		}
		if(controlState & odcsNoAccelerator) {
			fmt |= DT_HIDEPREFIX;
		}
		dc.DrawText(COLE2CT(controls.opt1->Text), controls.opt1->Text.length(), &rc, fmt | DT_CALCRECT);
		SIZE sz = rc.Size();
		rc.left = drawingRectangle->Left + (drawingRectangle->Right - drawingRectangle->Left - sz.cx) / 2;
		rc.top = drawingRectangle->Top + (drawingRectangle->Bottom - drawingRectangle->Top - sz.cy) / 2;
		if(sunken) {
			rc.left++;
			rc.top++;
		} else if(raised) {
			/*rc.left--;
			rc.top--;*/
		}
		rc.right = rc.left + sz.cx;
		rc.bottom = rc.top + sz.cy;

		dc.SetBkMode(TRANSPARENT);
		dc.DrawText(COLE2CT(controls.opt1->Text), controls.opt1->Text.length(), &rc, fmt);

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.SetRect(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}

	if(requiredAction & odaFocusChanged) {
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		// ensure border is drawn
		if(controlState & odcsFocus) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected)) {
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		}

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}
}

void CMainDlg::SelectionChanged1(VARIANT_BOOL /*previousSelectionState*/, VARIANT_BOOL newSelectionState)
{
	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt2->Selected = VARIANT_FALSE;
		opt1Wnd.ModifyStyle(0, WS_TABSTOP);
		opt2Wnd.ModifyStyle(WS_TABSTOP, 0);
	}
}

void CMainDlg::OwnerDrawOpt2(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, long hDC, RECTANGLE* drawingRectangle)
{
	CDCHandle dc = static_cast<HDC>(LongToHandle(hDC));
	if(requiredAction & odaDrawEntire) {
		// draw background
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		CBrush brush;
		if(((controlState & odcsSelected) == odcsSelected) && ((controlState & odcsHot) == 0) && ((controlState & odcsDisabled) == 0)) {
			CDC memDC;
			memDC.CreateCompatibleDC(dc);
			CBitmap memBMP;
			memBMP.CreateCompatibleBitmap(dc, 2, 2);
			HBITMAP hPreviousBitmap = memDC.SelectBitmap(memBMP);
			memDC.SetPixel(0, 0, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(1, 1, GetSysColor(COLOR_3DLIGHT));
			memDC.SetPixel(0, 1, GetSysColor(COLOR_3DHILIGHT));
			memDC.SetPixel(1, 0, GetSysColor(COLOR_3DHILIGHT));
			brush.CreatePatternBrush(memBMP);
			memDC.SelectBitmap(hPreviousBitmap);
		} else {
			COLORREF clr;
			OleTranslateColor(controls.opt2->BackColor, NULL, &clr);
			brush.CreateSolidBrush(clr);
		}
		dc.FillRect(&rc, brush);

		// draw border
		BOOL sunken = FALSE;
		BOOL raised = FALSE;
		if(((controlState & odcsHot) == odcsHot) || ((controlState & odcsFocus) == odcsFocus)) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected)) {
				sunken = TRUE;
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				raised = TRUE;
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		} else if(controlState & odcsSelected) {
			sunken = TRUE;
			dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
		}

		// draw text
		UINT fmt = DT_CENTER | DT_VCENTER | DT_EDITCONTROL;
		if(controls.opt2->RightToLeft & rtlText) {
			fmt |= DT_RTLREADING;
		}
		if(controls.opt2->MultiLine == VARIANT_FALSE) {
			fmt |= DT_SINGLELINE;
		}
		if(controlState & odcsNoAccelerator) {
			fmt |= DT_HIDEPREFIX;
		}
		dc.DrawText(COLE2CT(controls.opt2->Text), controls.opt2->Text.length(), &rc, fmt | DT_CALCRECT);
		SIZE sz = rc.Size();
		rc.left = drawingRectangle->Left + (drawingRectangle->Right - drawingRectangle->Left - sz.cx) / 2;
		rc.top = drawingRectangle->Top + (drawingRectangle->Bottom - drawingRectangle->Top - sz.cy) / 2;
		if(sunken) {
			rc.left++;
			rc.top++;
		} else if(raised) {
			/*rc.left--;
			rc.top--;*/
		}
		rc.right = rc.left + sz.cx;
		rc.bottom = rc.top + sz.cy;

		dc.SetBkMode(TRANSPARENT);
		dc.DrawText(COLE2CT(controls.opt2->Text), controls.opt2->Text.length(), &rc, fmt);

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.SetRect(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}

	if(requiredAction & odaFocusChanged) {
		WTL::CRect rc(drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
		// ensure border is drawn
		if(controlState & odcsFocus) {
			if(((controlState & odcsPushed) == odcsPushed) || ((controlState & odcsSelected) == odcsSelected)) {
				dc.DrawEdge(&rc, BDR_SUNKENINNER, BF_RECT);
			} else {
				dc.DrawEdge(&rc, BDR_RAISEDINNER, BF_RECT);
			}
		}

		// draw focus rectangle
		if(((controlState & odcsFocus) == odcsFocus) && ((controlState & odcsNoFocusRectangle) == 0)) {
			rc.InflateRect(-3, -3);
			dc.DrawFocusRect(&rc);
		}
	}
}

void CMainDlg::SelectionChanged2(VARIANT_BOOL /*previousSelectionState*/, VARIANT_BOOL newSelectionState)
{
	/* The COM container the controls sit in, is a very lightweight one. It doesn't recognize the exclusivity
	   of the OptionButtons' Selected property. */
	if(newSelectionState != VARIANT_FALSE) {
		controls.opt1->Selected = VARIANT_FALSE;
		opt2Wnd.ModifyStyle(0, WS_TABSTOP);
		opt1Wnd.ModifyStyle(WS_TABSTOP, 0);
	}
}
