// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));
	controls.useImprovedImgLstSupportCheck.Attach(GetDlgItem(IDC_USEIMPROVEDIMGLSTSUPPORTCHECK));
	controls.iconIndexesCombo.Attach(GetDlgItem(IDC_ICONINDEXESCOMBO));
	controls.iconIndexEdit.Attach(GetDlgItem(IDC_ICONINDEXEDIT));
	controls.iconIndexEdit.SetLimitText(4);
	controls.setIconIndexButton.Attach(GetDlgItem(IDC_SETICONINDEXBUTTON));

	// setup the toolbar
	WTL::CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(pNotificationDetails->hwndFrom != controls.disabledEventsList) {
		return 0;
	}

	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	switch(pDetails->iItem) {
		case 0:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, MouseMove"))));
			break;
		case 1:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
			break;
		case 2:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
			break;
		case 3:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CustomDraw"))));
			break;
		case 4:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CustomDropDownAreaSize"))));
			break;
	}
	ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnUseImprovedImageListSupport(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(properties.hWndCheckMarksAreSetFor != controls.useImprovedImgLstSupportCheck) {
		if(controls.useImprovedImgLstSupportCheck.GetCheck() == BST_INDETERMINATE) {
			controls.useImprovedImgLstSupportCheck.SetCheck(BST_UNCHECKED);
		}
	}

	BOOL checked = (controls.useImprovedImgLstSupportCheck.GetCheck() != BST_UNCHECKED);
	controls.iconIndexesCombo.EnableWindow(checked);
	controls.iconIndexEdit.EnableWindow(checked);
	controls.setIconIndexButton.EnableWindow(FALSE);
	SetDirty(TRUE);
	return 0;
}

LRESULT CommonProperties::OnSetIconIndex(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	TCHAR pBuffer[5];
	controls.iconIndexEdit.GetWindowText(pBuffer, 5);
	LONG iconIndex;
	ConvertStringToInteger(pBuffer, iconIndex);

	LONG itemIndex = controls.iconIndexesCombo.GetCurSel();
	if(itemIndex == 0) {
		for(int i = 0; i < 6; i++) {
			properties.iconIndexes[i] = iconIndex;
		}
	} else {
		properties.iconIndexes[itemIndex - 1] = iconIndex;
	}
	controls.setIconIndexButton.EnableWindow(FALSE);
	SetDirty(TRUE);
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_ICheckBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICheckBox*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_ICommandButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICommandButton*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IFrame, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IFrame*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IOptionButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IOptionButton*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_ICheckBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("CheckBox Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICheckBox*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_ICommandButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("CommandButton Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICommandButton*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IFrame, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("Frame Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IFrame*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IOptionButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("OptionButton Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IOptionButton*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnIconIndexSelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	LONG itemIndex = controls.iconIndexesCombo.GetCurSel();
	if(itemIndex == 0) {
		controls.iconIndexEdit.SetWindowText(_T(""));
	} else {
		CAtlString s;
		s.Format(_T("%i"), properties.iconIndexes[itemIndex - 1]);
		controls.iconIndexEdit.SetWindowText(s);
	}
	controls.setIconIndexButton.EnableWindow(FALSE);
	return 0;
}

LRESULT CommonProperties::OnIconIndexChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	controls.setIconIndexButton.EnableWindow(controls.iconIndexEdit.GetWindowTextLength() > 0);
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_ICheckBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<ICheckBox*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			reinterpret_cast<ICheckBox*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			int checkState = controls.useImprovedImgLstSupportCheck.GetCheck();
			if(checkState != BST_INDETERMINATE) {
				reinterpret_cast<ICheckBox*>(pControl)->put_UseImprovedImageListSupport(BOOL2VARIANTBOOL(checkState & BST_CHECKED));
			}
			for(LONG i = 0; i < 6; i++) {
				reinterpret_cast<ICheckBox*>(pControl)->put_IconIndex(i, properties.iconIndexes[i]);
			}
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_ICommandButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<ICommandButton*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deCustomDropDownAreaSize);
			reinterpret_cast<ICommandButton*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			int checkState = controls.useImprovedImgLstSupportCheck.GetCheck();
			if(checkState != BST_INDETERMINATE) {
				reinterpret_cast<ICommandButton*>(pControl)->put_UseImprovedImageListSupport(BOOL2VARIANTBOOL(checkState & BST_CHECKED));
			}
			for(LONG i = 0; i < 6; i++) {
				reinterpret_cast<ICommandButton*>(pControl)->put_IconIndex(i, properties.iconIndexes[i]);
			}
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IFrame, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IFrame*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deClickEvents);
			reinterpret_cast<IFrame*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			int checkState = controls.useImprovedImgLstSupportCheck.GetCheck();
			if(checkState != BST_INDETERMINATE) {
				reinterpret_cast<IFrame*>(pControl)->put_UseImprovedImageListSupport(BOOL2VARIANTBOOL(checkState & BST_CHECKED));
			}
			for(LONG i = 0; i < 6; i++) {
				reinterpret_cast<IFrame*>(pControl)->put_IconIndex(i, properties.iconIndexes[i]);
			}
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IOptionButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IOptionButton*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			reinterpret_cast<IOptionButton*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			int checkState = controls.useImprovedImgLstSupportCheck.GetCheck();
			if(checkState != BST_INDETERMINATE) {
				reinterpret_cast<IOptionButton*>(pControl)->put_UseImprovedImageListSupport(BOOL2VARIANTBOOL(checkState & BST_CHECKED));
			}
			for(LONG i = 0; i < 6; i++) {
				reinterpret_cast<IOptionButton*>(pControl)->put_IconIndex(i, properties.iconIndexes[i]);
			}
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	int numberOfCheckBoxes = 0;
	int numberOfCommandButtons = 0;
	int numberOfFrames = 0;
	int numberOfOptionButtons = 0;
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		VARIANT_BOOL* pUseImprovedImageListSupport = static_cast<VARIANT_BOOL*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(VARIANT_BOOL)));
		if(pUseImprovedImageListSupport) {
			ZeroMemory(pUseImprovedImageListSupport, m_nObjects * sizeof(VARIANT_BOOL));
			ZeroMemory(properties.iconIndexes, 6 * sizeof(LONG));
			for(UINT object = 0; object < m_nObjects; ++object) {
				LPUNKNOWN pControl = NULL;
				if(m_ppUnk[object]->QueryInterface(IID_ICheckBox, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfCheckBoxes;
					reinterpret_cast<ICheckBox*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
					if(object == 0) {
						for(LONG i = 0; i < 6; i++) {
							reinterpret_cast<ICheckBox*>(pControl)->get_IconIndex(i, &properties.iconIndexes[i]);
						}
					}
					reinterpret_cast<ICheckBox*>(pControl)->get_UseImprovedImageListSupport(&pUseImprovedImageListSupport[object]);
					pControl->Release();

				} else if(m_ppUnk[object]->QueryInterface(IID_ICommandButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfCommandButtons;
					reinterpret_cast<ICommandButton*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
					if(object == 0) {
						for(LONG i = 0; i < 6; i++) {
							reinterpret_cast<ICommandButton*>(pControl)->get_IconIndex(i, &properties.iconIndexes[i]);
						}
					}
					reinterpret_cast<ICommandButton*>(pControl)->get_UseImprovedImageListSupport(&pUseImprovedImageListSupport[object]);
					pControl->Release();

				} else if(m_ppUnk[object]->QueryInterface(IID_IFrame, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfFrames;
					reinterpret_cast<IFrame*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
					if(object == 0) {
						for(LONG i = 0; i < 6; i++) {
							reinterpret_cast<IFrame*>(pControl)->get_IconIndex(i, &properties.iconIndexes[i]);
						}
					}
					reinterpret_cast<IFrame*>(pControl)->get_UseImprovedImageListSupport(&pUseImprovedImageListSupport[object]);
					pControl->Release();

				} else if(m_ppUnk[object]->QueryInterface(IID_IOptionButton, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
					++numberOfOptionButtons;
					reinterpret_cast<IOptionButton*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
					if(object == 0) {
						for(LONG i = 0; i < 6; i++) {
							reinterpret_cast<IOptionButton*>(pControl)->get_IconIndex(i, &properties.iconIndexes[i]);
						}
					}
					reinterpret_cast<IOptionButton*>(pControl)->get_UseImprovedImageListSupport(&pUseImprovedImageListSupport[object]);
					pControl->Release();
				}
			}

			// fill the controls
			LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
			properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
			controls.disabledEventsList.DeleteAllItems();
			controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
			controls.disabledEventsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(1, 0, TEXT("Click events"));
			controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
			if((numberOfCheckBoxes > 0) || (numberOfCommandButtons > 0) || (numberOfOptionButtons > 0)) {
				controls.disabledEventsList.AddItem(2, 0, TEXT("Keyboard events"));
				controls.disabledEventsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(3, 0, TEXT("CustomDraw event"));
				controls.disabledEventsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, deCustomDraw), LVIS_STATEIMAGEMASK);
				if(numberOfCommandButtons > 0) {
					controls.disabledEventsList.AddItem(4, 0, TEXT("CustomDropDownAreaSize event"));
					controls.disabledEventsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, deCustomDropDownAreaSize), LVIS_STATEIMAGEMASK);
				}
			}
			controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

			properties.hWndCheckMarksAreSetFor = controls.useImprovedImgLstSupportCheck;
			int checkState = BST_UNCHECKED;
			for(UINT object = 0; object < m_nObjects; ++object) {
				if(pUseImprovedImageListSupport[object] != VARIANT_FALSE) {
					if(checkState == BST_UNCHECKED) {
						checkState = (object == 0 ? BST_CHECKED : BST_INDETERMINATE);
					}
				} else {
					if(checkState == BST_CHECKED) {
						checkState = (object == 0 ? BST_UNCHECKED : BST_INDETERMINATE);
					}
				}
			}
			controls.useImprovedImgLstSupportCheck.SetCheck(checkState);
			BOOL dummy;
			OnUseImprovedImageListSupport(BN_CLICKED, IDC_USEIMPROVEDIMGLSTSUPPORTCHECK, controls.useImprovedImgLstSupportCheck, dummy);

			properties.hWndCheckMarksAreSetFor = NULL;

			controls.iconIndexesCombo.ResetContent();
			controls.iconIndexesCombo.AddString(_T("(All States)"));
			controls.iconIndexesCombo.AddString(_T("Normal State"));
			controls.iconIndexesCombo.AddString(_T("Hovered (\"Hot\") State"));
			controls.iconIndexesCombo.AddString(_T("Pressed State"));
			controls.iconIndexesCombo.AddString(_T("Disabled State"));
			controls.iconIndexesCombo.AddString(_T("Default Button"));
			controls.iconIndexesCombo.AddString(_T("Stylus Hot State (Tablet PC)"));
			controls.iconIndexesCombo.SetCurSel(0);
			controls.iconIndexEdit.SetWindowText(_T(""));
			controls.iconIndexEdit.SetModify(FALSE);

			HeapFree(GetProcessHeap(), 0, pUseImprovedImageListSupport);
		}
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}