// Frame.cpp: Superclasses Button.

#include "stdafx.h"
#include "Frame.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC Frame::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


Frame::Frame()
{
	pSimpleFrameSite = NULL;

	properties.font.InitializePropertyWatcher(this, DISPID_FRM_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_FRM_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	SetRect(&currentIconRectangle, 0, 0, 0, 0);
	SetRect(&currentTextRectangle, 0, 0, 0, 0);

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
				CTheme themingEngine;
				themingEngine.OpenThemeData(*this, VSCLASS_BUTTON);
				flags.themeIsNull = themingEngine.IsThemeNull();
			}
			FreeLibrary(hThemeDLL);
		}
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP Frame::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IFrame, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP Frame::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<Frame>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		LONG iconIndexes[6];
		ZeroMemory(iconIndexes, 6 * sizeof(LONG));

		hr = pPropertyBag->Read(OLESTR("IconIndexes"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_I4) && propertyValue.parray) {
				LONG l = 0;
				SafeArrayGetLBound(propertyValue.parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(propertyValue.parray, 1, &u);
				if(u >= l) {
					ULONG arraySize = u - l + 1;
					if(arraySize <= 6) {
						for(LONG i = l; i <= u; ++i) {
							SafeArrayGetElement(propertyValue.parray, &i, &iconIndexes[i - l]);
						}
					} else {
						// invalid property value - raise VB runtime error 380
						hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
				}
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		for(LONG i = 0; i < 6; i++) {
			put_IconIndex(i, iconIndexes[i]);
		}
		VariantClear(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP Frame::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<Frame>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_I4;
		propertyValue.parray = SafeArrayCreateVectorEx(VT_I4, 0, 6, NULL);
		for(LONG i = 0; i < 6; i++) {
			SafeArrayPutElement(propertyValue.parray, &i, &properties.iconIndexes[i]);
		}
		hr = pPropertyBag->Write(OLESTR("IconIndexes"), &propertyValue);
		VariantClear(&propertyValue);

		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP Frame::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 16 VT_I4 properties...
	pSize->QuadPart += 16 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 7 VT_BOOL properties...
	pSize->QuadPart += 7 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 1 VT_BSTR property...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);
	// ...and 1 VT_ARRAY | VT_I4 property...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + 6 * sizeof(LONG);

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP Frame::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x09090909/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0700 || signature == 0x0710 || signature == 0x0800 || signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0102) {
			return E_FAIL;
		}
		LONG subSignature = 0;
		if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
			return hr;
		}
		if(subSignature != 0x03030303/*4x 0x03 (-> Frame)*/) {
			return E_FAIL;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL borderVisible = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ContentTypeConstants contentType = static_cast<ContentTypeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IconAlignmentConstants iconAlignment = static_cast<IconAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS iconMarginBottom = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS iconMarginLeft = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS iconMarginRight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS iconMarginTop = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StyleConstants style = static_cast<StyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR text;
	if(FAILED(hr = text.ReadFromStream(pStream))) {
		return hr;
	}
	if(!text) {
		text = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;

	LONG iconIndexes[6];
	VARIANT_BOOL useImprovedImageListSupport = VARIANT_FALSE;
	if(version >= 0x0102) {
		if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != (VT_ARRAY | VT_I4))) {
			return hr;
		}
		ULONG c = 0;
		if(FAILED(hr = pStream->Read(&c, sizeof(ULONG), NULL))) {
			return hr;
		}
		if(c != 6) {
			return E_FAIL;
		}
		// read the elements
		for(ULONG i = 0; i < c; i++) {
			if(FAILED(hr = pStream->Read(&iconIndexes[i], sizeof(LONG), NULL))) {
				return hr;
			}
		}

		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		useImprovedImageListSupport = propertyValue.boolVal;
	}


	/*hr = put_UseImprovedImageListSupport(useImprovedImageListSupport);
	if(FAILED(hr)) {
		return hr;
	}*/
	UINT b = VARIANTBOOL2BOOL(useImprovedImageListSupport);
	if(properties.useImprovedImageListSupport != b) {
		properties.useImprovedImageListSupport = b;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderVisible(borderVisible);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ContentType(contentType);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HAlignment(hAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IconAlignment(iconAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	for(LONG i = 0; i < 6; i++) {
		hr = put_IconIndex(i, iconIndexes[i]);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_IconMarginBottom(iconMarginBottom);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IconMarginLeft(iconMarginLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IconMarginRight(iconMarginRight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IconMarginTop(iconMarginTop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Style(style);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP Frame::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x09090909/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x03030303/*4x 0x03 (-> Frame)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.borderVisible);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.contentType;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconMarginBottom;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconMarginLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconMarginRight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.iconMarginTop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.style;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	// store some marker
	vt = VT_ARRAY | VT_I4;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	// store the number of elements
	ULONG c = 6;
	if(FAILED(hr = pStream->Write(&c, sizeof(ULONG), NULL))) {
		return hr;
	}
	// store the elements
	for(ULONG i = 0; i < c; i++) {
		if(FAILED(hr = pStream->Write(&properties.iconIndexes[i], sizeof(LONG), NULL))) {
			return hr;
		}
	}

	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useImprovedImageListSupport);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND Frame::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_STANDARD_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<Frame>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT Frame::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsWindow()) {
		/* Calling PrintClient causes huge drawing issues if the control contains our windowless label. It
		   does not seem to be a good idea to trigger new drawing messages from within a drawing cycle. */
		//PrintClient(drawInfo.hdcDraw, /*PRF_CHILDREN | */PRF_CLIENT | PRF_NONCLIENT);
		POINT previousViewPort = {0};
		if(GetExStyle() & WS_EX_LAYOUTRTL) {
			SetViewportOrgEx(drawInfo.hdcDraw, drawInfo.prcBounds->right + 1, drawInfo.prcBounds->top, &previousViewPort);
		} else {
			SetViewportOrgEx(drawInfo.hdcDraw, drawInfo.prcBounds->left, drawInfo.prcBounds->top, &previousViewPort);
		}

		BOOL wasHandled = TRUE;
		OnEraseBkGnd(WM_ERASEBKGND, reinterpret_cast<WPARAM>(drawInfo.hdcDraw), 0, wasHandled);

		if(properties.style == sOwnerDrawn) {
			RECT clientRectangle;
			GetClientRect(&clientRectangle);
			if(!wasHandled) {
				wasHandled = TRUE;
				HBRUSH hBrush = reinterpret_cast<HBRUSH>(OnReflectedCtlColorBtn(OCM_CTLCOLORBTN, reinterpret_cast<WPARAM>(drawInfo.hdcDraw), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)), wasHandled));
				if(wasHandled && hBrush) {
					FillRect(drawInfo.hdcDraw, &clientRectangle, hBrush);
				}
			}

			OwnerDrawControlStateConstants controlState = odcsNormal;
			DWORD state = SendMessage(BM_GETSTATE, 0, 0);
			DWORD uiState = SendMessage(WM_QUERYUISTATE, 0, 0);
			if(state & BST_CHECKED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsSelected);
			}
			if(GetStyle() & WS_DISABLED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsDisabled);
			}
			if(GetFocus() == *this) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsFocus);
			}
			if(uiState & UISF_HIDEACCEL) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoAccelerator);
			}
			if(uiState & UISF_HIDEFOCUS) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoFocusRectangle);
			}
			if(state & BST_PUSHED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsPushed);
			}
			if(state & BST_INDETERMINATE) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsIndeterminate);
			}
			if(mouseStatus.hoveredControl) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsHot);
			}

			Raise_OwnerDraw(odaDrawEntire, controlState, HandleToLong(drawInfo.hdcDraw), reinterpret_cast<RECTANGLE*>(&clientRectangle));
		} else {
			if(!wasHandled) {
				wasHandled = TRUE;
				if(properties.style == sOwnerDrawn) {
					OnReflectedCtlColorBtn(OCM_CTLCOLORBTN, reinterpret_cast<WPARAM>(drawInfo.hdcDraw), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)), wasHandled);
				} else {
					OnReflectedCtlColorStatic(OCM_CTLCOLORSTATIC, reinterpret_cast<WPARAM>(drawInfo.hdcDraw), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)), wasHandled);
				}
			}
			DrawFrame(drawInfo.hdcDraw);
		}

		SetViewportOrgEx(drawInfo.hdcDraw, previousViewPort.x, previousViewPort.y, NULL);
		return S_OK;
	} else if(IsInDesignMode()) {
		CAtlString text = TEXT("Frame ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void Frame::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP Frame::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	pSimpleFrameSite = NULL;
	if(pClientSite) {
		pClientSite->QueryInterface(IID_ISimpleFrameSite, reinterpret_cast<LPVOID*>(&pSimpleFrameSite));
	}

	HRESULT hr = CComControl<Frame>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	if(properties.resetTextToName) {
		properties.resetTextToName = FALSE;

		BSTR buffer = SysAllocString(L"");
		if(SUCCEEDED(GetAmbientDisplayName(buffer))) {
			properties.text.AssignBSTR(buffer);
		} else {
			SysFreeString(buffer);
		}
	}

	return hr;
}

STDMETHODIMP Frame::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL Frame::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT(dialogCode == DLGC_BUTTON);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<Frame>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP Frame::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP Frame::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP Frame::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP Frame::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP Frame::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP Frame::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_FRM_APPEARANCE:
		case DISPID_FRM_BORDERSTYLE:
		case DISPID_FRM_BORDERVISIBLE:
		case DISPID_FRM_CONTENTTYPE:
		case DISPID_FRM_HALIGNMENT:
		case DISPID_FRM_ICONALIGNMENT:
		case DISPID_FRM_ICONINDEX:
		case DISPID_FRM_ICONMARGINBOTTOM:
		case DISPID_FRM_ICONMARGINLEFT:
		case DISPID_FRM_ICONMARGINRIGHT:
		case DISPID_FRM_ICONMARGINTOP:
		case DISPID_FRM_IMAGE:
		case DISPID_FRM_MOUSEICON:
		case DISPID_FRM_MOUSEPOINTER:
		case DISPID_FRM_STYLE:
		case DISPID_FRM_TEXT:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_FRM_DISABLEDEVENTS:
		case DISPID_FRM_DONTREDRAW:
		case DISPID_FRM_HOVERTIME:
		case DISPID_FRM_RIGHTTOLEFT:
		case DISPID_FRM_USEIMPROVEDIMAGELISTSUPPORT:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_FRM_BACKCOLOR:
		case DISPID_FRM_FORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_FRM_APPID:
		case DISPID_FRM_APPNAME:
		case DISPID_FRM_APPSHORTNAME:
		case DISPID_FRM_BUILD:
		case DISPID_FRM_CHARSET:
		case DISPID_FRM_HIMAGELIST:
		case DISPID_FRM_HWND:
		case DISPID_FRM_ISRELEASE:
		case DISPID_FRM_PROGRAMMER:
		case DISPID_FRM_TESTER:
		case DISPID_FRM_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_FRM_REGISTERFOROLEDRAGDROP:
		case DISPID_FRM_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_FRM_FONT:
		case DISPID_FRM_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_FRM_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString Frame::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString Frame::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString Frame::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ButtonControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString Frame::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString Frame::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString Frame::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP Frame::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_FRM_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_FRM_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_FRM_CONTENTTYPE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.contentType), description);
			break;
		case DISPID_FRM_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_FRM_ICONALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.iconAlignment), description);
			break;
		case DISPID_FRM_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_FRM_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_FRM_STYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.style), description);
			break;
		case DISPID_FRM_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		default:
			return IPerPropertyBrowsingImpl<Frame>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP Frame::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_FRM_BORDERSTYLE:
			c = 2;
			break;
		case DISPID_FRM_APPEARANCE:
		case DISPID_FRM_CONTENTTYPE:
		case DISPID_FRM_HALIGNMENT:
		case DISPID_FRM_STYLE:
			c = 3;
			break;
		case DISPID_FRM_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_FRM_ICONALIGNMENT:
			c = 5;
			break;
		case DISPID_FRM_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<Frame>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_FRM_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP Frame::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_FRM_APPEARANCE:
		case DISPID_FRM_BORDERSTYLE:
		case DISPID_FRM_CONTENTTYPE:
		case DISPID_FRM_HALIGNMENT:
		case DISPID_FRM_ICONALIGNMENT:
		case DISPID_FRM_MOUSEPOINTER:
		case DISPID_FRM_RIGHTTOLEFT:
		case DISPID_FRM_STYLE:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<Frame>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP Frame::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_FRM_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<Frame>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT Frame::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_FRM_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_CONTENTTYPE:
			switch(cookie) {
				case ctText:
					description = GetResStringWithNumber(IDP_CONTENTTYPETEXT, ctText);
					break;
				case ctBitmap:
					description = GetResStringWithNumber(IDP_CONTENTTYPEBITMAP, ctBitmap);
					break;
				case ctIcon:
					description = GetResStringWithNumber(IDP_CONTENTTYPEICON, ctIcon);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_ICONALIGNMENT:
			switch(cookie) {
				case ialLeft:
					description = GetResStringWithNumber(IDP_ICONALIGNMENTLEFT, ialLeft);
					break;
				case ialRight:
					description = GetResStringWithNumber(IDP_ICONALIGNMENTRIGHT, ialRight);
					break;
				case ialTop:
					description = GetResStringWithNumber(IDP_ICONALIGNMENTTOP, ialTop);
					break;
				case ialBottom:
					description = GetResStringWithNumber(IDP_ICONALIGNMENTBOTTOM, ialBottom);
					break;
				case ialCenter:
					description = GetResStringWithNumber(IDP_ICONALIGNMENTCENTER, ialCenter);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_FRM_STYLE:
			switch(cookie) {
				case sNormal:
					description = GetResStringWithNumber(IDP_STYLENORMAL, sNormal);
					break;
				case sFlat:
					description = GetResStringWithNumber(IDP_STYLEFLAT, sFlat);
					break;
				case sOwnerDrawn:
					description = GetResStringWithNumber(IDP_STYLEOWNERDRAWN, sOwnerDrawn);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP Frame::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP Frame::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<Frame>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP Frame::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP Frame::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = properties.hAcceleratorTable;
	pControlInfo->cAccel = static_cast<USHORT>(properties.hAcceleratorTable ? CopyAcceleratorTable(properties.hAcceleratorTable, NULL, 0) : 0);
	pControlInfo->dwFlags = 0;
	return S_OK;
}

STDMETHODIMP Frame::OnMnemonic(LPMSG /*pMessage*/)
{
	if(GetStyle() & WS_DISABLED) {
		return S_OK;
	}

	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT Frame::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT Frame::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_FRM_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT Frame::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP Frame::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP Frame::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_APPEARANCE);
	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 9;
	return S_OK;
}

STDMETHODIMP Frame::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ButtonControls");
	return S_OK;
}

STDMETHODIMP Frame::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"BtnCtls");
	return S_OK;
}

STDMETHODIMP Frame::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP Frame::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_FRM_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_FRM_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP Frame::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_BorderVisible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((GetStyle() & BS_TYPEMASK) != BS_OWNERDRAW) {
			properties.borderVisible = TRUE;
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.borderVisible);
	return S_OK;
}

STDMETHODIMP Frame::put_BorderVisible(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_BORDERVISIBLE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.borderVisible != b) {
		properties.borderVisible = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.borderVisible) {
				switch(properties.style) {
					case sNormal:
						ModifyStyle(BS_FLAT | BS_TYPEMASK, (RunTimeHelper::IsCommCtrl6() ? BS_OWNERDRAW : BS_GROUPBOX));
						break;
					case sFlat:
						ModifyStyle(BS_TYPEMASK, BS_FLAT | (RunTimeHelper::IsCommCtrl6() ? BS_OWNERDRAW : BS_GROUPBOX));
						break;
					case sOwnerDrawn:
						ModifyStyle(BS_FLAT | BS_TYPEMASK, BS_OWNERDRAW);
						break;
				}
			} else {
				ModifyStyle(BS_TYPEMASK, BS_OWNERDRAW);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_BORDERVISIBLE);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP Frame::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP Frame::get_ContentType(ContentTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ContentTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (BS_BITMAP | BS_ICON | BS_TEXT)) {
			case BS_BITMAP:
				properties.contentType = ctBitmap;
				break;
			case BS_ICON:
				properties.contentType = ctIcon;
				break;
			case BS_TEXT:
				properties.contentType = ctText;
				break;
		}
	}

	*pValue = properties.contentType;
	return S_OK;
}

STDMETHODIMP Frame::put_ContentType(ContentTypeConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_CONTENTTYPE);
	if(properties.contentType != newValue) {
		properties.contentType = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.contentType) {
				case ctText:
					ModifyStyle(BS_BITMAP | BS_ICON, BS_TEXT);
					break;
				case ctBitmap:
					ModifyStyle(BS_ICON | BS_TEXT, BS_BITMAP);
					SendMessage(BM_SETIMAGE, IMAGE_BITMAP, reinterpret_cast<LPARAM>(properties.image));
					break;
				case ctIcon:
					ModifyStyle(BS_BITMAP | BS_TEXT, BS_ICON);
					SendMessage(BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(properties.image));
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_CONTENTTYPE);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP Frame::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(((static_cast<long>(newValue) & deMouseEvents) == deMouseEvents) && (properties.style != sOwnerDrawn)) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP Frame::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_FRM_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP Frame::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_FRM_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP Frame::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_FRM_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_FRM_FONT);
	return S_OK;
}

STDMETHODIMP Frame::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_FRM_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_FRM_FONT);
	return S_OK;
}

STDMETHODIMP Frame::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP Frame::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_FRM_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_FORECOLOR);
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP Frame::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (BS_LEFT | BS_CENTER | BS_RIGHT)) {
			case BS_CENTER:
				properties.hAlignment = halCenter;
				break;
			case BS_RIGHT:
				properties.hAlignment = halRight;
				break;
			case BS_LEFT:
			case 0:
				properties.hAlignment = halLeft;
				break;
		}
	}

	*pValue = properties.hAlignment;
	return S_OK;
}

STDMETHODIMP Frame::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.hAlignment) {
				case halLeft:
					ModifyStyle(0x00000F00, BS_LEFT);
					break;
				case halCenter:
					ModifyStyle(0x00000F00, BS_CENTER);
					break;
				case halRight:
					ModifyStyle(0x00000F00, BS_RIGHT);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_hImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			if(imageListSettings.himl != BCCL_NOGLYPH && properties.useImprovedImageListSupport && properties.pImageListWrapper) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.pImageListWrapper->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				if(pImgLstPriv) {
					pImgLstPriv->GetImageList(OIL_NORMAL, &imageListSettings.himl);
				}
			}
			properties.hImageList = (imageListSettings.himl == BCCL_NOGLYPH ? NULL : imageListSettings.himl);
		}
	}
	*pValue = HandleToLong(properties.hImageList);
	return S_OK;
}

STDMETHODIMP Frame::put_hImageList(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_FRM_HIMAGELIST);
	if(properties.hImageList != reinterpret_cast<HIMAGELIST>(LongToHandle(newValue))) {
		properties.hImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			if(properties.hImageList && properties.useImprovedImageListSupport) {
				if(!properties.pImageListWrapper) {
					OffsetImageList_CreateInstance(&properties.pImageListWrapper);
				}
				if(properties.pImageListWrapper) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pImageListWrapper->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					if(pImgLstPriv) {
						pImgLstPriv->SetImageList(OIL_NORMAL, (properties.hImageList ? properties.hImageList : NULL), NULL);
						for(int i = 0; i < 6; i++) {
							pImgLstPriv->SetIndexMapping(i, properties.iconIndexes[i]);
						}
					}
				}
				imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
			} else {
				imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
			}
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_HIMAGELIST);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP Frame::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_FRM_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP Frame::get_IconAlignment(IconAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IconAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			properties.iconAlignment = static_cast<IconAlignmentConstants>(imageListSettings.uAlign);
		}
	}

	*pValue = properties.iconAlignment;
	return S_OK;
}

STDMETHODIMP Frame::put_IconAlignment(IconAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ICONALIGNMENT);
	if(properties.iconAlignment != newValue) {
		properties.iconAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			imageListSettings.uAlign = properties.iconAlignment;
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_ICONALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_IconIndex(LONG controlState/* = -1*/, LONG* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(pValue == NULL) {
		return E_POINTER;
	}
	if(controlState == -1) {
		controlState = 0;
	}
	if(controlState < 0 || controlState >= 6) {
		return E_INVALIDARG;
	}

	if(properties.useImprovedImageListSupport) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		if(properties.pImageListWrapper) {
			properties.pImageListWrapper->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
		}
		*pValue = 0;
		if(pImgLstPriv) {
			pImgLstPriv->GetIndexMapping(controlState, reinterpret_cast<PINT>(pValue));
		} else {
			*pValue = properties.iconIndexes[controlState];
		}
	} else {
		*pValue = properties.iconIndexes[controlState];
	}
	return S_OK;
}

STDMETHODIMP Frame::put_IconIndex(LONG controlState/* = -1*/, LONG newValue/* = 0*/)
{
	PUTPROPPROLOG(DISPID_FRM_ICONINDEX);
	if(controlState < -1 || controlState >= 6) {
		return E_INVALIDARG;
	}

	for(LONG i = (controlState == -1 ? 0 : controlState); i < (controlState == -1 ? 6 : controlState + 1); i++) {
		properties.iconIndexes[i] = newValue;
	}
	if(properties.useImprovedImageListSupport) {
		if(properties.pImageListWrapper) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pImageListWrapper->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			if(pImgLstPriv) {
				for(int i = (controlState == -1 ? 0 : controlState); i < (controlState == -1 ? 6 : controlState + 1); i++) {
					pImgLstPriv->SetIndexMapping(i, newValue);
				}
			}
		}
	}
	FireOnChanged(DISPID_FRM_ICONINDEX);
	FireViewChange();
	return S_OK;
}


STDMETHODIMP Frame::get_IconMarginBottom(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			properties.iconMarginBottom = imageListSettings.margin.bottom;
		}
	}

	*pValue = properties.iconMarginBottom;
	return S_OK;
}

STDMETHODIMP Frame::put_IconMarginBottom(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ICONMARGINBOTTOM);
	if(properties.iconMarginBottom != newValue) {
		properties.iconMarginBottom = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			imageListSettings.margin.bottom = properties.iconMarginBottom;
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_ICONMARGINBOTTOM);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_IconMarginLeft(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			properties.iconMarginLeft = imageListSettings.margin.left;
		}
	}

	*pValue = properties.iconMarginLeft;
	return S_OK;
}

STDMETHODIMP Frame::put_IconMarginLeft(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ICONMARGINLEFT);
	if(properties.iconMarginLeft != newValue) {
		properties.iconMarginLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			imageListSettings.margin.left = properties.iconMarginLeft;
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_ICONMARGINLEFT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_IconMarginRight(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			properties.iconMarginRight = imageListSettings.margin.right;
		}
	}

	*pValue = properties.iconMarginRight;
	return S_OK;
}

STDMETHODIMP Frame::put_IconMarginRight(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ICONMARGINRIGHT);
	if(properties.iconMarginRight != newValue) {
		properties.iconMarginRight = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			imageListSettings.margin.right = properties.iconMarginRight;
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_ICONMARGINRIGHT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_IconMarginTop(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
			properties.iconMarginTop = imageListSettings.margin.top;
		}
	}

	*pValue = properties.iconMarginTop;
	return S_OK;
}

STDMETHODIMP Frame::put_IconMarginTop(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_FRM_ICONMARGINTOP);
	if(properties.iconMarginTop != newValue) {
		properties.iconMarginTop = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				if(properties.useImprovedImageListSupport) {
					imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
				} else {
					imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
				}
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
			}
			imageListSettings.margin.top = properties.iconMarginTop;
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_ICONMARGINTOP);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Image(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (BS_BITMAP | BS_ICON | BS_TEXT)) {
			case BS_BITMAP:
				properties.image = LongToHandle(SendMessage(BM_GETIMAGE, IMAGE_BITMAP, 0));
				break;
			case BS_ICON:
				properties.image = LongToHandle(SendMessage(BM_GETIMAGE, IMAGE_ICON, 0));
				break;
		}
		*pValue = HandleToLong(properties.image);
	} else {
		*pValue = NULL;
	}
	return S_OK;
}

STDMETHODIMP Frame::put_Image(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_FRM_IMAGE);
	HANDLE h = LongToHandle(newValue);
	if(h) {
		if(GetObjectType(h) == OBJ_BITMAP) {
			properties.image = h;
			SetDirty(TRUE);

			if(GetStyle() & BS_BITMAP) {
				SendMessage(BM_SETIMAGE, IMAGE_BITMAP, reinterpret_cast<LPARAM>(properties.image));
				FireViewChange();
			}
			FireOnChanged(DISPID_FRM_IMAGE);
			return S_OK;
		} else {
			ICONINFO iconInfo = {0};
			if(GetIconInfo(static_cast<HICON>(h), &iconInfo)) {
				DeleteObject(iconInfo.hbmColor);
				DeleteObject(iconInfo.hbmMask);

				properties.image = h;
				SetDirty(TRUE);

				if(GetStyle() & BS_ICON) {
					SendMessage(BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(properties.image));
					FireViewChange();
				}
				FireOnChanged(DISPID_FRM_IMAGE);
				return S_OK;
			} else {
				// invalid picture - raise VB runtime error 481
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 481);
			}
		}
	} else {
		properties.image = h;
		switch(GetStyle() & (BS_BITMAP | BS_ICON | BS_TEXT)) {
			case BS_BITMAP:
				SetDirty(TRUE);
				SendMessage(BM_SETIMAGE, IMAGE_BITMAP, reinterpret_cast<LPARAM>(properties.image));
				FireViewChange();
				FireOnChanged(DISPID_FRM_IMAGE);
				return S_OK;
			case BS_ICON:
				SetDirty(TRUE);
				SendMessage(BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(properties.image));
				FireViewChange();
				FireOnChanged(DISPID_FRM_IMAGE);
				return S_OK;
		}
	}

	// invalid picture - raise VB runtime error 481
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 481);
}

STDMETHODIMP Frame::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP Frame::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP Frame::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_FRM_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_FRM_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Frame::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_FRM_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_FRM_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Frame::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP Frame::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Frame::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP Frame::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_FRM_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetExStyle();
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP Frame::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			// TODO: Maybe this is not enough for systems older than Vista.
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Style(StyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && properties.borderVisible && !RunTimeHelper::IsCommCtrl6()) {
		DWORD style = GetStyle();
		if((style & BS_TYPEMASK) == BS_OWNERDRAW) {
			properties.style = sOwnerDrawn;
		} else if(style & BS_FLAT) {
			properties.style = sFlat;
		} else {
			properties.style = sNormal;
		}
	}

	*pValue = properties.style;
	return S_OK;
}

STDMETHODIMP Frame::put_Style(StyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_FRM_STYLE);
	if(properties.style != newValue) {
		properties.style = newValue;
		SetDirty(TRUE);

		if(IsWindow() && properties.borderVisible) {
			switch(properties.style) {
				case sNormal:
					ModifyStyle(BS_FLAT | BS_TYPEMASK, (RunTimeHelper::IsCommCtrl6() ? BS_OWNERDRAW : BS_GROUPBOX));
					break;
				case sFlat:
					ModifyStyle(BS_TYPEMASK, BS_FLAT | (RunTimeHelper::IsCommCtrl6() ? BS_OWNERDRAW : BS_GROUPBOX));
					break;
				case sOwnerDrawn:
					ModifyStyle(BS_FLAT | BS_TYPEMASK, BS_OWNERDRAW);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_STYLE);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP Frame::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Frame::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		GetWindowText(&properties.text);
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP Frame::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_FRM_TEXT);
	if(properties.text != newValue) {
		properties.text.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			RECT rc = currentTextRectangle;
			SetWindowText(COLE2CT(properties.text));
			InvalidateRect(&rc);
		}
		FireOnChanged(DISPID_FRM_TEXT);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP Frame::get_UseImprovedImageListSupport(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useImprovedImageListSupport);
	return S_OK;
}

STDMETHODIMP Frame::put_UseImprovedImageListSupport(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_USEIMPROVEDIMAGELISTSUPPORT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useImprovedImageListSupport != b) {
		properties.useImprovedImageListSupport = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			BUTTON_IMAGELIST imageListSettings = {0};
			if(!SendMessage(BCM_GETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings))) {
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			} else if(!imageListSettings.himl || (imageListSettings.himl == BCCL_NOGLYPH)) {
				imageListSettings.uAlign = properties.iconAlignment;
				imageListSettings.margin.bottom = properties.iconMarginBottom;
				imageListSettings.margin.left = properties.iconMarginLeft;
				imageListSettings.margin.right = properties.iconMarginRight;
				imageListSettings.margin.top = properties.iconMarginTop;
			}
			if(properties.hImageList && properties.useImprovedImageListSupport) {
				if(!properties.pImageListWrapper) {
					OffsetImageList_CreateInstance(&properties.pImageListWrapper);
				}
				if(properties.pImageListWrapper) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pImageListWrapper->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					if(pImgLstPriv) {
						pImgLstPriv->SetImageList(OIL_NORMAL, (properties.hImageList ? properties.hImageList : NULL), NULL);
						for(int i = 0; i < 6; i++) {
							pImgLstPriv->SetIndexMapping(i, properties.iconIndexes[i]);
						}
					}
				}
				imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
			} else {
				imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
			}
			SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
			FireViewChange();
		}
		FireOnChanged(DISPID_FRM_USEIMPROVEDIMAGELISTSUPPORT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP Frame::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_FRM_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_FRM_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP Frame::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP Frame::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP Frame::Click(void)
{
	if(IsWindow()) {
		SendMessage(BM_CLICK, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Frame::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP Frame::GetControlState(ControlStateConstants* pControlState)
{
	ATLASSERT_POINTER(pControlState, ControlStateConstants);
	if(!pControlState) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pControlState = static_cast<ControlStateConstants>(SendMessage(BM_GETSTATE, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Frame::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP Frame::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP Frame::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


LRESULT Frame::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT Frame::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/*DWORD style = GetClassLong(*this, GCL_STYLE);
	style &= ~(CS_HREDRAW | CS_VREDRAW);
	SetClassLong(*this, GCL_STYLE, style);*/

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT Frame::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT Frame::OnEraseBkGnd(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	#if FALSE
		wasHandled = (properties.style != sOwnerDrawn && RunTimeHelper::IsCommCtrl6());
	#else
		if(properties.style != sOwnerDrawn && RunTimeHelper::IsCommCtrl6()) {
			RECT clientRectangle;
			GetClientRect(&clientRectangle);
			if(flags.usingThemes) {
				int state = SaveDC(reinterpret_cast<HDC>(wParam));
				if(RunTimeHelper::IsVista()) {
					DrawThemeParentBackgroundEx(*this, reinterpret_cast<HDC>(wParam), DTPB_USECTLCOLORSTATIC, &clientRectangle);
				} else {
					FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, GetSysColorBrush(COLOR_BTNFACE));
					DrawThemeParentBackground(*this, reinterpret_cast<HDC>(wParam), &clientRectangle);
				}
				RestoreDC(reinterpret_cast<HDC>(wParam), state);
			} else {
				if(!(properties.backColor & 0x80000000)) {
					COLORREF previousColor = SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
					FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
					SetDCBrushColor(reinterpret_cast<HDC>(wParam), previousColor);
				} else {
					FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
				}
			}
		} else {
			wasHandled = FALSE;
		}
	#endif
	return 1;
}

LRESULT Frame::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	HWND hWndCapture = GetCapture();
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(hWndCapture == *this) {
		SetCapture();
	}
	return lr;
}

LRESULT Frame::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if((GetStyle() & BS_TYPEMASK) == BS_OWNERDRAW) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			Invalidate();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if((GetStyle() & BS_TYPEMASK) == BS_OWNERDRAW) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			Invalidate();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL inDesignMode = IsInDesignMode();
	if(!inDesignMode) {
		TCHAR pBuffer[200];
		ZeroMemory(pBuffer, 200 * sizeof(TCHAR));
		GetModuleFileName(NULL, pBuffer, 200);
		PathStripPath(pBuffer);
		if(lstrcmpi(pBuffer, TEXT("vb6.exe")) == 0) {
			HWND hWnd = GetParent();
			while(hWnd) {
				if(GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("DesignerWindow")) == 0) {
					hWnd = ::GetParent(hWnd);
					if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("MDIClient")) == 0) {
						hWnd = ::GetParent(hWnd);
						if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("wndclass_desked_gsk")) == 0) {
							inDesignMode = TRUE;
						}
					}
					break;
				}
				hWnd = ::GetParent(hWnd);
			}
		}
	}
	if(!inDesignMode) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		if(WindowFromPoint(mousePosition) == *this) {
			ScreenToClient(&mousePosition);

			/* Here we have 2 problems:
				* 1) Eating the mouse message (MA_NOACTIVATEANDEAT) prevents contained windowless controls
				*    from receiving the message, so those controls won't raise Click events anymore.
				*    Unfortunatly eating the message is required, otherwise the contained windowless control
				*    will lose focus if the Frame control is clicked.
				* 2) Raise_MouseDown calls SetCapture which also prevents contained windowless controls from
				*    raising Click events. Why is SetCapture required?
				* Idea: Detect that the message actually was for the contained windowless control. Treat the
				* message as if it was outside our window rectangle in such cases.
				*/

			CComPtr<IDispatch> pMyExtendedControl = NULL;
			if(m_spClientSite) {
				CComQIPtr<IOleControlSite> pControlSite = m_spClientSite;
				if(pControlSite) {
					pControlSite->GetExtendedControl(&pMyExtendedControl);
				}
			}
			BOOL hitWindowLessChild = FALSE;
			CComPtr<IOleContainer> pContainer = NULL;
			if(m_spClientSite && SUCCEEDED(m_spClientSite->GetContainer(&pContainer)) && pContainer) {
				CComPtr<IEnumUnknown> pEnum = NULL;
				if(SUCCEEDED(pContainer->EnumObjects(OLECONTF_EMBEDDINGS | OLECONTF_OTHERS, &pEnum)) && pEnum) {
					CComPtr<IUnknown> pControlUnk = NULL;
					while(pEnum->Next(1, &pControlUnk, NULL) == S_OK && !hitWindowLessChild) {
						CComPtr<IDispatch> pExtendedControl = NULL;
						CComQIPtr<IOleObject> pControl = pControlUnk;
						if(pControl) {
							CComPtr<IOleClientSite> pClientSite = NULL;
							if(SUCCEEDED(pControl->GetClientSite(&pClientSite)) && pClientSite) {
								CComQIPtr<IOleControlSite> pControlSite = pClientSite;
								if(pControlSite) {
									pControlSite->GetExtendedControl(&pExtendedControl);
								}
							}
						} else {
							pControlUnk->QueryInterface(IID_PPV_ARGS(&pExtendedControl));
						}

						if(pExtendedControl) {
							RECT controlRect = {0};
							VARIANT v;
							VariantInit(&v);
							DISPID dispid;
							DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
							BOOL isChildControl = FALSE;

							LPWSTR pName[1] = { L"Container" };
							if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
								if(SUCCEEDED(pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL))) {
									isChildControl = (v.vt == VT_DISPATCH && v.pdispVal == pMyExtendedControl);
									VariantClear(&v);
								}
							}

							if(isChildControl) {
								pName[0] = L"Left";
								if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
									if(SUCCEEDED(pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL))) {
										VARIANT v1;
										VariantInit(&v1);
										if(SUCCEEDED(VariantChangeType(&v1, &v, 0, VT_INT))) {
											controlRect.left = v1.intVal / 15;
										}
									}
								}
								pName[0] = L"Top";
								if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
									if(SUCCEEDED(pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL))) {
										VARIANT v1;
										VariantInit(&v1);
										if(SUCCEEDED(VariantChangeType(&v1, &v, 0, VT_INT))) {
											controlRect.top = v1.intVal / 15;
										}
									}
								}
								pName[0] = L"Height";
								if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
									if(SUCCEEDED(pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL))) {
										VARIANT v1;
										VariantInit(&v1);
										if(SUCCEEDED(VariantChangeType(&v1, &v, 0, VT_INT))) {
											controlRect.bottom = controlRect.top + v1.intVal / 15;
										}
									}
								}
								pName[0] = L"Width";
								if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
									if(SUCCEEDED(pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL))) {
										VARIANT v1;
										VariantInit(&v1);
										if(SUCCEEDED(VariantChangeType(&v1, &v, 0, VT_INT))) {
											controlRect.right = controlRect.left + v1.intVal / 15;
										}
									}
								}

								if(PtInRect(&controlRect, mousePosition)) {
									/*pName[0] = L"Name";
									if(SUCCEEDED(pExtendedControl->GetIDsOfNames(IID_NULL, pName, 1, LOCALE_SYSTEM_DEFAULT, &dispid))) {
										pExtendedControl->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParams, &v, NULL, NULL);
									}*/
									hitWindowLessChild = TRUE;
								}
							}
						}
						pControlUnk = NULL;
					}
				}
			}

			if(!hitWindowLessChild) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(HIWORD(lParam) == WM_LBUTTONDOWN) {
					button = 1/*MouseButtonConstants.vbLeftButton*/;
					Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
				} else if(HIWORD(lParam) == WM_MBUTTONDOWN) {
					button = 4/*MouseButtonConstants.vbMiddleButton*/;
					Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
				} else if(HIWORD(lParam) == WM_RBUTTONDOWN) {
					button = 2/*MouseButtonConstants.vbRightButton*/;
					Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
				} else if(HIWORD(lParam) == WM_XBUTTONDOWN) {
					Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
				}
				return MA_NOACTIVATEANDEAT;
			}
		}
	}

	/* NOTE: It's important to hide this message from CComControl. For instance, if you place a text box onto
	 *       the frame and click it twice, the second click generates a WM_MOUSEACTIVATE messages that is
	 *       forwarded upwards in the window hierarchy until a window handles the message. Therefore we'll
	 *       receive this message, too. Unfortunatly ComControl will handle the message in a way that makes
	 *       the text box lose the COM keyboard focus.
	 *       To make things more complicated, this happens only if the frame control sits on a MDI child
	 *       window.
	 */
	return DefWindowProc(message, wParam, lParam);
}

LRESULT Frame::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT Frame::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		if(properties.style == sOwnerDrawn) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
			trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
			TrackMouseEvent(&trackingOptions);
		}
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT Frame::OnNCHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return HTCLIENT;
}

LRESULT Frame::OnPrintClient(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(RunTimeHelper::IsCommCtrl6()) {
		if(properties.style == sOwnerDrawn) {
			OwnerDrawControlStateConstants controlState = odcsNormal;
			DWORD state = SendMessage(BM_GETSTATE, 0, 0);
			DWORD uiState = SendMessage(WM_QUERYUISTATE, 0, 0);
			if(state & BST_CHECKED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsSelected);
			}
			if(GetStyle() & WS_DISABLED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsDisabled);
			}
			if(GetFocus() == *this) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsFocus);
			}
			if(uiState & UISF_HIDEACCEL) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoAccelerator);
			}
			if(uiState & UISF_HIDEFOCUS) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoFocusRectangle);
			}
			if(state & BST_PUSHED) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsPushed);
			}
			if(state & BST_INDETERMINATE) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsIndeterminate);
			}
			if(mouseStatus.hoveredControl) {
				controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsHot);
			}

			RECT clientRectangle;
			GetClientRect(&clientRectangle);
			Raise_OwnerDraw(odaDrawEntire, controlState, HandleToLong(reinterpret_cast<HDC>(wParam)), reinterpret_cast<RECTANGLE*>(&clientRectangle));
		} else {
			if(!(lParam & PRF_CHECKVISIBLE) || IsWindowVisible()) {
				if(lParam & PRF_ERASEBKGND) {
					SendMessage(WM_ERASEBKGND, wParam, 0);
				}
				DrawFrame(reinterpret_cast<HDC>(wParam));
			}
		}
	} else {
		wasHandled = FALSE;
	}
	return 1;
}

LRESULT Frame::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT Frame::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_FRM_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_FRM_FONT);
	}

	return lr;
}

LRESULT Frame::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT Frame::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_FRM_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// update the accelerator table
	CComBSTR text;
	GetWindowText(&text);
	TCHAR newAccelerator = 0;
	for(long i = text.Length() - 1; i > 0; --i) {
		if((text[i - 1] == L'&') && (text[i] != L'&')) {
			COLE2T converter(text);
			newAccelerator = converter[i];
			break;
		}
	}

	if(newAccelerator != properties.accelerator) {
		properties.accelerator = newAccelerator;
		if(properties.hAcceleratorTable) {
			DestroyAcceleratorTable(properties.hAcceleratorTable);
		}

		// create a new accelerator table
		ACCEL accelerators[4] = {	{FVIRTKEY | FALT, LOBYTE(VkKeyScan(properties.accelerator)), 1},
															{FVIRTKEY | FALT | FSHIFT, LOBYTE(VkKeyScan(properties.accelerator)), 1},
															{FALT, static_cast<WORD>(tolower(properties.accelerator)), 1},
															{0, static_cast<WORD>(tolower(properties.accelerator)), 1} };
		properties.hAcceleratorTable = CreateAcceleratorTable(accelerators, _countof(accelerators));

		// report the new accelerator table to the container
		CComQIPtr<IOleControlSite, &IID_IOleControlSite> pSite(m_spClientSite);
		if(pSite) {
			pSite->OnControlInfoChanged();
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_FRM_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT Frame::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETNONCLIENTMETRICS) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT Frame::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
				CTheme themingEngine;
				themingEngine.OpenThemeData(*this, VSCLASS_BUTTON);
				flags.themeIsNull = themingEngine.IsThemeNull();
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT Frame::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT Frame::OnUpdateUIState(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	Invalidate();
	return lr;
}

LRESULT Frame::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT Frame::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Frame::OnReflectedCtlColorBtn(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(flags.themeIsNull) {
		SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
		SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
		if(!(properties.backColor & 0x80000000)) {
			SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
			return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
		} else {
			return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
		}
	} else {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
	}
}

LRESULT Frame::OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(flags.usingThemes) {
		wasHandled = FALSE;
		return 0;
	} else {
		SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
		SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
		RECT clientRectangle;
		::GetClientRect(reinterpret_cast<HWND>(lParam), &clientRectangle);
		HBRUSH hBrush = NULL;
		if(!(properties.backColor & 0x80000000)) {
			SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
			hBrush = static_cast<HBRUSH>(GetStockObject(DC_BRUSH));
			FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, hBrush);
		} else {
			hBrush = GetSysColorBrush(properties.backColor & 0x0FFFFFFF);
			FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, hBrush);
		}
		return reinterpret_cast<LRESULT>(hBrush);
	}
}

LRESULT Frame::OnReflectedDrawItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);
	ATLASSERT_POINTER(pDetails, DRAWITEMSTRUCT);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}

	ATLASSERT(pDetails->hwndItem == *this);
	ATLASSERT((pDetails->itemAction & ~(ODA_DRAWENTIRE | ODA_SELECT | ODA_FOCUS)) == 0);
	ATLASSERT((pDetails->itemState & ~(ODS_SELECTED | ODS_DISABLED | ODS_FOCUS | ODS_NOACCEL | ODS_NOFOCUSRECT)) == 0);

	if(properties.style == sOwnerDrawn) {
		DWORD state = SendMessage(BM_GETSTATE, 0, 0);
		DWORD uiState = SendMessage(WM_QUERYUISTATE, 0, 0);
		OwnerDrawControlStateConstants controlState = odcsNormal;
		if(pDetails->itemState & ODS_SELECTED) {
			ATLASSERT((state & BST_CHECKED) == BST_CHECKED);
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsSelected);
		} else {
			ATLASSERT((state & BST_CHECKED) == 0);
		}
		if(pDetails->itemState & ODS_DISABLED) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsDisabled);
		}
		if(GetFocus() == *this) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsFocus);
		}
		if(uiState & UISF_HIDEACCEL) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoAccelerator);
		}
		if(uiState & UISF_HIDEFOCUS) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsNoFocusRectangle);
		}
		if(state & BST_PUSHED) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsPushed);
		}
		if(state & BST_INDETERMINATE) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsIndeterminate);
		}
		if(mouseStatus.hoveredControl) {
			controlState = static_cast<OwnerDrawControlStateConstants>(static_cast<long>(controlState) | odcsHot);
		}
		Raise_OwnerDraw(static_cast<OwnerDrawActionConstants>(pDetails->itemAction), controlState, HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));
	} else if(RunTimeHelper::IsCommCtrl6()) {
		DrawFrame(pDetails->hDC);
	}

	return TRUE;
}

LRESULT Frame::OnGetImageList(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	PBUTTON_IMAGELIST pDetails = reinterpret_cast<PBUTTON_IMAGELIST>(lParam);
	ATLASSERT_POINTER(pDetails, BUTTON_IMAGELIST);
	if(!pDetails) {
		return FALSE;
	}

	pDetails->himl = properties.hImageList;
	pDetails->margin.bottom = properties.iconMarginBottom;
	pDetails->margin.left = properties.iconMarginLeft;
	pDetails->margin.right = properties.iconMarginRight;
	pDetails->margin.top = properties.iconMarginTop;
	pDetails->uAlign = properties.iconAlignment;

	return TRUE;
}

LRESULT Frame::OnSetImageList(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	PBUTTON_IMAGELIST pDetails = reinterpret_cast<PBUTTON_IMAGELIST>(lParam);
	ATLASSERT_POINTER(pDetails, BUTTON_IMAGELIST);
	if(!pDetails) {
		return FALSE;
	}

	properties.hImageList = (pDetails->himl == BCCL_NOGLYPH ? NULL : pDetails->himl);
	properties.iconMarginBottom = pDetails->margin.bottom;
	properties.iconMarginLeft = pDetails->margin.left;
	properties.iconMarginRight = pDetails->margin.right;
	properties.iconMarginTop = pDetails->margin.top;
	properties.iconAlignment = static_cast<IconAlignmentConstants>(pDetails->uAlign);

	return TRUE;
}

LRESULT Frame::OnClick(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	flags.handlingBM_CLICK = TRUE;
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	flags.handlingBM_CLICK = FALSE;
	return lr;
}


void Frame::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}

void Frame::CalculateCaptionRectangles(SIZE iconSize, SIZE textOrImageSize, LPRECT pIconRectangle, LPRECT pTextOrImageRectangle)
{
	ATLASSERT_POINTER(pIconRectangle, RECT);
	ATLASSERT_POINTER(pTextOrImageRectangle, RECT);
	if(!pIconRectangle || !pTextOrImageRectangle) {
		return;
	}

	int captionHorizontalMargin = 8 + GetSystemMetrics(SM_CXBORDER);
	CRect clientRectangle;
	GetClientRect(&clientRectangle);

	POINT iconPosition = {0};
	POINT imageOrTextPosition = {0};
	switch(GetStyle() & (BS_LEFT | BS_CENTER | BS_RIGHT)) {
		case BS_CENTER:
			if(iconSize.cx > 0) {
				switch(properties.iconAlignment) {
					case ialLeft:
						iconPosition.x = captionHorizontalMargin + properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							int xIconRight = iconPosition.x + iconSize.cx + properties.iconMarginRight;
							int remainingWidth = clientRectangle.Width() - captionHorizontalMargin - xIconRight;
							imageOrTextPosition.x = xIconRight + (max(remainingWidth - textOrImageSize.cx, 0) / 2);
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialRight:
						iconPosition.x = clientRectangle.Width() - captionHorizontalMargin - properties.iconMarginRight - iconSize.cx - properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							int remainingWidth = iconPosition.x - captionHorizontalMargin;
							if(textOrImageSize.cx > remainingWidth) {
								imageOrTextPosition.x = iconPosition.x - 1 - textOrImageSize.cx;
							} else {
								imageOrTextPosition.x = captionHorizontalMargin + (remainingWidth - textOrImageSize.cx) / 2;
							}
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialTop:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = (clientRectangle.Width() - textOrImageSize.cx) / 2;
							imageOrTextPosition.y = iconPosition.y + iconSize.cy - (properties.iconMarginTop / 2) + properties.iconMarginBottom;
						}
						break;
					case ialBottom:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = clientRectangle.Height() - properties.iconMarginBottom - iconSize.cy - properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = (clientRectangle.Width() - textOrImageSize.cx) / 2;
							imageOrTextPosition.y = min(0, iconPosition.y - textOrImageSize.cy);
						}
						break;
					case ialCenter:
						iconPosition.x = ((clientRectangle.Width() - iconSize.cx) + (properties.iconMarginLeft - properties.iconMarginRight)) / 2;
						iconPosition.y = properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						imageOrTextPosition.x = (clientRectangle.Width() - textOrImageSize.cx) / 2;
						imageOrTextPosition.y = 0;
						break;
				}
			} else if(textOrImageSize.cx > 0) {
				imageOrTextPosition.x = (clientRectangle.Width() - textOrImageSize.cx) / 2;
				imageOrTextPosition.y = 0;
			}
			break;
		case BS_RIGHT:
			if(iconSize.cx >= 0) {
				switch(properties.iconAlignment) {
					case ialLeft:
						iconPosition.x = captionHorizontalMargin + properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = max(clientRectangle.Width() - captionHorizontalMargin - textOrImageSize.cx, iconPosition.x + iconSize.cx + properties.iconMarginRight + GetSystemMetrics(SM_CXEDGE));
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialRight:
						iconPosition.x = clientRectangle.Width() - captionHorizontalMargin - properties.iconMarginRight - iconSize.cx - properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = iconPosition.x - 1 - textOrImageSize.cx;
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialTop:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = clientRectangle.Width() - captionHorizontalMargin - textOrImageSize.cx;
							imageOrTextPosition.y = iconPosition.y + iconSize.cy - (properties.iconMarginTop / 2) + properties.iconMarginBottom;
						}
						break;
					case ialBottom:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = clientRectangle.Height() - properties.iconMarginBottom - iconSize.cy - properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = clientRectangle.Width() - captionHorizontalMargin - textOrImageSize.cx;
							imageOrTextPosition.y = min(0, iconPosition.y - textOrImageSize.cy);
						}
						break;
					case ialCenter:
						iconPosition.x = ((clientRectangle.Width() - iconSize.cx) + (properties.iconMarginLeft - properties.iconMarginRight)) / 2;
						iconPosition.y = properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						imageOrTextPosition.x = clientRectangle.Width() - captionHorizontalMargin - textOrImageSize.cx;
						imageOrTextPosition.y = 0;
						break;
				}
			} else if(textOrImageSize.cx > 0) {
				imageOrTextPosition.x = clientRectangle.Width() - captionHorizontalMargin - textOrImageSize.cx;
				imageOrTextPosition.y = 0;
			}
			break;
		case BS_LEFT:
		default:
			if(iconSize.cx >= 0) {
				switch(properties.iconAlignment) {
					case ialLeft:
						iconPosition.x = captionHorizontalMargin + properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = iconPosition.x + iconSize.cx + properties.iconMarginRight;
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialRight:
						iconPosition.x = clientRectangle.Width() - captionHorizontalMargin - properties.iconMarginRight - iconSize.cx - properties.iconMarginLeft;
						iconPosition.y = 1 + properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = min(captionHorizontalMargin, iconPosition.x - textOrImageSize.cx - 1);
							imageOrTextPosition.y = (max(iconSize.cy, textOrImageSize.cy) - textOrImageSize.cy) / 2;
						}
						break;
					case ialTop:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = captionHorizontalMargin;
							imageOrTextPosition.y = iconPosition.y + iconSize.cy - (properties.iconMarginTop / 2) + properties.iconMarginBottom;
						}
						break;
					case ialBottom:
						iconPosition.x = (clientRectangle.Width() - iconSize.cx) / 2 + properties.iconMarginLeft - properties.iconMarginRight;
						iconPosition.y = clientRectangle.Height() - properties.iconMarginBottom - iconSize.cy - properties.iconMarginTop;
						if(textOrImageSize.cx > 0) {
							imageOrTextPosition.x = captionHorizontalMargin;
							imageOrTextPosition.y = min(0, iconPosition.y - textOrImageSize.cy);
						}
						break;
					case ialCenter:
						iconPosition.x = ((clientRectangle.Width() - iconSize.cx) + (properties.iconMarginLeft - properties.iconMarginRight)) / 2;
						iconPosition.y = properties.iconMarginTop - properties.iconMarginBottom + ((max(iconSize.cy, textOrImageSize.cy) - iconSize.cy) / 2);
						imageOrTextPosition.x = captionHorizontalMargin;
						imageOrTextPosition.y = 0;
						break;
				}
			} else if(textOrImageSize.cx > 0) {
				imageOrTextPosition.x = captionHorizontalMargin;
				imageOrTextPosition.y = 0;
			}
			break;
	}

	if(iconSize.cx >= 0) {
		pIconRectangle->left = clientRectangle.left + iconPosition.x;
		pIconRectangle->right = pIconRectangle->left + iconSize.cx;
		pIconRectangle->top = clientRectangle.top + iconPosition.y;
		pIconRectangle->bottom = pIconRectangle->top + iconSize.cy;
	}
	if(textOrImageSize.cx > 0) {
		pTextOrImageRectangle->left = clientRectangle.left + imageOrTextPosition.x;
		pTextOrImageRectangle->right = pTextOrImageRectangle->left + textOrImageSize.cx;
		pTextOrImageRectangle->top = clientRectangle.top + imageOrTextPosition.y;
		pTextOrImageRectangle->bottom = pTextOrImageRectangle->top + textOrImageSize.cy;
	}
}

void Frame::DrawFrame(HDC hTargetDC)
{
	CRect clientRectangle;
	GetClientRect(&clientRectangle);
	CDCHandle dcToDrawIn(hTargetDC);

	BOOL drawImage = FALSE;
	SIZE imageSize = {0};
	if(properties.image) {
		switch(properties.contentType) {
			case ctBitmap:
				if(GetObjectType(properties.image) == OBJ_BITMAP) {
					BITMAP bmp = {0};
					GetObject(properties.image, sizeof(BITMAP), &bmp);
					imageSize.cx = bmp.bmWidth;
					imageSize.cy = bmp.bmHeight;

					drawImage = TRUE;
				}
				break;
			case ctIcon: {
				ICONINFO dummy = {0};
				if(GetIconInfo(static_cast<HICON>(properties.image), &dummy)) {
					BITMAP bmp = {0};
					GetObject((dummy.hbmColor ? dummy.hbmColor : dummy.hbmMask), sizeof(BITMAP), &bmp);
					imageSize.cx = bmp.bmWidth;
					imageSize.cy = bmp.bmHeight;

					if(dummy.hbmColor) {
						DeleteObject(dummy.hbmColor);
					}
					if(dummy.hbmMask) {
						DeleteObject(dummy.hbmMask);
					}
					drawImage = TRUE;
				}
				break;
			}
		}
	}

	COLORREF previousBackColor = dcToDrawIn.SetBkColor(OLECOLOR2COLORREF(properties.backColor));
	COLORREF previousForeColor = dcToDrawIn.SetTextColor(OLECOLOR2COLORREF(properties.foreColor));
	HFONT hPreviousFont = dcToDrawIn.SelectFont(properties.font.currentFont);
	int previousBkMode = dcToDrawIn.SetBkMode(TRANSPARENT);

	if(drawImage || flags.themeIsNull) {
		// draw without themes
		if(properties.borderVisible && (properties.style != sOwnerDrawn)) {
			BOOL drawAccel = ((SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEACCEL) == 0);
			CComBSTR text;
			GetWindowText(&text);
			int textLength = text.Length();
			int iconToDraw = -1;

			// calculate sizes
			SIZE textSize = {0};
			SIZE iconSize = {0};
			if(!drawImage) {
				RECT textRectangle = clientRectangle;
				RECT rc = textRectangle;
				if(textLength) {
					dcToDrawIn.DrawText(COLE2CT(text), textLength, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX));
				} else {
					dcToDrawIn.DrawText(TEXT("A"), 1, &rc, DT_CALCRECT | DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX));
				}
				textSize.cx = rc.right - rc.left;
				textSize.cy = rc.bottom - rc.top;
			}
			if(properties.hImageList) {
				// Do we have to draw an icon?
				int imageCount = ImageList_GetImageCount(properties.hImageList);
				if(imageCount == 1) {
					iconToDraw = 0;
				} else if(imageCount > 1) {
					iconToDraw = 0;
					if(GetStyle() & WS_DISABLED) {
						iconToDraw = ((imageCount >= 4) ? 3 : -1);
					} else if(SendMessage(BM_GETSTATE, 0, 0) & BST_PUSHED) {
						iconToDraw = ((imageCount >= 3) ? 2 : -1);
					} else if(mouseStatus.hoveredControl) {
						iconToDraw = ((imageCount >= 2) ? 1 : -1);
					}
				}

				if(iconToDraw >= 0) {
					int cx = 0;
					int cy = 0;
					ImageList_GetIconSize(properties.hImageList, &cx, &cy);
					iconSize.cx = cx;
					iconSize.cy = cy;
				}
			}

			// calculate positions (based on CommandButton behavior, so don't wonder about strange behavior)
			RECT iconRectangle = {0};
			RECT imageOrTextRectangle = {0};
			CalculateCaptionRectangles(iconSize, (drawImage ? imageSize : textSize), &iconRectangle, &imageOrTextRectangle);
			currentIconRectangle = iconRectangle;
			currentTextRectangle = imageOrTextRectangle;

			WTL::CRgn clipRgn;
			dcToDrawIn.GetClipRgn(clipRgn);
			if(iconToDraw >= 0) {
				RECT rc = iconRectangle;
				rc.left -= GetSystemMetrics(SM_CXEDGE);
				rc.right += GetSystemMetrics(SM_CXEDGE);
				dcToDrawIn.ExcludeClipRect(&rc);
			}
			if(drawImage || (textLength > 0)) {
				RECT rc = imageOrTextRectangle;
				rc.left -= GetSystemMetrics(SM_CXEDGE);
				rc.right += GetSystemMetrics(SM_CXEDGE);
				dcToDrawIn.ExcludeClipRect(&rc);
			}

			// draw the border
			RECT borderRectangle = clientRectangle;
			borderRectangle.top += max((drawImage ? imageSize.cy : textSize.cy), iconSize.cy) / 2;
			switch(properties.style) {
				case sNormal:
					dcToDrawIn.DrawEdge(&borderRectangle, EDGE_ETCHED, BF_RECT);
					break;
				case sFlat:
					dcToDrawIn.DrawEdge(&borderRectangle, EDGE_ETCHED, BF_RECT | BF_FLAT | BF_MONO);
					break;
			}
			dcToDrawIn.SelectClipRgn(clipRgn);

			// draw the icon
			if(iconToDraw >= 0) {
				if(!(properties.backColor & 0x80000000)) {
					COLORREF previousColor = dcToDrawIn.SetDCBrushColor(OLECOLOR2COLORREF(properties.backColor));
					dcToDrawIn.FillRect(&iconRectangle, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
					dcToDrawIn.SetDCBrushColor(previousColor);
				} else {
					dcToDrawIn.FillRect(&iconRectangle, GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
				}
				ImageList_Draw(properties.hImageList, iconToDraw, dcToDrawIn, iconRectangle.left, iconRectangle.top, ILD_NORMAL);
			}

			// draw the image or text
			if(drawImage) {
				if(!(properties.backColor & 0x80000000)) {
					COLORREF previousColor = dcToDrawIn.SetDCBrushColor(OLECOLOR2COLORREF(properties.backColor));
					dcToDrawIn.FillRect(&imageOrTextRectangle, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
					dcToDrawIn.SetDCBrushColor(previousColor);
				} else {
					dcToDrawIn.FillRect(&imageOrTextRectangle, GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
				}
				switch(properties.contentType) {
					case ctBitmap: {
						CDC bitmapDC;
						bitmapDC.CreateCompatibleDC(dcToDrawIn);
						HBITMAP hPreviousBitmap = bitmapDC.SelectBitmap(static_cast<HBITMAP>(properties.image));
						dcToDrawIn.BitBlt(imageOrTextRectangle.left, imageOrTextRectangle.top, imageSize.cx, imageSize.cy, bitmapDC, 0, 0, SRCCOPY);
						bitmapDC.SelectBitmap(hPreviousBitmap);
						break;
					}
					case ctIcon:
						dcToDrawIn.DrawIcon(imageOrTextRectangle.left, imageOrTextRectangle.top, static_cast<HICON>(properties.image));
						break;
				}
			} else if(textLength) {
				dcToDrawIn.DrawText(COLE2CT(text), textLength, &imageOrTextRectangle, DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX));
			}
		}
	} else {
		// draw using themes
		if(properties.borderVisible && (properties.style != sOwnerDrawn)) {
			CTheme themingEngine;
			themingEngine.OpenThemeData(*this, VSCLASS_BUTTON);

			BOOL enabled = !(GetStyle() & WS_DISABLED);
			DWORD extendedStyle = GetExStyle();
			BOOL drawAccel = ((SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEACCEL) == 0);
			CComBSTR text;
			GetWindowText(&text);
			int textLength = text.Length();
			int iconToDraw = -1;

			// calculate sizes
			SIZE textSize = {0};
			SIZE iconSize = {0};
			RECT textRectangle = clientRectangle;
			RECT rc = {0};
			if(textLength > 0) {
				themingEngine.GetThemeTextExtent(dcToDrawIn, BP_GROUPBOX, (enabled ? GBS_NORMAL : GBS_DISABLED), CW2CW(text), textLength, DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX) | (extendedStyle & WS_EX_RTLREADING ? DT_RTLREADING : 0), &textRectangle, &rc);
			} else {
				themingEngine.GetThemeTextExtent(dcToDrawIn, BP_GROUPBOX, (enabled ? GBS_NORMAL : GBS_DISABLED), L"A", 1, DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX) | (extendedStyle & WS_EX_RTLREADING ? DT_RTLREADING : 0), &textRectangle, &rc);
			}
			textSize.cx = rc.right - rc.left;
			textSize.cy = rc.bottom - rc.top;
			if(properties.hImageList) {
				// Do we have to draw an icon?
				int imageCount = ImageList_GetImageCount(properties.hImageList);
				if(imageCount == 1) {
					iconToDraw = 0;
				} else if(imageCount > 1) {
					iconToDraw = 0;
					if(!enabled) {
						iconToDraw = ((imageCount >= 4) ? 3 : -1);
					} else if(SendMessage(BM_GETSTATE, 0, 0) & BST_PUSHED) {
						iconToDraw = ((imageCount >= 3) ? 2 : -1);
					} else if(mouseStatus.hoveredControl) {
						iconToDraw = ((imageCount >= 2) ? 1 : -1);
					}
				}

				if(iconToDraw >= 0) {
					int cx = 0;
					int cy = 0;
					ImageList_GetIconSize(properties.hImageList, &cx, &cy);
					iconSize.cx = cx;
					iconSize.cy = cy;
				}
			}

			// calculate positions (based on CommandButton behavior, so don't wonder about strange behavior)
			RECT iconRectangle = {0};
			CalculateCaptionRectangles(iconSize, textSize, &iconRectangle, &textRectangle);
			currentIconRectangle = iconRectangle;
			currentTextRectangle = textRectangle;

			WTL::CRgn clipRgn;
			dcToDrawIn.GetClipRgn(clipRgn);
			if(iconToDraw >= 0) {
				rc = iconRectangle;
				rc.left -= GetSystemMetrics(SM_CXEDGE);
				rc.right += GetSystemMetrics(SM_CXEDGE);
				dcToDrawIn.ExcludeClipRect(&rc);
			}
			if(textLength) {
				rc = textRectangle;
				rc.left -= GetSystemMetrics(SM_CXEDGE);
				rc.right += GetSystemMetrics(SM_CXEDGE);
				dcToDrawIn.ExcludeClipRect(&rc);
			}

			// draw the border
			RECT borderRectangle = clientRectangle;
			borderRectangle.top += max(textSize.cy, iconSize.cy) / 2;
			themingEngine.DrawThemeBackground(dcToDrawIn, BP_GROUPBOX, (enabled ? GBS_NORMAL : GBS_DISABLED), &borderRectangle, NULL);
			dcToDrawIn.SelectClipRgn(clipRgn);

			int state = dcToDrawIn.SaveDC();
			// draw the icon
			if(iconToDraw >= 0) {
				if(RunTimeHelper::IsVista()) {
					DrawThemeParentBackgroundEx(*this, dcToDrawIn, DTPB_USECTLCOLORSTATIC, &iconRectangle);
				} else {
					dcToDrawIn.FillRect(&iconRectangle, GetSysColorBrush(COLOR_BTNFACE));
					DrawThemeParentBackground(*this, dcToDrawIn, &iconRectangle);
				}
				ImageList_Draw(properties.hImageList, iconToDraw, dcToDrawIn, iconRectangle.left, iconRectangle.top, ILD_NORMAL);
			}

			// draw the text
			if(textLength) {
				if(RunTimeHelper::IsVista()) {
					DrawThemeParentBackgroundEx(*this, dcToDrawIn, DTPB_USECTLCOLORSTATIC, &textRectangle);
				} else {
					dcToDrawIn.FillRect(&textRectangle, GetSysColorBrush(COLOR_BTNFACE));
					DrawThemeParentBackground(*this, dcToDrawIn, &textRectangle);
				}
			}
			dcToDrawIn.RestoreDC(state);
			if(textLength) {
				themingEngine.DrawThemeText(dcToDrawIn, BP_GROUPBOX, (enabled ? GBS_NORMAL : GBS_DISABLED), CW2CW(text), textLength, DT_LEFT | DT_TOP | DT_SINGLELINE | (drawAccel ? 0 : DT_HIDEPREFIX) | (extendedStyle & WS_EX_RTLREADING ? DT_RTLREADING : 0), NULL, &textRectangle);
			}
		}
	}
	dcToDrawIn.SetBkMode(previousBkMode);
	dcToDrawIn.SelectFont(hPreviousFont);
	dcToDrawIn.SetTextColor(previousForeColor);
	dcToDrawIn.SetBkColor(previousBackColor);
}


inline HRESULT Frame::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if(flags.handlingBM_CLICK) {
			x = 0;
			y = 0;
		}

		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		/*if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
		}*/

		return Fire_ContextMenu(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_DestroyedControlWindow(HWND hWnd)
{
	KillTimer(timers.ID_REDRAW);

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(flags.handlingBM_CLICK) {
			x = 0;
			y = 0;
		}

		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		fireMouseDown = TRUE;
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
			if(properties.style == sOwnerDrawn) {
				Invalidate();
			}
		}
	}
	mouseStatus.StoreClickCandidate(button);
	if(!flags.handlingBM_CLICK) {
		SetCapture();
	}

	if(mouseStatus.IsDoubleClickCandidate(button)) {
		/* Watch for double-clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		if(!flags.handlingBM_CLICK) {
			DWORD position = GetMessagePos();
			POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			RECT clientArea = {0};
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
			if(PtInRect(&clientArea, cursorPosition)) {
				// maybe the control is overlapped by a foreign window
				if(WindowFromPoint(cursorPosition) != *this) {
					hasLeftControl = TRUE;
				}
			} else {
				hasLeftControl = TRUE;
			}
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the double-click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_DblClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RDblClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MDblClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XDblClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.RemoveAllDoubleClickCandidates();
		}

		mouseStatus.RemoveClickCandidate(button);
		if(GetCapture() == *this) {
			ReleaseCapture();
		}
		fireMouseDown = FALSE;
	} else {
		mouseStatus.RemoveAllDoubleClickCandidates();
	}

	HRESULT hr = S_OK;
	if(fireMouseDown) {
		hr = Fire_MouseDown(button, shift, x, y);
	}
	return hr;
}

inline HRESULT Frame::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		if(flags.handlingBM_CLICK) {
			x = 0;
			y = 0;
		}

		return Fire_MouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();
	if(properties.style == sOwnerDrawn) {
		Invalidate();
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(flags.handlingBM_CLICK) {
			x = 0;
			y = 0;
		}

		return Fire_MouseHover(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT Frame::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.LeaveControl();
	if(properties.style == sOwnerDrawn) {
		Invalidate();
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(flags.handlingBM_CLICK) {
			x = 0;
			y = 0;
		}

		return Fire_MouseLeave(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT Frame::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(flags.handlingBM_CLICK) {
		x = 0;
		y = 0;
	}

	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(flags.handlingBM_CLICK) {
		x = 0;
		y = 0;
	}

	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		if(!flags.handlingBM_CLICK) {
			DWORD position = GetMessagePos();
			POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			RECT clientArea = {0};
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
			if(PtInRect(&clientArea, cursorPosition)) {
				// maybe the control is overlapped by a foreign window
				if(WindowFromPoint(cursorPosition) != *this) {
					hasLeftControl = TRUE;
				}
			} else {
				hasLeftControl = TRUE;
			}
			if(GetCapture() == *this) {
				ReleaseCapture();
			}
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.StoreDoubleClickCandidate(button);
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT Frame::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Frame::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT Frame::Raise_OLEDragLeave(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Frame::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT Frame::Raise_OwnerDraw(OwnerDrawActionConstants requiredAction, OwnerDrawControlStateConstants controlState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDraw(requiredAction, controlState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_RecreatedControlWindow(HWND hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT Frame::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


void Frame::RecreateControlWindow(void)
{
	// This method shouldn't be used, because it will destroy all contained controls.
	ATLASSERT(FALSE);
	/*if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}*/
}

DWORD Frame::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD Frame::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	style |= BS_GROUPBOX | BS_NOTIFY;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.borderVisible) {
		style &= ~BS_TYPEMASK;
		style |= BS_OWNERDRAW;
	}
	switch(properties.contentType) {
		case ctText:
			style |= BS_TEXT;
			break;
		case ctBitmap:
			style |= BS_BITMAP;
			break;
		case ctIcon:
			style |= BS_ICON;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}
	switch(properties.hAlignment) {
		case halLeft:
			style |= BS_LEFT;
			break;
		case halCenter:
			style |= BS_CENTER;
			break;
		case halRight:
			style |= BS_RIGHT;
			break;
	}
	switch(properties.style) {
		case sFlat:
			style |= BS_FLAT;
			break;
		case sOwnerDrawn:
			style &= ~BS_TYPEMASK;
			style |= BS_OWNERDRAW;
			break;
	}
	if(RunTimeHelper::IsCommCtrl6()) {
		style &= ~BS_TYPEMASK;
		style |= BS_OWNERDRAW;
	}
	return style;
}

void Frame::SendConfigurationMessages(void)
{
	SetWindowText(COLE2CT(properties.text));
	switch(properties.contentType) {
		case ctBitmap:
			SendMessage(BM_SETIMAGE, IMAGE_BITMAP, reinterpret_cast<LPARAM>(properties.image));
			break;
		case ctIcon:
			SendMessage(BM_SETIMAGE, IMAGE_ICON, reinterpret_cast<LPARAM>(properties.image));
			break;
	}
	if(RunTimeHelper::IsCommCtrl6()) {
		BUTTON_IMAGELIST imageListSettings = {0};
		if(properties.useImprovedImageListSupport) {
			imageListSettings.himl = (properties.hImageList ? IImageListToHIMAGELIST(properties.pImageListWrapper) : BCCL_NOGLYPH);
		} else {
			imageListSettings.himl = (properties.hImageList ? properties.hImageList : BCCL_NOGLYPH);
		}
		imageListSettings.uAlign = properties.iconAlignment;
		imageListSettings.margin.bottom = properties.iconMarginBottom;
		imageListSettings.margin.left = properties.iconMarginLeft;
		imageListSettings.margin.right = properties.iconMarginRight;
		imageListSettings.margin.top = properties.iconMarginTop;
		SendMessage(BCM_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(&imageListSettings));
	}

	ApplyFont();
}

HCURSOR Frame::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

BOOL Frame::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}