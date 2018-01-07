// OffsetImageList.cpp: An implementation of the IImageList interface wrapping and customizing an image list

#include "stdafx.h"
#include "OffsetImageList.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE OffsetImageList_AddRef(IImageList2* pInterface)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	return InterlockedIncrement(&pThis->referenceCount);
}

ULONG STDMETHODCALLTYPE OffsetImageList_IImageListPrivate_AddRef(IImageListPrivate* pInterface)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromOffsetImageList(pThis);
	return OffsetImageList_AddRef(pBaseInterface);
}

ULONG STDMETHODCALLTYPE OffsetImageList_Release(IImageList2* pInterface)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	ULONG ret = InterlockedDecrement(&pThis->referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		if(pThis->pWrappedImageList) {
			pThis->pWrappedImageList->Release();
		}
		HeapFree(GetProcessHeap(), 0, pThis);
	}
	return ret;
}

ULONG STDMETHODCALLTYPE OffsetImageList_IImageListPrivate_Release(IImageListPrivate* pInterface)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromOffsetImageList(pThis);
	return OffsetImageList_Release(pBaseInterface);
}

STDMETHODIMP OffsetImageList_QueryInterface(IImageList2* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = reinterpret_cast<LPUNKNOWN>(&pThis->pImageList2Vtable);
	} else if(requiredInterface == IID_IImageList) {
		*ppInterfaceImpl = IImageListFromOffsetImageList(pThis);
	} else if(requiredInterface == IID_IImageList2) {
		*ppInterfaceImpl = IImageList2FromOffsetImageList(pThis);
	} else if(requiredInterface == IID_IImageListPrivate) {
		*ppInterfaceImpl = IImageListPrivateFromOffsetImageList(pThis);
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	OffsetImageList_AddRef(pInterface);
	return S_OK;
}

STDMETHODIMP OffsetImageList_IImageListPrivate_QueryInterface(IImageListPrivate* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromOffsetImageList(pThis);
	return OffsetImageList_QueryInterface(pBaseInterface, requiredInterface, ppInterfaceImpl);
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList
STDMETHODIMP OffsetImageList_Add(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "OffsetImageList_Add() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_ReplaceIcon(IImageList2* pInterface, int i, HICON hicon, int* pi)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->ReplaceIcon(pThis->iconIndexMappings[i], hicon, pi);
}

STDMETHODIMP OffsetImageList_SetOverlayImage(IImageList2* pInterface, int iImage, int iOverlay)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iImage < 0 || iImage > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->SetOverlayImage(pThis->iconIndexMappings[iImage], iOverlay);
}

STDMETHODIMP OffsetImageList_Replace(IImageList2* pInterface, int i, HBITMAP hbmImage, HBITMAP hbmMask)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->Replace(pThis->iconIndexMappings[i], hbmImage, hbmMask);
}

STDMETHODIMP OffsetImageList_AddMasked(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, COLORREF /*crMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "OffsetImageList_AddMasked() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_Draw(IImageList2* pInterface, IMAGELISTDRAWPARAMS* pimldp)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(pimldp->i < 0 || pimldp->i > 5) {
		return E_INVALIDARG;
	}
	IMAGELISTDRAWPARAMS params = *pimldp;
	params.himl = IImageListToHIMAGELIST(pThis->pWrappedImageList);
	params.i = pThis->iconIndexMappings[pimldp->i];
	return pThis->pWrappedImageList->Draw(&params);
}

STDMETHODIMP OffsetImageList_Remove(IImageList2* /*pInterface*/, int /*i*/)
{
	ATLASSERT(FALSE && "OffsetImageList_Remove() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_GetIcon(IImageList2* pInterface, int i, UINT flags, HICON* picon)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetIcon(pThis->iconIndexMappings[i], flags, picon);
}

STDMETHODIMP OffsetImageList_GetImageInfo(IImageList2* pInterface, int i, IMAGEINFO* pImageInfo)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetImageInfo(pThis->iconIndexMappings[i], pImageInfo);
}

STDMETHODIMP OffsetImageList_Copy(IImageList2* pInterface, int iDst, IUnknown* punkSrc, int iSrc, UINT uFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iDst < 0 || iDst > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->Copy(pThis->iconIndexMappings[iDst], punkSrc, iSrc, uFlags);
}

STDMETHODIMP OffsetImageList_Merge(IImageList2* pInterface, int i1, IUnknown* punk2, int i2, int dx, int dy, REFIID riid, void** ppv)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i1 < 0 || i1 > 5) {
		return E_INVALIDARG;
	}
	if(punk2 == pInterface) {
		if(i2 < 0 || i2 > 5) {
			return E_INVALIDARG;
		}
		return pThis->pWrappedImageList->Merge(pThis->iconIndexMappings[i1], punk2, pThis->iconIndexMappings[i2], dx, dy, riid, ppv);
	} else {
		return pThis->pWrappedImageList->Merge(pThis->iconIndexMappings[i1], punk2, i2, dx, dy, riid, ppv);
	}
}

STDMETHODIMP OffsetImageList_Clone(IImageList2* pInterface, REFIID riid, void** ppv)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}

	IImageList* pClone = NULL;
	HRESULT hr = OffsetImageList_CreateInstance(&pClone);
	if(SUCCEEDED(hr)) {
		OffsetImageList* pCloneInstance = OffsetImageListFromIImageList(pClone);
		CopyMemory(pCloneInstance->iconIndexMappings, pThis->iconIndexMappings, 6 * sizeof(int));
		hr = pClone->QueryInterface(riid, ppv);
		if(FAILED(hr)) {
			pClone->Release();
		}
	}
	return hr;
}

STDMETHODIMP OffsetImageList_GetImageRect(IImageList2* pInterface, int i, RECT* prc)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetImageRect(pThis->iconIndexMappings[i], prc);
}

STDMETHODIMP OffsetImageList_GetIconSize(IImageList2* pInterface, int* cx, int* cy)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetIconSize(cx, cy);
}

STDMETHODIMP OffsetImageList_SetIconSize(IImageList2* pInterface, int cx, int cy)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->SetIconSize(cx, cy);
}

STDMETHODIMP OffsetImageList_GetImageCount(IImageList2* pInterface, int* pi)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetImageCount(pi);
}

STDMETHODIMP OffsetImageList_SetImageCount(IImageList2* pInterface, UINT uNewCount)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList || uNewCount > 6) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->SetImageCount(min(6, uNewCount));
}

STDMETHODIMP OffsetImageList_SetBkColor(IImageList2* pInterface, COLORREF clrBk, COLORREF* pclr)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->SetBkColor(clrBk, pclr);
}

STDMETHODIMP OffsetImageList_GetBkColor(IImageList2* pInterface, COLORREF* pclr)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetBkColor(pclr);
}

STDMETHODIMP OffsetImageList_BeginDrag(IImageList2* pInterface, int iTrack, int dxHotspot, int dyHotspot)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iTrack < 0 || iTrack > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->BeginDrag(pThis->iconIndexMappings[iTrack], dxHotspot, dyHotspot);
}

STDMETHODIMP OffsetImageList_EndDrag(IImageList2* pInterface)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->EndDrag();
}

STDMETHODIMP OffsetImageList_DragEnter(IImageList2* pInterface, HWND hwndLock, int x, int y)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->DragEnter(hwndLock, x, y);
}

STDMETHODIMP OffsetImageList_DragLeave(IImageList2* pInterface, HWND hwndLock)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->DragLeave(hwndLock);
}

STDMETHODIMP OffsetImageList_DragMove(IImageList2* pInterface, int x, int y)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->DragMove(x, y);
}

STDMETHODIMP OffsetImageList_SetDragCursorImage(IImageList2* pInterface, IUnknown* punk, int iDrag, int dxHotspot, int dyHotspot)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iDrag < 0 || iDrag > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->SetDragCursorImage(punk, pThis->iconIndexMappings[iDrag], dxHotspot, dyHotspot);
}

STDMETHODIMP OffsetImageList_DragShowNolock(IImageList2* pInterface, BOOL fShow)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->DragShowNolock(fShow);
}

STDMETHODIMP OffsetImageList_GetDragImage(IImageList2* pInterface, POINT* ppt, POINT* pptHotspot, REFIID riid, void** ppv)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetDragImage(ppt, pptHotspot, riid, ppv);
}

STDMETHODIMP OffsetImageList_GetItemFlags(IImageList2* pInterface, int i, DWORD* dwFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetItemFlags(pThis->iconIndexMappings[i], dwFlags);
}

STDMETHODIMP OffsetImageList_GetOverlayImage(IImageList2* pInterface, int iOverlay, int* piIndex)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iOverlay < 0 || iOverlay > 5) {
		return E_INVALIDARG;
	}
	return pThis->pWrappedImageList->GetOverlayImage(pThis->iconIndexMappings[iOverlay], piIndex);
}
// implementation of IImageList
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList2
STDMETHODIMP OffsetImageList_Resize(IImageList2* pInterface, int cxNewIconSize, int cyNewIconSize)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->Resize(cxNewIconSize, cyNewIconSize);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_GetOriginalSize(IImageList2* pInterface, int iImage, DWORD dwFlags, int* pcx, int* pcy)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iImage < 0 || iImage > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->GetOriginalSize(pThis->iconIndexMappings[iImage], dwFlags, pcx, pcy);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_SetOriginalSize(IImageList2* pInterface, int iImage, int cx, int cy)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iImage < 0 || iImage > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->SetOriginalSize(pThis->iconIndexMappings[iImage], cx, cy);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_SetCallback(IImageList2* pInterface, IUnknown* punk)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->SetCallback(punk);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_GetCallback(IImageList2* pInterface, REFIID riid, void** ppv)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->GetCallback(riid, ppv);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_ForceImagePresent(IImageList2* pInterface, int iImage, DWORD dwFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iImage < 0 || iImage > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->ForceImagePresent(pThis->iconIndexMappings[iImage], dwFlags);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_DiscardImages(IImageList2* pInterface, int iFirstImage, int iLastImage, DWORD dwFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(iFirstImage < 0 || iFirstImage > 5) {
		return E_INVALIDARG;
	}
	if(iLastImage < 0 || iLastImage > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		HRESULT hr = S_OK;
		for(int i = iFirstImage; i <= iLastImage && SUCCEEDED(hr); i++) {
			hr = pWrappedImageList->DiscardImages(pThis->iconIndexMappings[i], pThis->iconIndexMappings[i], dwFlags);
		}
		return hr;
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_PreloadImages(IImageList2* pInterface, IMAGELISTDRAWPARAMS* pimldp)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(pimldp->i < 0 || pimldp->i > 5) {
		return E_INVALIDARG;
	}
	IMAGELISTDRAWPARAMS params = *pimldp;
	params.himl = IImageListToHIMAGELIST(pThis->pWrappedImageList);
	params.i = pThis->iconIndexMappings[pimldp->i];
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->PreloadImages(&params);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_GetStatistics(IImageList2* pInterface, IMAGELISTSTATS* pils)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->GetStatistics(pils);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_Initialize(IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/)
{
	ATLASSERT(FALSE && "OffsetImageList_Initialize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_Replace2(IImageList2* pInterface, int i, HBITMAP hbmImage, HBITMAP hbmMask, IUnknown* punk, DWORD dwFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->Replace2(pThis->iconIndexMappings[i], hbmImage, hbmMask, punk, dwFlags);
	}
	return E_INVALIDARG;
}

STDMETHODIMP OffsetImageList_ReplaceFromImageList(IImageList2* pInterface, int i, IImageList* pil, int iSrc, IUnknown* punk, DWORD dwFlags)
{
	OffsetImageList* pThis = OffsetImageListFromIImageList2(pInterface);
	if(!pThis->pWrappedImageList) {
		return E_INVALIDARG;
	}
	if(i < 0 || i > 5) {
		return E_INVALIDARG;
	}
	CComQIPtr<IImageList2> pWrappedImageList = pThis->pWrappedImageList;
	if(pWrappedImageList) {
		return pWrappedImageList->ReplaceFromImageList(pThis->iconIndexMappings[i], pil, iSrc, punk, dwFlags);
	}
	return E_INVALIDARG;
}
// implementation of IImageList2
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageListPrivate
STDMETHODIMP OffsetImageList_GetImageList(IImageListPrivate* pInterface, UINT imageListType, HIMAGELIST* pHImageList)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	switch(imageListType) {
		case OIL_NORMAL:
			*pHImageList = IImageListToHIMAGELIST(pThis->pWrappedImageList);
			return S_OK;
			break;
	}
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_SetImageList(IImageListPrivate* pInterface, UINT imageListType, HIMAGELIST hImageList, HIMAGELIST* pHPreviousImageList)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	switch(imageListType) {
		case OIL_NORMAL:
			IImageList* pTmp = NULL;
			HRESULT hr = WrappedHIMAGELIST_QueryInterface(hImageList, IID_PPV_ARGS(&pTmp));
			if(SUCCEEDED(hr)) {
				if(pHPreviousImageList) {
					*pHPreviousImageList = IImageListToHIMAGELIST(pThis->pWrappedImageList);
				}
				if(pThis->pWrappedImageList) {
					pThis->pWrappedImageList->Release();
				}
				pThis->pWrappedImageList = pTmp;
				return S_OK;
			} else {
				return hr;
			}
			break;
	}
	return E_NOTIMPL;
}

STDMETHODIMP OffsetImageList_GetIndexMapping(IImageListPrivate* pInterface, int externalIconIndex, PINT pRealIconIndex)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	if(externalIconIndex < 0 || externalIconIndex > 5) {
		return E_INVALIDARG;
	}
	*pRealIconIndex = pThis->iconIndexMappings[externalIconIndex];
	return S_OK;
}

STDMETHODIMP OffsetImageList_SetIndexMapping(IImageListPrivate* pInterface, int externalIconIndex, int realIconIndex)
{
	OffsetImageList* pThis = OffsetImageListFromIImageListPrivate(pInterface);
	if(externalIconIndex < 0 || externalIconIndex > 5) {
		return E_INVALIDARG;
	}
	pThis->iconIndexMappings[externalIconIndex] = realIconIndex;
	return S_OK;
}
// implementation of IImageListPrivate
//////////////////////////////////////////////////////////////////////


STDMETHODIMP OffsetImageList_CreateInstance(IImageList** ppObject)
{
	if(!ppObject) {
		return E_POINTER;
	}

	OffsetImageList* pInstance = static_cast<OffsetImageList*>(HeapAlloc(GetProcessHeap(), 0, sizeof(OffsetImageList)));
	if(!pInstance) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pInstance, sizeof(*pInstance));
	pInstance->magicValue = 0x4C4D4948;
	pInstance->pImageList2Vtable = &IImageList2Impl_Vtbl;
	pInstance->pImageListPrivateVtable = &IImageListPrivateImpl_Vtbl;
	pInstance->referenceCount = 1;

	*ppObject = IImageListFromOffsetImageList(pInstance);
	return S_OK;
}


HRESULT WrappedHIMAGELIST_QueryInterface(HIMAGELIST hImageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST himl, __in REFIID riid, __deref_out void** ppid);
	HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = NULL;

	HMODULE hComctl32 = GetModuleHandle(TEXT("comctl32.dll"));
	if(hComctl32) {
		pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hComctl32, "HIMAGELIST_QueryInterface"));
	}
	if(pfnHIMAGELIST_QueryInterface) {
		return pfnHIMAGELIST_QueryInterface(hImageList, requiredInterface, ppInterfaceImpl);
	} else {
		return E_NOTIMPL;
	}
}