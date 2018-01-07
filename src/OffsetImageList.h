//////////////////////////////////////////////////////////////////////
/// \file OffsetImageList.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements a class (C style) that wraps an image list allowing to adjust the icon indexes</em>
///
/// The pseudo class provided by this file wraps an image list and implements the \c IImageList interface
/// through which the wrapped and customized image list can be accessed like a normal image list. This is
/// used to provide a mechanism that allows usage of a single image list for different button controls.
///
/// \sa CheckBox, CommandButton, Frame, OptionButton, IImageListPrivate
//////////////////////////////////////////////////////////////////////


#pragma once

#include <commoncontrols.h>
#include "IImageListPrivate.h"
#include "helpers.h"


/// \brief <em>The layout of the vtable of the \c IImageList2 interface</em>
///
/// \sa IImageList2Impl_Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
typedef struct IImageList2Vtbl
{
	HRESULT(STDMETHODCALLTYPE* QueryInterface)(IImageList2* This, REFIID riid, __RPC__deref_out void** ppvObject);
	ULONG(STDMETHODCALLTYPE* AddRef)(IImageList2* This);
	ULONG(STDMETHODCALLTYPE* Release)(IImageList2* This);
	HRESULT(STDMETHODCALLTYPE* Add)(IImageList2* This, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* ReplaceIcon)(IImageList2* This, int i, __in HICON hicon, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetOverlayImage)(IImageList2* This, int iImage, int iOverlay);
	HRESULT(STDMETHODCALLTYPE* Replace)(IImageList2* This, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask);
	HRESULT(STDMETHODCALLTYPE* AddMasked)(IImageList2* This, __in HBITMAP hbmImage, COLORREF crMask, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* Draw)(IImageList2* This, __in IMAGELISTDRAWPARAMS* pimldp);
	HRESULT(STDMETHODCALLTYPE* Remove)(IImageList2* This, int i);
	HRESULT(STDMETHODCALLTYPE* GetIcon)(IImageList2* This, int i, UINT flags, __out HICON* picon);
	HRESULT(STDMETHODCALLTYPE* GetImageInfo)(IImageList2* This, int i, __out IMAGEINFO* pImageInfo);
	HRESULT(STDMETHODCALLTYPE* Copy)(IImageList2* This, int iDst, __in IUnknown* punkSrc, int iSrc, UINT uFlags);
	HRESULT(STDMETHODCALLTYPE* Merge)(IImageList2* This, int i1, __in IUnknown* punk2, int i2, int dx, int dy, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* Clone)(IImageList2* This, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetImageRect)(IImageList2* This, int i, __out RECT* prc);
	HRESULT(STDMETHODCALLTYPE* GetIconSize)(IImageList2* This, __out int* cx, __out int* cy);
	HRESULT(STDMETHODCALLTYPE* SetIconSize)(IImageList2* This, int cx, int cy);
	HRESULT(STDMETHODCALLTYPE* GetImageCount)(IImageList2* This, __out int* pi);
	HRESULT(STDMETHODCALLTYPE* SetImageCount)(IImageList2* This, UINT uNewCount);
	HRESULT(STDMETHODCALLTYPE* SetBkColor)(IImageList2* This, COLORREF clrBk, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* GetBkColor)(IImageList2* This, __out COLORREF* pclr);
	HRESULT(STDMETHODCALLTYPE* BeginDrag)(IImageList2* This, int iTrack, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* EndDrag)(IImageList2* This);
	HRESULT(STDMETHODCALLTYPE* DragEnter)(IImageList2* This, __in_opt HWND hwndLock, int x, int y);
	HRESULT(STDMETHODCALLTYPE* DragLeave)(IImageList2* This, __in_opt HWND hwndLock);
	HRESULT(STDMETHODCALLTYPE* DragMove)(IImageList2* This, int x, int y);
	HRESULT(STDMETHODCALLTYPE* SetDragCursorImage)(IImageList2* This, __in IUnknown* punk, int iDrag, int dxHotspot, int dyHotspot);
	HRESULT(STDMETHODCALLTYPE* DragShowNolock)(IImageList2* This, BOOL fShow);
	HRESULT(STDMETHODCALLTYPE* GetDragImage)(IImageList2* This, __out_opt POINT* ppt, __out_opt POINT* pptHotspot, REFIID riid, __deref_out void** ppv);
	HRESULT(STDMETHODCALLTYPE* GetItemFlags)(IImageList2* This, int i, __out DWORD* dwFlags);
	HRESULT(STDMETHODCALLTYPE* GetOverlayImage)(IImageList2* This, int iOverlay, __out int* piIndex);
	HRESULT(STDMETHODCALLTYPE* Resize)(IImageList2* This, int cxNewIconSize, int cyNewIconSize);
	HRESULT(STDMETHODCALLTYPE* GetOriginalSize)(IImageList2* This, int iImage, DWORD dwFlags, __out int* pcx, __out int* pcy);
 	HRESULT(STDMETHODCALLTYPE* SetOriginalSize)(IImageList2* This, int iImage, int cx, int cy);
 	HRESULT(STDMETHODCALLTYPE* SetCallback)(IImageList2* This, __in_opt IUnknown* punk);
 	HRESULT(STDMETHODCALLTYPE* GetCallback)(IImageList2* This, REFIID riid, __deref_out void** ppv);
 	HRESULT(STDMETHODCALLTYPE* ForceImagePresent)(IImageList2* This, int iImage, DWORD dwFlags);
 	HRESULT(STDMETHODCALLTYPE* DiscardImages)(IImageList2* This, int iFirstImage, int iLastImage, DWORD dwFlags);
 	HRESULT(STDMETHODCALLTYPE* PreloadImages)(IImageList2* This, __in IMAGELISTDRAWPARAMS* pimldp);
 	HRESULT(STDMETHODCALLTYPE* GetStatistics)(IImageList2* This, __inout IMAGELISTSTATS* pils);
 	HRESULT(STDMETHODCALLTYPE* Initialize)(IImageList2* This, int cx, int cy, UINT flags, int cInitial, int cGrow);
	HRESULT(STDMETHODCALLTYPE* Replace2)(IImageList2* This, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __in_opt IUnknown* punk, DWORD dwFlags);
	HRESULT(STDMETHODCALLTYPE* ReplaceFromImageList)(IImageList2* This, int i, __in IImageList* pil, int iSrc, __in_opt IUnknown* punk, DWORD dwFlags);
} IImageList2Vtbl;


/// \brief <em>Creates a new \c OffsetImageList instance</em>
///
/// \param[out] ppObject Receives a pointer to the created instance's \c IImageList interface.
///
/// \return An \c HRESULT error code.
///
/// \sa OffsetImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_CreateInstance(__deref_out IImageList** ppObject);


//////////////////////////////////////////////////////////////////////
/// \name Implementation of IUnknown
///
//@{
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa OffsetImageList_IImageListPrivate_AddRef, OffsetImageList_Release, OffsetImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE OffsetImageList_AddRef(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::AddRef</em>
///
/// \return The new reference count.
///
/// \sa OffsetImageList_AddRef, OffsetImageList_IImageListPrivate_Release,
///     OffsetImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
extern ULONG STDMETHODCALLTYPE OffsetImageList_IImageListPrivate_AddRef(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa OffsetImageList_IImageListPrivate_Release, OffsetImageList_AddRef, OffsetImageList_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE OffsetImageList_Release(__in IImageList2* pInterface);
/// \brief <em>Implements \c IUnknown::Release</em>
///
/// \return The new reference count.
///
/// \sa OffsetImageList_Release, OffsetImageList_IImageListPrivate_AddRef,
///     OffsetImageList_IImageListPrivate_QueryInterface,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
extern ULONG STDMETHODCALLTYPE OffsetImageList_IImageListPrivate_Release(__in IImageListPrivate* pInterface);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa OffsetImageList_IImageListPrivate_QueryInterface, OffsetImageList_AddRef, OffsetImageList_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_QueryInterface(__in IImageList2* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
/// \brief <em>Implements \c IUnknown::QueryInterface</em>
///
/// \return An \c HRESULT error code.
///
/// \sa OffsetImageList_QueryInterface, OffsetImageList_IImageListPrivate_AddRef,
///     OffsetImageList_IImageListPrivate_Release,
///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_IImageListPrivate_QueryInterface(__in IImageListPrivate* pInterface, REFIID requiredInterface, __RPC__deref_out LPVOID* ppInterfaceImpl);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageList
///
//@{
/// \brief <em>Wraps \c IImageList::Add</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761435.aspx">IImageList::Add</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Add(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, __in_opt HBITMAP /*hbmMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::ReplaceIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761498.aspx">IImageList::ReplaceIcon</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_ReplaceIcon(__in IImageList2* pInterface, int i, __in HICON hicon, __out int* pi);
/// \brief <em>Wraps \c IImageList::SetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761508.aspx">IImageList::SetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetOverlayImage(__in IImageList2* pInterface, int iImage, int iOverlay);
/// \brief <em>Wraps \c IImageList::Replace</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761496.aspx">IImageList::Replace</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Replace(__in IImageList2* pInterface, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask);
/// \brief <em>Wraps \c IImageList::AddMasked</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList::AddMasked</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_AddMasked(__in IImageList2* /*pInterface*/, __in HBITMAP /*hbmImage*/, COLORREF /*crMask*/, __out int* /*pi*/);
/// \brief <em>Wraps \c IImageList::Draw</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761455.aspx">IImageList::Draw</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Draw(__in IImageList2* pInterface, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Wraps \c IImageList::Remove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761494.aspx">IImageList::Remove</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Remove(__in IImageList2* /*pInterface*/, int /*i*/);
/// \brief <em>Wraps \c IImageList::GetIcon</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761463.aspx">IImageList::GetIcon</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetIcon(__in IImageList2* pInterface, int i, UINT flags, __out HICON* picon);
/// \brief <em>Wraps \c IImageList::GetImageInfo</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761482.aspx">IImageList::GetImageInfo</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetImageInfo(__in IImageList2* pInterface, int i, __out IMAGEINFO* pImageInfo);
/// \brief <em>Wraps \c IImageList::Copy</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761443.aspx">IImageList::Copy</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Copy(__in IImageList2* pInterface, int iDst, __in IUnknown* punkSrc, int iSrc, UINT uFlags);
/// \brief <em>Wraps \c IImageList::Merge</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761492.aspx">IImageList::Merge</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Merge(__in IImageList2* pInterface, int i1, __in IUnknown* punk2, int i2, int dx, int dy, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::Clone</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761442.aspx">IImageList::Clone</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Clone(__in IImageList2* pInterface, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::GetImageRect</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761484.aspx">IImageList::GetImageRect</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetImageRect(__in IImageList2* pInterface, int i, __out RECT* prc);
/// \brief <em>Wraps \c IImageList::GetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761478.aspx">IImageList::GetIconSize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetIconSize(__in IImageList2* pInterface, __out int* cx, __out int* cy);
/// \brief <em>Wraps \c IImageList::SetIconSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761504.aspx">IImageList::SetIconSize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetIconSize(__in IImageList2* pInterface, int cx, int cy);
/// \brief <em>Wraps \c IImageList::GetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761480.aspx">IImageList::GetImageCount</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetImageCount(__in IImageList2* pInterface, __out int* pi);
/// \brief <em>Wraps \c IImageList::SetImageCount</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761506.aspx">IImageList::SetImageCount</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetImageCount(__in IImageList2* pInterface, UINT uNewCount);
/// \brief <em>Wraps \c IImageList::SetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761500.aspx">IImageList::SetBkColor</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetBkColor(__in IImageList2* pInterface, COLORREF clrBk, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::GetBkColor</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761459.aspx">IImageList::GetBkColor</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetBkColor(__in IImageList2* pInterface, __out COLORREF* pclr);
/// \brief <em>Wraps \c IImageList::BeginDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761440.aspx">IImageList::BeginDrag</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_BeginDrag(__in IImageList2* pInterface, int iTrack, int dxHotspot, int dyHotspot);
/// \brief <em>Wraps \c IImageList::EndDrag</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761457.aspx">IImageList::EndDrag</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_EndDrag(__in IImageList2* pInterface);
/// \brief <em>Wraps \c IImageList::DragEnter</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761446.aspx">IImageList::DragEnter</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_DragEnter(__in IImageList2* pInterface, __in_opt HWND hwndLock, int x, int y);
/// \brief <em>Wraps \c IImageList::DragLeave</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761448.aspx">IImageList::DragLeave</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_DragLeave(__in IImageList2* pInterface, __in_opt HWND hwndLock);
/// \brief <em>Wraps \c IImageList::DragMove</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761451.aspx">IImageList::DragMove</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_DragMove(__in IImageList2* pInterface, int x, int y);
/// \brief <em>Wraps \c IImageList::SetDragCursorImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761502.aspx">IImageList::SetDragCursorImage</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetDragCursorImage(__in IImageList2* pInterface, __in IUnknown* punk, int iDrag, int dxHotspot, int dyHotspot);
/// \brief <em>Wraps \c IImageList::DragShowNolock</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761453.aspx">IImageList::DragShowNolock</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_DragShowNolock(__in IImageList2* pInterface, BOOL fShow);
/// \brief <em>Wraps \c IImageList::GetDragImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761461.aspx">IImageList::GetDragImage</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetDragImage(__in IImageList2* pInterface, __out_opt POINT* ppt, __out_opt POINT* pptHotspot, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList::GetItemFlags</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761486.aspx">IImageList::GetItemFlags</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetItemFlags(__in IImageList2* pInterface, int i, __out DWORD* dwFlags);
/// \brief <em>Wraps \c IImageList::GetOverlayImage</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761488.aspx">IImageList::GetOverlayImage</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetOverlayImage(__in IImageList2* pInterface, int iOverlay, __out int* piIndex);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageList2
///
//@{
/// \brief <em>Wraps \c IImageList2::Resize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761429.aspx">IImageList2::Resize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Resize(__in IImageList2* pInterface, int cxNewIconSize, int cyNewIconSize);
/// \brief <em>Wraps \c IImageList2::GetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761415.aspx">IImageList2::GetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetOriginalSize(__in IImageList2* pInterface, int iImage, DWORD dwFlags, __out int* pcx, __out int* pcy);
/// \brief <em>Wraps \c IImageList2::SetOriginalSize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761433.aspx">IImageList2::SetOriginalSize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetOriginalSize(__in IImageList2* pInterface, int iImage, int cx, int cy);
/// \brief <em>Wraps \c IImageList2::SetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761431.aspx">IImageList2::SetCallback</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetCallback(__in IImageList2* pInterface, __in_opt IUnknown* punk);
/// \brief <em>Wraps \c IImageList2::GetCallback</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761438.aspx">IImageList2::GetCallback</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetCallback(__in IImageList2* pInterface, REFIID riid, __deref_out void** ppv);
/// \brief <em>Wraps \c IImageList2::ForceImagePresent</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761411.aspx">IImageList2::ForceImagePresent</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_ForceImagePresent(__in IImageList2* pInterface, int iImage, DWORD dwFlags);
/// \brief <em>Wraps \c IImageList2::DiscardImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761409.aspx">IImageList2::DiscardImages</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_DiscardImages(__in IImageList2* pInterface, int iFirstImage, int iLastImage, DWORD dwFlags);
/// \brief <em>Wraps \c IImageList2::PreloadImages</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761423.aspx">IImageList2::PreloadImages</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_PreloadImages(__in IImageList2* pInterface, __in IMAGELISTDRAWPARAMS* pimldp);
/// \brief <em>Wraps \c IImageList2::GetStatistics</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761417.aspx">IImageList2::GetStatistics</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetStatistics(__in IImageList2* pInterface, __inout IMAGELISTSTATS* pils);
/// \brief <em>Wraps \c IImageList2::Initialize</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761421.aspx">IImageList2::Initialize</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Initialize(__in IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/);
/// \brief <em>Wraps \c IImageList2::Replace2</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761425.aspx">IImageList2::Replace2</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_Replace2(__in IImageList2* pInterface, int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __in_opt IUnknown* punk, DWORD dwFlags);
/// \brief <em>Wraps \c IImageList2::ReplaceFromImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761427.aspx">IImageList2::ReplaceFromImageList</a>
extern HRESULT STDMETHODCALLTYPE OffsetImageList_ReplaceFromImageList(__in IImageList2* pInterface, int i, __in IImageList* pil, int iSrc, __in_opt IUnknown* punk, DWORD dwFlags);
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Implementation of IImageListPrivate
///
//@{
/// \brief <em>Wraps \c IImageListPrivate::GetImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetImageList
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetImageList(__in IImageListPrivate* pInterface, UINT imageListType, __out HIMAGELIST* pHImageList);
/// \brief <em>Wraps \c IImageListPrivate::SetImageList</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetImageList
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetImageList(__in IImageListPrivate* pInterface, UINT imageListType, __in_opt HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList);
/// \brief <em>Wraps \c IImageListPrivate::GetIndexMapping</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::GetIndexMapping
extern HRESULT STDMETHODCALLTYPE OffsetImageList_GetIndexMapping(__in IImageListPrivate* pInterface, int externalIconIndex, __out PINT pRealIconIndex);
/// \brief <em>Wraps \c IImageListPrivate::SetIndexMapping</em>
///
/// \return An \c HRESULT error code.
///
/// \sa IImageListPrivate::SetIndexMapping
extern HRESULT STDMETHODCALLTYPE OffsetImageList_SetIndexMapping(__in IImageListPrivate* pInterface, int externalIconIndex, int realIconIndex);
//@}
//////////////////////////////////////////////////////////////////////


/// \brief <em>The vtable of our implementation of the \c IImageList2 interface</em>
///
/// \sa IImageList2Vtbl,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static const IImageList2Vtbl IImageList2Impl_Vtbl =
{
	// IUnknown methods
	OffsetImageList_QueryInterface,
	OffsetImageList_AddRef,
	OffsetImageList_Release,

	// IImageList methods
	OffsetImageList_Add,
	OffsetImageList_ReplaceIcon,
	OffsetImageList_SetOverlayImage,
	OffsetImageList_Replace,
	OffsetImageList_AddMasked,
	OffsetImageList_Draw,
	OffsetImageList_Remove,
	OffsetImageList_GetIcon,
	OffsetImageList_GetImageInfo,
	OffsetImageList_Copy,
	OffsetImageList_Merge,
	OffsetImageList_Clone,
	OffsetImageList_GetImageRect,
	OffsetImageList_GetIconSize,
	OffsetImageList_SetIconSize,
	OffsetImageList_GetImageCount,
	OffsetImageList_SetImageCount,
	OffsetImageList_SetBkColor,
	OffsetImageList_GetBkColor,
	OffsetImageList_BeginDrag,
	OffsetImageList_EndDrag,
	OffsetImageList_DragEnter,
	OffsetImageList_DragLeave,
	OffsetImageList_DragMove,
	OffsetImageList_SetDragCursorImage,
	OffsetImageList_DragShowNolock,
	OffsetImageList_GetDragImage,
	OffsetImageList_GetItemFlags,
	OffsetImageList_GetOverlayImage,

	// IImageList2 methods
	OffsetImageList_Resize,
	OffsetImageList_GetOriginalSize,
	OffsetImageList_SetOriginalSize,
	OffsetImageList_SetCallback,
	OffsetImageList_GetCallback,
	OffsetImageList_ForceImagePresent,
	OffsetImageList_DiscardImages,
	OffsetImageList_PreloadImages,
	OffsetImageList_GetStatistics,
	OffsetImageList_Initialize,
	OffsetImageList_Replace2,
	OffsetImageList_ReplaceFromImageList,
};

/// \brief <em>The vtable of our implementation of the \c IImageListPrivate interface</em>
///
/// \sa IImageListPrivateVtbl, IImageListPrivate
static const IImageListPrivateVtbl IImageListPrivateImpl_Vtbl =
{
	// IUnknown methods
	OffsetImageList_IImageListPrivate_QueryInterface,
	OffsetImageList_IImageListPrivate_AddRef,
	OffsetImageList_IImageListPrivate_Release,

	// IImageListPrivate methods
	OffsetImageList_GetImageList,
	OffsetImageList_SetImageList,
	OffsetImageList_GetIndexMapping,
	OffsetImageList_SetIndexMapping,
};

/// \brief <em>Holds any vtables and instance specific data of the \c OffsetImageList pseudo class</em>
///
/// \sa OffsetImageList_CreateInstance
typedef struct OffsetImageList
{
	/// \brief <em>Holds the magic value \c 0x4C4D4948 ('HIML') used by comctl32.dll to identify the object as an image list</em>
	///
	/// \remarks This magic value must stand right before the \c IImageList vtable.
	DWORD magicValue;
	/// \brief <em>Holds the \c IImageList2 vtable</em>
	///
	/// \sa IImageList2Impl_Vtbl
	const IImageList2Vtbl* pImageList2Vtable;
	/// \brief <em>Holds the \c IImageListPrivate vtable</em>
	///
	/// \sa IImageListPrivateImpl_Vtbl
	const IImageListPrivateVtbl* pImageListPrivateVtable;
	/// \brief <em>Holds the object's reference count</em>
	LONG referenceCount;
	/// \brief <em>Holds the icon index mappings</em>
	int iconIndexMappings[6];

	/// \brief <em>Holds the wrapped image list</em>
	IImageList* pWrappedImageList;
} OffsetImageList;


/// \brief <em>Casts an \c OffsetImageList pointer to an \c IUnknown pointer</em>
///
/// \sa OffsetImageList, IImageListFromOffsetImageList, IImageList2FromOffsetImageList,
///     IImageListPrivateFromOffsetImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline IUnknown* IUnknownFromOffsetImageList(OffsetImageList* pThis)
{
	return reinterpret_cast<IUnknown*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c OffsetImageList pointer to an \c IImageList pointer</em>
///
/// \sa OffsetImageList, IUnknownFromOffsetImageList, IImageList2FromOffsetImageList,
///     IImageListPrivateFromOffsetImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline IImageList* IImageListFromOffsetImageList(OffsetImageList* pThis)
{
	return reinterpret_cast<IImageList*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c OffsetImageList pointer to an \c IImageList2 pointer</em>
///
/// \sa OffsetImageList, IUnknownFromOffsetImageList, IImageListFromOffsetImageList,
///     IImageListPrivateFromOffsetImageList,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline IImageList2* IImageList2FromOffsetImageList(OffsetImageList* pThis)
{
	return reinterpret_cast<IImageList2*>(&pThis->pImageList2Vtable);
}

/// \brief <em>Casts an \c OffsetImageList pointer to an \c IImageListPrivate pointer</em>
///
/// \sa OffsetImageList, IImageListPrivate, IUnknownFromOffsetImageList,
///     IImageListFromOffsetImageList, IImageList2FromOffsetImageList
static __inline IImageListPrivate* IImageListPrivateFromOffsetImageList(OffsetImageList* pThis)
{
	return reinterpret_cast<IImageListPrivate*>(&pThis->pImageListPrivateVtable);
}

/// \brief <em>Casts an \c IUnknown pointer to an \c OffsetImageList pointer</em>
///
/// \sa OffsetImageList, OffsetImageListFromIImageList, OffsetImageListFromIImageList2,
///     OffsetImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
static __inline OffsetImageList* OffsetImageListFromIUnknown(IUnknown* pInterface)
{
	return reinterpret_cast<OffsetImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(OffsetImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList pointer to an \c OffsetImageList pointer</em>
///
/// \sa OffsetImageList, OffsetImageListFromIUnknown, OffsetImageListFromIImageList2,
///     OffsetImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761490.aspx">IImageList</a>
static __inline OffsetImageList* OffsetImageListFromIImageList(IImageList* pInterface)
{
	return reinterpret_cast<OffsetImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(OffsetImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageList2 pointer to an \c OffsetImageList pointer</em>
///
/// \sa OffsetImageList, OffsetImageListFromIUnknown, OffsetImageListFromIImageList,
///     OffsetImageListFromIImageListPrivate,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761419.aspx">IImageList2</a>
static __inline OffsetImageList* OffsetImageListFromIImageList2(IImageList2* pInterface)
{
	return reinterpret_cast<OffsetImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(OffsetImageList, pImageList2Vtable));
}

/// \brief <em>Casts an \c IImageListPrivate pointer to an \c OffsetImageList pointer</em>
///
/// \sa OffsetImageList, OffsetImageListFromIUnknown, OffsetImageListFromIImageList,
///     OffsetImageListFromIImageList2, IImageListPrivate
static __inline OffsetImageList* OffsetImageListFromIImageListPrivate(IImageListPrivate* pInterface)
{
	return reinterpret_cast<OffsetImageList*>(reinterpret_cast<ULONG_PTR>(pInterface) - FIELD_OFFSET(OffsetImageList, pImageListPrivateVtable));
}


/// \brief <em>Wraps the \c HIMAGELIST_QueryInterface function and emulates it if it is not available</em>
///
/// \param[in] hImageList The \c himl parameter of the wrapped function.
/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
/// \param[in] ppInterfaceImpl The \c ppv parameter of the wrapped function.
///
/// \return An \c HRESULT error code.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761510.aspx">HIMAGELIST_QueryInterface</a>
static HRESULT WrappedHIMAGELIST_QueryInterface(HIMAGELIST hImageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl);