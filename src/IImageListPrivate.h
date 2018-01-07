//////////////////////////////////////////////////////////////////////
/// \class IImageListPrivate
/// \author Timo "TimoSoft" Kunze
/// \brief <em>This interface is implemented by \c OffsetImageList to control things like which image list is wrapped</em>
///
/// This interface is implemented by the \c OffsetImageList class. It allows the controls to set the image
/// list that is wrapped and to set the icon index mappings.
///
/// \sa OffsetImageList_CreateInstance
//////////////////////////////////////////////////////////////////////


#pragma once

// the interface's GUID {79465103-AE71-4c21-A90D-0488C740B330}
const IID IID_IImageListPrivate = {0x79465103, 0xAE71, 0x4c21, {0xA9, 0x0D, 0x04, 0x88, 0xC7, 0x40, 0xB3, 0x30}};


/// \brief <em>Sets the wrapped image list</em>
///
/// Sets the wrapped image list.
///
/// \sa IImageListPrivate::SetImageList
#define OIL_NORMAL										0

class IImageListPrivate :
    public IUnknown
{
public:
	/// \brief <em>Retrieves the specified wrapped image list</em>
	///
	/// This method is used to retrieve the wrapped image list.
	///
	/// \param[in] imageListType The image list to retrieve. Any of the \c OIL_* values is valid.
	/// \param[out] pHImageList Receives the wrapped image list.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetImageList
	virtual HRESULT STDMETHODCALLTYPE GetImageList(UINT imageListType, __out HIMAGELIST* pHImageList) = 0;
	/// \brief <em>Sets the specified wrapped image list</em>
	///
	/// This method is used to set the image list that shall be wrapped.
	///
	/// \param[in] imageListType The image list to set. Any of the \c OIL_* values is valid.
	/// \param[in] hImageList The new image list.
	/// \param[out] pHPreviousImageList Receives the old image list. May be \c NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetImageList
	virtual HRESULT STDMETHODCALLTYPE SetImageList(UINT imageListType, __in HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList) = 0;
	/// \brief <em>Retrieves the real icon index for an icon</em>
	///
	/// An OffsetImageList wraps an image list and maps the icon indexes, that are used by the control using
	/// this image list, to the icon indexes that shall be used. This method retrieves the icon index that an
	/// icon index used by the controll will be mapped to.
	///
	/// \param[in] externalIconIndex The zero-based index of the icon as it is used by the control.
	/// \param[out] pRealIconIndex Receives the zero-based index of the icon in the wrapped image list.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetIndexMapping
	virtual HRESULT STDMETHODCALLTYPE GetIndexMapping(int externalIconIndex, __out PINT pRealIconIndex) = 0;
	/// \brief <em>Sets the real icon index for an icon</em>
	///
	/// An OffsetImageList wraps an image list and maps the icon indexes, that are used by the control using
	/// this image list, to the icon indexes that shall be used. This method sets the icon index that an icon
	/// index used by the controll will be mapped to.
	///
	/// \param[in] externalIconIndex The zero-based index of the icon as it is used by the control.
	/// \param[in] realIconIndex The zero-based index of the icon in the wrapped image list.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIndexMapping
	virtual HRESULT STDMETHODCALLTYPE SetIndexMapping(int externalIconIndex, int realIconIndex) = 0;
};


typedef struct IImageListPrivateVtbl
{
	HRESULT(STDMETHODCALLTYPE* QueryInterface)(IImageListPrivate* pThis, REFIID requiredInterface, __RPC__deref_out void** ppInterfaceImpl);
	ULONG(STDMETHODCALLTYPE* AddRef)(IImageListPrivate* pThis);
	ULONG(STDMETHODCALLTYPE* Release)(IImageListPrivate* pThis);
	HRESULT(STDMETHODCALLTYPE* GetImageList)(IImageListPrivate* pThis, UINT imageListType, __out HIMAGELIST* pHImageList);
	HRESULT(STDMETHODCALLTYPE* SetImageList)(IImageListPrivate* pThis, UINT imageListType, __in HIMAGELIST hImageList, __out_opt HIMAGELIST* pHPreviousImageList);
	HRESULT(STDMETHODCALLTYPE* GetIndexMapping)(IImageListPrivate* pThis, int externalIconIndex, __out_opt PINT pRealIconIndex);
	HRESULT(STDMETHODCALLTYPE* SetIndexMapping)(IImageListPrivate* pThis, int externalIconIndex, int realIconIndex);
} IImageListPrivateVtbl;