// LegacyImageListWrapper.cpp: An implementation of the IImageList interface wrapping an image list for version 5 of comctl32.dll

#include "stdafx.h"
#include "LegacyImageListWrapper.h"
#include "../ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE LegacyImageListWrapper_AddRef(IImageList* pInterface)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	return InterlockedIncrement(&pThis->referenceCount);
}

ULONG STDMETHODCALLTYPE LegacyImageListWrapper_Release(IImageList* pInterface)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	ULONG ret = InterlockedDecrement(&pThis->referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		ImageList_Destroy(pThis->hImageList);
		HeapFree(GetProcessHeap(), 0, pThis);
	}
	return ret;
}

STDMETHODIMP LegacyImageListWrapper_QueryInterface(IImageList* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = reinterpret_cast<LPUNKNOWN>(&pThis->pImageListVtable);
	} else if(requiredInterface == IID_IImageList) {
		*ppInterfaceImpl = IImageListFromLegacyImageListWrapper(pThis);
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	LegacyImageListWrapper_AddRef(pInterface);
	return S_OK;
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList
STDMETHODIMP LegacyImageListWrapper_Add(IImageList* pInterface, HBITMAP hbmImage, HBITMAP hbmMask, int* pi)
{
	if(!pi) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pi = ImageList_Add(pThis->hImageList, hbmImage, hbmMask);
	return (*pi != -1 ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_ReplaceIcon(IImageList* pInterface, int i, HICON hicon, int* pi)
{
	if(!pi) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pi = ImageList_ReplaceIcon(pThis->hImageList, i, hicon);
	return (*pi != -1 ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_SetOverlayImage(IImageList* pInterface, int iImage, int iOverlay)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_SetOverlayImage(pThis->hImageList, iImage, iOverlay);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_Replace(IImageList* pInterface, int i, HBITMAP hbmImage, HBITMAP hbmMask)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_Replace(pThis->hImageList, i, hbmImage, hbmMask);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_AddMasked(IImageList* pInterface, HBITMAP hbmImage, COLORREF crMask, int* pi)
{
	if(!pi) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pi = ImageList_AddMasked(pThis->hImageList, hbmImage, crMask);
	return (*pi != -1 ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_Draw(IImageList* pInterface, IMAGELISTDRAWPARAMS* pimldp)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	IMAGELISTDRAWPARAMS params = *pimldp;
	params.himl = pThis->hImageList;
	BOOL b = ImageList_DrawIndirect(&params);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_Remove(IImageList* pInterface, int i)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_Remove(pThis->hImageList, i);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_GetIcon(IImageList* pInterface, int i, UINT flags, HICON* picon)
{
	if(!picon) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*picon = ImageList_GetIcon(pThis->hImageList, i, flags);
	return (*picon ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_GetImageInfo(IImageList* pInterface, int i, IMAGEINFO* pImageInfo)
{
	if(!pImageInfo) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_GetImageInfo(pThis->hImageList, i, pImageInfo);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_Copy(IImageList* pInterface, int iDst, IUnknown* /*punkSrc*/, int iSrc, UINT uFlags)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_Copy(pThis->hImageList, iDst, pThis->hImageList, iSrc, uFlags);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_Merge(IImageList* pInterface, int i1, IUnknown* punk2, int i2, int dx, int dy, REFIID riid, void** ppv)
{
	if(!ppv) {
		return E_POINTER;
	}
	if(riid != IID_IImageList) {
		return E_NOINTERFACE;
	}
	CComQIPtr<IImageList> pImgLstOfSecondImage = punk2;
	if(!pImgLstOfSecondImage) {
		return E_INVALIDARG;
	}
	HIMAGELIST hImgLstOfSecondImage = IImageListToHIMAGELIST(pImgLstOfSecondImage);

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	HIMAGELIST hImgLstResult = ImageList_Merge(pThis->hImageList, i1, hImgLstOfSecondImage, i2, dx, dy);
	if(hImgLstResult) {
		return LegacyImageListWrapper_CreateInstance(hImgLstResult, reinterpret_cast<IImageList**>(ppv));
	}
	return E_FAIL;
}

STDMETHODIMP LegacyImageListWrapper_Clone(IImageList* pInterface, REFIID riid, void** ppv)
{
	if(!ppv) {
		return E_POINTER;
	}
	if(riid != IID_IImageList) {
		return E_NOINTERFACE;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	HIMAGELIST hImgLstClone = ImageList_Duplicate(pThis->hImageList);
	if(hImgLstClone) {
		return LegacyImageListWrapper_CreateInstance(hImgLstClone, reinterpret_cast<IImageList**>(ppv));
	}
	return E_FAIL;
}

STDMETHODIMP LegacyImageListWrapper_GetImageRect(IImageList* pInterface, int i, RECT* prc)
{
	if(!prc) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	int iconWidth = 0;
	int iconHeight = 0;
	if(ImageList_GetIconSize(pThis->hImageList, &iconWidth, &iconHeight)) {
		prc->left = i * iconWidth;
		prc->top = 0;
		prc->right = prc->left + iconWidth;
		prc->bottom = prc->top + iconHeight;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP LegacyImageListWrapper_GetIconSize(IImageList* pInterface, int* cx, int* cy)
{
	if(!cx || !cy) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_GetIconSize(pThis->hImageList, cx, cy);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_SetIconSize(IImageList* pInterface, int cx, int cy)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_SetIconSize(pThis->hImageList, cx, cy);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_GetImageCount(IImageList* pInterface, int* pi)
{
	if(!pi) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pi = ImageList_GetImageCount(pThis->hImageList);
	return S_OK;
}

STDMETHODIMP LegacyImageListWrapper_SetImageCount(IImageList* pInterface, UINT uNewCount)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_SetImageCount(pThis->hImageList, uNewCount);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_SetBkColor(IImageList* pInterface, COLORREF clrBk, COLORREF* pclr)
{
	if(!pclr) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pclr = ImageList_SetBkColor(pThis->hImageList, clrBk);
	return (*pclr != CLR_NONE ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_GetBkColor(IImageList* pInterface, COLORREF* pclr)
{
	if(!pclr) {
		return E_POINTER;
	}

	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	*pclr = ImageList_GetBkColor(pThis->hImageList);
	return S_OK;
}

STDMETHODIMP LegacyImageListWrapper_BeginDrag(IImageList* pInterface, int iTrack, int dxHotspot, int dyHotspot)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_BeginDrag(pThis->hImageList, iTrack, dxHotspot, dyHotspot);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_EndDrag(IImageList* /*pInterface*/)
{
	ImageList_EndDrag();
	return S_OK;
}

STDMETHODIMP LegacyImageListWrapper_DragEnter(IImageList* /*pInterface*/, HWND hwndLock, int x, int y)
{
	BOOL b = ImageList_DragEnter(hwndLock, x, y);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_DragLeave(IImageList* /*pInterface*/, HWND hwndLock)
{
	BOOL b = ImageList_DragLeave(hwndLock);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_DragMove(IImageList* /*pInterface*/, int x, int y)
{
	BOOL b = ImageList_DragMove(x, y);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_SetDragCursorImage(IImageList* pInterface, IUnknown* /*punk*/, int iDrag, int dxHotspot, int dyHotspot)
{
	LegacyImageListWrapper* pThis = LegacyImageListWrapperFromIImageList(pInterface);
	BOOL b = ImageList_SetDragCursorImage(pThis->hImageList, iDrag, dxHotspot, dyHotspot);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_DragShowNolock(IImageList* /*pInterface*/, BOOL fShow)
{
	BOOL b = ImageList_DragShowNolock(fShow);
	return (b ? S_OK : E_FAIL);
}

STDMETHODIMP LegacyImageListWrapper_GetDragImage(IImageList* /*pInterface*/, POINT* ppt, POINT* pptHotspot, REFIID riid, void** ppv)
{
	if(!ppv) {
		return E_POINTER;
	}
	if(riid != IID_IImageList) {
		return E_NOINTERFACE;
	}

	HIMAGELIST hImgLstResult = ImageList_GetDragImage(ppt, pptHotspot);
	if(hImgLstResult) {
		return LegacyImageListWrapper_CreateInstance(hImgLstResult, reinterpret_cast<IImageList**>(ppv));
	}
	return E_FAIL;
}

STDMETHODIMP LegacyImageListWrapper_GetItemFlags(IImageList* /*pInterface*/, int /*i*/, DWORD* /*dwFlags*/)
{
	ATLASSERT(FALSE && "LegacyImageListWrapper_GetItemFlags() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP LegacyImageListWrapper_GetOverlayImage(IImageList* /*pInterface*/, int /*iOverlay*/, int* /*piIndex*/)
{
	ATLASSERT(FALSE && "LegacyImageListWrapper_GetOverlayImage() should never be called!");
	return E_NOTIMPL;
}
// implementation of IImageList
//////////////////////////////////////////////////////////////////////


STDMETHODIMP LegacyImageListWrapper_CreateInstance(HIMAGELIST hImageListToWrap, IImageList** ppObject)
{
	if(!ppObject) {
		return E_POINTER;
	}
	if(!hImageListToWrap) {
		return E_INVALIDARG;
	}

	LegacyImageListWrapper* pInstance = static_cast<LegacyImageListWrapper*>(HeapAlloc(GetProcessHeap(), 0, sizeof(LegacyImageListWrapper)));
	if(!pInstance) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pInstance, sizeof(*pInstance));
	pInstance->magicValue = 0x4C4D4948;
	pInstance->pImageListVtable = &IImageListImpl_Vtbl;
	pInstance->referenceCount = 1;
	pInstance->hImageList = hImageListToWrap;

	*ppObject = IImageListFromLegacyImageListWrapper(pInstance);
	return S_OK;
}