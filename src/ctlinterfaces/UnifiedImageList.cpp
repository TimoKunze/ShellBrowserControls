// UnifiedImageList.cpp: An implementation of the IImageList interface wrapping multiple image lists

#include "stdafx.h"
#include "UnifiedImageList.h"
#include "../ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE UnifiedImageList_AddRef(IImageList2* pInterface)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	return InterlockedIncrement(&pThis->referenceCount);
}

ULONG STDMETHODCALLTYPE UnifiedImageList_IImageListPrivate_AddRef(IImageListPrivate* pInterface)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromUnifiedImageList(pThis);
	return UnifiedImageList_AddRef(pBaseInterface);
}

ULONG STDMETHODCALLTYPE UnifiedImageList_Release(IImageList2* pInterface)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(pThis->magicValue != 0x4C4D4948) {
		return 0;
	}

	ULONG ret = InterlockedDecrement(&pThis->referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		UnifiedImageList_RemoveIconInfo(IImageListPrivateFromUnifiedImageList(pThis), 0xFFFFFFFF);
		delete pThis->pIconInfos;
		if(pThis->pWICImagingFactory) {
			pThis->pWICImagingFactory->Release();
		}
		pThis->magicValue = 0x00000000;
		HeapFree(GetProcessHeap(), 0, pThis);
	}
	return ret;
}

ULONG STDMETHODCALLTYPE UnifiedImageList_IImageListPrivate_Release(IImageListPrivate* pInterface)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromUnifiedImageList(pThis);
	return UnifiedImageList_Release(pBaseInterface);
}

STDMETHODIMP UnifiedImageList_QueryInterface(IImageList2* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = reinterpret_cast<LPUNKNOWN>(&pThis->pImageList2Vtable);
	} else if(requiredInterface == IID_IImageList) {
		*ppInterfaceImpl = IImageListFromUnifiedImageList(pThis);
	} else if(requiredInterface == IID_IImageList2) {
		*ppInterfaceImpl = IImageList2FromUnifiedImageList(pThis);
	} else if(requiredInterface == IID_IImageListPrivate) {
		*ppInterfaceImpl = IImageListPrivateFromUnifiedImageList(pThis);
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	UnifiedImageList_AddRef(pInterface);
	return S_OK;
}

STDMETHODIMP UnifiedImageList_IImageListPrivate_QueryInterface(IImageListPrivate* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromUnifiedImageList(pThis);
	return UnifiedImageList_QueryInterface(pBaseInterface, requiredInterface, ppInterfaceImpl);
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList
STDMETHODIMP UnifiedImageList_Add(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Add() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_ReplaceIcon(IImageList2* /*pInterface*/, int /*i*/, HICON /*hicon*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_ReplaceIcon() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_SetOverlayImage(IImageList2* /*pInterface*/, int /*iImage*/, int /*iOverlay*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetOverlayImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Replace(IImageList2* /*pInterface*/, int /*i*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Replace() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_AddMasked(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, COLORREF /*crMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_AddMasked() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Draw(IImageList2* pInterface, IMAGELISTDRAWPARAMS* pimldp)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(pimldp->i);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(pimldp->i);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	HIMAGELIST hImageListToUse = (pIconInfo->flags & AII_NONSHELLITEM ? pThis->wrappedImageLists.hNonShellItems : pThis->wrappedImageLists.hShellItems);
	int iconToUse = -1;
	switch(pIconInfo->flags & (AII_USESELECTEDICON | AII_USEEXPANDEDICON)) {
		case AII_USESELECTEDICON:
			iconToUse = pIconInfo->selectedIconIndex;
			break;
		case AII_USEEXPANDEDICON:
			iconToUse = pIconInfo->expandedIconIndex;
			break;
		default:
			iconToUse = pIconInfo->iconIndex;
			break;
	}
	if(pimldp->cx == 0 && pimldp->cy == 0) {
		// center icon
		int iconWidth = 0;
		int iconHeight = 0;
		ImageList_GetIconSize(hImageListToUse, &iconWidth, &iconHeight);
		UnifiedImageList_GetIconSize(pInterface, &pimldp->cx, &pimldp->cy);
		pimldp->x += (pimldp->cx - iconWidth) >> 1;
		pimldp->y += (pimldp->cy - iconHeight) >> 1;
	}
	if(!pThis->pWICImagingFactory || (pIconInfo->flags & AII_USELEGACYDISPLAYCODE)) {
		return UnifiedImageList_Draw_Legacy(pThis, pIconInfo, pimldp);
	}

	HRESULT hr = E_FAIL;
	// TODO: Handling of pimldp->cy == 0 is probably wrong!
	int desiredImageHeight = (pimldp->cy > 0 ? pimldp->cy : pimldp->cx);
	BOOL ghosted = ((pimldp->fStyle & ILD_BLEND) == ILD_BLEND);
	if(pimldp->rgbFg == CLR_DEFAULT) {
		/* TODO: Properly support standard theme. This won't be possible as long as we draw and blend each
		 *       component for its own. The big problem is, that IImageList::GetIcon doesn't handle icons
		 *       that are too small (and therefore get framed) well.
		 */
		ghosted = FALSE;
	}

	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));
	if(iconToUse >= 0 && hImageListToUse) {
		IMAGELISTDRAWPARAMS params = *pimldp;
		params.himl = hImageListToUse;
		params.i = iconToUse;
		ImageList_GetIconSize(hImageListToUse, &params.cx, &params.cy);
		params.cx = min(params.cx, pimldp->cx);
		params.cy = min(params.cy, desiredImageHeight);
		params.x += (pimldp->cx - params.cx) >> 1;
		params.y += (desiredImageHeight - params.cy) >> 1;
		if(ghosted) {
			params.fState |= ILS_ALPHA;
			params.Frame = 128;
		}
		params.fStyle &= ~(ILD_BLEND | ILD_SCALE);					// NOTE: If we activate ILD_SCALE, overlays are drawn too large.
		if(drawUACStamp) {
			if(ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags, FALSE)) {
				hr = S_OK;
			}
		} else if(ghosted && UnifiedImageList_IsComctl32Version600()) {
			if(ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags, FALSE)) {
				hr = S_OK;
			}
		} else if(ImageList_DrawIndirect(&params)) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP UnifiedImageList_Draw_Legacy(UnifiedImageList* pThis, UnifiedIconInformation* pIconInfo, IMAGELISTDRAWPARAMS* pimldp)
{
	HRESULT hr = E_FAIL;

	HIMAGELIST hImageListToUse = (pIconInfo->flags & AII_NONSHELLITEM ? pThis->wrappedImageLists.hNonShellItems : pThis->wrappedImageLists.hShellItems);
	int iconToUse = -1;
	switch(pIconInfo->flags & (AII_USESELECTEDICON | AII_USEEXPANDEDICON)) {
		case AII_USESELECTEDICON:
			iconToUse = pIconInfo->selectedIconIndex;
			break;
		case AII_USEEXPANDEDICON:
			iconToUse = pIconInfo->expandedIconIndex;
			break;
		default:
			iconToUse = pIconInfo->iconIndex;
			break;
	}
	// TODO: Handling of pimldp->cy == 0 is probably wrong!
	int desiredImageHeight = (pimldp->cy > 0 ? pimldp->cy : pimldp->cx);

	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));

	HBITMAP hMainBitmap = NULL;
	HBITMAP hMaskBitmap = NULL;
	BOOL doAlphaChannelPostProcessing = FALSE;
	if(iconToUse >= 0 && hImageListToUse) {
		int unifiedIconWidth = 0;
		int unifiedIconHeight = 0;
		UnifiedImageList_GetIconSize(IImageList2FromUnifiedImageList(pThis), &unifiedIconWidth, &unifiedIconHeight);
		SIZE imageSize = {0};
		ImageList_GetIconSize(hImageListToUse, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
		imageSize.cx = min(imageSize.cx, unifiedIconWidth);
		imageSize.cy = min(imageSize.cy, unifiedIconHeight);

		if(RunTimeHelper::IsVista()) {
			HICON hIcon = ImageList_GetIcon(hImageListToUse, iconToUse, /*ILD_SCALE | */ILD_TRANSPARENT | (pimldp->fStyle & ILD_OVERLAYMASK));					// NOTE: If we activate ILD_SCALE, overlays are drawn too large.
			ATLASSERT(hIcon);
			if(hIcon) {
				hMainBitmap = StampIconForElevation(hIcon, &imageSize, drawUACStamp);
				DestroyIcon(hIcon);
				drawUACStamp = FALSE;
			}
		} else {
			// On Windows XP with comctl32.dll 5.x, ImageList_GetIcon produces bad looking icons.
			hMaskBitmap = CreateBitmap(imageSize.cx, imageSize.cy, 1, 1, NULL);
			BITMAPINFO bitmapInfo = {0};
			bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			bitmapInfo.bmiHeader.biWidth = imageSize.cx;
			bitmapInfo.bmiHeader.biHeight = imageSize.cy;
			bitmapInfo.bmiHeader.biPlanes = 1;
			bitmapInfo.bmiHeader.biBitCount = 32;
			bitmapInfo.bmiHeader.biSizeImage = imageSize.cy * imageSize.cx * 4;
			LPRGBQUAD pBits;
			hMainBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
			if(pBits) {
				ZeroMemory(pBits, bitmapInfo.bmiHeader.biSizeImage);
				WTL::CDC dc;
				dc.CreateCompatibleDC();
				HBITMAP hOldBmp = dc.SelectBitmap(hMainBitmap);
				ImageList_Draw(hImageListToUse, iconToUse, dc, 0, 0, ILD_TRANSPARENT);
				dc.SelectBitmap(hMaskBitmap);
				ImageList_Draw(hImageListToUse, iconToUse, dc, 0, 0, ILD_MASK);
				dc.SelectBitmap(hOldBmp);

				IImageList* pImgLst = NULL;
				if(APIWrapper::IsSupported_HIMAGELIST_QueryInterface()) {
					APIWrapper::HIMAGELIST_QueryInterface(hImageListToUse, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst), NULL);
				} else {
					pImgLst = reinterpret_cast<IImageList*>(hImageListToUse);
					pImgLst->AddRef();
				}
				ATLASSUME(pImgLst);

				DWORD flags = 0;
				pImgLst->GetItemFlags(iconToUse, &flags);
				pImgLst->Release();
				doAlphaChannelPostProcessing = !(flags & ILIF_ALPHA);
			}
		}
	}
	if(!hMainBitmap) {
		return hr;
	}
	BITMAP bitmap = {0};
	GetObject(hMainBitmap, sizeof(bitmap), &bitmap);
	SIZE iconSize = {bitmap.bmWidth, bitmap.bmHeight};

	if(doAlphaChannelPostProcessing) {
		BITMAPINFO bitmapInfo = {0};
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = iconSize.cx;
		bitmapInfo.bmiHeader.biHeight = -iconSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
		bitmapInfo.bmiHeader.biSizeImage = iconSize.cx * iconSize.cy * 4;
		LPRGBQUAD pBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
		if(pBits) {
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(NULL);
			if(GetDIBits(memoryDC, hMainBitmap, 0, iconSize.cy, pBits, &bitmapInfo, DIB_RGB_COLORS)) {
				HBITMAP hOldBmp = memoryDC.SelectBitmap(hMaskBitmap);
				//#pragma omp parallel for -> produces artefacts, probably conflicts with GetPixel
				for(int i = 0; i < iconSize.cy * iconSize.cx; i++) {
					if(memoryDC.GetPixel(i % iconSize.cy, i / iconSize.cy) == 0x00000000) {
						pBits[i].rgbReserved = 0xFF;
					}
				}
				memoryDC.SelectBitmap(hOldBmp);
				SetDIBits(memoryDC, hMainBitmap, 0, iconSize.cy, pBits, &bitmapInfo, DIB_RGB_COLORS);
			}
			HeapFree(GetProcessHeap(), 0, pBits);
		}
	}
	if(hMaskBitmap) {
		DeleteObject(hMaskBitmap);
	}

	IMAGELISTDRAWPARAMS params = *pimldp;
	params.cx = iconSize.cx;
	params.cy = iconSize.cy;
	params.x += (pimldp->cx - iconSize.cx) >> 1;
	// center vertically
	params.y += (desiredImageHeight - iconSize.cy) >> 1;
	// draw using AlphaBlend
	if(pThis->flags & ILOF_IGNOREEXTRAALPHA) {
		params.fState &= ~ILS_ALPHA;
	}
	params.Frame = 255 - (params.fState & ILS_ALPHA ? params.Frame : 0);

	WTL::CDC dc;
	dc.CreateCompatibleDC();
	HBITMAP hOldBmp = NULL;
	HBITMAP hBlendedBitmap = NULL;
	BOOL blend = (params.fStyle & ILD_BLEND50) == ILD_BLEND50 || (params.fStyle & ILD_BLEND25) == ILD_BLEND25;
	if(blend && params.rgbFg != CLR_NONE) {
		COLORREF blendColor = (params.rgbFg == CLR_DEFAULT ? GetSysColor(COLOR_HIGHLIGHT) : params.rgbFg);
		hBlendedBitmap = BlendBitmap(hMainBitmap, blendColor, (params.fStyle & ILD_BLEND50) == ILD_BLEND50 ? 50 : 25, TRUE);
		hOldBmp = dc.SelectBitmap(hBlendedBitmap);
	} else {
		PremultiplyBitmap(hMainBitmap);
		hOldBmp = dc.SelectBitmap(hMainBitmap);
	}
	BLENDFUNCTION blendFunction = {0};
	blendFunction.BlendOp = AC_SRC_OVER;
	blendFunction.SourceConstantAlpha = static_cast<BYTE>(params.Frame);
	blendFunction.AlphaFormat = AC_SRC_ALPHA;
	if(AlphaBlend(params.hdcDst, params.x, params.y, params.cx, params.cy, dc, 0, 0, params.cx, params.cy, blendFunction)) {
		hr = S_OK;
	}
	dc.SelectBitmap(hOldBmp);
	if(hBlendedBitmap) {
		DeleteObject(hBlendedBitmap);
	}
	DeleteObject(hMainBitmap);
	return hr;
}

STDMETHODIMP UnifiedImageList_Remove(IImageList2* /*pInterface*/, int /*i*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Remove() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetIcon(IImageList2* pInterface, int i, UINT flags, HICON* picon)
{
	/* TODO: This method is pretty much broken. Also it should be split into 2 methods, one for legacy
	 *       drawing mode and one for modern mode - just like the Draw method.
	 */

	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	// TODO: support ILD_BLEND
	//BOOL ghosted = ((flags & ILD_BLEND) == ILD_BLEND);
	flags &= ~(ILD_BLEND | ILD_SCALE);

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(i);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(i);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	HIMAGELIST hImageListToUse = (pIconInfo->flags & AII_NONSHELLITEM ? pThis->wrappedImageLists.hNonShellItems : pThis->wrappedImageLists.hShellItems);
	int iconToUse = -1;
	switch(pIconInfo->flags & (AII_USESELECTEDICON | AII_USEEXPANDEDICON)) {
		case AII_USESELECTEDICON:
			iconToUse = pIconInfo->selectedIconIndex;
			break;
		case AII_USEEXPANDEDICON:
			iconToUse = pIconInfo->expandedIconIndex;
			break;
		default:
			iconToUse = pIconInfo->iconIndex;
			break;
	}

	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));
	HBITMAP hMainBitmap = NULL;
	if(iconToUse >= 0 && hImageListToUse) {
		int unifiedIconWidth = 0;
		int unifiedIconHeight = 0;
		UnifiedImageList_GetIconSize(pInterface, &unifiedIconWidth, &unifiedIconHeight);
		SIZE imageSize = {0};
		ImageList_GetIconSize(hImageListToUse, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
		imageSize.cx = min(imageSize.cx, unifiedIconWidth);
		imageSize.cy = min(imageSize.cy, unifiedIconHeight);

		HICON hIcon = ImageList_GetIcon(hImageListToUse, iconToUse, /*ILD_SCALE | */ILD_TRANSPARENT | (flags & ILD_OVERLAYMASK));					// NOTE: If we activate ILD_SCALE, overlays are drawn too large.
		ATLASSERT(hIcon);
		if(hIcon) {
			hMainBitmap = StampIconForElevation(hIcon, &imageSize, drawUACStamp);
			DestroyIcon(hIcon);
		}
	}
	if(!hMainBitmap) {
		return hr;
	}

	BITMAP bmp = {0};
	GetObject(hMainBitmap, sizeof(bmp), &bmp);

	BOOL useFallback = TRUE;
	if(RunTimeHelper::IsCommCtrl6()) {
		CComPtr<IImageList2> pImgLst = NULL;
		ATLVERIFY(SUCCEEDED(APIWrapper::ImageList_CoCreateInstance(CLSID_ImageList, NULL, IID_PPV_ARGS(&pImgLst), NULL)));
		if(pImgLst) {
			ATLVERIFY(SUCCEEDED(pImgLst->Initialize(bmp.bmWidth, bmp.bmHeight, ILC_COLOR32 | ILC_HIGHQUALITYSCALE, 1, 0)));
			int iconIndex = -1;
			ATLVERIFY(SUCCEEDED(pImgLst->Add(hMainBitmap, NULL, &iconIndex)));
			if(iconIndex >= 0) {
				hr = pImgLst->GetIcon(iconIndex, flags, picon);
				useFallback = FALSE;
			}
		}
	}
	if(useFallback) {
		hr = E_FAIL;
		HIMAGELIST h = NULL;
		if(RunTimeHelper::IsCommCtrl6()) {
			h = ImageList_Create(bmp.bmWidth, bmp.bmHeight, ILC_COLOR32, 1, 0);
		} else {
			// TODO: Specify ILC_MASK?
			h = ImageList_Create(bmp.bmWidth, bmp.bmHeight, ILC_COLOR24, 1, 0);
		}
		if(h) {
			int iconIndex = ImageList_Add(h, hMainBitmap, NULL);
			if(iconIndex >= 0) {
				*picon = ImageList_GetIcon(h, iconIndex, flags);
				hr = S_OK;
			}
			ImageList_Destroy(h);
		}
	}
	DeleteObject(hMainBitmap);
	return hr;
}

STDMETHODIMP UnifiedImageList_GetImageInfo(IImageList2* /*pInterface*/, int /*i*/, IMAGEINFO* /*pImageInfo*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetImageInfo() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Copy(IImageList2* /*pInterface*/, int /*iDst*/, IUnknown* /*punkSrc*/, int /*iSrc*/, UINT /*uFlags*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Copy() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Merge(IImageList2* /*pInterface*/, int /*i1*/, IUnknown* /*punk2*/, int /*i2*/, int /*dx*/, int /*dy*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Merge() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Clone(IImageList2* /*pInterface*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Clone() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetImageRect(IImageList2* /*pInterface*/, int /*i*/, RECT* /*prc*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetImageRect() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetIconSize(IImageList2* pInterface, int* cx, int* cy)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	// return the larger size
	int cxShellItems = 0;
	int cyShellItems = 0;
	if(pThis->wrappedImageLists.hShellItems) {
		ImageList_GetIconSize(pThis->wrappedImageLists.hShellItems, &cxShellItems, &cyShellItems);
	}
	int cxNonShellItems = 0;
	int cyNonShellItems = 0;
	if(pThis->wrappedImageLists.hNonShellItems) {
		ImageList_GetIconSize(pThis->wrappedImageLists.hNonShellItems, &cxNonShellItems, &cyNonShellItems);
	}
	*cx = max(cxShellItems, cxNonShellItems);
	*cy = max(cyShellItems, cyNonShellItems);
	return S_OK;
}

STDMETHODIMP UnifiedImageList_SetIconSize(IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetIconSize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetImageCount(IImageList2* /*pInterface*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetImageCount() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_SetImageCount(IImageList2* /*pInterface*/, UINT /*uNewCount*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetImageCount() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_SetBkColor(IImageList2* /*pInterface*/, COLORREF /*clrBk*/, COLORREF* /*pclr*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetBkColor() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetBkColor(IImageList2* /*pInterface*/, COLORREF* pclr)
{
	*pclr = CLR_NONE;
	return S_OK;
}

STDMETHODIMP UnifiedImageList_BeginDrag(IImageList2* /*pInterface*/, int /*iTrack*/, int /*dxHotspot*/, int /*dyHotspot*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_BeginDrag() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_EndDrag(IImageList2* /*pInterface*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_EndDrag() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_DragEnter(IImageList2* /*pInterface*/, HWND /*hwndLock*/, int /*x*/, int /*y*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_DragEnter() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_DragLeave(IImageList2* /*pInterface*/, HWND /*hwndLock*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_DragLeave() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_DragMove(IImageList2* /*pInterface*/, int /*x*/, int /*y*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_DragMove() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_SetDragCursorImage(IImageList2* /*pInterface*/, IUnknown* /*punk*/, int /*iDrag*/, int /*dxHotspot*/, int /*dyHotspot*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetDragCursorImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_DragShowNolock(IImageList2* /*pInterface*/, BOOL /*fShow*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_DragShowNolock() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetDragImage(IImageList2* /*pInterface*/, POINT* /*ppt*/, POINT* /*pptHotspot*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetDragImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetItemFlags(IImageList2* /*pInterface*/, int /*i*/, DWORD* /*dwFlags*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetItemFlags() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetOverlayImage(IImageList2* /*pInterface*/, int /*iOverlay*/, int* /*piIndex*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetOverlayImage() should never be called!");
	return E_NOTIMPL;
}
// implementation of IImageList
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList2
STDMETHODIMP UnifiedImageList_Resize(IImageList2* /*pInterface*/, int /*cxNewIconSize*/, int /*cyNewIconSize*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Resize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetOriginalSize(IImageList2* pInterface, int iImage, DWORD dwFlags, int* pcx, int* pcy)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(iImage);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(iImage);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	CComPtr<IImageList2> pImgLst = NULL;
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hNonShellItems, IID_PPV_ARGS(&pImgLst), NULL);
		}
	} else {
		if(pThis->wrappedImageLists.hShellItems) {
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hShellItems, IID_PPV_ARGS(&pImgLst), NULL);
		}
	}
	if(pImgLst) {
		// TODO: AII_SELECTED?
		return pImgLst->GetOriginalSize(pIconInfo->iconIndex, dwFlags, pcx, pcy);
	}
	return E_FAIL;
}

STDMETHODIMP UnifiedImageList_SetOriginalSize(IImageList2* /*pInterface*/, int /*iImage*/, int /*cx*/, int /*cy*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetOriginalSize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_SetCallback(IImageList2* /*pInterface*/, IUnknown* /*punk*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_SetCallback() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetCallback(IImageList2* /*pInterface*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetCallback() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_ForceImagePresent(IImageList2* pInterface, int iImage, DWORD dwFlags)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(iImage);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(iImage);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		UnifiedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	CComPtr<IImageList2> pImgLst = NULL;
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hNonShellItems, IID_PPV_ARGS(&pImgLst), NULL);
		}
	} else {
		if(pThis->wrappedImageLists.hShellItems) {
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hShellItems, IID_PPV_ARGS(&pImgLst), NULL);
		}
	}
	if(pImgLst) {
		ATLVERIFY(SUCCEEDED(pImgLst->ForceImagePresent(pIconInfo->iconIndex, dwFlags)));
		if(pIconInfo->selectedIconIndex >= 0 && pIconInfo->selectedIconIndex != pIconInfo->iconIndex) {
			ATLVERIFY(SUCCEEDED(pImgLst->ForceImagePresent(pIconInfo->selectedIconIndex, dwFlags)));
		}
	}
	return S_OK;
}

STDMETHODIMP UnifiedImageList_DiscardImages(IImageList2* /*pInterface*/, int /*iFirstImage*/, int /*iLastImage*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_DiscardImages() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_PreloadImages(IImageList2* /*pInterface*/, IMAGELISTDRAWPARAMS* /*pimldp*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_PreloadImages() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetStatistics(IImageList2* /*pInterface*/, IMAGELISTSTATS* /*pils*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_GetStatistics() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Initialize(IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Initialize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_Replace2(IImageList2* /*pInterface*/, int /*i*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/, IUnknown* /*punk*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_Replace2() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_ReplaceFromImageList(IImageList2* /*pInterface*/, int /*i*/, IImageList* /*pil*/, int /*iSrc*/, IUnknown* /*punk*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "UnifiedImageList_ReplaceFromImageList() should never be called!");
	return E_NOTIMPL;
}
// implementation of IImageList2
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageListPrivate
STDMETHODIMP UnifiedImageList_SetImageList(IImageListPrivate* pInterface, UINT imageListType, HIMAGELIST hImageList, HIMAGELIST* pHPreviousImageList)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	switch(imageListType) {
		case AIL_SHELLITEMS:
			if(pHPreviousImageList) {
				*pHPreviousImageList = pThis->wrappedImageLists.hShellItems;
			}
			pThis->wrappedImageLists.hShellItems = hImageList;
			return S_OK;
			break;
		case AIL_NONSHELLITEMS:
			if(pHPreviousImageList) {
				*pHPreviousImageList = pThis->wrappedImageLists.hNonShellItems;
			}
			pThis->wrappedImageLists.hNonShellItems = hImageList;
			return S_OK;
			break;
	}
	return E_NOTIMPL;
}

STDMETHODIMP UnifiedImageList_GetIconInfo(IImageListPrivate* pInterface, UINT itemID, UINT mask, PINT pSystemOrNonShellIconIndex, PINT /*pExecutableIconIndex*/, LPWSTR* /*ppOverlayImageResource*/, PUINT pFlags)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
		if(iter != pThis->pIconInfos->cend()) {
			if(pSystemOrNonShellIconIndex) {
				if(iter->second->flags & AII_NONSHELLITEM) {
					if(mask & SIIF_NONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->iconIndex;
					} else if(mask & SIIF_SELECTEDNONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->selectedIconIndex;
					} else if(mask & SIIF_EXPANDEDNONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->expandedIconIndex;
					}
				} else {
					if(mask & SIIF_SYSTEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->iconIndex;
					} else if(mask & SIIF_SELECTEDSYSTEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->selectedIconIndex;
					} else if(mask & SIIF_EXPANDEDSYSTEMICON) {
						*pSystemOrNonShellIconIndex = iter->second->expandedIconIndex;
					}
				}
			}
			if(mask & SIIF_FLAGS && pFlags) {
				*pFlags = iter->second->flags;
			}
		} else {
			return E_INVALIDARG;
		}
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pIconInfo = pThis->pIconInfos->Lookup(itemID);
		if(pIconInfo) {
			if(pSystemOrNonShellIconIndex) {
				if(pIconInfo->m_value->flags & AII_NONSHELLITEM) {
					if(mask & SIIF_NONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->iconIndex;
					} else if(mask & SIIF_SELECTEDNONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->selectedIconIndex;
					} else if(mask & SIIF_EXPANDEDNONSHELLITEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->expandedIconIndex;
					}
				} else {
					if(mask & SIIF_SYSTEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->iconIndex;
					} else if(mask & SIIF_SELECTEDSYSTEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->selectedIconIndex;
					} else if(mask & SIIF_EXPANDEDSYSTEMICON) {
						*pSystemOrNonShellIconIndex = pIconInfo->m_value->expandedIconIndex;
					}
				}
			}
			if(mask & SIIF_FLAGS && pFlags) {
				*pFlags = pIconInfo->m_value->flags;
			}
		} else {
			return E_INVALIDARG;
		}
	#endif
	return S_OK;
}

STDMETHODIMP UnifiedImageList_SetIconInfo(IImageListPrivate* pInterface, UINT itemID, UINT mask, HBITMAP /*hImage*/, int systemOrNonShellIconIndex, int /*executableIconIndex*/, LPWSTR /*pOverlayImageResource*/, UINT flags)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
		if(iter == pThis->pIconInfos->cend()) {
	#else
		CAtlMap<UINT, UnifiedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(itemID);
		if(!pEntry) {
	#endif
		UnifiedIconInformation* pIconInfo = new UnifiedIconInformation;
		if(mask & SIIF_FLAGS) {
			pIconInfo->flags = flags & ~(AII_ADORNMENTMASK | AII_FILETYPEOVERLAYMASK | AII_HASTHUMBNAIL);
		}
		if(pIconInfo->flags & AII_NONSHELLITEM) {
			if(mask & SIIF_NONSHELLITEMICON) {
				pIconInfo->iconIndex = systemOrNonShellIconIndex;
			} else if(mask & SIIF_SELECTEDNONSHELLITEMICON) {
				pIconInfo->selectedIconIndex = systemOrNonShellIconIndex;
			} else if(mask & SIIF_EXPANDEDNONSHELLITEMICON) {
				pIconInfo->expandedIconIndex = systemOrNonShellIconIndex;
			}
		} else {
			if(mask & SIIF_SYSTEMICON) {
				pIconInfo->iconIndex = systemOrNonShellIconIndex;
			} else if(mask & SIIF_SELECTEDSYSTEMICON) {
				pIconInfo->selectedIconIndex = systemOrNonShellIconIndex;
			} else if(mask & SIIF_EXPANDEDSYSTEMICON) {
				pIconInfo->expandedIconIndex = systemOrNonShellIconIndex;
			}
		}
		(*(pThis->pIconInfos))[itemID] = pIconInfo;
	} else {
		#ifdef USE_STL
			if(mask & SIIF_FLAGS) {
				iter->second->flags = flags & ~(AII_ADORNMENTMASK | AII_FILETYPEOVERLAYMASK | AII_HASTHUMBNAIL);
			}
			if(iter->second->flags & AII_NONSHELLITEM) {
				if(mask & SIIF_NONSHELLITEMICON) {
					iter->second->iconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_SELECTEDNONSHELLITEMICON) {
					iter->second->selectedIconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_EXPANDEDNONSHELLITEMICON) {
					iter->second->expandedIconIndex = systemOrNonShellIconIndex;
				}
			} else {
				if(mask & SIIF_SYSTEMICON) {
					iter->second->iconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_SELECTEDSYSTEMICON) {
					iter->second->selectedIconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_EXPANDEDSYSTEMICON) {
					iter->second->expandedIconIndex = systemOrNonShellIconIndex;
				}
			}
		#else
			if(mask & SIIF_FLAGS) {
				pEntry->m_value->flags = flags & ~(AII_ADORNMENTMASK | AII_FILETYPEOVERLAYMASK | AII_HASTHUMBNAIL);
			}
			if(pEntry->m_value->flags & AII_NONSHELLITEM) {
				if(mask & SIIF_NONSHELLITEMICON) {
					pEntry->m_value->iconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_SELECTEDNONSHELLITEMICON) {
					pEntry->m_value->selectedIconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_EXPANDEDNONSHELLITEMICON) {
					pEntry->m_value->expandedIconIndex = systemOrNonShellIconIndex;
				}
			} else {
				if(mask & SIIF_SYSTEMICON) {
					pEntry->m_value->iconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_SELECTEDSYSTEMICON) {
					pEntry->m_value->selectedIconIndex = systemOrNonShellIconIndex;
				} else if(mask & SIIF_EXPANDEDSYSTEMICON) {
					pEntry->m_value->expandedIconIndex = systemOrNonShellIconIndex;
				}
			}
		#endif
	}
	return S_OK;
}

STDMETHODIMP UnifiedImageList_RemoveIconInfo(IImageListPrivate* pInterface, UINT itemID)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		if(itemID == 0xFFFFFFFF) {
			std::unordered_map<UINT, UnifiedIconInformation*>::iterator iter = pThis->pIconInfos->begin();
			while(iter != pThis->pIconInfos->end()) {
				UnifiedIconInformation* pEntry = iter->second;
				iter = pThis->pIconInfos->erase(iter);
				delete pEntry;
			}
		} else {
			std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
			if(iter != pThis->pIconInfos->cend()) {
				UnifiedIconInformation* pEntry = iter->second;
				pThis->pIconInfos->erase(itemID);
				delete pEntry;
			}
		}
	#else
		if(itemID == 0xFFFFFFFF) {
			POSITION p = pThis->pIconInfos->GetStartPosition();
			while(p) {
				UnifiedIconInformation* pEntry = pThis->pIconInfos->GetValueAt(p);
				pThis->pIconInfos->RemoveAtPos(p);
				p = pThis->pIconInfos->GetStartPosition();
				delete pEntry;
			}
		} else {
			CAtlMap<UINT, UnifiedIconInformation*>::CPair* pIconInfo = pThis->pIconInfos->Lookup(itemID);
			if(pIconInfo) {
				UnifiedIconInformation* pEntry = pIconInfo->m_value;
				pThis->pIconInfos->RemoveKey(itemID);
				delete pEntry;
			}
		}
	#endif
	return S_OK;
}

STDMETHODIMP UnifiedImageList_GetOptions(IImageListPrivate* pInterface, UINT mask, UINT* pFlags)
{
	if(!pFlags) {
		return E_POINTER;
	}

	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(mask) {
		*pFlags = pThis->flags & mask;
	} else {
		*pFlags = pThis->flags;
	}
	return S_OK;
}

STDMETHODIMP UnifiedImageList_SetOptions(IImageListPrivate* pInterface, UINT mask, UINT flags)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(mask) {
		flags |= (pThis->flags & ~mask);
	}
	pThis->flags = flags;
	return S_OK;
}

STDMETHODIMP UnifiedImageList_TransferNonShellItems(IImageListPrivate* pInterface, IImageListPrivate* pTarget)
{
	UnifiedImageList* pThis = UnifiedImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		for(std::unordered_map<UINT, UnifiedIconInformation*>::const_iterator iter = pThis->pIconInfos->cbegin(); iter != pThis->pIconInfos->cend(); iter++) {
			if(iter->second->flags & AII_NONSHELLITEM) {
				// TODO: The icon for selected state gets lost here.
				pTarget->SetIconInfo(iter->first, SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, iter->second->iconIndex, 0, NULL, iter->second->flags);
			}
		}
	#else
		POSITION p = pThis->pIconInfos->GetStartPosition();
		while(p) {
			UnifiedIconInformation* pIconInfo = pThis->pIconInfos->GetValueAt(p);
			if(pIconInfo->flags & AII_NONSHELLITEM) {
				// TODO: The icon for selected state gets lost here.
				pTarget->SetIconInfo(pThis->pIconInfos->GetKeyAt(p), SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, pIconInfo->iconIndex, 0, NULL, pIconInfo->flags);
			}
			pThis->pIconInfos->GetNext(p);
		}
	#endif
	return S_OK;
}
// implementation of IImageListPrivate
//////////////////////////////////////////////////////////////////////


STDMETHODIMP UnifiedImageList_CreateInstance(IImageList** ppObject)
{
	if(!ppObject) {
		return E_POINTER;
	}

	UnifiedImageList* pInstance = static_cast<UnifiedImageList*>(HeapAlloc(GetProcessHeap(), 0, sizeof(UnifiedImageList)));
	if(!pInstance) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pInstance, sizeof(*pInstance));
	#ifdef USE_STL
		pInstance->pIconInfos = new std::unordered_map<UINT, UnifiedIconInformation*>();
	#else
		pInstance->pIconInfos = new CAtlMap<UINT, UnifiedIconInformation*>();
	#endif
	pInstance->magicValue = 0x4C4D4948;
	pInstance->pImageList2Vtable = &UnifiedImageList_IImageList2Impl_Vtbl;
	pInstance->pImageListPrivateVtable = &UnifiedImageList_IImageListPrivateImpl_Vtbl;
	pInstance->referenceCount = 1;
	CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pInstance->pWICImagingFactory));

	*ppObject = IImageListFromUnifiedImageList(pInstance);
	return S_OK;
}

BOOL ImageList_DrawIndirect_HQScaling(UnifiedImageList* pInstance, IMAGELISTDRAWPARAMS* pParams, UINT flags, BOOL forceSize)
{
	BOOL b = FALSE;

	HICON hIconToScale = ImageList_GetIcon(pParams->himl, pParams->i, pParams->fStyle);
	ATLASSERT(hIconToScale);
	if((flags & AII_DISPLAYELEVATIONOVERLAY) && hIconToScale && APIWrapper::IsSupported_StampIconForElevation()) {
		HRESULT hr = E_FAIL;
		HICON hStampedIcon = NULL;
		APIWrapper::StampIconForElevation(hIconToScale, FALSE, &hStampedIcon, &hr);
		if(SUCCEEDED(hr)) {
			DestroyIcon(hIconToScale);
			hIconToScale = hStampedIcon;

			if(!forceSize) {
				SIZE iconSize = {0};
				if(GetIconSize(hIconToScale, &iconSize)) {
					if(iconSize.cx > pParams->cx || iconSize.cy > pParams->cy) {
						pParams->x -= iconSize.cx - pParams->cx;
						pParams->y -= iconSize.cy - pParams->cy;
						pParams->cx = iconSize.cx;
						pParams->cy = iconSize.cy;
					}
				}
			}
		}
	}
	if(hIconToScale) {
		if(pInstance->pWICImagingFactory) {
			CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
			HRESULT hr = pInstance->pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pWICBitmapScaler);
				CComPtr<IWICBitmap> pWICBitmap = NULL;
				hr = pInstance->pWICImagingFactory->CreateBitmapFromHICON(hIconToScale, &pWICBitmap);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					ATLASSUME(pWICBitmap);
					ATLVERIFY(SUCCEEDED(pWICBitmapScaler->Initialize(pWICBitmap, pParams->cx, pParams->cy, WICBitmapInterpolationModeFant)));

					// create target bitmap
					BITMAPINFO bitmapInfo = {0};
					bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
					bitmapInfo.bmiHeader.biWidth = pParams->cx;
					bitmapInfo.bmiHeader.biHeight = -pParams->cy;
					bitmapInfo.bmiHeader.biPlanes = 1;
					bitmapInfo.bmiHeader.biBitCount = 32;
					bitmapInfo.bmiHeader.biCompression = BI_RGB;
					LPRGBQUAD pBits = NULL;
					CBitmap scaledBitmap;
					scaledBitmap.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
					if(!scaledBitmap.IsNull()) {
						ATLASSERT_POINTER(pBits, RGBQUAD);
						if(pBits) {
							ATLVERIFY(SUCCEEDED(pWICBitmapScaler->CopyPixels(NULL, pParams->cx * sizeof(RGBQUAD), pParams->cy * pParams->cx * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits))));

							IMAGELISTDRAWPARAMS params = *pParams;
							//if((params.fState & ILS_ALPHA) && UnifiedImageList_IsComctl32Version600()) {
							if(params.fState & ILS_ALPHA) {
								// blend pBits by the value specified in params.Frame
								LPRGBQUAD p = pBits;
								for(int i = 0; i < pParams->cy * pParams->cx; i++, p++) {
									p->rgbReserved = static_cast<BYTE>((static_cast<DWORD>(p->rgbReserved) * params.Frame) / 0xFF);
								}
								params.fState &= ~ILS_ALPHA;
								params.Frame = 0;
							}

							if(flags & AII_USELEGACYDISPLAYCODE) {
								params.Frame = 255 - params.Frame;
								WTL::CDC dc;
								dc.CreateCompatibleDC();
								HBITMAP hOldBmp = NULL;
								HBITMAP hBlendedBitmap = NULL;
								BOOL blend = (params.fStyle & (ILD_BLEND50 | ILD_BLEND25)) != 0;
								if(blend && params.rgbFg != CLR_NONE) {
									COLORREF blendColor = (params.rgbFg == CLR_DEFAULT ? GetSysColor(COLOR_HIGHLIGHT) : params.rgbFg);
									hBlendedBitmap = BlendBitmap(scaledBitmap, blendColor, (params.fStyle & ILD_BLEND50) ? 50 : 25, TRUE);
									hOldBmp = dc.SelectBitmap(hBlendedBitmap);
								} else {
									PremultiplyBitmap(scaledBitmap);
									hOldBmp = dc.SelectBitmap(scaledBitmap);
								}
								BLENDFUNCTION blendFunction = {0};
								blendFunction.BlendOp = AC_SRC_OVER;
								blendFunction.SourceConstantAlpha = static_cast<BYTE>(params.Frame);
								blendFunction.AlphaFormat = AC_SRC_ALPHA;
								AlphaBlend(params.hdcDst, params.x, params.y, params.cx, params.cy, dc, 0, 0, params.cx, params.cy, blendFunction);
								dc.SelectBitmap(hOldBmp);
								if(hBlendedBitmap) {
									DeleteObject(hBlendedBitmap);
								}
							} else {
								// use an imagelist to draw the thumbnail (works around alpha channel problems)
								params.himl = ImageList_Create(pParams->cx, pParams->cy, ILC_COLOR32, 1, 0);
								ATLASSERT(params.himl);
								if(params.himl) {
									params.i = ImageList_Add(params.himl, scaledBitmap, NULL);
									if(params.i >= 0) {
										b = ImageList_DrawIndirect(&params);
									}
									ImageList_Destroy(params.himl);
								}
							}
						}
					}
				}
			}
		#if FALSE
		} else {
			// This code has two issues: 1) For some icons it needs alpha-channel post-processing. 2) Some icons (folder icons on Windows XP) have some missing pixels.
			BITMAPINFO bitmapInfo = {0};
			bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			bitmapInfo.bmiHeader.biWidth = pParams->cx;
			bitmapInfo.bmiHeader.biHeight = pParams->cy;
			bitmapInfo.bmiHeader.biPlanes = 1;
			bitmapInfo.bmiHeader.biBitCount = 32;
			bitmapInfo.bmiHeader.biCompression = BI_RGB;
			LPRGBQUAD pBits = NULL;
			CBitmap scaledBitmap;
			scaledBitmap.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
			if(!scaledBitmap.IsNull()) {
				ATLASSERT_POINTER(pBits, RGBQUAD);
				if(pBits) {
					IMAGELISTDRAWPARAMS params = *pParams;
					SIZE iconSize = {0};
					if(GetIconSize(hIconToScale, &iconSize)) {
						LPRGBQUAD pIconBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, iconSize.cx * iconSize.cy * sizeof(RGBQUAD)));
						if(pIconBits) {
							ICONINFO iconInfo = {0};
							GetIconInfo(hIconToScale, &iconInfo);
							if(iconInfo.hbmColor) {
								CDC memoryDC;
								memoryDC.CreateCompatibleDC(NULL);
								BITMAP bmp = {0};
								GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bmp);
								ATLASSERT(bmp.bmHeight >= 0);

								BITMAPINFO bitmapInfo;
								bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
								bitmapInfo.bmiHeader.biWidth = bmp.bmWidth;
								bitmapInfo.bmiHeader.biHeight = bmp.bmHeight;
								bitmapInfo.bmiHeader.biPlanes = 1;
								bitmapInfo.bmiHeader.biBitCount = 32;
								bitmapInfo.bmiHeader.biCompression = 0;
								bitmapInfo.bmiHeader.biSizeImage = bmp.bmHeight * bmp.bmWidth * 4;
								bitmapInfo.bmiHeader.biXPelsPerMeter = 0;
								bitmapInfo.bmiHeader.biYPelsPerMeter = 0;
								bitmapInfo.bmiHeader.biClrUsed = 0;
								bitmapInfo.bmiHeader.biClrImportant = 0;
								LPRGBQUAD pColorBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
								if(pColorBits) {
									GetDIBits(memoryDC, iconInfo.hbmColor, 0, bmp.bmHeight, pColorBits, &bitmapInfo, DIB_RGB_COLORS);
									if(iconInfo.hbmMask) {
										CDC maskDC;
										maskDC.CreateCompatibleDC(NULL);
										HBITMAP hOldBmp = maskDC.SelectBitmap(iconInfo.hbmMask);

										//#pragma omp parallel for -> produces artefacts, probably conflicts with GetPixel
										for(int i = 0; i < iconSize.cy * iconSize.cx; i++) {
											if(maskDC.GetPixel(i % iconSize.cx, i / iconSize.cx) == 0x00000000) {
												pIconBits[i] = pColorBits[i];
											} else {
												pIconBits[i].rgbRed = 0;
												pIconBits[i].rgbGreen = 0;
												pIconBits[i].rgbBlue = 0;
												pIconBits[i].rgbReserved = 0;
											}
										}

										maskDC.SelectBitmap(hOldBmp);
										DeleteObject(iconInfo.hbmMask);
									}
									HeapFree(GetProcessHeap(), 0, pColorBits);
								}
								DeleteObject(iconInfo.hbmColor);
							}

							CDC memoryDC;
							memoryDC.CreateCompatibleDC(NULL);
							HBITMAP hOldBmp = memoryDC.SelectBitmap(scaledBitmap);
							BITMAPINFO bitmapInfo;
							bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
							bitmapInfo.bmiHeader.biWidth = pParams->cx;
							bitmapInfo.bmiHeader.biHeight = pParams->cy;
							bitmapInfo.bmiHeader.biPlanes = 1;
							bitmapInfo.bmiHeader.biBitCount = 32;
							bitmapInfo.bmiHeader.biCompression = 0;
							bitmapInfo.bmiHeader.biSizeImage = pParams->cx * pParams->cy * 4;
							bitmapInfo.bmiHeader.biXPelsPerMeter = 0;
							bitmapInfo.bmiHeader.biYPelsPerMeter = 0;
							bitmapInfo.bmiHeader.biClrUsed = 0;
							bitmapInfo.bmiHeader.biClrImportant = 0;
							memoryDC.StretchDIBits(0, 0, pParams->cx, pParams->cy, 0, 0, iconSize.cx, iconSize.cy, reinterpret_cast<LPVOID*>(pIconBits), &bitmapInfo, DIB_RGB_COLORS, SRCCOPY);
							memoryDC.SelectBitmap(hOldBmp);
							/*// Scale pIconBits to pBits
							float blockWidth = static_cast<float>(iconSize.cx) / static_cast<float>(pParams->cx);
							float blockHeight = static_cast<float>(iconSize.cy) / static_cast<float>(pParams->cy);
							for(int yDst = 0; yDst < pParams->cy; yDst++) {
								for(int xDst = 0; xDst < pParams->cx; xDst++) {
									int dst = yDst * pParams->cx + xDst;
									int red = 0;
									int green = 0;
									int blue = 0;
									int alpha = 0;
									int c = 0;
									for(int ySrc = static_cast<int>(yDst * blockHeight); ySrc < (yDst + 1) * blockHeight; ySrc++) {
										for(int xSrc = static_cast<int>(xDst * blockWidth); xSrc < (xDst + 1) * blockWidth; xSrc++) {
											int src = ySrc * iconSize.cx + xSrc;
											red += pIconBits[src].rgbRed;
											green += pIconBits[src].rgbGreen;
											blue += pIconBits[src].rgbBlue;
											alpha += pIconBits[src].rgbReserved;
											c++;
										}
									}
									pBits[dst].rgbRed = static_cast<BYTE>(red / c);
									pBits[dst].rgbGreen = static_cast<BYTE>(green / c);
									pBits[dst].rgbBlue = static_cast<BYTE>(blue / c);
									pBits[dst].rgbReserved = static_cast<BYTE>(alpha / c);
								}
							}*/

							HeapFree(GetProcessHeap(), 0, pIconBits);
						}
					}

					if(params.fState & ILS_ALPHA) {
						// blend pBits by the value specified in params.Frame
						LPRGBQUAD p = pBits;
						for(int i = 0; i < pParams->cy * pParams->cx; i++, p++) {
							p->rgbReserved = static_cast<BYTE>((static_cast<DWORD>(p->rgbReserved) * params.Frame) / 0xFF);
						}
						params.fState &= ~ILS_ALPHA;
						params.Frame = 0;
					}

					if(flags & AII_USELEGACYDISPLAYCODE) {
						params.Frame = 255 - params.Frame;
						WTL::CDC dc;
						dc.CreateCompatibleDC();
						HBITMAP hOldBmp = NULL;
						HBITMAP hBlendedBitmap = NULL;
						BOOL blend = (params.fStyle & (ILD_BLEND50 | ILD_BLEND25)) != 0;
						if(blend && params.rgbFg != CLR_NONE) {
							COLORREF blendColor = (params.rgbFg == CLR_DEFAULT ? GetSysColor(COLOR_HIGHLIGHT) : params.rgbFg);
							hBlendedBitmap = BlendBitmap(scaledBitmap, blendColor, (params.fStyle & ILD_BLEND50) ? 50 : 25, TRUE);
							hOldBmp = dc.SelectBitmap(hBlendedBitmap);
						} else {
							PremultiplyBitmap(scaledBitmap);
							hOldBmp = dc.SelectBitmap(scaledBitmap);
						}
						BLENDFUNCTION blendFunction = {0};
						blendFunction.BlendOp = AC_SRC_OVER;
						blendFunction.SourceConstantAlpha = static_cast<BYTE>(params.Frame);
						blendFunction.AlphaFormat = AC_SRC_ALPHA;
						AlphaBlend(params.hdcDst, params.x, params.y, params.cx, params.cy, dc, 0, 0, params.cx, params.cy, blendFunction);
						dc.SelectBitmap(hOldBmp);
						if(hBlendedBitmap) {
							DeleteObject(hBlendedBitmap);
						}
					} else {
						// use an imagelist to draw the thumbnail (works around alpha channel problems)
						params.himl = ImageList_Create(pParams->cx, pParams->cy, ILC_COLOR32, 1, 0);
						ATLASSERT(params.himl);
						if(params.himl) {
							params.i = ImageList_Add(params.himl, scaledBitmap, NULL);
							if(params.i >= 0) {
								BitBlt(params.hdcDst, params.x, params.y, params.cx, params.cy, pParams->hdcDst, 0, 0, WHITENESS);
								b = ImageList_DrawIndirect(&params);
							}
							ImageList_Destroy(params.himl);
						}
					}
				}
			}
		#endif
		}
		DestroyIcon(hIconToScale);
	}

	return b;
}


BOOL UnifiedImageList_IsComctl32Version600()
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hRet = ATL::AtlGetCommCtrlVersion(&major, &minor);
	return (SUCCEEDED(hRet) && (major == 6) && (minor == 0));
}