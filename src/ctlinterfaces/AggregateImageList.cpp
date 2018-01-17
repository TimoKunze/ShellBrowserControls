// AggregateImageList.cpp: An implementation of the IImageList interface wrapping multiple image lists

#include "stdafx.h"
#include "AggregateImageList.h"
#include "../ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE AggregateImageList_AddRef(IImageList2* pInterface)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	return InterlockedIncrement(&pThis->referenceCount);
}

ULONG STDMETHODCALLTYPE AggregateImageList_IImageListPrivate_AddRef(IImageListPrivate* pInterface)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromAggregateImageList(pThis);
	return AggregateImageList_AddRef(pBaseInterface);
}

ULONG STDMETHODCALLTYPE AggregateImageList_Release(IImageList2* pInterface)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(pThis->magicValue != 0x4C4D4948) {
		return 0;
	}

	ULONG ret = InterlockedDecrement(&pThis->referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		pThis->destroyingIconInfo = TRUE;
		AggregateImageList_RemoveIconInfo(IImageListPrivateFromAggregateImageList(pThis), 0xFFFFFFFF);
		EnterCriticalSection(&pThis->freeSlotsLock);
		if(pThis->hFreeSlots) {
			DPA_Destroy(pThis->hFreeSlots);
		}
		LeaveCriticalSection(&pThis->freeSlotsLock);
		DeleteCriticalSection(&pThis->freeSlotsLock);
		delete pThis->pIconInfos;
		if(pThis->pWICImagingFactory) {
			pThis->pWICImagingFactory->Release();
		}
		if(pThis->wrappedImageLists.pThumbnails) {
			pThis->wrappedImageLists.pThumbnails->Release();
		}
		pThis->magicValue = 0x00000000;
		HeapFree(GetProcessHeap(), 0, pThis);
	}
	return ret;
}

ULONG STDMETHODCALLTYPE AggregateImageList_IImageListPrivate_Release(IImageListPrivate* pInterface)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromAggregateImageList(pThis);
	return AggregateImageList_Release(pBaseInterface);
}

STDMETHODIMP AggregateImageList_QueryInterface(IImageList2* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = reinterpret_cast<LPUNKNOWN>(&pThis->pImageList2Vtable);
	} else if(requiredInterface == IID_IImageList) {
		*ppInterfaceImpl = IImageListFromAggregateImageList(pThis);
	} else if(requiredInterface == IID_IImageList2) {
		*ppInterfaceImpl = IImageList2FromAggregateImageList(pThis);
	} else if(requiredInterface == IID_IImageListPrivate) {
		*ppInterfaceImpl = IImageListPrivateFromAggregateImageList(pThis);
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	AggregateImageList_AddRef(pInterface);
	return S_OK;
}

STDMETHODIMP AggregateImageList_IImageListPrivate_QueryInterface(IImageListPrivate* pInterface, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	IImageList2* pBaseInterface = IImageList2FromAggregateImageList(pThis);
	return AggregateImageList_QueryInterface(pBaseInterface, requiredInterface, ppInterfaceImpl);
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList
STDMETHODIMP AggregateImageList_Add(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Add() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_ReplaceIcon(IImageList2* /*pInterface*/, int /*i*/, HICON /*hicon*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "AggregateImageList_ReplaceIcon() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_SetOverlayImage(IImageList2* /*pInterface*/, int /*iImage*/, int /*iOverlay*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetOverlayImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Replace(IImageList2* /*pInterface*/, int /*i*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Replace() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_AddMasked(IImageList2* /*pInterface*/, HBITMAP /*hbmImage*/, COLORREF /*crMask*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "AggregateImageList_AddMasked() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Draw(IImageList2* pInterface, IMAGELISTDRAWPARAMS* pimldp)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(pimldp->i);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(pimldp->i);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			IMAGELISTDRAWPARAMS params = *pimldp;
			params.himl = pThis->wrappedImageLists.hNonShellItems;
			params.i = pIconInfo->mainIconIndex;
			// center icon
			ImageList_GetIconSize(params.himl, &params.cx, &params.cy);
			params.x += ((pimldp->cx == 0 ? pThis->imageSize.cx : pimldp->cx) - params.cx) >> 1;
			params.y += ((pimldp->cy == 0 ? pThis->imageSize.cy : pimldp->cy) - params.cy) >> 1;
			if(ImageList_DrawIndirect(&params)) {
				return S_OK;
			} else {
				return E_FAIL;
			}
		}
	}

	if(pimldp->cx == 0 && pimldp->cy == 0) {
		pimldp->cx = pThis->imageSize.cx;
		pimldp->cy = pThis->imageSize.cy;
	}
	if(!pThis->pWICImagingFactory || (pIconInfo->flags & AII_USELEGACYDISPLAYCODE)) {
		return AggregateImageList_Draw_Legacy(pThis, pIconInfo, pimldp);
	}

	HRESULT hr = E_FAIL;
	// TODO: Handling of pimldp->cy == 0 is probably wrong!
	int desiredImageHeight = (pimldp->cy > 0 ? pimldp->cy : pimldp->cx);
	int overlayIndex = (pimldp->fStyle & ILD_OVERLAYMASK) >> 8;
	int overlayLeft = pimldp->x;
	int imageRight = pimldp->x + pimldp->cx;
	BOOL ghosted = ((pimldp->fStyle & ILD_BLEND) == ILD_BLEND);
	if(pimldp->rgbFg == CLR_DEFAULT) {
		/* TODO: Properly support standard theme. This won't be possible as long as we draw and blend each
		 *       component for its own. The big problem is, that IImageList::GetIcon doesn't handle icons
		 *       that are too small (and therefore get framed) well.
		 */
		ghosted = FALSE;
	}

	BOOL hasThumbnail = (pIconInfo->flags & AII_HASTHUMBNAIL);
	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));
	BOOL drawFileTypeOverlay = FALSE;
	HICON hOverlayImage = NULL;
	if(hasThumbnail) {
		switch(pIconInfo->flags & AII_FILETYPEOVERLAYMASK) {
			case AII_EXECUTABLEICONOVERLAY:
				// ensure the icon is valid
				drawFileTypeOverlay = (pThis->wrappedImageLists.hExecutableOverlays && pIconInfo->overlay.iconIndex >= 0);
				break;
			case AII_IMAGERESOURCEOVERLAY:
				// load the image resource
				drawFileTypeOverlay = (pIconInfo->overlay.pImageResource != NULL && lstrlenW(pIconInfo->overlay.pImageResource) > 0);
				if(drawFileTypeOverlay) {
					//TODO: hOverlayImage = LoadIconResource(pIconInfo->overlay.pImageResource);
				}
				break;
		}
	}

	if(pThis->wrappedImageLists.pThumbnails && pIconInfo->mainIconIndex >= 0) {
		int imageHeight = 0;
		int imageWidth = 0;
		if(pThis->wrappedImageLists.pThumbnails2 && pIconInfo->mainIconIndex >= 0) {
			hr = pThis->wrappedImageLists.pThumbnails2->GetOriginalSize(pIconInfo->mainIconIndex, ILGOS_ALWAYS, &imageWidth, &imageHeight);
		} else {
			hr = pThis->wrappedImageLists.pThumbnails->GetIconSize(&imageWidth, &imageHeight);
		}

		IMAGELISTDRAWPARAMS params = *pimldp;
		params.himl = IImageListToHIMAGELIST(pThis->wrappedImageLists.pThumbnails);
		params.i = pIconInfo->mainIconIndex;
		if(ghosted) {
			params.fState |= ILS_ALPHA;
			params.Frame = 128;
		}

		params.x += (pimldp->cx - imageWidth) >> 1;
		overlayLeft = params.x;
		imageRight = params.x + imageWidth;
		if(hasThumbnail && RunTimeHelper::IsVista()) {
			// bottom-align
			params.y += (desiredImageHeight - imageHeight);
		} else {
			// center vertically
			params.y += (desiredImageHeight - imageHeight) >> 1;
		}
		params.cx = imageWidth;
		params.cy = imageHeight;
		params.fStyle &= ~(ILD_BLEND | ILD_SCALE | ILD_OVERLAYMASK);

		if(!drawFileTypeOverlay && drawUACStamp) {
			if(ImageList_DrawIndirect_UACStamped(&params, pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY)) {
				hr = S_OK;
				drawUACStamp = FALSE;
			}
		} else {
			hr = pThis->wrappedImageLists.pThumbnails->Draw(&params);
		}
	} else if(pIconInfo->systemIconIndex >= 0 && pThis->wrappedImageLists.hFallback) {
		IMAGELISTDRAWPARAMS params = *pimldp;
		params.himl = pThis->wrappedImageLists.hFallback;
		params.i = pIconInfo->systemIconIndex;
		ImageList_GetIconSize(pThis->wrappedImageLists.hFallback, &params.cx, &params.cy);
		params.cx = min(params.cx, pimldp->cx);
		params.cy = min(params.cy, desiredImageHeight);
		params.x += (pimldp->cx - params.cx) >> 1;
		params.y += (desiredImageHeight - params.cy) >> 1;
		if(ghosted) {
			params.fState |= ILS_ALPHA;
			params.Frame = 128;
		}
		params.fStyle &= ~(ILD_BLEND | ILD_OVERLAYMASK);
		params.fStyle |= ILD_SCALE;					// NOTE: If we ever get problems with overlays being too large, then this is the reason.
		if(!drawFileTypeOverlay && drawUACStamp) {
			if(ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags, FALSE)) {
				hr = S_OK;
			}
		} else if(ghosted && AggregateImageList_IsComctl32Version600()) {
			if(ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags, FALSE)) {
				hr = S_OK;
			}
		} else if(ImageList_DrawIndirect(&params)) {
			hr = S_OK;
		}
	}

	if(drawFileTypeOverlay && pIconInfo->mainIconIndex >= 0) {
		// draw file type overlay
		// TODO: image resource mode
		IMAGELISTDRAWPARAMS params = *pimldp;
		int maxOverlayWidth;
		int maxOverlayHeight;
		ImageList_GetIconSize(pThis->wrappedImageLists.hExecutableOverlays, &maxOverlayWidth, &maxOverlayHeight);
		params.cx = min(maxOverlayWidth, static_cast<int>(static_cast<float>(pimldp->cx) * 0.35f));
		params.cy = min(maxOverlayHeight, static_cast<int>(static_cast<float>(desiredImageHeight) * 0.35f));
		/*params.cx = static_cast<int>(static_cast<float>(pimldp->cx) / 4.15f);
		params.cy = static_cast<int>(static_cast<float>(desiredImageHeight) / 4.15f);*/
		params.cx = params.cy = max(params.cx, params.cy);
		params.x = imageRight - params.cx;
		params.y = pimldp->y + desiredImageHeight - params.cy;
		if(ghosted) {
			// TODO: We should not blend each component, but create the combined icon (without normal overlay) and blend this one.
			params.fState |= ILS_ALPHA;
			params.Frame = 192;
		}
		params.fStyle &= ~(ILD_BLEND | ILD_SCALE | ILD_OVERLAYMASK);
		params.fStyle |= ILD_TRANSPARENT;
		params.himl = pThis->wrappedImageLists.hExecutableOverlays;
		params.i = pIconInfo->overlay.iconIndex;
		ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags & (drawUACStamp ? 0xFFFFFFFF : ~AII_DISPLAYELEVATIONOVERLAY), TRUE);
	}

	if(pThis->wrappedImageLists.hOverlays && overlayIndex >= 1) {
		// draw overlay
		overlayIndex--;
		PINT pOverlayData = ImageList_GetOverlayData(pThis->wrappedImageLists.hOverlays);
		if(pOverlayData) {
			int maxOverlayWidth;
			int maxOverlayHeight;
			ImageList_GetIconSize(pThis->wrappedImageLists.hOverlays, &maxOverlayWidth, &maxOverlayHeight);
			IMAGELISTDRAWPARAMS params = *pimldp;
			params.himl = pThis->wrappedImageLists.hOverlays;
			params.i = pOverlayData[overlayIndex];
			params.cx = min(maxOverlayWidth, static_cast<int>(static_cast<float>(pimldp->cx) * 0.7f));
			params.cy = min(maxOverlayHeight, static_cast<int>(static_cast<float>(desiredImageHeight) * 0.7f));
			params.cx = params.cy = max(params.cx, params.cy);
			params.x = overlayLeft;
			params.y = pimldp->y + desiredImageHeight - params.cy;
			params.fStyle &= ~(ILD_BLEND | ILD_SCALE | ILD_OVERLAYMASK);
			params.fStyle |= ILD_TRANSPARENT;
			ImageList_DrawIndirect_HQScaling(pThis, &params, pIconInfo->flags & ~AII_DISPLAYELEVATIONOVERLAY, TRUE);
		}
	}
	if(hOverlayImage) {
		DestroyIcon(hOverlayImage);
	}
	return hr;
}

STDMETHODIMP AggregateImageList_Draw_Legacy(AggregateImageList* pThis, AggregatedIconInformation* pIconInfo, IMAGELISTDRAWPARAMS* pimldp)
{
	HRESULT hr = E_FAIL;

	// TODO: Handling of pimldp->cy == 0 is probably wrong!
	int desiredImageHeight = (pimldp->cy > 0 ? pimldp->cy : pimldp->cx);
	int overlayIndex = (pimldp->fStyle & ILD_OVERLAYMASK) >> 8;

	BOOL hasThumbnail = (pIconInfo->flags & AII_HASTHUMBNAIL);
	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));
	BOOL drawFileTypeOverlay = FALSE;
	HICON hOverlayImage = NULL;
	if(hasThumbnail) {
		switch(pIconInfo->flags & AII_FILETYPEOVERLAYMASK) {
			case AII_EXECUTABLEICONOVERLAY:
				// ensure the icon is valid
				drawFileTypeOverlay = (pThis->wrappedImageLists.hExecutableOverlays && pIconInfo->overlay.iconIndex >= 0);
				break;
			case AII_IMAGERESOURCEOVERLAY:
				// load the image resource
				drawFileTypeOverlay = (pIconInfo->overlay.pImageResource != NULL && lstrlenW(pIconInfo->overlay.pImageResource) > 0);
				if(drawFileTypeOverlay) {
					//TODO: hOverlayImage = LoadIconResource(pIconInfo->overlay.pImageResource);
				}
				break;
		}
	}

	HBITMAP hMainBitmap = NULL;
	HBITMAP hMaskBitmap = NULL;
	BOOL doAlphaChannelPostProcessing = FALSE;
	if(pIconInfo->hBmp) {
		BITMAP bmp = {0};
		GetObject(pIconInfo->hBmp, sizeof(bmp), &bmp);
		if(RunTimeHelper::IsVista()) {
			ICONINFO iconInfo = {0};
			iconInfo.fIcon = TRUE;
			iconInfo.hbmColor = pIconInfo->hBmp;
			HBITMAP hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, NULL);
			if(hMaskBmp) {
				WTL::CDC dc;
				dc.CreateCompatibleDC();
				HBITMAP hOldBitmap = dc.SelectBitmap(hMaskBmp);
				RECT rc = {0, 0, bmp.bmWidth, bmp.bmHeight};
				dc.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
				dc.SelectBitmap(hOldBitmap);
				iconInfo.hbmMask = hMaskBmp;
				HICON hIcon = CreateIconIndirect(&iconInfo);
				if(hIcon) {
					hMainBitmap = StampIconForElevation(hIcon, NULL, !drawFileTypeOverlay && drawUACStamp);
					DestroyIcon(hIcon);
					if(!drawFileTypeOverlay && drawUACStamp) {
						drawUACStamp = FALSE;
					}
				}
				DeleteObject(hMaskBmp);
			}
		} else {
			// Windows XP and older
			hMainBitmap = CopyBitmap(pIconInfo->hBmp);
			if(!RunTimeHelper::IsVista()) {
				hMaskBitmap = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, NULL);
			}
		}
	} else if(pIconInfo->systemIconIndex >= 0 && pThis->wrappedImageLists.hFallback) {
		SIZE imageSize = {0};
		ImageList_GetIconSize(pThis->wrappedImageLists.hFallback, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
		imageSize.cx = min(imageSize.cx, pThis->imageSize.cx);
		imageSize.cy = min(imageSize.cy, pThis->imageSize.cy);

		if(RunTimeHelper::IsVista()) {
			HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hFallback, pIconInfo->systemIconIndex, /*ILD_SCALE | */ILD_TRANSPARENT);					// NOTE: If we activate ILD_SCALE, overlays are drawn too large.
			ATLASSERT(hIcon);
			if(hIcon) {
				hMainBitmap = StampIconForElevation(hIcon, &imageSize, !drawFileTypeOverlay && drawUACStamp);
				DestroyIcon(hIcon);
				if(!drawFileTypeOverlay && drawUACStamp) {
					drawUACStamp = FALSE;
				}
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
				ImageList_Draw(pThis->wrappedImageLists.hFallback, pIconInfo->systemIconIndex, dc, 0, 0, ILD_TRANSPARENT);
				dc.SelectBitmap(hMaskBitmap);
				ImageList_Draw(pThis->wrappedImageLists.hFallback, pIconInfo->systemIconIndex, dc, 0, 0, ILD_MASK);
				dc.SelectBitmap(hOldBmp);

				IImageList* pImgLst = NULL;
				if(APIWrapper::IsSupported_HIMAGELIST_QueryInterface()) {
					APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hFallback, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst), NULL);
				} else {
					pImgLst = reinterpret_cast<IImageList*>(pThis->wrappedImageLists.hFallback);
					pImgLst->AddRef();
				}
				ATLASSUME(pImgLst);

				DWORD flags = 0;
				pImgLst->GetItemFlags(pIconInfo->systemIconIndex, &flags);
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

	if(drawFileTypeOverlay && pIconInfo->hBmp) {
		// merge file type overlay
		// TODO: image resource mode
		SIZE overlaySize = {0};
		ImageList_GetIconSize(pThis->wrappedImageLists.hExecutableOverlays, reinterpret_cast<PINT>(&overlaySize.cx), reinterpret_cast<PINT>(&overlaySize.cy));
		overlaySize.cx = min(overlaySize.cx, static_cast<int>(static_cast<float>(pThis->imageSize.cx) * 0.35f));
		overlaySize.cy = min(overlaySize.cy, static_cast<int>(static_cast<float>(pThis->imageSize.cy) * 0.35f));
		overlaySize.cx = overlaySize.cy = max(overlaySize.cx, overlaySize.cy);

		if(RunTimeHelper::IsVista()) {
			HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hExecutableOverlays, pIconInfo->overlay.iconIndex, ILD_TRANSPARENT);
			ATLASSERT(hIcon);
			if(hIcon) {
				HBITMAP hOverlayBitmap = StampIconForElevation(hIcon, &overlaySize, drawUACStamp);
				drawUACStamp = FALSE;
				if(hOverlayBitmap) {
					/* If the thumbnail size is changed, it may happen that the item needs to be drawn when the new
					 * executable overlay image list is already set, but the item's thumbnail has not yet been
					 * updated. The overlay may be larger than the icon then, so make sure we don't get offsets < 0.
					 */
					MergeBitmaps(hMainBitmap, hOverlayBitmap, max(iconSize.cx - overlaySize.cx, 0), max(iconSize.cy - overlaySize.cy, 0));
					DeleteObject(hOverlayBitmap);
				}
				DestroyIcon(hIcon);
			}
		} else {
			// On Windows XP with comctl32.dll 5.x, ImageList_GetIcon produces bad looking icons.
			BITMAPINFO bitmapInfo = {0};
			bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			bitmapInfo.bmiHeader.biWidth = overlaySize.cx;
			bitmapInfo.bmiHeader.biHeight = overlaySize.cy;
			bitmapInfo.bmiHeader.biPlanes = 1;
			bitmapInfo.bmiHeader.biBitCount = 32;
			bitmapInfo.bmiHeader.biSizeImage = overlaySize.cy * overlaySize.cx * 4;
			LPRGBQUAD pBits;
			HBITMAP hOverlayBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
			if(pBits) {
				ZeroMemory(pBits, bitmapInfo.bmiHeader.biSizeImage);
				WTL::CDC dc;
				dc.CreateCompatibleDC();
				HBITMAP hOldBmp = dc.SelectBitmap(hOverlayBitmap);
				ImageList_Draw(pThis->wrappedImageLists.hExecutableOverlays, pIconInfo->overlay.iconIndex, dc, 0, 0, ILD_TRANSPARENT);
				dc.SelectBitmap(hMaskBitmap);
				ImageList_Draw(pThis->wrappedImageLists.hExecutableOverlays, pIconInfo->overlay.iconIndex, dc, max(iconSize.cx - overlaySize.cx, 0), max(iconSize.cy - overlaySize.cy, 0), ILD_MASK);
				dc.SelectBitmap(hOldBmp);
				/* If the thumbnail size is changed, it may happen that the item needs to be drawn when the new
				 * executable overlay image list is already set, but the item's thumbnail has not yet been
				 * updated. The overlay may be larger than the icon then, so make sure we don't get offsets < 0.
				 */
				MergeBitmaps(hMainBitmap, hOverlayBitmap, max(iconSize.cx - overlaySize.cx, 0), max(iconSize.cy - overlaySize.cy, 0));
				DeleteObject(hOverlayBitmap);

				if(!doAlphaChannelPostProcessing) {
					IImageList* pImgLst = NULL;
					if(APIWrapper::IsSupported_HIMAGELIST_QueryInterface()) {
						APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hExecutableOverlays, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst), NULL);
					} else {
						pImgLst = reinterpret_cast<IImageList*>(pThis->wrappedImageLists.hExecutableOverlays);
						pImgLst->AddRef();
					}
					ATLASSUME(pImgLst);

					DWORD flags = 0;
					pImgLst->GetItemFlags(pIconInfo->overlay.iconIndex, &flags);
					pImgLst->Release();
					doAlphaChannelPostProcessing = !(flags & ILIF_ALPHA);
				}
			}
		}
	}
	if(hOverlayImage) {
		DestroyIcon(hOverlayImage);
	}

	if(pThis->wrappedImageLists.hOverlays && overlayIndex >= 1) {
		// draw overlay
		overlayIndex--;
		PINT pOverlayData = ImageList_GetOverlayData(pThis->wrappedImageLists.hOverlays);
		if(pOverlayData) {
			SIZE overlaySize = {0};
			ImageList_GetIconSize(pThis->wrappedImageLists.hOverlays, reinterpret_cast<PINT>(&overlaySize.cx), reinterpret_cast<PINT>(&overlaySize.cy));
			overlaySize.cx = min(overlaySize.cx, static_cast<int>(static_cast<float>(pThis->imageSize.cx) * 0.7f));
			overlaySize.cy = min(overlaySize.cy, static_cast<int>(static_cast<float>(pThis->imageSize.cy) * 0.7f));
			overlaySize.cx = overlaySize.cy = max(overlaySize.cx, overlaySize.cy);

			if(RunTimeHelper::IsVista()) {
				HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hOverlays, pOverlayData[overlayIndex], ILD_TRANSPARENT);
				ATLASSERT(hIcon);
				if(hIcon) {
					HBITMAP hOverlayBitmap = StampIconForElevation(hIcon, &overlaySize, FALSE);
					if(hOverlayBitmap) {
						/* If the thumbnail size is changed, it may happen that the item needs to be drawn when the new
						 * executable overlay image list is already set, but the item's thumbnail has not yet been
						 * updated. The overlay may be larger than the icon then, so make sure we don't get offsets < 0.
						 */
						MergeBitmaps(hMainBitmap, hOverlayBitmap, 0, max(iconSize.cy - overlaySize.cy, 0));
						DeleteObject(hOverlayBitmap);
					}
					DestroyIcon(hIcon);
				}
			} else {
				// On Windows XP with comctl32.dll 5.x, ImageList_GetIcon produces bad looking icons.
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = overlaySize.cx;
				bitmapInfo.bmiHeader.biHeight = overlaySize.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biSizeImage = overlaySize.cy * overlaySize.cx * 4;
				LPRGBQUAD pBits;
				HBITMAP hOverlayBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
				if(pBits) {
					ZeroMemory(pBits, bitmapInfo.bmiHeader.biSizeImage);
					WTL::CDC dc;
					dc.CreateCompatibleDC();
					HBITMAP hOldBmp = dc.SelectBitmap(hOverlayBitmap);
					ImageList_Draw(pThis->wrappedImageLists.hOverlays, pOverlayData[overlayIndex], dc, 0, 0, ILD_TRANSPARENT);
					dc.SelectBitmap(hMaskBitmap);
					ImageList_Draw(pThis->wrappedImageLists.hOverlays, pOverlayData[overlayIndex], dc, 0, max(iconSize.cy - overlaySize.cy, 0), ILD_MASK);
					dc.SelectBitmap(hOldBmp);
					/* If the thumbnail size is changed, it may happen that the item needs to be drawn when the new
					 * executable overlay image list is already set, but the item's thumbnail has not yet been
					 * updated. The overlay may be larger than the icon then, so make sure we don't get offsets < 0.
					 */
					MergeBitmaps(hMainBitmap, hOverlayBitmap, 0, max(iconSize.cy - overlaySize.cy, 0));
					DeleteObject(hOverlayBitmap);

					if(!doAlphaChannelPostProcessing) {
						IImageList* pImgLst = NULL;
						if(APIWrapper::IsSupported_HIMAGELIST_QueryInterface()) {
							APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hOverlays, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst), NULL);
						} else {
							pImgLst = reinterpret_cast<IImageList*>(pThis->wrappedImageLists.hOverlays);
							pImgLst->AddRef();
						}
						ATLASSUME(pImgLst);

						DWORD flags = 0;
						pImgLst->GetItemFlags(pOverlayData[overlayIndex], &flags);
						pImgLst->Release();
						doAlphaChannelPostProcessing = !(flags & ILIF_ALPHA);
					}
				}
			}
		}
	}

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
	if(hasThumbnail && RunTimeHelper::IsVista()) {
		// bottom-align
		params.y += (desiredImageHeight - iconSize.cy);
	} else {
		// center vertically
		params.y += (desiredImageHeight - iconSize.cy) >> 1;
	}
	params.fStyle &= ~ILD_OVERLAYMASK;
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

STDMETHODIMP AggregateImageList_Remove(IImageList2* /*pInterface*/, int /*i*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Remove() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetIcon(IImageList2* pInterface, int i, UINT flags, HICON* picon)
{
	/* TODO: This method is pretty much broken. Also it should be split into 2 methods, one for legacy
	 *       drawing mode and one for modern mode - just like the Draw method.
	 */

	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	int overlayIndex = (flags & ILD_OVERLAYMASK) >> 8;
	flags &= ~ILD_OVERLAYMASK;
	int imageRight = pThis->imageSize.cx;
	int imageBottom = pThis->imageSize.cy;
	// TODO: support ILD_BLEND
	//BOOL ghosted = ((flags & ILD_BLEND) == ILD_BLEND);
	flags &= ~(ILD_BLEND | ILD_SCALE);

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(i);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(i);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			*picon = ImageList_GetIcon(pThis->wrappedImageLists.hNonShellItems, pIconInfo->mainIconIndex, flags);
			return S_OK;
		}
	}

	BOOL hasThumbnail = (pIconInfo->flags & AII_HASTHUMBNAIL);
	BOOL drawUACStamp = ((pThis->flags & ILOF_DISPLAYELEVATIONSHIELDS) && (pIconInfo->flags & AII_DISPLAYELEVATIONOVERLAY));
	BOOL drawFileTypeOverlay = FALSE;
	HICON hOverlayImage = NULL;
	if(hasThumbnail) {
		switch(pIconInfo->flags & AII_FILETYPEOVERLAYMASK) {
			case AII_EXECUTABLEICONOVERLAY:
				// ensure the icon is valid
				drawFileTypeOverlay = (pThis->wrappedImageLists.hExecutableOverlays && pIconInfo->overlay.iconIndex >= 0);
				break;
			case AII_IMAGERESOURCEOVERLAY:
				// load the image resource
				drawFileTypeOverlay = (pIconInfo->overlay.pImageResource != NULL && lstrlenW(pIconInfo->overlay.pImageResource) > 0);
				if(drawFileTypeOverlay) {
					//TODO: hOverlayImage = LoadIconResource(pIconInfo->overlay.pImageResource);
				}
				break;
		}
	}

	HBITMAP hMainBitmap = NULL;
	if((pThis->wrappedImageLists.pThumbnails && pIconInfo->mainIconIndex >= 0) || ((pIconInfo->flags & AII_USELEGACYDISPLAYCODE) && pIconInfo->hBmp)) {
		if(pIconInfo->flags & AII_USELEGACYDISPLAYCODE) {
			ATLASSERT(FALSE && "AggregateImageList_GetIcon isn't fully supported for comctl32.dll 5.x");
			// TODO: Set imageRight and imageBottom.
			hMainBitmap = CopyBitmap(pIconInfo->hBmp);
			// TODO: StampIconForElevation
		} else {
			if(pThis->wrappedImageLists.pThumbnails2) {
				hr = pThis->wrappedImageLists.pThumbnails2->GetOriginalSize(pIconInfo->mainIconIndex, ILGOS_ALWAYS, &imageRight, &imageBottom);
			} else {
				hr = pThis->wrappedImageLists.pThumbnails->GetIconSize(&imageRight, &imageBottom);
			}
			HICON hIcon = NULL;
			ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->GetIcon(pIconInfo->mainIconIndex, 0, &hIcon)));
			if(hIcon) {
				hMainBitmap = StampIconForElevation(hIcon, NULL, drawUACStamp);
				DestroyIcon(hIcon);
			}
		}
	} else if(pIconInfo->systemIconIndex >= 0 && pThis->wrappedImageLists.hFallback) {
		SIZE imageSize = {0};
		ImageList_GetIconSize(pThis->wrappedImageLists.hFallback, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
		imageSize.cx = min(imageSize.cx, pThis->imageSize.cx);
		imageSize.cy = min(imageSize.cy, pThis->imageSize.cy);

		HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hFallback, pIconInfo->systemIconIndex, /*ILD_SCALE | */ILD_TRANSPARENT);					// NOTE: If we activate ILD_SCALE, overlays are drawn too large.
		ATLASSERT(hIcon);
		if(hIcon) {
			hMainBitmap = StampIconForElevation(hIcon, &imageSize, drawUACStamp);
			DestroyIcon(hIcon);
		}
	}
	if(!hMainBitmap) {
		return hr;
	}

	if(drawFileTypeOverlay && (pIconInfo->mainIconIndex >= 0 || ((pIconInfo->flags & AII_USELEGACYDISPLAYCODE) && pIconInfo->hBmp))) {
		// merge file type overlay
		// TODO: image resource mode
		SIZE overlaySize = {0};
		ImageList_GetIconSize(pThis->wrappedImageLists.hExecutableOverlays, reinterpret_cast<PINT>(&overlaySize.cx), reinterpret_cast<PINT>(&overlaySize.cy));
		overlaySize.cx = min(overlaySize.cx, static_cast<int>(static_cast<float>(pThis->imageSize.cx) * 0.35f));
		overlaySize.cy = min(overlaySize.cy, static_cast<int>(static_cast<float>(pThis->imageSize.cy) * 0.35f));
		overlaySize.cx = overlaySize.cy = max(overlaySize.cx, overlaySize.cy);

		HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hExecutableOverlays, pIconInfo->overlay.iconIndex, ILD_TRANSPARENT);
		ATLASSERT(hIcon);
		if(hIcon) {
			HBITMAP hOverlayBitmap = StampIconForElevation(hIcon, &overlaySize, drawUACStamp);
			if(hOverlayBitmap) {
				MergeBitmaps(hMainBitmap, hOverlayBitmap, max(imageRight - overlaySize.cx, 0), max(imageBottom - overlaySize.cy, 0));
				DeleteObject(hOverlayBitmap);
			}
			DestroyIcon(hIcon);
		}
	}
	if(hOverlayImage) {
		DestroyIcon(hOverlayImage);
	}

	if(pThis->wrappedImageLists.hOverlays && overlayIndex >= 1) {
		// draw overlay
		overlayIndex--;
		PINT pOverlayData = ImageList_GetOverlayData(pThis->wrappedImageLists.hOverlays);
		if(pOverlayData) {
			SIZE overlaySize = {0};
			ImageList_GetIconSize(pThis->wrappedImageLists.hOverlays, reinterpret_cast<PINT>(&overlaySize.cx), reinterpret_cast<PINT>(&overlaySize.cy));
			overlaySize.cx = min(overlaySize.cx, static_cast<int>(static_cast<float>(pThis->imageSize.cx) * 0.7f));
			overlaySize.cy = min(overlaySize.cy, static_cast<int>(static_cast<float>(pThis->imageSize.cy) * 0.7f));
			overlaySize.cx = overlaySize.cy = max(overlaySize.cx, overlaySize.cy);

			HICON hIcon = ImageList_GetIcon(pThis->wrappedImageLists.hOverlays, pOverlayData[overlayIndex], ILD_TRANSPARENT);
			ATLASSERT(hIcon);
			if(hIcon) {
				HBITMAP hOverlayBitmap = StampIconForElevation(hIcon, &overlaySize, drawUACStamp);
				if(hOverlayBitmap) {
					MergeBitmaps(hMainBitmap, hOverlayBitmap, 0, max(imageBottom - overlaySize.cy, 0));
					DeleteObject(hOverlayBitmap);
				}
				DestroyIcon(hIcon);
			}
		}
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

STDMETHODIMP AggregateImageList_GetImageInfo(IImageList2* /*pInterface*/, int /*i*/, IMAGEINFO* /*pImageInfo*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetImageInfo() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Copy(IImageList2* /*pInterface*/, int /*iDst*/, IUnknown* /*punkSrc*/, int /*iSrc*/, UINT /*uFlags*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Copy() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Merge(IImageList2* /*pInterface*/, int /*i1*/, IUnknown* /*punk2*/, int /*i2*/, int /*dx*/, int /*dy*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Merge() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Clone(IImageList2* /*pInterface*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Clone() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetImageRect(IImageList2* /*pInterface*/, int /*i*/, RECT* /*prc*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetImageRect() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetIconSize(IImageList2* pInterface, int* cx, int* cy)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	*cx = pThis->imageSize.cx;
	*cy = pThis->imageSize.cy;
	return S_OK;
}

STDMETHODIMP AggregateImageList_SetIconSize(IImageList2* pInterface, int cx, int cy)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	SIZE tmp = pThis->imageSize;
	pThis->imageSize.cx = cx;
	pThis->imageSize.cy = cy;

	if(pThis->wrappedImageLists.pThumbnails) {
		if(pThis->hasCreatedMainImageList) {
			if(pThis->wrappedImageLists.pThumbnails2) {
				ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails2->Resize(cx, cy)));
			} else {
				pThis->imageSize = tmp;
				return E_NOTIMPL;
			}
		} else {
			if(pThis->wrappedImageLists.pThumbnails2) {
				ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails2->Initialize(cx, cy, ILC_COLOR32 | ILC_ORIGINALSIZE | ILC_HIGHQUALITYSCALE, 5, 5)));
				pThis->hasCreatedMainImageList = TRUE;
			} else {
				// should never happen
				ATLASSERT(FALSE);
				// TODO: Shouldn't we first release pThis->wrappedImageLists.pThumbnails?
				HIMAGELIST h = NULL;
				if(RunTimeHelper::IsCommCtrl6()) {
					h = ImageList_Create(cx, cy, ILC_COLOR32, 5, 5);
					ATLASSERT(h);
					HRESULT hr = E_FAIL;
					APIWrapper::HIMAGELIST_QueryInterface(h, IID_PPV_ARGS(&pThis->wrappedImageLists.pThumbnails), &hr);
					pThis->hasCreatedMainImageList = (SUCCEEDED(hr) && pThis->wrappedImageLists.pThumbnails);
				}
			}
		}
	} else {
		if(RunTimeHelper::IsCommCtrl6()) {
			HIMAGELIST h = ImageList_Create(cx, cy, ILC_COLOR32, 5, 5);
			ATLASSERT(h);
			HRESULT hr = E_FAIL;
			APIWrapper::HIMAGELIST_QueryInterface(h, IID_PPV_ARGS(&pThis->wrappedImageLists.pThumbnails), &hr);
			pThis->hasCreatedMainImageList = (SUCCEEDED(hr) && pThis->wrappedImageLists.pThumbnails);
		}
	}
	return S_OK;
}

STDMETHODIMP AggregateImageList_GetImageCount(IImageList2* /*pInterface*/, int* /*pi*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetImageCount() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_SetImageCount(IImageList2* /*pInterface*/, UINT /*uNewCount*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetImageCount() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_SetBkColor(IImageList2* /*pInterface*/, COLORREF /*clrBk*/, COLORREF* /*pclr*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetBkColor() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetBkColor(IImageList2* /*pInterface*/, COLORREF* pclr)
{
	*pclr = CLR_NONE;
	return S_OK;
}

STDMETHODIMP AggregateImageList_BeginDrag(IImageList2* /*pInterface*/, int /*iTrack*/, int /*dxHotspot*/, int /*dyHotspot*/)
{
	ATLASSERT(FALSE && "AggregateImageList_BeginDrag() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_EndDrag(IImageList2* /*pInterface*/)
{
	ATLASSERT(FALSE && "AggregateImageList_EndDrag() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_DragEnter(IImageList2* /*pInterface*/, HWND /*hwndLock*/, int /*x*/, int /*y*/)
{
	ATLASSERT(FALSE && "AggregateImageList_DragEnter() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_DragLeave(IImageList2* /*pInterface*/, HWND /*hwndLock*/)
{
	ATLASSERT(FALSE && "AggregateImageList_DragLeave() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_DragMove(IImageList2* /*pInterface*/, int /*x*/, int /*y*/)
{
	ATLASSERT(FALSE && "AggregateImageList_DragMove() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_SetDragCursorImage(IImageList2* /*pInterface*/, IUnknown* /*punk*/, int /*iDrag*/, int /*dxHotspot*/, int /*dyHotspot*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetDragCursorImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_DragShowNolock(IImageList2* /*pInterface*/, BOOL /*fShow*/)
{
	ATLASSERT(FALSE && "AggregateImageList_DragShowNolock() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetDragImage(IImageList2* /*pInterface*/, POINT* /*ppt*/, POINT* /*pptHotspot*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetDragImage() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetItemFlags(IImageList2* /*pInterface*/, int /*i*/, DWORD* /*dwFlags*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetItemFlags() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetOverlayImage(IImageList2* /*pInterface*/, int /*iOverlay*/, int* /*piIndex*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetOverlayImage() should never be called!");
	return E_NOTIMPL;
}
// implementation of IImageList
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageList2
STDMETHODIMP AggregateImageList_Resize(IImageList2* pInterface, int cxNewIconSize, int cyNewIconSize)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);

	if(RunTimeHelper::IsCommCtrl6()) {
		if(pThis->wrappedImageLists.pThumbnails && pThis->hasCreatedMainImageList) {
			if(pThis->wrappedImageLists.pThumbnails2) {
				ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails2->Resize(cxNewIconSize, cyNewIconSize)));
			} else {
				return E_NOTIMPL;
			}
			pThis->imageSize.cx = cxNewIconSize;
			pThis->imageSize.cy = cyNewIconSize;
			return S_OK;
		}
	} else {
		pThis->imageSize.cx = cxNewIconSize;
		pThis->imageSize.cy = cyNewIconSize;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP AggregateImageList_GetOriginalSize(IImageList2* pInterface, int iImage, DWORD dwFlags, int* pcx, int* pcy)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(iImage);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(iImage);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			CComPtr<IImageList2> pImgLst = NULL;
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hNonShellItems, IID_PPV_ARGS(&pImgLst), NULL);
			if(pImgLst) {
				return pImgLst->GetOriginalSize(pIconInfo->mainIconIndex, dwFlags, pcx, pcy);
			}
			return E_FAIL;
		}
	}

	if((pThis->wrappedImageLists.pThumbnails && pIconInfo->mainIconIndex >= 0) || ((pIconInfo->flags & AII_USELEGACYDISPLAYCODE) && pIconInfo->hBmp)) {
		if(RunTimeHelper::IsCommCtrl6()) {
			if(pThis->hasCreatedMainImageList) {
				if(pThis->wrappedImageLists.pThumbnails2) {
					return pThis->wrappedImageLists.pThumbnails2->GetOriginalSize(pIconInfo->mainIconIndex, dwFlags, pcx, pcy);
				} else {
					return E_NOTIMPL;
				}
			}
		} else if(pIconInfo->hBmp) {
			BITMAP bmp = {0};
			GetObject(pIconInfo->hBmp, sizeof(bmp), &bmp);
			*pcx = bmp.bmWidth;
			*pcy = bmp.bmHeight;
		}
	} else if(pThis->wrappedImageLists.hFallback && pIconInfo->systemIconIndex >= 0) {
		CComPtr<IImageList2> pImgLst = NULL;
		APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hFallback, IID_PPV_ARGS(&pImgLst), NULL);
		if(pImgLst) {
			return pImgLst->GetOriginalSize(pIconInfo->mainIconIndex, dwFlags, pcx, pcy);
		}
		return E_FAIL;
	}
	return E_INVALIDARG;
}

STDMETHODIMP AggregateImageList_SetOriginalSize(IImageList2* /*pInterface*/, int /*iImage*/, int /*cx*/, int /*cy*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetOriginalSize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_SetCallback(IImageList2* /*pInterface*/, IUnknown* /*punk*/)
{
	ATLASSERT(FALSE && "AggregateImageList_SetCallback() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetCallback(IImageList2* /*pInterface*/, REFIID /*riid*/, void** /*ppv*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetCallback() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_ForceImagePresent(IImageList2* pInterface, int iImage, DWORD dwFlags)
{
	AggregateImageList* pThis = AggregateImageListFromIImageList2(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(iImage);
		if(iter == pThis->pIconInfos->cend()) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = iter->second;
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(iImage);
		if(!pEntry) {
			return E_INVALIDARG;
		}
		AggregatedIconInformation* pIconInfo = pEntry->m_value;
	#endif
	if(pIconInfo->flags & AII_NONSHELLITEM) {
		if(pThis->wrappedImageLists.hNonShellItems) {
			CComPtr<IImageList2> pImgLst = NULL;
			APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hNonShellItems, IID_PPV_ARGS(&pImgLst), NULL);
			if(pImgLst) {
				ATLVERIFY(SUCCEEDED(pImgLst->ForceImagePresent(pIconInfo->mainIconIndex, dwFlags)));
			}
		}
	}

	if(pThis->wrappedImageLists.pThumbnails && pIconInfo->mainIconIndex >= 0) {
		if(pThis->hasCreatedMainImageList) {
			if(pThis->wrappedImageLists.pThumbnails2) {
				ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails2->ForceImagePresent(pIconInfo->mainIconIndex, dwFlags)));
			}
		}
		CComPtr<IImageList2> pImgLst = NULL;
		APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hFallback, IID_PPV_ARGS(&pImgLst), NULL);
		if(pImgLst) {
			ATLVERIFY(SUCCEEDED(pImgLst->ForceImagePresent(pIconInfo->systemIconIndex, dwFlags)));
		}
		pImgLst = NULL;
		APIWrapper::HIMAGELIST_QueryInterface(pThis->wrappedImageLists.hExecutableOverlays, IID_PPV_ARGS(&pImgLst), NULL);
		if(pImgLst) {
			ATLVERIFY(SUCCEEDED(pImgLst->ForceImagePresent(pIconInfo->overlay.iconIndex, dwFlags)));
		}
	}
	return S_OK;
}

STDMETHODIMP AggregateImageList_DiscardImages(IImageList2* /*pInterface*/, int /*iFirstImage*/, int /*iLastImage*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "AggregateImageList_DiscardImages() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_PreloadImages(IImageList2* /*pInterface*/, IMAGELISTDRAWPARAMS* /*pimldp*/)
{
	ATLASSERT(FALSE && "AggregateImageList_PreloadImages() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_GetStatistics(IImageList2* /*pInterface*/, IMAGELISTSTATS* /*pils*/)
{
	ATLASSERT(FALSE && "AggregateImageList_GetStatistics() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Initialize(IImageList2* /*pInterface*/, int /*cx*/, int /*cy*/, UINT /*flags*/, int /*cInitial*/, int /*cGrow*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Initialize() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_Replace2(IImageList2* /*pInterface*/, int /*i*/, HBITMAP /*hbmImage*/, HBITMAP /*hbmMask*/, IUnknown* /*punk*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "AggregateImageList_Replace2() should never be called!");
	return E_NOTIMPL;
}

STDMETHODIMP AggregateImageList_ReplaceFromImageList(IImageList2* /*pInterface*/, int /*i*/, IImageList* /*pil*/, int /*iSrc*/, IUnknown* /*punk*/, DWORD /*dwFlags*/)
{
	ATLASSERT(FALSE && "AggregateImageList_ReplaceFromImageList() should never be called!");
	return E_NOTIMPL;
}
// implementation of IImageList2
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IImageListPrivate
STDMETHODIMP AggregateImageList_SetImageList(IImageListPrivate* pInterface, UINT imageListType, HIMAGELIST hImageList, HIMAGELIST* pHPreviousImageList)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	switch(imageListType) {
		case AIL_SHELLITEMS:
			if(pHPreviousImageList) {
				*pHPreviousImageList = pThis->wrappedImageLists.hFallback;
			}
			pThis->wrappedImageLists.hFallback = hImageList;
			return S_OK;
			break;
		case AIL_EXECUTABLEOVERLAYS:
			if(pHPreviousImageList) {
				*pHPreviousImageList = pThis->wrappedImageLists.hExecutableOverlays;
			}
			pThis->wrappedImageLists.hExecutableOverlays = hImageList;
			return S_OK;
			break;
		case AIL_OVERLAYS:
			if(pHPreviousImageList) {
				*pHPreviousImageList = pThis->wrappedImageLists.hOverlays;
			}
			pThis->wrappedImageLists.hOverlays = hImageList;
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

STDMETHODIMP AggregateImageList_GetIconInfo(IImageListPrivate* pInterface, UINT itemID, UINT mask, PINT pSystemOrNonShellIconIndex, PINT pExecutableIconIndex, LPWSTR* ppOverlayImageResource, PUINT pFlags)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
		if(iter != pThis->pIconInfos->cend()) {
			if(iter->second->flags & AII_NONSHELLITEM) {
				if((mask & SIIF_NONSHELLITEMICON) && pSystemOrNonShellIconIndex) {
					*pSystemOrNonShellIconIndex = iter->second->mainIconIndex;
				}
			} else {
				if((mask & SIIF_SYSTEMICON) && pSystemOrNonShellIconIndex) {
					*pSystemOrNonShellIconIndex = iter->second->systemIconIndex;
				}
			}
			if(mask & SIIF_OVERLAYIMAGE) {
				if(pExecutableIconIndex) {
					*pExecutableIconIndex = ((iter->second->flags & AII_FILETYPEOVERLAYMASK) == AII_EXECUTABLEICONOVERLAY ? iter->second->overlay.iconIndex : -1);
				} else if(ppOverlayImageResource) {
					*ppOverlayImageResource = ((iter->second->flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY ? iter->second->overlay.pImageResource : NULL);
				}
			}
			if(mask & SIIF_FLAGS && pFlags) {
				*pFlags = iter->second->flags;
			}
		} else {
			return E_INVALIDARG;
		}
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(itemID);
		if(pEntry) {
			if(pEntry->m_value->flags & AII_NONSHELLITEM) {
				if((mask & SIIF_NONSHELLITEMICON) && pSystemOrNonShellIconIndex) {
					*pSystemOrNonShellIconIndex = pEntry->m_value->mainIconIndex;
				}
			} else {
				if((mask & SIIF_SYSTEMICON) && pSystemOrNonShellIconIndex) {
					*pSystemOrNonShellIconIndex = pEntry->m_value->systemIconIndex;
				}
			}
			if(mask & SIIF_OVERLAYIMAGE) {
				if(pExecutableIconIndex) {
					*pExecutableIconIndex = ((pEntry->m_value->flags & AII_FILETYPEOVERLAYMASK) == AII_EXECUTABLEICONOVERLAY ? pEntry->m_value->overlay.iconIndex : -1);
				} else if(ppOverlayImageResource) {
					*ppOverlayImageResource = ((pEntry->m_value->flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY ? pEntry->m_value->overlay.pImageResource : NULL);
				}
			}
			if(mask & SIIF_FLAGS && pFlags) {
				*pFlags = pEntry->m_value->flags;
			}
		} else {
			return E_INVALIDARG;
		}
	#endif
	return S_OK;
}

STDMETHODIMP AggregateImageList_SetIconInfo(IImageListPrivate* pInterface, UINT itemID, UINT mask, HBITMAP hImage, int systemOrNonShellIconIndex, int executableIconIndex, LPWSTR pOverlayImageResource, UINT flags)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(pThis->destroyingIconInfo || !pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
		if(iter == pThis->pIconInfos->cend()) {
	#else
		CAtlMap<UINT, AggregatedIconInformation*>::CPair* pEntry = pThis->pIconInfos->Lookup(itemID);
		if(!pEntry) {
	#endif
		AggregatedIconInformation* pIconInfo = new AggregatedIconInformation;
		if(mask & SIIF_FLAGS) {
			// maybe we have to free pIconInfo->overlay.pImageResource
			if((pIconInfo->flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY && (flags & AII_FILETYPEOVERLAYMASK) != AII_IMAGERESOURCEOVERLAY) {
				// AII_IMAGERESOURCEOVERLAY is being removed
				if(pIconInfo->overlay.pImageResource) {
					LPWSTR p = pIconInfo->overlay.pImageResource;
					pIconInfo->overlay.pImageResource = NULL;
					CoTaskMemFree(p);
				}
			}
			pIconInfo->flags = flags;
		}
		if(mask & SIIF_IMAGE) {
			ATLASSERT(pThis->hasCreatedMainImageList || (pIconInfo->flags & AII_USELEGACYDISPLAYCODE));
			if(hImage) {
				if(pIconInfo->flags & AII_USELEGACYDISPLAYCODE) {
					pIconInfo->hBmp = hImage;
				} else {
					BITMAP bitmap = {0};
					GetObject(hImage, sizeof(BITMAP), &bitmap);
					EnterCriticalSection(&pThis->freeSlotsLock);
					pIconInfo->mainIconIndex = AggregateImageList_PopEmptySlot(pThis);
					LeaveCriticalSection(&pThis->freeSlotsLock);
					if(pIconInfo->mainIconIndex == -1) {
						ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Add(hImage, NULL, &pIconInfo->mainIconIndex)));
					} else {
						ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Replace(pIconInfo->mainIconIndex, hImage, NULL)));
					}
					if(pThis->wrappedImageLists.pThumbnails2) {
						pThis->wrappedImageLists.pThumbnails2->SetOriginalSize(pIconInfo->mainIconIndex, bitmap.bmWidth, bitmap.bmHeight);
					}
					DeleteObject(hImage);
				}
			}
		}
		if(pIconInfo->flags & AII_NONSHELLITEM) {
			if(mask & SIIF_NONSHELLITEMICON) {
				pIconInfo->mainIconIndex = systemOrNonShellIconIndex;
			}
		} else {
			if(mask & SIIF_SYSTEMICON) {
				pIconInfo->systemIconIndex = systemOrNonShellIconIndex;
			}
		}
		if(mask & SIIF_OVERLAYIMAGE) {
			switch(pIconInfo->flags & AII_FILETYPEOVERLAYMASK) {
				case AII_EXECUTABLEICONOVERLAY:
					pIconInfo->overlay.iconIndex = executableIconIndex;
					break;
				case AII_IMAGERESOURCEOVERLAY:
					if(pIconInfo->overlay.pImageResource != pOverlayImageResource) {
						LPWSTR p = pIconInfo->overlay.pImageResource;
						pIconInfo->overlay.pImageResource = pOverlayImageResource;
						CoTaskMemFree(p);
					}
					break;
			}
		}
		(*(pThis->pIconInfos))[itemID] = pIconInfo;
	} else {
		#ifdef USE_STL
			if(mask & SIIF_FLAGS) {
				ATLASSERT((iter->second->mainIconIndex == -1 && !iter->second->hBmp) || ((iter->second->flags & AII_USELEGACYDISPLAYCODE) == (flags & AII_USELEGACYDISPLAYCODE)));
				// maybe we have to free pIconInfo->overlay.pImageResource
				if((iter->second->flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY && (flags & AII_FILETYPEOVERLAYMASK) != AII_IMAGERESOURCEOVERLAY) {
					// AII_IMAGERESOURCEOVERLAY is being removed
					if(iter->second->overlay.pImageResource) {
						LPWSTR p = iter->second->overlay.pImageResource;
						iter->second->overlay.pImageResource = NULL;
						CoTaskMemFree(p);
					}
				}
				iter->second->flags = flags;
			}
			if(mask & SIIF_IMAGE) {
				ATLASSERT(pThis->hasCreatedMainImageList || (iter->second->flags & AII_USELEGACYDISPLAYCODE));
				if(hImage) {
					if(iter->second->flags & AII_USELEGACYDISPLAYCODE) {
						if(iter->second->hBmp) {
							DeleteObject(iter->second->hBmp);
						}
						iter->second->hBmp = hImage;
					} else {
						BITMAP bitmap = {0};
						GetObject(hImage, sizeof(BITMAP), &bitmap);
						if(iter->second->mainIconIndex == -1) {
							EnterCriticalSection(&pThis->freeSlotsLock);
							iter->second->mainIconIndex = AggregateImageList_PopEmptySlot(pThis);
							LeaveCriticalSection(&pThis->freeSlotsLock);
						}
						if(iter->second->mainIconIndex == -1) {
							ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Add(hImage, NULL, &iter->second->mainIconIndex)));
						} else {
							ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Replace(iter->second->mainIconIndex, hImage, NULL)));
						}
						if(pThis->wrappedImageLists.pThumbnails2) {
							pThis->wrappedImageLists.pThumbnails2->SetOriginalSize(iter->second->mainIconIndex, bitmap.bmWidth, bitmap.bmHeight);
						}
						DeleteObject(hImage);
					}
				} else {
					if(iter->second->hBmp) {
						DeleteObject(iter->second->hBmp);
						iter->second->hBmp = NULL;
					}
					iter->second->mainIconIndex = -1;
				}
			}
			if(iter->second->flags & AII_NONSHELLITEM) {
				if(mask & SIIF_NONSHELLITEMICON) {
					iter->second->mainIconIndex = systemOrNonShellIconIndex;
				}
			} else {
				if(mask & SIIF_SYSTEMICON) {
					iter->second->systemIconIndex = systemOrNonShellIconIndex;
				}
			}
			if(mask & SIIF_OVERLAYIMAGE) {
				switch(iter->second->flags & AII_FILETYPEOVERLAYMASK) {
					case AII_EXECUTABLEICONOVERLAY:
						iter->second->overlay.iconIndex = executableIconIndex;
						break;
					case AII_IMAGERESOURCEOVERLAY:
						if(iter->second->overlay.pImageResource != pOverlayImageResource) {
							LPWSTR p = iter->second->overlay.pImageResource;
							iter->second->overlay.pImageResource = pOverlayImageResource;
							CoTaskMemFree(p);
						}
						break;
				}
			}
		#else
			if(mask & SIIF_FLAGS) {
				ATLASSERT((pEntry->m_value->mainIconIndex == -1 && !pEntry->m_value->hBmp) || ((pEntry->m_value->flags & AII_USELEGACYDISPLAYCODE) == (flags & AII_USELEGACYDISPLAYCODE)));
				// maybe we have to free pIconInfo->overlay.pImageResource
				if((pEntry->m_value->flags & AII_FILETYPEOVERLAYMASK) == AII_IMAGERESOURCEOVERLAY && (flags & AII_FILETYPEOVERLAYMASK) != AII_IMAGERESOURCEOVERLAY) {
					// AII_IMAGERESOURCEOVERLAY is being removed
					if(pEntry->m_value->overlay.pImageResource) {
						LPWSTR p = pEntry->m_value->overlay.pImageResource;
						pEntry->m_value->overlay.pImageResource = NULL;
						CoTaskMemFree(p);
					}
				}
				pEntry->m_value->flags = flags;
			}
			if(mask & SIIF_IMAGE) {
				ATLASSERT(pThis->hasCreatedMainImageList || (pEntry->m_value->flags & AII_USELEGACYDISPLAYCODE));
				if(hImage) {
					if(pEntry->m_value->flags & AII_USELEGACYDISPLAYCODE) {
						if(pEntry->m_value->hBmp) {
							DeleteObject(pEntry->m_value->hBmp);
						}
						pEntry->m_value->hBmp = hImage;
					} else {
						BITMAP bitmap = {0};
						GetObject(hImage, sizeof(BITMAP), &bitmap);
						if(pEntry->m_value->mainIconIndex == -1) {
							EnterCriticalSection(&pThis->freeSlotsLock);
							pEntry->m_value->mainIconIndex = AggregateImageList_PopEmptySlot(pThis);
							LeaveCriticalSection(&pThis->freeSlotsLock);
						}
						if(pEntry->m_value->mainIconIndex == -1) {
							ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Add(hImage, NULL, &pEntry->m_value->mainIconIndex)));
						} else {
							ATLVERIFY(SUCCEEDED(pThis->wrappedImageLists.pThumbnails->Replace(pEntry->m_value->mainIconIndex, hImage, NULL)));
						}
						if(pThis->wrappedImageLists.pThumbnails2) {
							pThis->wrappedImageLists.pThumbnails2->SetOriginalSize(pEntry->m_value->mainIconIndex, bitmap.bmWidth, bitmap.bmHeight);
						}
						DeleteObject(hImage);
					}
				} else {
					if(pEntry->m_value->hBmp) {
						DeleteObject(pEntry->m_value->hBmp);
						pEntry->m_value->hBmp = NULL;
					}
					pEntry->m_value->mainIconIndex = -1;
				}
			}
			if(pEntry->m_value->flags & AII_NONSHELLITEM) {
				if(mask & SIIF_NONSHELLITEMICON) {
					pEntry->m_value->mainIconIndex = systemOrNonShellIconIndex;
				}
			} else {
				if(mask & SIIF_SYSTEMICON) {
					pEntry->m_value->systemIconIndex = systemOrNonShellIconIndex;
				}
			}
			if(mask & SIIF_OVERLAYIMAGE) {
				switch(pEntry->m_value->flags & AII_FILETYPEOVERLAYMASK) {
					case AII_EXECUTABLEICONOVERLAY:
						pEntry->m_value->overlay.iconIndex = executableIconIndex;
						break;
					case AII_IMAGERESOURCEOVERLAY:
						if(pEntry->m_value->overlay.pImageResource != pOverlayImageResource) {
							LPWSTR p = pEntry->m_value->overlay.pImageResource;
							pEntry->m_value->overlay.pImageResource = pOverlayImageResource;
							CoTaskMemFree(p);
						}
						break;
				}
			}
		#endif
	}
	return S_OK;
}

STDMETHODIMP AggregateImageList_RemoveIconInfo(IImageListPrivate* pInterface, UINT itemID)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	int mainIconIndex = -1;
	#ifdef USE_STL
		if(itemID == 0xFFFFFFFF) {
			std::unordered_map<UINT, AggregatedIconInformation*>::iterator iter = pThis->pIconInfos->begin();
			while(iter != pThis->pIconInfos->end()) {
				AggregatedIconInformation* pEntry = iter->second;
				iter = pThis->pIconInfos->erase(iter);
				delete pEntry;
			}
		} else {
			std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->find(itemID);
			if(iter != pThis->pIconInfos->cend()) {
				AggregatedIconInformation* pEntry = iter->second;
				if(!(pEntry->flags & AII_NONSHELLITEM)) {
					mainIconIndex = pEntry->mainIconIndex;
				}
				pThis->pIconInfos->erase(itemID);
				delete pEntry;
			}
		}
	#else
		if(itemID == 0xFFFFFFFF) {
			POSITION p = pThis->pIconInfos->GetStartPosition();
			while(p) {
				AggregatedIconInformation* pEntry = pThis->pIconInfos->GetValueAt(p);
				pThis->pIconInfos->RemoveAtPos(p);
				p = pThis->pIconInfos->GetStartPosition();
				delete pEntry;
			}
		} else {
			CAtlMap<UINT, AggregatedIconInformation*>::CPair* pIconInfo = pThis->pIconInfos->Lookup(itemID);
			if(pIconInfo) {
				AggregatedIconInformation* pEntry = pIconInfo->m_value;
				if(!(pEntry->flags & AII_NONSHELLITEM)) {
					mainIconIndex = pEntry->mainIconIndex;
				}
				pThis->pIconInfos->RemoveKey(itemID);
				delete pEntry;
			}
		}
	#endif
	if(itemID == 0xFFFFFFFF) {
		if(pThis->wrappedImageLists.pThumbnails) {
			/* TODO: What's the better solution here? Clearing the DPA saves memory, but marking all slots as
			 *       empty accelerates future seeks for empty slots. On the other hand DPA_DeleteAllPtrs is fast
			 *       and a loop of n calls of AggregateImageList_PushEmptySlot isn't.
			 *       The advantage of a fast clean-up is probably bigger than the advantage of a fast slot
			 *       seeking.
			 */
			pThis->wrappedImageLists.pThumbnails->Remove(-1);
			EnterCriticalSection(&pThis->freeSlotsLock);
			DPA_DeleteAllPtrs(pThis->hFreeSlots);
			LeaveCriticalSection(&pThis->freeSlotsLock);
			/*int numberOfImages = 0;
			pThis->wrappedImageLists.pThumbnails->GetImageCount(&numberOfImages);
			EnterCriticalSection(&pThis->freeSlotsLock);
			for(int i = 0; i < numberOfImages; i++) {
				AggregateImageList_PushEmptySlot(pThis, i);
			}
			LeaveCriticalSection(&pThis->freeSlotsLock);*/
		}
	} else if(mainIconIndex != -1) {
		EnterCriticalSection(&pThis->freeSlotsLock);
		AggregateImageList_PushEmptySlot(pThis, mainIconIndex);
		LeaveCriticalSection(&pThis->freeSlotsLock);
	}
	return S_OK;
}

STDMETHODIMP AggregateImageList_GetOptions(IImageListPrivate* pInterface, UINT mask, UINT* pFlags)
{
	if(!pFlags) {
		return E_POINTER;
	}

	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(mask) {
		*pFlags = pThis->flags & mask;
	} else {
		*pFlags = pThis->flags;
	}
	return S_OK;
}

STDMETHODIMP AggregateImageList_SetOptions(IImageListPrivate* pInterface, UINT mask, UINT flags)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(mask) {
		flags |= (pThis->flags & ~mask);
	}
	pThis->flags = flags;
	return S_OK;
}

STDMETHODIMP AggregateImageList_TransferNonShellItems(IImageListPrivate* pInterface, IImageListPrivate* pTarget)
{
	AggregateImageList* pThis = AggregateImageListFromIImageListPrivate(pInterface);
	if(!pThis->pIconInfos) {
		return E_FAIL;
	}

	#ifdef USE_STL
		for(std::unordered_map<UINT, AggregatedIconInformation*>::const_iterator iter = pThis->pIconInfos->cbegin(); iter != pThis->pIconInfos->cend(); iter++) {
			if(iter->second->flags & AII_NONSHELLITEM) {
				// TODO: The icon for selected state gets lost here.
				pTarget->SetIconInfo(iter->first, SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, iter->second->mainIconIndex, 0, NULL, iter->second->flags);
			}
		}
	#else
		POSITION p = pThis->pIconInfos->GetStartPosition();
		while(p) {
			AggregatedIconInformation* pIconInfo = pThis->pIconInfos->GetValueAt(p);
			if(pIconInfo->flags & AII_NONSHELLITEM) {
				// TODO: The icon for selected state gets lost here.
				pTarget->SetIconInfo(pThis->pIconInfos->GetKeyAt(p), SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, pIconInfo->mainIconIndex, 0, NULL, pIconInfo->flags);
			}
			pThis->pIconInfos->GetNext(p);
		}
	#endif
	return S_OK;
}
// implementation of IImageListPrivate
//////////////////////////////////////////////////////////////////////


STDMETHODIMP AggregateImageList_CreateInstance(IImageList** ppObject)
{
	if(!ppObject) {
		return E_POINTER;
	}

	AggregateImageList* pInstance = static_cast<AggregateImageList*>(HeapAlloc(GetProcessHeap(), 0, sizeof(AggregateImageList)));
	if(!pInstance) {
		return E_OUTOFMEMORY;
	}
	ZeroMemory(pInstance, sizeof(*pInstance));
	#ifdef USE_STL
		pInstance->pIconInfos = new std::unordered_map<UINT, AggregatedIconInformation*>();
	#else
		pInstance->pIconInfos = new CAtlMap<UINT, AggregatedIconInformation*>();
	#endif
	pInstance->magicValue = 0x4C4D4948;
	pInstance->pImageList2Vtable = &AggregateImageList_IImageList2Impl_Vtbl;
	pInstance->pImageListPrivateVtable = &AggregateImageList_IImageListPrivateImpl_Vtbl;
	InitializeCriticalSection(&pInstance->freeSlotsLock);
	pInstance->hFreeSlots = DPA_Create(30);
	pInstance->referenceCount = 1;
	APIWrapper::ImageList_CoCreateInstance(CLSID_ImageList, NULL, IID_PPV_ARGS(&pInstance->wrappedImageLists.pThumbnails2), NULL);
	if(pInstance->wrappedImageLists.pThumbnails2) {
		pInstance->wrappedImageLists.pThumbnails2->QueryInterface(IID_PPV_ARGS(&pInstance->wrappedImageLists.pThumbnails));
	}
	CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pInstance->pWICImagingFactory));

	*ppObject = IImageListFromAggregateImageList(pInstance);
	return S_OK;
}

PINT ImageList_GetOverlayData(HIMAGELIST hImageList)
{
	static DWORD offset = 0xFFFFFFFF;
	static HIMAGELIST hLastImageList = NULL;

	LPBYTE pInternalImageListStruct = reinterpret_cast<LPBYTE>(hImageList);
	if(APIWrapper::IsSupported_SHGetStockIconInfo()) {
		if(hImageList != hLastImageList) {
			hLastImageList = hImageList;
			offset = 0xFFFFFFFF;
			// NOTE: We assume that hImageList is a system image list!
			// Idea: Use SHGetStockIconInfo and SHGetIconOverlayIndex to find the three default overlays.
			SHSTOCKICONINFO stockIconInfo = {0};
			stockIconInfo.cbSize = sizeof(SHSTOCKICONINFO);
			stockIconInfo.iSysImageIndex = -1;
			APIWrapper::SHGetStockIconInfo(SIID_SHARE, SHGSI_SYSICONINDEX, &stockIconInfo, NULL);
			int shareIconIndex = stockIconInfo.iSysImageIndex;
			stockIconInfo.iSysImageIndex = -1;
			APIWrapper::SHGetStockIconInfo(SIID_LINK, SHGSI_SYSICONINDEX, &stockIconInfo, NULL);
			int linkIconIndex = stockIconInfo.iSysImageIndex;
			stockIconInfo.iSysImageIndex = -1;
			APIWrapper::SHGetStockIconInfo(SIID_SLOWFILE, SHGSI_SYSICONINDEX, &stockIconInfo, NULL);
			int slowFileIconIndex = stockIconInfo.iSysImageIndex;

			int shareOverlayIndex = max(SHGetIconOverlayIndex(NULL, IDO_SHGIOI_SHARE) - 1, 0);
			int linkOverlayIndex = max(SHGetIconOverlayIndex(NULL, IDO_SHGIOI_LINK) - 1, 0);
			int slowFileOverlayIndex = max(SHGetIconOverlayIndex(NULL, IDO_SHGIOI_SLOWFILE) - 1, 0);
			DWORD o = 0;
			do {
				if(!memcmp(pInternalImageListStruct + o + shareOverlayIndex * sizeof(int), &shareIconIndex, sizeof(int)) && !memcmp(pInternalImageListStruct + o + linkOverlayIndex * sizeof(int), &linkIconIndex, sizeof(int)) && !memcmp(pInternalImageListStruct + o + slowFileOverlayIndex * sizeof(int), &slowFileIconIndex, sizeof(int))) {
					offset = o;
					break;
				}
			} while(o++ < 250);
		}
	}

	if(offset == 0xFFFFFFFF) {
		// we're either on XP or older or we're not dealing with a system image list
		/* Idea: Create an image list, add five icons and define these icons as the first five overlays. Then
		 *       scan the memory for the sequence 0x0000 0x0001 0x0002 0x0003 0x0004.
		 */
		CDCHandle compatibleDC = GetDC(NULL);
		CBitmap dummyIcon;
		dummyIcon.CreateCompatibleBitmap(compatibleDC, 16, 16);
		ReleaseDC(NULL, compatibleDC);

		HIMAGELIST h = ImageList_Create(16, 16, ILC_COLOR24, 5, 0);
		ATLASSERT(h);
		if(h) {
			// add the icons
			for(int i = 0; i < 5; i++) {
				ATLVERIFY(ImageList_Add(h, dummyIcon, NULL) == i);
				ATLVERIFY(ImageList_SetOverlayImage(h, i, i + 1));
			}
			int p[5] = {0, 1, 2, 3, 4};

			// scan for the overlays
			LPBYTE pInternalStruct = reinterpret_cast<LPBYTE>(h);
			DWORD o = 0;
			do {
				if(!memcmp(pInternalStruct + o, p, 5 * sizeof(int))) {
					offset = o;
					break;
				}
			} while(o++ < 250);

			ImageList_Destroy(h);
		}
	}
	ATLASSERT(offset != 0xFFFFFFFF);
	if(offset != 0xFFFFFFFF) {
		return reinterpret_cast<PINT>(pInternalImageListStruct + offset);
	}
	return NULL;
}

BOOL ImageList_DrawIndirect_HQScaling(AggregateImageList* pInstance, IMAGELISTDRAWPARAMS* pParams, UINT flags, BOOL forceSize)
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
							//if((params.fState & ILS_ALPHA) && AggregateImageList_IsComctl32Version600()) {
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
		}
		DestroyIcon(hIconToScale);
	}

	return b;
}


int AggregateImageList_PopEmptySlot(AggregateImageList* pInstance)
{
	if(DPA_GetPtrCount(pInstance->hFreeSlots) > 0) {
		return PtrToInt(DPA_DeletePtr(pInstance->hFreeSlots, 0));
	}
	return -1;
}

void AggregateImageList_PushEmptySlot(AggregateImageList* pInstance, int slot)
{
	// ensure that the slot is not already in the list
	for(int i = 0; i < DPA_GetPtrCount(pInstance->hFreeSlots); i++) {
		if(PtrToInt(DPA_GetPtr(pInstance->hFreeSlots, i)) == slot) {
			return;
		}
	}
	DPA_AppendPtr(pInstance->hFreeSlots, IntToPtr(slot));
}


BOOL AggregateImageList_IsComctl32Version600()
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hRet = ATL::AtlGetCommCtrlVersion(&major, &minor);
	return (SUCCEEDED(hRet) && (major == 6) && (minor == 0));
}