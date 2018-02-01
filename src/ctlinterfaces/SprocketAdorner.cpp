// SprocketAdorner.cpp: An implementation of IThumbnailAdorner for drawing sprockets around thumbnails

#include "stdafx.h"
#include "SprocketAdorner.h"


SprocketAdorner::SprocketAdorner()
{
	properties.hAdornmentIcon = NULL;
	properties.pScaledAdornmentIconBits = NULL;
}

SprocketAdorner::~SprocketAdorner()
{
	if(properties.hAdornmentIcon) {
		DestroyIcon(properties.hAdornmentIcon);
		properties.hAdornmentIcon = NULL;
	}
	if(properties.pScaledAdornmentIconBits) {
		HeapFree(GetProcessHeap(), 0, properties.pScaledAdornmentIconBits);
		properties.pScaledAdornmentIconBits = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IThumbnailAdorner
STDMETHODIMP SprocketAdorner::SetIconSize(SIZE& targetIconSize, double contentAspectRatio)
{
	this->properties.targetIconSize = targetIconSize;
	BOOL wideScreen = (contentAspectRatio < 0.6);

	HMODULE hImageRes = LoadLibraryEx(TEXT("imageres.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hImageRes) {
		UINT iconToLoad = FindBestMatchingIconResource(hImageRes, MAKEINTRESOURCE(wideScreen ? 194 : 193), max(targetIconSize.cx, targetIconSize.cy), &properties.loadedIconSize);
		if(iconToLoad != 0xFFFFFFFF) {
			properties.extraMargins.left = -5;
			properties.extraMargins.right = -5;
			properties.extraMargins.top = -3;
			properties.extraMargins.bottom = -3;

			if(properties.hAdornmentIcon) {
				DestroyIcon(properties.hAdornmentIcon);
				properties.hAdornmentIcon = NULL;
			}
			HRSRC hResource = FindResource(hImageRes, MAKEINTRESOURCE(iconToLoad), RT_ICON);
			if(hResource) {
				HGLOBAL hMem = LoadResource(hImageRes, hResource);
				if(hMem) {
					LPVOID pIconData = LockResource(hMem);
					if(pIconData) {
						properties.hAdornmentIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hImageRes, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
					}
				}
			}
		}
		FreeLibrary(hImageRes);
	}

	if(properties.hAdornmentIcon && FindContentAreaMargins(&properties.contentMargins)) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP SprocketAdorner::GetMaxContentSize(double contentAspectRatio, LPSIZE pContentSize)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	SIZE contentAreaSize;
	contentAreaSize.cx = properties.targetIconSize.cx - properties.contentMargins.left - properties.contentMargins.right;
	contentAreaSize.cy = properties.targetIconSize.cy - properties.contentMargins.top - properties.contentMargins.bottom;
	contentAreaSize.cx -= (properties.extraMargins.left + properties.extraMargins.right);
	contentAreaSize.cy -= (properties.extraMargins.top + properties.extraMargins.bottom);

	if(contentAspectRatio > 1.0) {
		// the image's height is larger than its width
		pContentSize->cx = contentAreaSize.cx;
		pContentSize->cy = static_cast<LONG>(contentAspectRatio * pContentSize->cx);
	} else {
		// the image's width is larger than its height
		pContentSize->cy = contentAreaSize.cy;
		pContentSize->cx = static_cast<LONG>(static_cast<double>(pContentSize->cy) / contentAspectRatio);
	}
	return S_OK;
}

STDMETHODIMP SprocketAdorner::GetAdornedSize(SIZE& contentSize, LPSIZE pAdornedSize)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	pAdornedSize->cx = contentSize.cx + properties.contentMargins.left + properties.contentMargins.right;
	pAdornedSize->cy = contentSize.cy + properties.contentMargins.top + properties.contentMargins.bottom;
	pAdornedSize->cx += (properties.extraMargins.left + properties.extraMargins.right);
	pAdornedSize->cy += (properties.extraMargins.top + properties.extraMargins.bottom);
	return S_OK;
}

STDMETHODIMP SprocketAdorner::SetAdornedSize(SIZE& adornedSize)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	if(properties.pScaledAdornmentIconBits) {
		HeapFree(GetProcessHeap(), 0, properties.pScaledAdornmentIconBits);
		properties.pScaledAdornmentIconBits = NULL;
	}

	properties.scaledAdornmentIconSize = adornedSize;

	CComPtr<IWICImagingFactory> pWICImagingFactory = NULL;
	HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
	if(pWICImagingFactory) {
		CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pWICBitmapScaler);
			CComPtr<IWICBitmap> pWICBitmap = NULL;
			hr = pWICImagingFactory->CreateBitmapFromHICON(properties.hAdornmentIcon, &pWICBitmap);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pWICBitmap);
				hr = pWICBitmapScaler->Initialize(pWICBitmap, properties.scaledAdornmentIconSize.cx, properties.scaledAdornmentIconSize.cy, WICBitmapInterpolationModeFant);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					properties.pScaledAdornmentIconBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, properties.scaledAdornmentIconSize.cx * properties.scaledAdornmentIconSize.cy * sizeof(RGBQUAD)));
					if(properties.pScaledAdornmentIconBits) {
						hr = pWICBitmapScaler->CopyPixels(NULL, properties.scaledAdornmentIconSize.cx * sizeof(RGBQUAD), properties.scaledAdornmentIconSize.cy * properties.scaledAdornmentIconSize.cx * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(properties.pScaledAdornmentIconBits));
						if(SUCCEEDED(hr)) {
							FindContentAreaMargins(&properties.scaledContentMargins);
						}
					}
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP SprocketAdorner::GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, LPRECT pContentRectangle, LPRECT pContentAreaRectangle, LPRECT pContentToDrawRectangle)
{
	RECT margins = (properties.pScaledAdornmentIconBits ? properties.scaledContentMargins : properties.contentMargins);

	CRect marginedContentRectangle(adornedRectangle.left + margins.left, adornedRectangle.top + margins.top, adornedRectangle.right - margins.right, adornedRectangle.bottom - margins.bottom);
	*pContentAreaRectangle = marginedContentRectangle;
	marginedContentRectangle.left += properties.extraMargins.left;
	marginedContentRectangle.top += properties.extraMargins.top;
	marginedContentRectangle.right -= properties.extraMargins.right;
	marginedContentRectangle.bottom -= properties.extraMargins.bottom;

	double aspectRatio = static_cast<double>(contentSize.cy) / static_cast<double>(contentSize.cx);
	int debtY = marginedContentRectangle.Height() - contentSize.cy;
	int debtX = static_cast<int>(static_cast<double>(debtY) / aspectRatio);

	// center the content within the adorner
	pContentRectangle->left = marginedContentRectangle.left + ((marginedContentRectangle.Width() - (contentSize.cx + debtX)) >> 1);
	pContentRectangle->right = pContentRectangle->left + contentSize.cx + debtX;
	pContentRectangle->top = marginedContentRectangle.top + ((marginedContentRectangle.Height() - (contentSize.cy + debtY)) >> 1);
	pContentRectangle->bottom = pContentRectangle->top + contentSize.cy + debtY;

	pContentToDrawRectangle->left = max(marginedContentRectangle.left, pContentRectangle->left);
	pContentToDrawRectangle->right = min(marginedContentRectangle.right, pContentRectangle->right);
	pContentToDrawRectangle->top = max(marginedContentRectangle.top, pContentRectangle->top);
	pContentToDrawRectangle->bottom = min(marginedContentRectangle.bottom, pContentRectangle->bottom);
	return S_OK;
}

STDMETHODIMP SprocketAdorner::GetDrawOrder(LPBOOL pAdornmentFirst)
{
	*pAdornmentFirst = TRUE;
	return S_OK;
}

STDMETHODIMP SprocketAdorner::DrawIntoDIB(LPRECT /*pBoundingRectangle*/, LPRGBQUAD pBits)
{
	if(!properties.hAdornmentIcon || !properties.pScaledAdornmentIconBits) {
		return E_FAIL;
	}

	CopyMemory(pBits, properties.pScaledAdornmentIconBits, properties.scaledAdornmentIconSize.cx * properties.scaledAdornmentIconSize.cy * sizeof(RGBQUAD));
	return S_OK;
}

STDMETHODIMP SprocketAdorner::PostProcess(HBITMAP hSourceBitmap, LPRGBQUAD pSourceBits/* = NULL*/, HBITMAP* pHCroppedBitmap/* = NULL*/)
{
	*pHCroppedBitmap = hSourceBitmap;

	BITMAP bmp = {0};
	GetObject(hSourceBitmap, sizeof(BITMAP), &bmp);
	SIZE bitmapSize = {bmp.bmWidth, bmp.bmHeight};
	ATLASSERT(bitmapSize.cy >= 0);

	BOOL freeMem = FALSE;
	if(!pSourceBits) {
		CDC memoryDC;
		memoryDC.CreateCompatibleDC(NULL);

		BITMAPINFO bitmapInfo;
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = bitmapSize.cx;
		bitmapInfo.bmiHeader.biHeight = -bitmapSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = 0;
		bitmapInfo.bmiHeader.biSizeImage = bitmapSize.cy * bitmapSize.cx * 4;
		bitmapInfo.bmiHeader.biXPelsPerMeter = 0;
		bitmapInfo.bmiHeader.biYPelsPerMeter = 0;
		bitmapInfo.bmiHeader.biClrUsed = 0;
		bitmapInfo.bmiHeader.biClrImportant = 0;

		pSourceBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
		if(pSourceBits) {
			GetDIBits(memoryDC, hSourceBitmap, 0, bitmapSize.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
		}
		freeMem = TRUE;
	}

	RECT adornmentMargins = {0};
	if(FindAdornmentMargins(bitmapSize, pSourceBits, &adornmentMargins)) {
		SIZE croppedSize = bitmapSize;
		croppedSize.cx -= adornmentMargins.left + adornmentMargins.right;
		croppedSize.cy -= adornmentMargins.top + adornmentMargins.bottom;

		BITMAPINFO bitmapInfo = {0};
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = croppedSize.cx;
		bitmapInfo.bmiHeader.biHeight = -croppedSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
		LPRGBQUAD pCroppedBits = NULL;
		HBITMAP hCroppedBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pCroppedBits), NULL, 0);
		if(hCroppedBitmap) {
			ATLASSERT_POINTER(pCroppedBits, RGBQUAD);
			if(pCroppedBits) {
				#pragma omp parallel for
				for(int y = 0; y < croppedSize.cy; y++) {
					UINT offset1 = (y + adornmentMargins.top) * bitmapSize.cx + adornmentMargins.left;
					UINT offset2 = y * croppedSize.cx;
					CopyMemory(&pCroppedBits[offset2], &pSourceBits[offset1], croppedSize.cx * sizeof(RGBQUAD));
				}

				*pHCroppedBitmap = hCroppedBitmap;
				DeleteObject(hSourceBitmap);
			}
		}
	}

	if(freeMem) {
		HeapFree(GetProcessHeap(), 0, pSourceBits);
	}
	return S_OK;
}
// implementation of IThumbnailAdorner
//////////////////////////////////////////////////////////////////////


BOOL SprocketAdorner::FindAdornmentMargins(SIZE bitmapSize, LPRGBQUAD pBits, LPRECT pMargins)
{
	LPDWORD p = reinterpret_cast<LPDWORD>(pBits);
	POINT pt;
	// start with the left and right margins
	for(pt.x = 0; pt.x < bitmapSize.cx; pt.x++) {
		BOOL columnIsEmpty = TRUE;
		for(pt.y = 0; pt.y < bitmapSize.cy && columnIsEmpty; pt.y++) {
			columnIsEmpty = (p[pt.y * bitmapSize.cx + pt.x] == 0x00000000);
		}
		if(!columnIsEmpty) {
			pMargins->left = pt.x;
			break;
		}
	}
	for(pt.x = bitmapSize.cx - 1; pt.x >= 0; pt.x--) {
		BOOL columnIsEmpty = TRUE;
		for(pt.y = 0; pt.y < bitmapSize.cy && columnIsEmpty; pt.y++) {
			columnIsEmpty = (p[pt.y * bitmapSize.cx + pt.x] == 0x00000000);
		}
		if(!columnIsEmpty) {
			pMargins->right = bitmapSize.cx - 1 - pt.x;
			break;
		}
	}

	// top and bottom margins
	for(pt.y = 0; pt.y < bitmapSize.cy; pt.y++) {
		BOOL lineIsEmpty = TRUE;
		for(pt.x = pMargins->left; pt.x < (bitmapSize.cx - pMargins->right) && lineIsEmpty; pt.x++) {
			lineIsEmpty = (p[pt.y * bitmapSize.cx + pt.x] == 0x00000000);
		}
		if(!lineIsEmpty) {
			pMargins->top = pt.y;
			break;
		}
	}
	for(pt.y = bitmapSize.cy - 1; pt.y >= 0; pt.y--) {
		BOOL lineIsEmpty = TRUE;
		for(pt.x = pMargins->left; pt.x < (bitmapSize.cx - pMargins->right) && lineIsEmpty; pt.x++) {
			lineIsEmpty = (p[pt.y * bitmapSize.cx + pt.x] == 0x00000000);
		}
		if(!lineIsEmpty) {
			pMargins->bottom = bitmapSize.cy - 1 - pt.y;
			break;
		}
	}

	return TRUE;
}

BOOL SprocketAdorner::FindContentAreaMargins(LPRECT pMargins)
{
	BOOL foundIt = FALSE;

	BOOL freeMem = FALSE;
	LPDWORD pBits = NULL;
	SIZE bitmapSize = {0};

	if(properties.pScaledAdornmentIconBits) {
		pBits = reinterpret_cast<LPDWORD>(properties.pScaledAdornmentIconBits);
		bitmapSize = properties.scaledAdornmentIconSize;
	} else {
		ICONINFO iconInfo = {0};
		GetIconInfo(properties.hAdornmentIcon, &iconInfo);
		if(iconInfo.hbmMask) {
			DeleteObject(iconInfo.hbmMask);
		}
		if(iconInfo.hbmColor) {
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(NULL);
			BITMAP bmp = {0};
			GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bmp);
			ATLASSERT(bmp.bmHeight >= 0);

			BITMAPINFO bitmapInfo;
			bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			bitmapInfo.bmiHeader.biWidth = bmp.bmWidth;
			bitmapInfo.bmiHeader.biHeight = -bmp.bmHeight;
			bitmapInfo.bmiHeader.biPlanes = 1;
			bitmapInfo.bmiHeader.biBitCount = 32;
			bitmapInfo.bmiHeader.biCompression = 0;
			bitmapInfo.bmiHeader.biSizeImage = bmp.bmHeight * bmp.bmWidth * 4;
			bitmapInfo.bmiHeader.biXPelsPerMeter = 0;
			bitmapInfo.bmiHeader.biYPelsPerMeter = 0;
			bitmapInfo.bmiHeader.biClrUsed = 0;
			bitmapInfo.bmiHeader.biClrImportant = 0;

			pBits = static_cast<LPDWORD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
			if(pBits) {
				GetDIBits(memoryDC, iconInfo.hbmColor, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS);
			}
			freeMem = TRUE;
			bitmapSize.cx = bmp.bmWidth;
			bitmapSize.cy = bmp.bmHeight;

			DeleteObject(iconInfo.hbmColor);
		}
	}

	if(pBits) {
		int yMaxTop = bitmapSize.cy >> 1;
		int yMinBottom = (bitmapSize.cy >> 1) - 2;
		int xMaxLeft = bitmapSize.cx >> 1;
		int xMinRight = bitmapSize.cx >> 1;

		// search top-left corner
		POINT pt = {0};
		BOOL hasFoundColoredLine = FALSE;
		for(pt.y = 0; pt.y < yMaxTop && !foundIt; pt.y++) {
			// the line must "end" with 0x00000000 and it mustn't be the first line of this kind
			if(pBits[pt.y * bitmapSize.cx + xMaxLeft - 1] == 0x00000000) {
				if(hasFoundColoredLine) {
					foundIt = TRUE;
					for(pt.x = xMaxLeft - 1; pt.x >= 0; pt.x--) {
						if(pBits[pt.y * bitmapSize.cx + pt.x] == 0x00000000) {
							pMargins->left = pt.x;
							pMargins->top = pt.y;
						} else {
							break;
						}
					}
				}
			} else {
				hasFoundColoredLine = TRUE;
			}
		}
		if(foundIt) {
			// search bottom-right corner
			foundIt = FALSE;
			hasFoundColoredLine = FALSE;
			for(pt.y = bitmapSize.cy - 1; pt.y >= yMinBottom && !foundIt; pt.y--) {
				// the line must "end" with 0x00000000 and it mustn't be the first line of this kind
				if(pBits[pt.y * bitmapSize.cx + xMinRight] == 0x00000000) {
					if(hasFoundColoredLine) {
						foundIt = TRUE;
						pMargins->bottom = bitmapSize.cy - 1 - pt.y;
						for(pt.x = xMinRight; pt.x < bitmapSize.cx - 1; pt.x++) {
							if(pBits[pt.y * bitmapSize.cx + pt.x] == 0x00000000) {
								pMargins->right = bitmapSize.cx - 1 - pt.x;
							} else {
								break;
							}
						}
					}
				} else {
					hasFoundColoredLine = TRUE;
				}
			}
		}
	}

	if(freeMem) {
		HeapFree(GetProcessHeap(), 0, pBits);
	}
	return foundIt;
}