// DropShadowAdorner.cpp: An implementation of IThumbnailAdorner for drawing drop shadows around thumbnails

#include "stdafx.h"
#include "DropShadowAdorner.h"


DropShadowAdorner::DropShadowAdorner()
{
	hAdornmentIcon = NULL;
}

DropShadowAdorner::~DropShadowAdorner()
{
	if(hAdornmentIcon) {
		DestroyIcon(hAdornmentIcon);
		hAdornmentIcon = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IThumbnailAdorner
STDMETHODIMP DropShadowAdorner::SetIconSize(SIZE& targetIconSize, double /*contentAspectRatio*/)
{
	this->targetIconSize = targetIconSize;

	HMODULE hImageRes = LoadLibraryEx(TEXT("imageres.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hImageRes) {
		int sz = max(targetIconSize.cx, targetIconSize.cy);
		UINT iconToLoad = FindBestMatchingIconResource(hImageRes, MAKEINTRESOURCE(191), sz, &loadedIconSize);
		if(iconToLoad != 0xFFFFFFFF) {
			extraMargins.left = 5;
			extraMargins.right = 5;
			extraMargins.top = 5;
			extraMargins.bottom = 5;
			if(sz <= 96) {
				extraMargins.left = 2;
				extraMargins.right = 2;
				extraMargins.top = 2;
				extraMargins.bottom = 2;
			}
			if(sz <= 48) {
				extraMargins.left = 1;
				extraMargins.right = 1;
				extraMargins.top = 1;
				extraMargins.bottom = 1;
			}

			if(hAdornmentIcon) {
				DestroyIcon(hAdornmentIcon);
				hAdornmentIcon = NULL;
			}
			HRSRC hResource = FindResource(hImageRes, MAKEINTRESOURCE(iconToLoad), RT_ICON);
			if(hResource) {
				HGLOBAL hMem = LoadResource(hImageRes, hResource);
				if(hMem) {
					LPVOID pIconData = LockResource(hMem);
					if(pIconData) {
						hAdornmentIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hImageRes, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
					}
				}
			}
		}
		FreeLibrary(hImageRes);
	}

	if(hAdornmentIcon && FindContentAreaMargins(&contentMargins)) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DropShadowAdorner::GetMaxContentSize(double contentAspectRatio, LPSIZE pContentSize)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	SIZE contentAreaSize;
	contentAreaSize.cx = targetIconSize.cx - contentMargins.left - contentMargins.right;
	contentAreaSize.cy = targetIconSize.cy - contentMargins.top - contentMargins.bottom;

	if(contentAspectRatio > 1.0) {
		// the image's height is larger than its width
		pContentSize->cy = contentAreaSize.cy;
		pContentSize->cx = static_cast<LONG>(static_cast<double>(pContentSize->cy) / contentAspectRatio);
	} else {
		// the image's width is larger than its height
		pContentSize->cx = contentAreaSize.cx;
		pContentSize->cy = static_cast<LONG>(contentAspectRatio * pContentSize->cx);
	}
	return S_OK;
}

STDMETHODIMP DropShadowAdorner::GetAdornedSize(SIZE& contentSize, LPSIZE pAdornedSize)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	pAdornedSize->cx = contentSize.cx + contentMargins.left + contentMargins.right;
	pAdornedSize->cy = contentSize.cy + contentMargins.top + contentMargins.bottom;
	return S_OK;
}

STDMETHODIMP DropShadowAdorner::SetAdornedSize(SIZE& /*adornedSize*/)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP DropShadowAdorner::GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, LPRECT pContentRectangle, LPRECT pContentAreaRectangle, LPRECT pContentToDrawRectangle)
{
	CRect marginedContentRectangle(adornedRectangle.left + contentMargins.left, adornedRectangle.top + contentMargins.top, adornedRectangle.right - contentMargins.right, adornedRectangle.bottom - contentMargins.bottom);

	// center the content within the adorner
	pContentRectangle->left = marginedContentRectangle.left + ((marginedContentRectangle.Width() - contentSize.cx) >> 1);
	pContentRectangle->right = pContentRectangle->left + contentSize.cx;
	pContentRectangle->top = marginedContentRectangle.top + ((marginedContentRectangle.Height() - contentSize.cy) >> 1);
	pContentRectangle->bottom = pContentRectangle->top + contentSize.cy;

	pContentToDrawRectangle->left = max(marginedContentRectangle.left, pContentRectangle->left);
	pContentToDrawRectangle->right = min(marginedContentRectangle.right, pContentRectangle->right);
	pContentToDrawRectangle->top = max(marginedContentRectangle.top, pContentRectangle->top);
	pContentToDrawRectangle->bottom = min(marginedContentRectangle.bottom, pContentRectangle->bottom);

	*pContentAreaRectangle = marginedContentRectangle;
	return S_OK;
}

STDMETHODIMP DropShadowAdorner::GetDrawOrder(LPBOOL pAdornmentFirst)
{
	*pAdornmentFirst = TRUE;
	return S_OK;
}

STDMETHODIMP DropShadowAdorner::DrawIntoDIB(LPRECT pBoundingRectangle, LPRGBQUAD pBits)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	SIZE boundingSize = CRect(pBoundingRectangle).Size();

	CComPtr<IWICImagingFactory> pWICImagingFactory = NULL;
	hr = CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
	if(pWICImagingFactory) {
		CComPtr<IWICBitmap> pWICBitmap = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hAdornmentIcon, &pWICBitmap);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pWICBitmap);

			// top-left corner
			WICRect rectangleToCopy = {0};
			rectangleToCopy.Width = contentMargins.left + extraMargins.left;
			rectangleToCopy.Height = contentMargins.top + extraMargins.top;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), boundingSize.cx * boundingSize.cy * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits))));
			// top edge
			rectangleToCopy.X = contentMargins.left + extraMargins.left;
			rectangleToCopy.Width = boundingSize.cx - rectangleToCopy.X - (contentMargins.right + extraMargins.right);
			int offset = rectangleToCopy.X;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// top-right corner
			rectangleToCopy.X = loadedIconSize - 1 - (contentMargins.right + extraMargins.right);
			rectangleToCopy.Width = contentMargins.right + extraMargins.right;
			offset = boundingSize.cx - 1 - rectangleToCopy.Width;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// left edge
			rectangleToCopy.X = 0;
			rectangleToCopy.Y = contentMargins.top + extraMargins.top;
			rectangleToCopy.Width = contentMargins.left + extraMargins.left;
			rectangleToCopy.Height = boundingSize.cy - rectangleToCopy.Y - (contentMargins.bottom + extraMargins.bottom);
			offset = rectangleToCopy.Y * boundingSize.cx;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// right edge
			rectangleToCopy.X = loadedIconSize - 1 - (contentMargins.right + extraMargins.right);
			rectangleToCopy.Width = contentMargins.right + extraMargins.right;
			offset += boundingSize.cx - 1 - rectangleToCopy.Width;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom-left corner
			rectangleToCopy.X = 0;
			rectangleToCopy.Y = loadedIconSize - 1 - (contentMargins.bottom + extraMargins.bottom);
			rectangleToCopy.Width = contentMargins.left + extraMargins.left;
			rectangleToCopy.Height = contentMargins.bottom + extraMargins.bottom;
			offset = (boundingSize.cy - 1 - (contentMargins.bottom + extraMargins.bottom)) * boundingSize.cx;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom edge
			rectangleToCopy.X = contentMargins.left + extraMargins.left;
			rectangleToCopy.Width = boundingSize.cx - rectangleToCopy.X - (contentMargins.right + extraMargins.right);
			offset += rectangleToCopy.X;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom-right corner
			rectangleToCopy.X = loadedIconSize - 1 - (contentMargins.right + extraMargins.right);
			rectangleToCopy.Width = contentMargins.right + extraMargins.right;
			offset = (boundingSize.cy - (contentMargins.bottom + extraMargins.bottom)) * boundingSize.cx - 1 - rectangleToCopy.Width;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
		}
	}

	return hr;
}

STDMETHODIMP DropShadowAdorner::PostProcess(HBITMAP /*hSourceBitmap*/, LPRGBQUAD /*pSourceBits*//* = NULL*/, HBITMAP* /*pHCroppedBitmap*//* = NULL*/)
{
	return E_NOTIMPL;
}
// implementation of IThumbnailAdorner
//////////////////////////////////////////////////////////////////////


BOOL DropShadowAdorner::FindContentAreaMargins(LPRECT pMargins)
{
	BOOL foundIt = FALSE;

	ICONINFO iconInfo = {0};
	GetIconInfo(hAdornmentIcon, &iconInfo);
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
		LPDWORD pBits = static_cast<LPDWORD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
		if(pBits) {
			GetDIBits(memoryDC, iconInfo.hbmColor, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS);

			// search top-left corner
			POINT pt = {bmp.bmWidth >> 1, bmp.bmHeight >> 1};
			pMargins->left = 0;
			for(pt.x = min(20, bmp.bmWidth >> 1); pt.x > 0 && !foundIt; pt.x--) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->left = pt.x + 1;
				}
			}
			foundIt = FALSE;
			pt.x = bmp.bmWidth >> 1;
			pMargins->top = 0;
			for(pt.y = min(20, bmp.bmHeight >> 1); pt.y > 0 && !foundIt; pt.y--) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->top = pt.y + 1;
				}
			}
			// search bottom-right corner
			foundIt = FALSE;
			pt.y = bmp.bmHeight >> 1;
			pMargins->right = 0;
			for(pt.x = min(bmp.bmWidth - 20, bmp.bmWidth >> 1); pt.x < bmp.bmWidth && !foundIt; pt.x++) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->right = bmp.bmWidth - pt.x;
				}
			}
			foundIt = FALSE;
			pt.x = bmp.bmWidth >> 1;
			pMargins->bottom = 0;
			for(pt.y = min(bmp.bmHeight - 20, bmp.bmHeight >> 1); pt.y < bmp.bmHeight && !foundIt; pt.y++) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->bottom = bmp.bmHeight - pt.y;
				}
			}
			HeapFree(GetProcessHeap(), 0, pBits);
			foundIt = TRUE;
		}
		DeleteObject(iconInfo.hbmColor);
	}
	return foundIt;
}