// DropShadowAdorner.cpp: An implementation of IThumbnailAdorner for drawing drop shadows around thumbnails

#include "stdafx.h"
#include "DropShadowAdorner.h"


DropShadowAdorner::DropShadowAdorner()
{
	properties.hAdornmentIcon = NULL;
}

DropShadowAdorner::~DropShadowAdorner()
{
	if(properties.hAdornmentIcon) {
		DestroyIcon(properties.hAdornmentIcon);
		properties.hAdornmentIcon = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IThumbnailAdorner
STDMETHODIMP DropShadowAdorner::SetIconSize(SIZE& targetIconSize, double /*contentAspectRatio*/)
{
	this->properties.targetIconSize = targetIconSize;

	HMODULE hImageRes = LoadLibraryEx(TEXT("imageres.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hImageRes) {
		int sz = max(targetIconSize.cx, targetIconSize.cy);
		UINT iconToLoad = FindBestMatchingIconResource(hImageRes, MAKEINTRESOURCE(191), sz, &properties.loadedIconSize);
		if(iconToLoad != 0xFFFFFFFF) {
			properties.extraMargins.left = 5;
			properties.extraMargins.right = 5;
			properties.extraMargins.top = 5;
			properties.extraMargins.bottom = 5;
			if(sz <= 96) {
				properties.extraMargins.left = 2;
				properties.extraMargins.right = 2;
				properties.extraMargins.top = 2;
				properties.extraMargins.bottom = 2;
			}
			if(sz <= 48) {
				properties.extraMargins.left = 1;
				properties.extraMargins.right = 1;
				properties.extraMargins.top = 1;
				properties.extraMargins.bottom = 1;
			}

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

STDMETHODIMP DropShadowAdorner::GetMaxContentSize(double contentAspectRatio, LPSIZE pContentSize)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	SIZE contentAreaSize;
	contentAreaSize.cx = properties.targetIconSize.cx - properties.contentMargins.left - properties.contentMargins.right;
	contentAreaSize.cy = properties.targetIconSize.cy - properties.contentMargins.top - properties.contentMargins.bottom;

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
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	pAdornedSize->cx = contentSize.cx + properties.contentMargins.left + properties.contentMargins.right;
	pAdornedSize->cy = contentSize.cy + properties.contentMargins.top + properties.contentMargins.bottom;
	return S_OK;
}

STDMETHODIMP DropShadowAdorner::SetAdornedSize(SIZE& /*adornedSize*/)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP DropShadowAdorner::GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, LPRECT pContentRectangle, LPRECT pContentAreaRectangle, LPRECT pContentToDrawRectangle)
{
	CRect marginedContentRectangle(adornedRectangle.left + properties.contentMargins.left, adornedRectangle.top + properties.contentMargins.top, adornedRectangle.right - properties.contentMargins.right, adornedRectangle.bottom - properties.contentMargins.bottom);

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
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	SIZE boundingSize = CRect(pBoundingRectangle).Size();

	CComPtr<IWICImagingFactory> pWICImagingFactory = NULL;
	hr = CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
	if(pWICImagingFactory) {
		CComPtr<IWICBitmap> pWICBitmap = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(properties.hAdornmentIcon, &pWICBitmap);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pWICBitmap);

			// top-left corner
			WICRect rectangleToCopy = {0};
			rectangleToCopy.Width = properties.contentMargins.left + properties.extraMargins.left;
			rectangleToCopy.Height = properties.contentMargins.top + properties.extraMargins.top;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), boundingSize.cx * boundingSize.cy * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits))));
			// top edge
			rectangleToCopy.X = properties.contentMargins.left + properties.extraMargins.left;
			rectangleToCopy.Width = boundingSize.cx - rectangleToCopy.X - (properties.contentMargins.right + properties.extraMargins.right);
			int offset = rectangleToCopy.X;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// top-right corner
			rectangleToCopy.X = properties.loadedIconSize - 1 - (properties.contentMargins.right + properties.extraMargins.right);
			rectangleToCopy.Width = properties.contentMargins.right + properties.extraMargins.right;
			offset = boundingSize.cx - 1 - rectangleToCopy.Width;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// left edge
			rectangleToCopy.X = 0;
			rectangleToCopy.Y = properties.contentMargins.top + properties.extraMargins.top;
			rectangleToCopy.Width = properties.contentMargins.left + properties.extraMargins.left;
			rectangleToCopy.Height = boundingSize.cy - rectangleToCopy.Y - (properties.contentMargins.bottom + properties.extraMargins.bottom);
			offset = rectangleToCopy.Y * boundingSize.cx;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// right edge
			rectangleToCopy.X = properties.loadedIconSize - 1 - (properties.contentMargins.right + properties.extraMargins.right);
			rectangleToCopy.Width = properties.contentMargins.right + properties.extraMargins.right;
			offset += boundingSize.cx - 1 - rectangleToCopy.Width;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom-left corner
			rectangleToCopy.X = 0;
			rectangleToCopy.Y = properties.loadedIconSize - 1 - (properties.contentMargins.bottom + properties.extraMargins.bottom);
			rectangleToCopy.Width = properties.contentMargins.left + properties.extraMargins.left;
			rectangleToCopy.Height = properties.contentMargins.bottom + properties.extraMargins.bottom;
			offset = (boundingSize.cy - 1 - (properties.contentMargins.bottom + properties.extraMargins.bottom)) * boundingSize.cx;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom edge
			rectangleToCopy.X = properties.contentMargins.left + properties.extraMargins.left;
			rectangleToCopy.Width = boundingSize.cx - rectangleToCopy.X - (properties.contentMargins.right + properties.extraMargins.right);
			offset += rectangleToCopy.X;
			ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(&rectangleToCopy, boundingSize.cx * sizeof(RGBQUAD), (boundingSize.cx * boundingSize.cy - offset) * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits + offset))));
			// bottom-right corner
			rectangleToCopy.X = properties.loadedIconSize - 1 - (properties.contentMargins.right + properties.extraMargins.right);
			rectangleToCopy.Width = properties.contentMargins.right + properties.extraMargins.right;
			offset = (boundingSize.cy - (properties.contentMargins.bottom + properties.extraMargins.bottom)) * boundingSize.cx - 1 - rectangleToCopy.Width;
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
		LPDWORD pBits = static_cast<LPDWORD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
		if(pBits) {
			GetDIBits(memoryDC, iconInfo.hbmColor, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS);

			// search top-left corner
			POINT pt = {bmp.bmWidth >> 1, bmp.bmHeight >> 1};
			pMargins->left = 0;
			for(pt.x = min(20, bmp.bmWidth >> 1); pt.x > 0 && !foundIt; pt.x--) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(properties.loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->left = pt.x + 1;
				}
			}
			foundIt = FALSE;
			pt.x = bmp.bmWidth >> 1;
			pMargins->top = 0;
			for(pt.y = min(20, bmp.bmHeight >> 1); pt.y > 0 && !foundIt; pt.y--) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(properties.loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
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
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(properties.loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
					foundIt = TRUE;
					pMargins->right = bmp.bmWidth - pt.x;
				}
			}
			foundIt = FALSE;
			pt.x = bmp.bmWidth >> 1;
			pMargins->bottom = 0;
			for(pt.y = min(bmp.bmHeight - 20, bmp.bmHeight >> 1); pt.y < bmp.bmHeight && !foundIt; pt.y++) {
				// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
				if(pBits[pt.y * bmp.bmWidth + pt.x] != static_cast<DWORD>(properties.loadedIconSize == 16 ? 0x00FFFFFF : 0x00000000)) {
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