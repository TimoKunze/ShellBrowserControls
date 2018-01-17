// PhotoBorderAdorner.cpp: An implementation of IThumbnailAdorner for drawing photo borders around thumbnails

#include "stdafx.h"
#include "PhotoBorderAdorner.h"


PhotoBorderAdorner::PhotoBorderAdorner()
{
	hAdornmentIcon = NULL;
}

PhotoBorderAdorner::~PhotoBorderAdorner()
{
	if(hAdornmentIcon) {
		DestroyIcon(hAdornmentIcon);
		hAdornmentIcon = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IThumbnailAdorner
STDMETHODIMP PhotoBorderAdorner::SetIconSize(SIZE& targetIconSize, double /*contentAspectRatio*/)
{
	this->targetIconSize = targetIconSize;

	HMODULE hImageRes = LoadLibraryEx(TEXT("imageres.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hImageRes) {
		int sz = max(targetIconSize.cx, targetIconSize.cy);
		UINT iconToLoad = FindBestMatchingIconResource(hImageRes, MAKEINTRESOURCE(192), sz, &loadedIconSize);
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

STDMETHODIMP PhotoBorderAdorner::GetMaxContentSize(double contentAspectRatio, LPSIZE pContentSize)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	SIZE contentAreaSize;
	contentAreaSize.cx = targetIconSize.cx - contentMargins.left - contentMargins.right;
	contentAreaSize.cy = targetIconSize.cy - contentMargins.top - contentMargins.bottom;
	contentAreaSize.cx -= (extraMargins.left + extraMargins.right);
	contentAreaSize.cy -= (extraMargins.top + extraMargins.bottom);

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

STDMETHODIMP PhotoBorderAdorner::GetAdornedSize(SIZE& contentSize, LPSIZE pAdornedSize)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	pAdornedSize->cx = contentSize.cx + contentMargins.left + contentMargins.right;
	pAdornedSize->cy = contentSize.cy + contentMargins.top + contentMargins.bottom;
	pAdornedSize->cx += (extraMargins.left + extraMargins.right);
	pAdornedSize->cy += (extraMargins.top + extraMargins.bottom);
	return S_OK;
}

STDMETHODIMP PhotoBorderAdorner::SetAdornedSize(SIZE& /*adornedSize*/)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP PhotoBorderAdorner::GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, LPRECT pContentRectangle, LPRECT pContentAreaRectangle, LPRECT pContentToDrawRectangle)
{
	WTL::CRect marginedContentRectangle(adornedRectangle.left + contentMargins.left, adornedRectangle.top + contentMargins.top, adornedRectangle.right - contentMargins.right, adornedRectangle.bottom - contentMargins.bottom);
	marginedContentRectangle.left += extraMargins.left;
	marginedContentRectangle.top += extraMargins.top;
	marginedContentRectangle.right -= extraMargins.right;
	marginedContentRectangle.bottom -= extraMargins.bottom;

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

STDMETHODIMP PhotoBorderAdorner::GetDrawOrder(LPBOOL pAdornmentFirst)
{
	*pAdornmentFirst = TRUE;
	return S_OK;
}

STDMETHODIMP PhotoBorderAdorner::DrawIntoDIB(LPRECT pBoundingRectangle, LPRGBQUAD pBits)
{
	if(!hAdornmentIcon) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	SIZE boundingSize = WTL::CRect(pBoundingRectangle).Size();

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

STDMETHODIMP PhotoBorderAdorner::PostProcess(HBITMAP /*hSourceBitmap*/, LPRGBQUAD /*pSourceBits*//* = NULL*/, HBITMAP* /*pHCroppedBitmap*//* = NULL*/)
{
	return E_NOTIMPL;
}
// implementation of IThumbnailAdorner
//////////////////////////////////////////////////////////////////////


BOOL PhotoBorderAdorner::FindContentAreaMargins(LPRECT pMargins)
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
			POINT pt = {0};
			for(pt.y = 0; pt.y < min(10, bmp.bmHeight) && !foundIt; pt.y++) {
				for(pt.x = 0; pt.x < min(10, bmp.bmWidth) && !foundIt; pt.x++) {
					// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
					if(pBits[pt.y * bmp.bmWidth + pt.x] == (loadedIconSize == 16 ? 0x00FFFFFF : 0xFFFFFFFF)) {
						foundIt = TRUE;
						pMargins->left = pt.x;
						pMargins->top = pt.y;
					}
				}
			}
			if(foundIt) {
				// search bottom-right corner
				foundIt = FALSE;
				for(pt.y = bmp.bmHeight - 1; pt.y >= max(0, bmp.bmHeight - 15) && !foundIt; pt.y--) {
					for(pt.x = bmp.bmWidth - 1; pt.x >= max(0, bmp.bmWidth - 15) && !foundIt; pt.x--) {
						// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
						// HACK: The last white column in the 48x48 icon isn't plain white!
						if((pBits[pt.y * bmp.bmWidth + pt.x] & 0xFF000000) == (loadedIconSize == 16 ? 0x00000000 : 0xFF000000)) {
							LPRGBQUAD p = reinterpret_cast<LPRGBQUAD>(&pBits[pt.y * bmp.bmWidth + pt.x]);
							if(p->rgbBlue == p->rgbGreen && p->rgbBlue == p->rgbRed && p->rgbBlue >= 250) {
								foundIt = TRUE;
								pMargins->right = bmp.bmWidth - 1 - pt.x;
								pMargins->bottom = bmp.bmHeight - 1 - pt.y;
							}
						}
					}
				}
			}
			HeapFree(GetProcessHeap(), 0, pBits);
		}
		DeleteObject(iconInfo.hbmColor);
	}
	return foundIt;
}