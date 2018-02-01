// PhotoBorderAdorner.cpp: An implementation of IThumbnailAdorner for drawing photo borders around thumbnails

#include "stdafx.h"
#include "PhotoBorderAdorner.h"


PhotoBorderAdorner::PhotoBorderAdorner()
{
	properties.hAdornmentIcon = NULL;
}

PhotoBorderAdorner::~PhotoBorderAdorner()
{
	if(properties.hAdornmentIcon) {
		DestroyIcon(properties.hAdornmentIcon);
		properties.hAdornmentIcon = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IThumbnailAdorner
STDMETHODIMP PhotoBorderAdorner::SetIconSize(SIZE& targetIconSize, double /*contentAspectRatio*/)
{
	this->properties.targetIconSize = targetIconSize;

	HMODULE hImageRes = LoadLibraryEx(TEXT("imageres.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
	if(hImageRes) {
		int sz = max(targetIconSize.cx, targetIconSize.cy);
		UINT iconToLoad = FindBestMatchingIconResource(hImageRes, MAKEINTRESOURCE(192), sz, &properties.loadedIconSize);
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

STDMETHODIMP PhotoBorderAdorner::GetMaxContentSize(double contentAspectRatio, LPSIZE pContentSize)
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
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	pAdornedSize->cx = contentSize.cx + properties.contentMargins.left + properties.contentMargins.right;
	pAdornedSize->cy = contentSize.cy + properties.contentMargins.top + properties.contentMargins.bottom;
	pAdornedSize->cx += (properties.extraMargins.left + properties.extraMargins.right);
	pAdornedSize->cy += (properties.extraMargins.top + properties.extraMargins.bottom);
	return S_OK;
}

STDMETHODIMP PhotoBorderAdorner::SetAdornedSize(SIZE& /*adornedSize*/)
{
	if(!properties.hAdornmentIcon) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP PhotoBorderAdorner::GetContentRectangle(RECT& adornedRectangle, SIZE& contentSize, LPRECT pContentRectangle, LPRECT pContentAreaRectangle, LPRECT pContentToDrawRectangle)
{
	CRect marginedContentRectangle(adornedRectangle.left + properties.contentMargins.left, adornedRectangle.top + properties.contentMargins.top, adornedRectangle.right - properties.contentMargins.right, adornedRectangle.bottom - properties.contentMargins.bottom);
	marginedContentRectangle.left += properties.extraMargins.left;
	marginedContentRectangle.top += properties.extraMargins.top;
	marginedContentRectangle.right -= properties.extraMargins.right;
	marginedContentRectangle.bottom -= properties.extraMargins.bottom;

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
			POINT pt = {0};
			for(pt.y = 0; pt.y < min(10, bmp.bmHeight) && !foundIt; pt.y++) {
				for(pt.x = 0; pt.x < min(10, bmp.bmWidth) && !foundIt; pt.x++) {
					// HACK: 16x16 icon is 24bpp only, but bmp.bmBitsPixel is 32?!
					if(pBits[pt.y * bmp.bmWidth + pt.x] == (properties.loadedIconSize == 16 ? 0x00FFFFFF : 0xFFFFFFFF)) {
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
						if((pBits[pt.y * bmp.bmWidth + pt.x] & 0xFF000000) == (properties.loadedIconSize == 16 ? 0x00000000 : 0xFF000000)) {
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