// helpers.cpp: global helper functions, macros and other stuff

#include "stdafx.h"
#include "helpers.h"


LPTSTR ConvertIntegerToString(char value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_itot_s(static_cast<int>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(int value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_itot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(long value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ltot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(__int64 value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_i64tot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned char value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(static_cast<unsigned long>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned int value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(static_cast<unsigned long>(value), pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned long value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ultot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

LPTSTR ConvertIntegerToString(unsigned __int64 value)
{
	LPTSTR pBuffer = new TCHAR[70];
	if(_ui64tot_s(value, pBuffer, 70, 10)) {
		delete[] pBuffer;
	}
	return pBuffer;
}

void ConvertStringToInteger(LPTSTR str, char& value)
{
	value = static_cast<char>(_ttoi(str) & 0xFF);
}

void ConvertStringToInteger(LPTSTR str, int& value)
{
	value = _ttoi(str);
}

void ConvertStringToInteger(LPTSTR str, long& value)
{
	value = _ttol(str);
}

void ConvertStringToInteger(LPTSTR str, __int64& value)
{
	value = _ttoi64(str);
}

long ConvertPixelsToTwips(HDC hDC, long pixels, BOOL vertical)
{
	const int twipsPerInch = 1440;
	int pixelsPerInch = 0;
	if(vertical) {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
	} else {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX);
	}
	return (pixels * twipsPerInch) / pixelsPerInch;
}

POINT ConvertPixelsToTwips(HDC hDC, POINT &pixels)
{
	POINT converted = {0};
	converted.x = ConvertPixelsToTwips(hDC, pixels.x, FALSE);
	converted.y = ConvertPixelsToTwips(hDC, pixels.y, TRUE);
	return converted;
}

RECT ConvertPixelsToTwips(HDC hDC, RECT &pixels)
{
	RECT converted = {0};
	converted.bottom = ConvertPixelsToTwips(hDC, pixels.bottom, TRUE);
	converted.left = ConvertPixelsToTwips(hDC, pixels.left, FALSE);
	converted.right = ConvertPixelsToTwips(hDC, pixels.right, FALSE);
	converted.top = ConvertPixelsToTwips(hDC, pixels.top, TRUE);
	return converted;
}

SIZE ConvertPixelsToTwips(HDC hDC, SIZE &pixels)
{
	SIZE converted = {0};
	converted.cx = ConvertPixelsToTwips(hDC, pixels.cx, FALSE);
	converted.cy = ConvertPixelsToTwips(hDC, pixels.cy, TRUE);
	return converted;
}

long ConvertTwipsToPixels(HDC hDC, long twips, BOOL vertical)
{
	const int twipsPerInch = 1440;
	int pixelsPerInch = 0;
	if(vertical) {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
	} else {
		pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX);
	}
	return (twips * pixelsPerInch) / twipsPerInch;
}

POINT ConvertTwipsToPixels(HDC hDC, POINT &twips)
{
	POINT converted = {0};
	converted.x = ConvertTwipsToPixels(hDC, twips.x, FALSE);
	converted.y = ConvertTwipsToPixels(hDC, twips.y, TRUE);
	return converted;
}

RECT ConvertTwipsToPixels(HDC hDC, RECT &twips)
{
	RECT converted = {0};
	converted.bottom = ConvertTwipsToPixels(hDC, twips.bottom, TRUE);
	converted.left = ConvertTwipsToPixels(hDC, twips.left, FALSE);
	converted.right = ConvertTwipsToPixels(hDC, twips.right, FALSE);
	converted.top = ConvertTwipsToPixels(hDC, twips.top, TRUE);
	return converted;
}

SIZE ConvertTwipsToPixels(HDC hDC, SIZE &twips)
{
	SIZE converted = {0};
	converted.cx = ConvertTwipsToPixels(hDC, twips.cx, FALSE);
	converted.cy = ConvertTwipsToPixels(hDC, twips.cy, TRUE);
	return converted;
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, LPTSTR description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	// convert the description string to an OLE string
	LPOLESTR oleDescription = NULL;
	if(description) {
		oleDescription = SysAllocString(CT2COLE(description));
	} else {
		if(HRESULT_FACILITY(hError) == FACILITY_WIN32) {
			// get the error from the system
			LPTSTR win32Description = NULL;
			if(!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER, NULL, HRESULT_CODE(hError), MAKELANGID(LANG_USER_DEFAULT, SUBLANG_DEFAULT), reinterpret_cast<LPTSTR>(&win32Description), 0, NULL)) {
				return HRESULT_FROM_WIN32(GetLastError());
			}

			// convert the multibyte string to an OLE string
			if(win32Description) {
				oleDescription = SysAllocString(CT2COLE(win32Description));
				LocalFree(win32Description);
			}
		}
	}

	// convert the title string to an OLE string
	LPOLESTR oleTitle = NULL;
	if(title) {
		oleTitle = SysAllocString(CT2COLE(title));
	}

	// convert the help filename string to an OLE string
	LPOLESTR oleHelpFileName = NULL;
	if(helpFileName) {
		oleHelpFileName = SysAllocString(CT2COLE(helpFileName));
	}

	// retrieve the ICreateErrorInfo interface
	LPCREATEERRORINFO pErrorInfoCreator = NULL;
	ATLVERIFY(SUCCEEDED(CreateErrorInfo(&pErrorInfoCreator)));

	// fill the error information into it
	pErrorInfoCreator->SetGUID(source);
	if(oleDescription) {
		pErrorInfoCreator->SetDescription(oleDescription);
	}
	if(oleTitle) {
		pErrorInfoCreator->SetSource(oleTitle);
	}
	if(oleHelpFileName) {
		pErrorInfoCreator->SetHelpFile(oleHelpFileName);
	}
	pErrorInfoCreator->SetHelpContext(helpContext);

	SysFreeString(oleDescription);
	SysFreeString(oleTitle);
	SysFreeString(oleHelpFileName);

	// retrieve the IErrorInfo interface
	LPERRORINFO pErrorInfo = NULL;
	HRESULT hSuccess = pErrorInfoCreator->QueryInterface(IID_IErrorInfo, reinterpret_cast<LPVOID*>(&pErrorInfo));
	if(FAILED(hSuccess)) {
		pErrorInfoCreator->Release();
		return hSuccess;
	}

	// set this error information in the current thread
	hSuccess = SetErrorInfo(0, pErrorInfo);
	if(FAILED(hSuccess)) {
		pErrorInfoCreator->Release();
		pErrorInfo->Release();
		return hSuccess;
	}

	// finally release the interfaces
	pErrorInfoCreator->Release();
	pErrorInfo->Release();

	// return the error code that was asked to be dispatched
	return hError;
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, LPTSTR title, UINT description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, title, COLE2T(GetResString(description)), helpContext, helpFileName);
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, LPTSTR description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, COLE2T(GetResString(title)), description, helpContext, helpFileName);
}

HRESULT DispatchError(HRESULT hError, REFCLSID source, UINT title, UINT description, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(hError, source, COLE2T(GetResString(title)), COLE2T(GetResString(description)), helpContext, helpFileName);
}

HRESULT DispatchError(DWORD errorCode, REFCLSID source, LPTSTR title, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	// dispatch the requested error message
	return DispatchError(HRESULT_FROM_WIN32(errorCode), source, title, static_cast<LPTSTR>(NULL), helpContext, helpFileName);
}

HRESULT DispatchError(DWORD errorCode, REFCLSID source, UINT title, DWORD helpContext/* = 0*/, LPTSTR helpFileName/* = NULL*/)
{
	return DispatchError(errorCode, source, COLE2T(GetResString(title)), helpContext, helpFileName);
}


HRESULT ReadVariantFromStream(LPSTREAM pStream, VARTYPE expectedVarType, LPVARIANT pVariant)
{
	ATLASSERT(expectedVarType == VT_BOOL || expectedVarType == VT_DATE || expectedVarType == VT_I2 || expectedVarType == VT_I4);

	HRESULT hr = pStream->Read(&pVariant->vt, sizeof(VARTYPE), NULL);
	if(SUCCEEDED(hr)) {
		if(pVariant->vt != expectedVarType) {
			return DISP_E_TYPEMISMATCH;
		}
		switch(pVariant->vt) {
			case VT_BOOL:
				hr = pStream->Read(&pVariant->boolVal, sizeof(VARIANT_BOOL), NULL);
				break;
			case VT_DATE:
				hr = pStream->Read(&pVariant->date, sizeof(DATE), NULL);
				break;
			case VT_I2:
				hr = pStream->Read(&pVariant->iVal, sizeof(SHORT), NULL);
				break;
			case VT_I4:
				hr = pStream->Read(&pVariant->lVal, sizeof(LONG), NULL);
				break;
		}
	}
	return hr;
}

HRESULT WriteVariantToStream(LPSTREAM pStream, LPVARIANT pVariant)
{
	ATLASSERT(pVariant->vt == VT_BOOL || pVariant->vt == VT_DATE || pVariant->vt == VT_I2 || pVariant->vt == VT_I4);

	HRESULT hr = pStream->Write(&pVariant->vt, sizeof(VARTYPE), NULL);
	if(SUCCEEDED(hr)) {
		switch(pVariant->vt) {
			case VT_BOOL:
				hr = pStream->Write(&pVariant->boolVal, sizeof(VARIANT_BOOL), NULL);
				break;
			case VT_DATE:
				hr = pStream->Write(&pVariant->date, sizeof(DATE), NULL);
				break;
			case VT_I2:
				hr = pStream->Write(&pVariant->iVal, sizeof(SHORT), NULL);
				break;
			case VT_I4:
				hr = pStream->Write(&pVariant->lVal, sizeof(LONG), NULL);
				break;
		}
	}
	return hr;
}

CComBSTR GetResString(UINT stringToLoad)
{
	CComBSTR comString;
	comString.LoadString(stringToLoad);
	return comString;
}

CComBSTR GetResStringWithNumber(UINT stringToLoad, int numberToInsert)
{
	CComBSTR result = GetResString(stringToLoad);
	int arraySize = result.Length() + 50;
	LPTSTR pBuffer = new TCHAR[arraySize];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, arraySize * sizeof(TCHAR), COLE2CT(result), numberToInsert)));
	result = CComBSTR(pBuffer);
	delete[] pBuffer;
	return result;
}

LPWSTR MLLoadDialogTitle(HINSTANCE hModule, LPTSTR pResourceName)
{
	ATLASSERT(hModule);

	LPWSTR pTitle = NULL;

	HRSRC hResource = FindResource(hModule, pResourceName, RT_DIALOG);
	if(hResource) {
		HGLOBAL hDialogTemplate = LoadResource(hModule, hResource);
		if(hDialogTemplate) {
			LPBYTE pDialogTemplate = reinterpret_cast<LPBYTE>(LockResource(hDialogTemplate));
			if(pDialogTemplate) {
				ATLASSERT(reinterpret_cast<DLGTEMPLATEEX*>(pDialogTemplate)->signature == 0xFFFF);
				LPWSTR p = reinterpret_cast<LPWSTR>(pDialogTemplate + 26);
				// skip menu
				if(*p == 0xFFFF) {
					p += 2;
				} else {
					while(*p++);
				}
				// skip window class
				if(*p == 0xFFFF) {
					p += 2;
				} else {
					while(*p++);
				}
				if(*p != 0x0000) {
					// copy the title
					int titleLength = lstrlenW(p) + 1;
					pTitle = new WCHAR[titleLength];
					lstrcpynW(pTitle, p, titleLength);
				}

				UnlockResource(hDialogTemplate);
			}
			FreeResource(hDialogTemplate);
		}
	}

	return pTitle;
}

HBITMAP LoadJPGResource(UINT jpgToLoad)
{
	HRSRC hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(jpgToLoad), TEXT("JPG"));
	if(hResource) {
		HGLOBAL hResourceMem = LoadResource(ModuleHelper::GetResourceInstance(), hResource);
		if(hResourceMem) {
			LPVOID pResourceData = LockResource(hResourceMem);
			if(pResourceData) {
				DWORD resourceSize = SizeofResource(ModuleHelper::GetResourceInstance(), hResource);
				HGLOBAL hJPGMem = GlobalAlloc(GMEM_MOVEABLE, resourceSize);
				if(hJPGMem) {
					LPVOID pJPGData = GlobalLock(hJPGMem);
					if(pJPGData) {
						CopyMemory(pJPGData, pResourceData, resourceSize);

						CComPtr<IStream> pStream;
						if(SUCCEEDED(CreateStreamOnHGlobal(hJPGMem, TRUE, &pStream))) {
							CComPtr<IPictureDisp> pPictureDisp;
							if(SUCCEEDED(OleLoadPicture(pStream, 0, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&pPictureDisp)))) {
								CComQIPtr<IPicture, &IID_IPicture> pPicture(pPictureDisp);
								if(pPicture) {
									HBITMAP hImage = NULL;
									pPicture->get_Handle(reinterpret_cast<OLE_HANDLE*>(&hImage));
									if(hImage) {
										return CopyBitmap(hImage);
									}
								}
							}
						}
						GlobalUnlock(hJPGMem);
					}
					GlobalFree(hJPGMem);
				}
			}
		}
	}
	return NULL;
}

BOOL IsMenuSeparator(HMENU hMenu, int position)
{
	CMenuHandle menu = hMenu;
	ATLASSERT(menu.IsMenu());

	CMenuItemInfo menuItemInfo;
	menuItemInfo.fMask = MIIM_FTYPE;
	menu.GetMenuItemInfo(position, TRUE, &menuItemInfo);
	return (menuItemInfo.fType & MFT_SEPARATOR);
}

void PrettifyMenu(HMENU hMenu)
{
	CMenuHandle menu = hMenu;
	ATLASSERT(menu.IsMenu());

	while(IsMenuSeparator(hMenu, 0)) {
		menu.RemoveMenu(0, MF_BYPOSITION);
	}

	int i = 0;
	int iPreviousSeparator = -2;
	while(i < menu.GetMenuItemCount()) {
		if(IsMenuSeparator(hMenu, i)) {
			if(i == iPreviousSeparator + 1) {
				menu.RemoveMenu(i, MF_BYPOSITION);
			} else {
				iPreviousSeparator = i++;
			}
		} else {
			++i;
		}
	}

	--i;
	while(IsMenuSeparator(hMenu, i)) {
		menu.RemoveMenu(i--, MF_BYPOSITION);
	}
}


HBITMAP CopyBitmap(HBITMAP hSourceBitmap, BOOL allowNullHandle/* = FALSE*/)
{
	if(!hSourceBitmap && !allowNullHandle) {
		return NULL;
	}

	CDC sourceDC;
	sourceDC.CreateCompatibleDC(NULL);
	CBitmap sourceBitmap;
	HBITMAP hPreviousBitmap1 = NULL;
	SIZE bitmapSize = {0};
	if(hSourceBitmap) {
		sourceBitmap.Attach(hSourceBitmap);
		hPreviousBitmap1 = sourceDC.SelectBitmap(sourceBitmap);
		sourceBitmap.GetSize(bitmapSize);
	}

	CDC targetDC;
	targetDC.CreateCompatibleDC(NULL);
	HDC hScreenDC = GetDC(HWND_DESKTOP);
	HBITMAP hTargetBitmap = CreateCompatibleBitmap(hScreenDC, bitmapSize.cx, bitmapSize.cy);
	ReleaseDC(HWND_DESKTOP, hScreenDC);
	HBITMAP hPreviousBitmap2 = targetDC.SelectBitmap(hTargetBitmap);

	targetDC.BitBlt(0, 0, bitmapSize.cx, bitmapSize.cy, sourceDC, 0, 0, SRCCOPY);

	if(hSourceBitmap) {
		sourceDC.SelectBitmap(hPreviousBitmap1);
		sourceBitmap.Detach();
	}
	targetDC.SelectBitmap(hPreviousBitmap2);

	return hTargetBitmap;
}

BOOL IsPremultiplied(HBITMAP hBitmap)
{
	BOOL premultiplied = FALSE;

	BITMAP bmp = {0};
	GetObject(hBitmap, sizeof(bmp), &bmp);

	BITMAPINFO bitmapInfo = {0};
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
	LPRGBQUAD pBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
	if(pBits) {
		CDC memoryDC;
		memoryDC.CreateCompatibleDC(NULL);
		if(GetDIBits(memoryDC, hBitmap, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS)) {
			premultiplied = IsPremultiplied(pBits, bmp.bmHeight * bmp.bmWidth);
		}
		HeapFree(GetProcessHeap(), 0, pBits);
	}
	return premultiplied;
}

BOOL IsPremultiplied(LPRGBQUAD pBits, int numberOfPixels)
{
	BOOL premultiplied = TRUE;

	/*// version without OpenMP
	for(int i = 0; i < numberOfPixels && premultiplied; ++i) {
		premultiplied = !(pBits[i].rgbBlue > pBits[i].rgbReserved || pBits[i].rgbGreen > pBits[i].rgbReserved || pBits[i].rgbRed > pBits[i].rgbReserved);
	}*/
	// TODO: Maybe it's faster to not use OpenMP here?
	#pragma omp parallel for reduction(&& : premultiplied)
	for(int i = 0; i < numberOfPixels; ++i) {
		premultiplied = premultiplied && !(pBits[i].rgbBlue > pBits[i].rgbReserved || pBits[i].rgbGreen > pBits[i].rgbReserved || pBits[i].rgbRed > pBits[i].rgbReserved);
	}
	return premultiplied;
}

HBITMAP BlendBitmap(HBITMAP hSourceBitmap, COLORREF blendColor, UINT intensity, BOOL premultiply)
{
	BITMAP bmp = {0};
	GetObject(hSourceBitmap, sizeof(bmp), &bmp);

	RGBQUAD blendClr;
	ATLASSERT(sizeof(COLORREF) == sizeof(RGBQUAD));
	CopyMemory(&blendClr, &blendColor, sizeof(blendColor));

	BITMAPINFO bitmapInfo = {0};
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
	LPRGBQUAD pBits;
	HBITMAP hTargetBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
	if(pBits) {
		CDC dc;
		dc.CreateCompatibleDC(NULL);
		if(GetDIBits(dc, hSourceBitmap, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS)) {
			if(premultiply) {
				#pragma omp parallel for
				for(int i = 0; i < bmp.bmHeight * bmp.bmWidth; i++) {
					pBits[i].rgbRed = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbRed) * (100 - intensity) + static_cast<UINT>(blendClr.rgbBlue) * intensity) / 100) & 0xFF);
					pBits[i].rgbGreen = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbGreen) * (100 - intensity) + static_cast<UINT>(blendClr.rgbGreen) * intensity) / 100) & 0xFF);
					pBits[i].rgbBlue = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbBlue) * (100 - intensity) + static_cast<UINT>(blendClr.rgbRed) * intensity) / 100) & 0xFF);
					if(pBits[i].rgbReserved) {
						pBits[i].rgbBlue = static_cast<BYTE>((static_cast<int>(pBits[i].rgbBlue) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
						pBits[i].rgbGreen = static_cast<BYTE>((static_cast<int>(pBits[i].rgbGreen) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
						pBits[i].rgbRed = static_cast<BYTE>((static_cast<int>(pBits[i].rgbRed) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
					}
				}
			} else {
				#pragma omp parallel for
				for(int i = 0; i < bmp.bmHeight * bmp.bmWidth; i++) {
					pBits[i].rgbRed = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbRed) * (100 - intensity) + static_cast<UINT>(blendClr.rgbBlue) * intensity) / 100) & 0xFF);
					pBits[i].rgbGreen = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbGreen) * (100 - intensity) + static_cast<UINT>(blendClr.rgbGreen) * intensity) / 100) & 0xFF);
					pBits[i].rgbBlue = static_cast<BYTE>(((static_cast<UINT>(pBits[i].rgbBlue) * (100 - intensity) + static_cast<UINT>(blendClr.rgbRed) * intensity) / 100) & 0xFF);
				}
			}
		}
	}
	return hTargetBitmap;
}

void PremultiplyBitmap(HBITMAP hBitmap)
{
	BITMAP bmp = {0};
	GetObject(hBitmap, sizeof(bmp), &bmp);

	BITMAPINFO bitmapInfo = {0};
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
	LPRGBQUAD pBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, bitmapInfo.bmiHeader.biSizeImage));
	if(pBits) {
		CDC memoryDC;
		memoryDC.CreateCompatibleDC(NULL);
		if(GetDIBits(memoryDC, hBitmap, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS)) {
			#pragma omp parallel for
			for(int i = 0; i < bmp.bmHeight * bmp.bmWidth; i++) {
				if(pBits[i].rgbReserved) {
					pBits[i].rgbBlue = static_cast<BYTE>((static_cast<int>(pBits[i].rgbBlue) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
					pBits[i].rgbGreen = static_cast<BYTE>((static_cast<int>(pBits[i].rgbGreen) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
					pBits[i].rgbRed = static_cast<BYTE>((static_cast<int>(pBits[i].rgbRed) * static_cast<int>(pBits[i].rgbReserved) / 255) & 0xFF);
				}
			}
			SetDIBits(memoryDC, hBitmap, 0, bmp.bmHeight, pBits, &bitmapInfo, DIB_RGB_COLORS);
		}
		HeapFree(GetProcessHeap(), 0, pBits);
	}
}

HIMAGELIST SetupStateImageList(HIMAGELIST hStateImageList)
{
	BOOL usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	CTheme themingEngine;
	if(usingThemes) {
		themingEngine.OpenThemeData(NULL, VSCLASS_BUTTON);
	}

	CImageList imageList = hStateImageList;
	SIZE iconSize;
	imageList.GetIconSize(iconSize);

	if(themingEngine.IsThemeNull()) {
		CIcon icon = imageList.ExtractIcon(1);
		if(!icon.IsNull()) {
			ICONINFO iconInfo;
			icon.GetIconInfo(&iconInfo);

			if(iconInfo.hbmColor) {
				CDCHandle compatibleDC = GetDC(HWND_DESKTOP);
				if(compatibleDC) {
					CDC memoryDC;
					memoryDC.CreateCompatibleDC(compatibleDC);
					if(memoryDC) {
						HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(iconInfo.hbmColor);

						for(int y = 0; y < iconSize.cy; ++y) {
							for(int x = 0; x < iconSize.cx; ++x) {
								if(memoryDC.GetPixel(x, y) == RGB(255, 255, 255)) {
									memoryDC.SetPixelV(x, y, RGB(192, 192, 192));
								}
							}
						}

						memoryDC.SelectBitmap(hPreviousBitmap);

						imageList.Add(iconInfo.hbmColor, iconInfo.hbmMask);
					}
					ReleaseDC(HWND_DESKTOP, compatibleDC.Detach());
				}
			}

			if(iconInfo.hbmColor) {
				DeleteObject(iconInfo.hbmColor);
			}
			if(iconInfo.hbmMask) {
				DeleteObject(iconInfo.hbmMask);
			}
			icon.DestroyIcon();
		}
	} else {
		CDCHandle compatibleDC = GetDC(HWND_DESKTOP);
		if(compatibleDC) {
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(compatibleDC);
			if(memoryDC) {
				CBitmap bitmap;
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = iconSize.cx;
				bitmapInfo.bmiHeader.biHeight = -iconSize.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pPixel;
				bitmap.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pPixel), NULL, 0);
				if(bitmap) {
					ZeroMemory(pPixel, iconSize.cx * iconSize.cy * sizeof(RGBQUAD));

					HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(bitmap);

					WTL::CRect rc(0, 0, iconSize.cx, iconSize.cy);
					SIZE partSize;
					themingEngine.GetThemePartSize(memoryDC, BP_CHECKBOX, CBS_MIXEDNORMAL, NULL, TS_TRUE, &partSize);
					rc.OffsetRect((iconSize.cx - partSize.cx) / 2, (iconSize.cy - partSize.cy) / 2);
					themingEngine.DrawThemeBackground(memoryDC, BP_CHECKBOX, CBS_MIXEDNORMAL, &rc);

					memoryDC.SelectBitmap(hPreviousBitmap);

					imageList.Add(bitmap);
				}
			}
			ReleaseDC(HWND_DESKTOP, compatibleDC.Detach());
		}
	}

	if(!themingEngine.IsThemeNull()) {
		themingEngine.CloseThemeData();
	}
	return imageList.Detach();
}

BOOL GetIconSize(HICON hIcon, LPSIZE pSize)
{
	BOOL succeeded = FALSE;

	ICONINFO iconInfo = {0};
	if(GetIconInfo(hIcon, &iconInfo)) {
		HBITMAP hBitmap = (iconInfo.hbmColor ? iconInfo.hbmColor : iconInfo.hbmMask);
		if(hBitmap) {
			BITMAP bmp = {0};
			if(GetObject(hBitmap, sizeof(BITMAP), &bmp)) {
				succeeded = TRUE;
				pSize->cx = bmp.bmWidth;
				pSize->cy = abs(bmp.bmHeight);
			}
		}

		if(iconInfo.hbmColor) {
			DeleteObject(iconInfo.hbmColor);
		}
		if(iconInfo.hbmMask) {
			DeleteObject(iconInfo.hbmMask);
		}
	}
	return succeeded;
}

BOOL MergeBitmaps(HBITMAP hFirstBitmap, HBITMAP hSecondBitmap, int xOffset, int yOffset)
{
	ATLASSERT(xOffset >= 0);
	ATLASSERT(yOffset >= 0);

	BOOL succeeded = FALSE;

	BITMAPINFO firstBitmapInfo;
	BITMAP firstBitmapDetails = {0};
	BITMAP secondBitmapDetails = {0};
	LPRGBQUAD pFirstBitmapBits = NULL;
	LPRGBQUAD pSecondBitmapBits = NULL;

	CDC memoryDC;
	memoryDC.CreateCompatibleDC(NULL);
	if(!memoryDC.IsNull()) {
		if(GetObject(hFirstBitmap, sizeof(BITMAP), &firstBitmapDetails) && GetObject(hSecondBitmap, sizeof(BITMAP), &secondBitmapDetails)) {
			ATLASSERT(firstBitmapDetails.bmHeight >= 0);
			ATLASSERT(secondBitmapDetails.bmHeight >= 0);

			firstBitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			firstBitmapInfo.bmiHeader.biWidth = firstBitmapDetails.bmWidth;
			firstBitmapInfo.bmiHeader.biHeight = -firstBitmapDetails.bmHeight;
			firstBitmapInfo.bmiHeader.biPlanes = 1;
			firstBitmapInfo.bmiHeader.biBitCount = 32;
			firstBitmapInfo.bmiHeader.biCompression = 0;
			firstBitmapInfo.bmiHeader.biSizeImage = firstBitmapDetails.bmHeight * firstBitmapDetails.bmWidth * 4;
			firstBitmapInfo.bmiHeader.biXPelsPerMeter = 0;
			firstBitmapInfo.bmiHeader.biYPelsPerMeter = 0;
			firstBitmapInfo.bmiHeader.biClrUsed = 0;
			firstBitmapInfo.bmiHeader.biClrImportant = 0;
			BITMAPINFO secondBitmapInfo = firstBitmapInfo;
			secondBitmapInfo.bmiHeader.biWidth = secondBitmapDetails.bmWidth;
			secondBitmapInfo.bmiHeader.biHeight = -secondBitmapDetails.bmHeight;
			secondBitmapInfo.bmiHeader.biSizeImage = secondBitmapDetails.bmHeight * secondBitmapDetails.bmWidth * 4;

			pFirstBitmapBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, firstBitmapInfo.bmiHeader.biSizeImage));
			if(pFirstBitmapBits) {
				if(GetDIBits(memoryDC, hFirstBitmap, 0, firstBitmapDetails.bmHeight, pFirstBitmapBits, &firstBitmapInfo, DIB_RGB_COLORS)) {
					pSecondBitmapBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, secondBitmapInfo.bmiHeader.biSizeImage));
					if(pSecondBitmapBits) {
						succeeded = GetDIBits(memoryDC, hSecondBitmap, 0, secondBitmapDetails.bmHeight, pSecondBitmapBits, &secondBitmapInfo, DIB_RGB_COLORS);
					}
				}
			}
		}
	}

	if(succeeded) {
		int maxX = min(firstBitmapDetails.bmWidth, xOffset + secondBitmapDetails.bmWidth);
		#pragma omp parallel for
		for(int y = yOffset; y < firstBitmapDetails.bmHeight; y++) {
			UINT offset1 = y * firstBitmapDetails.bmWidth;
			UINT offset2 = (y - yOffset) * secondBitmapDetails.bmWidth;
			for(int x = xOffset; x < maxX; x++) {
				// merge the pixels
				pFirstBitmapBits[offset1 + x].rgbRed = pSecondBitmapBits[offset2 + x - xOffset].rgbRed * pSecondBitmapBits[offset2 + x - xOffset].rgbReserved / 0xFF + (0xFF - pSecondBitmapBits[offset2 + x - xOffset].rgbReserved) * pFirstBitmapBits[offset1 + x].rgbRed / 0xFF;
				pFirstBitmapBits[offset1 + x].rgbGreen = pSecondBitmapBits[offset2 + x - xOffset].rgbGreen * pSecondBitmapBits[offset2 + x - xOffset].rgbReserved / 0xFF + (0xFF - pSecondBitmapBits[offset2 + x - xOffset].rgbReserved) * pFirstBitmapBits[offset1 + x].rgbGreen / 0xFF;
				pFirstBitmapBits[offset1 + x].rgbBlue = pSecondBitmapBits[offset2 + x - xOffset].rgbBlue * pSecondBitmapBits[offset2 + x - xOffset].rgbReserved / 0xFF + (0xFF - pSecondBitmapBits[offset2 + x - xOffset].rgbReserved) * pFirstBitmapBits[offset1 + x].rgbBlue / 0xFF;
				pFirstBitmapBits[offset1 + x].rgbReserved = pSecondBitmapBits[offset2 + x - xOffset].rgbReserved + (0xFF - pSecondBitmapBits[offset2 + x - xOffset].rgbReserved) * pFirstBitmapBits[offset1 + x].rgbReserved / 0xFF;
			}
		}
	}

	if(pSecondBitmapBits) {
		HeapFree(GetProcessHeap(), 0, pSecondBitmapBits);
	}
	if(pFirstBitmapBits) {
		if(succeeded) {
			succeeded = SetDIBits(memoryDC, hFirstBitmap, 0, firstBitmapDetails.bmHeight, pFirstBitmapBits, &firstBitmapInfo, DIB_RGB_COLORS);
		}
		HeapFree(GetProcessHeap(), 0, pFirstBitmapBits);
	}
	return succeeded;
}

HBITMAP StampIconForElevation(HICON hIcon, LPSIZE pTargetSize, BOOL stampIcon)
{
	HBITMAP hTargetImage = NULL;
	BOOL destroyIcon = FALSE;

	if(stampIcon && hIcon && APIWrapper::IsSupported_StampIconForElevation()) {
		HRESULT hr = E_FAIL;
		HICON hStampedIcon = NULL;
		APIWrapper::StampIconForElevation(hIcon, FALSE, &hStampedIcon, &hr);
		if(SUCCEEDED(hr)) {
			hIcon = hStampedIcon;
			destroyIcon = TRUE;

			if(pTargetSize) {
				SIZE iconSize = {0};
				if(GetIconSize(hIcon, &iconSize)) {
					if(iconSize.cx > pTargetSize->cx || iconSize.cy > pTargetSize->cy) {
						pTargetSize->cx = iconSize.cx;
						pTargetSize->cy = iconSize.cy;
					}
				}
			}
		}
	}

	if(pTargetSize) {
		CComPtr<IWICImagingFactory> pWICImagingFactory;
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		if(pWICImagingFactory && hIcon) {
			CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
			HRESULT hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pWICBitmapScaler);
				CComPtr<IWICBitmap> pWICBitmap = NULL;
				hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmap);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					ATLASSUME(pWICBitmap);
					ATLVERIFY(SUCCEEDED(pWICBitmapScaler->Initialize(pWICBitmap, pTargetSize->cx, pTargetSize->cy, WICBitmapInterpolationModeFant)));

					// create target bitmap
					BITMAPINFO bitmapInfo = {0};
					bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
					bitmapInfo.bmiHeader.biWidth = pTargetSize->cx;
					bitmapInfo.bmiHeader.biHeight = -pTargetSize->cy;
					bitmapInfo.bmiHeader.biPlanes = 1;
					bitmapInfo.bmiHeader.biBitCount = 32;
					bitmapInfo.bmiHeader.biCompression = BI_RGB;
					LPRGBQUAD pBits = NULL;
					hTargetImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
					if(hTargetImage) {
						ATLASSERT_POINTER(pBits, RGBQUAD);
						if(pBits) {
							ATLVERIFY(SUCCEEDED(pWICBitmapScaler->CopyPixels(NULL, pTargetSize->cx * sizeof(RGBQUAD), pTargetSize->cy * pTargetSize->cx * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits))));
						}
					}
				}
			}
		} else if(hIcon) {
			// no scaling
			ICONINFO iconInfo = {0};
			if(GetIconInfo(hIcon, &iconInfo)) {
				if(iconInfo.hbmMask) {
					DeleteObject(iconInfo.hbmMask);
				}
				hTargetImage = iconInfo.hbmColor;
			}
		}
	} else if(hIcon) {
		// no scaling
		ICONINFO iconInfo = {0};
		if(GetIconInfo(hIcon, &iconInfo)) {
			if(iconInfo.hbmMask) {
				DeleteObject(iconInfo.hbmMask);
			}
			hTargetImage = iconInfo.hbmColor;
		}
	}

	if(destroyIcon) {
		DestroyIcon(hIcon);
	}
	return hTargetImage;
}

BOOL ImageList_DrawIndirect_UACStamped(IMAGELISTDRAWPARAMS* pParams, BOOL drawUACStamp)
{
	BOOL b = FALSE;

	HICON hIcon = ImageList_GetIcon(pParams->himl, pParams->i, pParams->fStyle);
	ATLASSERT(hIcon);
	if(drawUACStamp && hIcon && APIWrapper::IsSupported_StampIconForElevation()) {
		HRESULT hr = E_FAIL;
		HICON hStampedIcon = NULL;
		APIWrapper::StampIconForElevation(hIcon, FALSE, &hStampedIcon, &hr);
		if(SUCCEEDED(hr)) {
			DestroyIcon(hIcon);
			hIcon = hStampedIcon;
		}
	}
	if(hIcon) {
		SIZE iconSize = {0};
		if(GetIconSize(hIcon, &iconSize)) {
			// use an imagelist to draw the thumbnail (works around alpha channel problems)
			IMAGELISTDRAWPARAMS params = *pParams;
			params.himl = ImageList_Create(iconSize.cx, iconSize.cy, ILC_COLOR32, 1, 0);
			ATLASSERT(params.himl);
			if(params.himl) {
				params.i = ImageList_AddIcon(params.himl, hIcon);
				if(params.i >= 0) {
					if(iconSize.cx > params.cx || iconSize.cy > params.cy) {
						params.x -= iconSize.cx - params.cx;
						params.y -= iconSize.cy - params.cy;
						params.cx = iconSize.cx;
						params.cy = iconSize.cy;
					}
					b = ImageList_DrawIndirect(&params);
				}
				ImageList_Destroy(params.himl);
			}
		}
		DestroyIcon(hIcon);
	}

	return b;
}

void CalculateAspectScaledRectangle(const LPSIZE pOuterSize, LPRECT pInnerRectangle)
{
	ATLASSERT(pInnerRectangle->left == 0);
	ATLASSERT(pInnerRectangle->top == 0);

	int width = pInnerRectangle->right;
	int height = pInnerRectangle->bottom;
	int xRatio = (width * 1000) / pOuterSize->cx;
	int yRatio = (height * 1000) / pOuterSize->cy;

	if(xRatio > yRatio) {
		pInnerRectangle->right = pOuterSize->cx;
		// work out the blank space and split it evenly between the top and the bottom
		int newHeight = ((height * 1000) / xRatio);
		if(newHeight == 0) {
			newHeight = 1;
		}
		int remainder = pOuterSize->cy - newHeight;
		pInnerRectangle->top = remainder >> 1;
		pInnerRectangle->bottom = newHeight + pInnerRectangle->top;
	} else {
		pInnerRectangle->bottom = pOuterSize->cy;
		// work out the blank space and split it evenly between the left and the right
		int newWidth = ((width * 1000) / yRatio);
		if(newWidth == 0) {
			newWidth = 1;
		}
		int remainder = pOuterSize->cx - newWidth;
		pInnerRectangle->left = remainder >> 1;
		pInnerRectangle->right = newWidth + pInnerRectangle->left;
	}
}

void CalculateAspectRatio(const LPSIZE pOuterSize, LPRECT pInnerRectangle)
{
	int height = abs(pInnerRectangle->bottom - pInnerRectangle->top);
	int width = abs(pInnerRectangle->right - pInnerRectangle->left);

	// check if the initial bitmap is larger than the size of the thumbnail
	if(width > pOuterSize->cx || height > pOuterSize->cy) {
		pInnerRectangle->left = 0;
		pInnerRectangle->top = 0;
		pInnerRectangle->right = width;
		pInnerRectangle->bottom = height;
		CalculateAspectScaledRectangle(pOuterSize, pInnerRectangle);
	} else {
		// if the bitmap was smaller than the thumbnail, just center it
		pInnerRectangle->left = (pOuterSize->cx - width) >> 1;
		pInnerRectangle->top = (pOuterSize->cy - height) >> 1;
		pInnerRectangle->right = pInnerRectangle->left + width;
		pInnerRectangle->bottom = pInnerRectangle->top + height;
	}
}

BOOL DrawOntoWhiteBackground(LPBITMAPINFO pSourceBitmapInfo, LPVOID pSourceBits, const LPSIZE pTargetSize, LPRECT pBoundingRectangle, HBITMAP* phBitmap)
{
	*phBitmap = NULL;
	int result = GDI_ERROR;

	// create thumbnail bitmap and copy image into it
	CDC memoryDC;
	memoryDC.CreateCompatibleDC();
	if(!memoryDC.IsNull()) {
		BITMAPINFO bitmapInfo = {0};
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = pTargetSize->cx;
		bitmapInfo.bmiHeader.biHeight = -pTargetSize->cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
		LPVOID pTargetBits = NULL;
		*phBitmap = CreateDIBSection(memoryDC, &bitmapInfo, DIB_RGB_COLORS, &pTargetBits, NULL, 0);

		HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(*phBitmap);
		memoryDC.SetStretchBltMode(COLORONCOLOR);
		HBRUSH hPreviousBrush = memoryDC.SelectBrush(static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
		HPEN hPreviousPen = memoryDC.SelectPen(static_cast<HPEN>(GetStockObject(WHITE_PEN)));
		memoryDC.SetMapMode(MM_TEXT);

		memoryDC.Rectangle(0, 0, pTargetSize->cx, pTargetSize->cy);

		int destinationHeight = pBoundingRectangle->bottom - pBoundingRectangle->top;
		int destinationTop = pBoundingRectangle->top;
		int sourceTop = 0;
		if(pSourceBitmapInfo->bmiHeader.biHeight < 0) {
			destinationHeight *= -1;
			destinationTop = pBoundingRectangle->bottom;
			sourceTop = -pSourceBitmapInfo->bmiHeader.biHeight;
		}

		result = memoryDC.StretchDIBits(pBoundingRectangle->left, destinationTop, pBoundingRectangle->right - pBoundingRectangle->left, destinationHeight, 0, sourceTop, pSourceBitmapInfo->bmiHeader.biWidth, pSourceBitmapInfo->bmiHeader.biHeight, pSourceBits, pSourceBitmapInfo, DIB_RGB_COLORS, SRCCOPY);

		memoryDC.SelectPen(hPreviousPen);
		memoryDC.SelectBrush(hPreviousBrush);
		memoryDC.SelectBitmap(hPreviousBitmap);
	}
	return (result != GDI_ERROR);
}

UINT FindBestMatchingIconResource(HMODULE hSource, LPCTSTR pResource, int desiredSize, int* pSize)
{
	#pragma pack(push)
	#pragma pack(2)
	typedef struct
	{
		BYTE bWidth;
		BYTE bHeight;
		BYTE bColorCount;
		BYTE bReserved;
		WORD wPlanes;
		WORD wBitCount;
		DWORD dwBytesInRes;
		WORD nID;
	} GRPICONDIRENTRY, *LPGRPICONDIRENTRY;

	typedef struct
	{
		WORD idReserved;
		WORD idType;
		WORD idCount;
		GRPICONDIRENTRY idEntries[1];
	} GRPICONDIR, *LPGRPICONDIR;
	#pragma pack(pop)

	UINT ret = 0;
	*pSize = 0;

	HRSRC hResource = FindResource(hSource, pResource, RT_GROUP_ICON);
	if(hResource) {
		HGLOBAL hMem = LoadResource(hSource, hResource);
		if(hMem) {
			LPGRPICONDIR pIconData = reinterpret_cast<LPGRPICONDIR>(LockResource(hMem));
			if(pIconData) {
				int bestMatchIndex = -1;
				int bestMatchSize = 0x7FFFFFFF;
				for(int i = 0; i < pIconData->idCount; ++i) {
					int entrySize = (pIconData->idEntries[i].bWidth == 0 ? 256 : pIconData->idEntries[i].bWidth);
					if(desiredSize == entrySize) {
						// perfect match
						bestMatchIndex = i;
						bestMatchSize = entrySize;
						break;
					} else if(desiredSize < entrySize) {
						if(entrySize < bestMatchSize) {
							// current entry is a better match
							bestMatchIndex = i;
							bestMatchSize = entrySize;
						}
					}
				}
				if(bestMatchIndex == -1) {
					bestMatchIndex = 0;
					bestMatchSize = (pIconData->idEntries[0].bWidth == 0 ? 256 : pIconData->idEntries[0].bWidth);
				}
				ret = pIconData->idEntries[bestMatchIndex].nID;
				*pSize = bestMatchSize;
			}
		}
	}
	return ret;
}