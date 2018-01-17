// SecurityZone.cpp: A wrapper for Internet Explorer security zones.

#include "stdafx.h"
#include "SecurityZone.h"


#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	// implementation of ISupportErrorInfo
	STDMETHODIMP SecurityZone::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
	{
		if(InlineIsEqualGUID(IID_ISecurityZone, interfaceToCheck)) {
			return S_OK;
		}
		return S_FALSE;
	}
	// implementation of ISupportErrorInfo
	//////////////////////////////////////////////////////////////////////


	void SecurityZone::Attach(LONG zoneIndex, LPZONEATTRIBUTES pZoneAttributes)
	{
		properties.zoneIndex = zoneIndex;
		properties.zoneAttributes = *pZoneAttributes;
	}

	void SecurityZone::Detach(void)
	{
		properties.zoneIndex = -1;
	}


	STDMETHODIMP SecurityZone::get_Description(BSTR* pValue)
	{
		ATLASSERT_POINTER(pValue, BSTR);
		if(!pValue) {
			return E_POINTER;
		}

		*pValue = _bstr_t(properties.zoneAttributes.szDescription).Detach();
		return S_OK;
	}

	STDMETHODIMP SecurityZone::get_DisplayName(BSTR* pValue)
	{
		ATLASSERT_POINTER(pValue, BSTR);
		if(!pValue) {
			return E_POINTER;
		}

		*pValue = _bstr_t(properties.zoneAttributes.szDisplayName).Detach();
		return S_OK;
	}

	STDMETHODIMP SecurityZone::get_hIcon(OLE_HANDLE* pValue)
	{
		ATLASSERT_POINTER(pValue, OLE_HANDLE);
		if(!pValue) {
			return E_POINTER;
		}

		WORD icon = 0;
		HICON hIcon = NULL;
		WCHAR pBuffer[MAX_PATH];
		lstrcpynW(pBuffer, properties.zoneAttributes.szIconPath, min(260, MAX_PATH));
		LPWSTR pSharp = StrChrW(pBuffer, L'#');
		if(pSharp) {
			pSharp[0] = L'\0';
			icon = static_cast<WORD>(StrToIntW(pSharp + 1));
			if(lstrcmpiW(pBuffer, L"explorer.exe") == 0) {
				/* HACK: The app itself is likely to be called explorer.exe. But here C:\Windows\explorer.exe is
								 meant! */
				SHGetFolderPathW(NULL, CSIDL_WINDOWS, NULL, SHGFP_TYPE_CURRENT, pBuffer);
				PathAddBackslashW(pBuffer);
				StringCchCatW(pBuffer, MAX_PATH, L"explorer.exe");
			}
			ExtractIconExW(pBuffer, static_cast<UINT>(-1 * icon), NULL, &hIcon, 1);
		} else {
			HINSTANCE hBrowseUI = LoadLibraryEx(TEXT("browseui.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
			if(hBrowseUI) {
				hIcon = ExtractAssociatedIconExW(hBrowseUI, pBuffer, &icon, &icon);
				FreeLibrary(hBrowseUI);
			}
		}
		*pValue = HandleToLong(hIcon);
		return S_OK;
	}

	STDMETHODIMP SecurityZone::get_IconPath(BSTR* pValue)
	{
		ATLASSERT_POINTER(pValue, BSTR);
		if(!pValue) {
			return E_POINTER;
		}

		*pValue = _bstr_t(properties.zoneAttributes.szIconPath).Detach();
		return S_OK;
	}
#endif