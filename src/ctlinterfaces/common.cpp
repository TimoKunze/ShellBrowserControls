// common.cpp: common control interfaces helper methods, macros and other stuff

#include "stdafx.h"
#include "common.h"
#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
	#include "exlvw/definitions.h"
#endif


SHCONTF GetEnumerationFlags(INamespaceEnumSettings* pEnumerationSettings)
{
	ATLASSUME(pEnumerationSettings);

	ShNamespaceEnumerationConstants enumerationFlags;
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_EnumerationFlags(&enumerationFlags)));

	SHCONTF enumFlags = static_cast<SHCONTF>(enumerationFlags & ~0x1);
	if(enumerationFlags & snseMayIncludeHiddenItems) {
		SHELLFLAGSTATE shellSettings = {0};
		SHGetSettings(&shellSettings, SSF_SHOWALLOBJECTS);
		if(!shellSettings.fShowAllObjects) {
			enumFlags = static_cast<SHCONTF>(enumFlags & ~SHCONTF_INCLUDEHIDDEN);
		}
	}
	return enumFlags;
}

CachedEnumSettings CacheEnumSettings(INamespaceEnumSettings* pEnumerationSettings)
{
	ATLASSUME(pEnumerationSettings);

	CachedEnumSettings settings;
	settings.enumerationFlags = GetEnumerationFlags(pEnumerationSettings);
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_ExcludedFileItemFileAttributes(&settings.excludedFileItemFileAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_ExcludedFileItemShellAttributes(&settings.excludedFileItemShellAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_ExcludedFolderItemFileAttributes(&settings.excludedFolderItemFileAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_ExcludedFolderItemShellAttributes(&settings.excludedFolderItemShellAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_IncludedFileItemFileAttributes(&settings.includedFileItemFileAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_IncludedFileItemShellAttributes(&settings.includedFileItemShellAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_IncludedFolderItemFileAttributes(&settings.includedFolderItemFileAttributes)));
	ATLVERIFY(SUCCEEDED(pEnumerationSettings->get_IncludedFolderItemShellAttributes(&settings.includedFolderItemShellAttributes)));
	settings.fileAttributesMask = static_cast<DWORD>(settings.includedFileItemFileAttributes) | static_cast<DWORD>(settings.includedFolderItemFileAttributes) | static_cast<DWORD>(settings.excludedFileItemFileAttributes) | static_cast<DWORD>(settings.excludedFolderItemFileAttributes);
	if(settings.fileAttributesMask != 0) {
		settings.includedFolderItemFileAttributes = static_cast<ItemFileAttributesConstants>(settings.includedFolderItemShellAttributes | ifaDirectory);
	}
	settings.shellAttributesMask = static_cast<SFGAOF>(settings.includedFileItemShellAttributes) | static_cast<SFGAOF>(settings.includedFolderItemShellAttributes) | static_cast<SFGAOF>(settings.excludedFileItemShellAttributes) | static_cast<SFGAOF>(settings.excludedFolderItemShellAttributes);
	if(settings.shellAttributesMask != 0) {
		settings.includedFolderItemShellAttributes = static_cast<ItemShellAttributesConstants>(settings.includedFolderItemShellAttributes | isaIsFolder);
	}

	return settings;
}

BOOL ShouldShowItem(PCIDLIST_ABSOLUTE pIDL, CachedEnumSettings* pCachedEnumSettings)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		return ShouldShowItem(pParentISF, const_cast<PITEMID_CHILD>(pRelativePIDL), pCachedEnumSettings);
	}
	return FALSE;
}

BOOL ShouldShowItem(LPSHELLFOLDER pParentISF, PITEMID_CHILD pRelativePIDL, CachedEnumSettings* pCachedEnumSettings)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(pCachedEnumSettings, CachedEnumSettings);
	if(!pParentISF || !pRelativePIDL || !pCachedEnumSettings) {
		return FALSE;
	}

	// we'll check this flag, so we need to query it
	SFGAOF shellAttributes = SFGAO_FOLDER;
	if(pCachedEnumSettings->shellAttributesMask != 0) {
		shellAttributes |= pCachedEnumSettings->shellAttributesMask;
	}
	if(pCachedEnumSettings->fileAttributesMask != 0) {
		// we'll check this flag, so we need to query it
		shellAttributes |= SFGAO_FILESYSTEM;
	}

	BOOL show = TRUE;
	if(shellAttributes != 0) {
		ATLVERIFY(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &shellAttributes)));
		if(pCachedEnumSettings->shellAttributesMask != 0) {
			if(shellAttributes & SFGAO_FOLDER) {
				if(shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->excludedFolderItemShellAttributes)) {
					show = FALSE;
				} else if((shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->includedFolderItemShellAttributes)) != static_cast<SFGAOF>(pCachedEnumSettings->includedFolderItemShellAttributes)) {
					show = FALSE;
				}
			} else {
				if(shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->excludedFileItemShellAttributes)) {
					show = FALSE;
				} else if((shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->includedFileItemShellAttributes)) != static_cast<SFGAOF>(pCachedEnumSettings->includedFileItemShellAttributes)) {
					show = FALSE;
				}
			}
		}

		if(show && (shellAttributes & SFGAO_FILESYSTEM) && (pCachedEnumSettings->fileAttributesMask != 0)) {
			LPWSTR pPath = NULL;
			ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
			if(pPath) {
				BOOL isDrive = FALSE;
				if(PathIsRootW(pPath)) {
					int len = lstrlenW(pPath);
					if(len >= 2 && pPath[1] == L':') {
						isDrive = (len == 2 || (len == 3 && pPath[2] == L'\\'));
					}
				}
				DWORD fileAttributes = INVALID_FILE_ATTRIBUTES;
				if(!isDrive) {
					/* The values returned by GetFileAttributes for drives are a bit strange, e.g. it makes a
					   difference whether you pass C: or C:\. Also we don't really want to apply attribute filters to
					   drives. */
					fileAttributes = GetFileAttributesW(pPath);
				}
				CoTaskMemFree(pPath);

				if(fileAttributes != INVALID_FILE_ATTRIBUTES) {
					if(fileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
						if(fileAttributes & static_cast<DWORD>(pCachedEnumSettings->excludedFolderItemFileAttributes)) {
							show = FALSE;
						} else if((fileAttributes & static_cast<DWORD>(pCachedEnumSettings->includedFolderItemFileAttributes)) != static_cast<DWORD>(pCachedEnumSettings->includedFolderItemFileAttributes)) {
							show = FALSE;
						}
					} else {
						if(fileAttributes & static_cast<DWORD>(pCachedEnumSettings->excludedFileItemFileAttributes)) {
							show = FALSE;
						} else if((fileAttributes & static_cast<DWORD>(pCachedEnumSettings->includedFileItemFileAttributes)) != static_cast<DWORD>(pCachedEnumSettings->includedFileItemFileAttributes)) {
							show = FALSE;
						}
					}
				}
			}
		}
	}
	return show;
}

BOOL ShouldShowItem(IShellItem* pShellItem, CachedEnumSettings* pCachedEnumSettings)
{
	ATLASSUME(pShellItem);
	ATLASSERT_POINTER(pCachedEnumSettings, CachedEnumSettings);
	if(!pShellItem || !pCachedEnumSettings) {
		return FALSE;
	}

	// we'll check this flag, so we need to query it
	SFGAOF shellAttributesMask = SFGAO_FOLDER;
	if(pCachedEnumSettings->shellAttributesMask != 0) {
		shellAttributesMask |= pCachedEnumSettings->shellAttributesMask;
	}
	if(pCachedEnumSettings->fileAttributesMask != 0) {
		// we'll check this flag, so we need to query it
		shellAttributesMask |= SFGAO_FILESYSTEM;
	}

	BOOL show = TRUE;
	if(shellAttributesMask != 0) {
		SFGAOF shellAttributes = 0;
		ATLVERIFY(SUCCEEDED(pShellItem->GetAttributes(shellAttributesMask, &shellAttributes)));
		if(pCachedEnumSettings->shellAttributesMask != 0) {
			if(shellAttributes & SFGAO_FOLDER) {
				if(shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->excludedFolderItemShellAttributes)) {
					show = FALSE;
				} else if((shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->includedFolderItemShellAttributes)) != static_cast<SFGAOF>(pCachedEnumSettings->includedFolderItemShellAttributes)) {
					show = FALSE;
				}
			} else {
				if(shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->excludedFileItemShellAttributes)) {
					show = FALSE;
				} else if((shellAttributes & static_cast<SFGAOF>(pCachedEnumSettings->includedFileItemShellAttributes)) != static_cast<SFGAOF>(pCachedEnumSettings->includedFileItemShellAttributes)) {
					show = FALSE;
				}
			}
		}

		if(show && (shellAttributes & SFGAO_FILESYSTEM) && (pCachedEnumSettings->fileAttributesMask != 0)) {
			LPWSTR pPath = NULL;
			BOOL useFallBack = TRUE;
			if(FAILED(pShellItem->GetDisplayName(SIGDN_FILESYSPATH, &pPath))) {
				useFallBack = FALSE;				// SIC!
			} else {
				useFallBack = (pPath == NULL);
			}
			if(useFallBack) {
				if(pPath) {
					CoTaskMemFree(pPath);
					pPath = NULL;
				}
				PIDLIST_ABSOLUTE pIDL = NULL;
				CComQIPtr<IPersistIDList>(pShellItem)->GetIDList(&pIDL);
				if(pIDL) {
					pPath = static_cast<LPWSTR>(CoTaskMemAlloc(2048 * sizeof(WCHAR)));
					if(pPath) {
						ZeroMemory(pPath, 2048 * sizeof(WCHAR));
						SHGetPathFromIDListW(pIDL, pPath);
					}
					ILFree(pIDL);
				}
			}
			if(pPath) {
				BOOL isDrive = FALSE;
				if(PathIsRootW(pPath)) {
					int len = lstrlenW(pPath);
					if(len >= 2 && pPath[1] == L':') {
						isDrive = (len == 2 || (len == 3 && pPath[2] == L'\\'));
					}
				}
				DWORD fileAttributes = INVALID_FILE_ATTRIBUTES;
				if(!isDrive) {
					/* The values returned by GetFileAttributes for drives are a bit strange, e.g. it makes a
					   difference whether you pass C: or C:\. Also we don't really want to apply attribute filters to
					   drives. */
					fileAttributes = GetFileAttributesW(pPath);
				}
				CoTaskMemFree(pPath);

				if(fileAttributes != INVALID_FILE_ATTRIBUTES) {
					if(fileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
						if(fileAttributes & static_cast<DWORD>(pCachedEnumSettings->excludedFolderItemFileAttributes)) {
							show = FALSE;
						} else if((fileAttributes & static_cast<DWORD>(pCachedEnumSettings->includedFolderItemFileAttributes)) != static_cast<DWORD>(pCachedEnumSettings->includedFolderItemFileAttributes)) {
							show = FALSE;
						}
					} else {
						if(fileAttributes & static_cast<DWORD>(pCachedEnumSettings->excludedFileItemFileAttributes)) {
							show = FALSE;
						} else if((fileAttributes & static_cast<DWORD>(pCachedEnumSettings->includedFileItemFileAttributes)) != static_cast<DWORD>(pCachedEnumSettings->includedFileItemFileAttributes)) {
							show = FALSE;
						}
					}
				}
			}
		}
	}
	return show;
}

HasSubItems HasAtLeastOneSubItem(PIDLIST_ABSOLUTE pAbsolutePIDL, CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pCachedEnumSettings, CachedEnumSettings);

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	if(IsSlowItem(pAbsolutePIDL, TRUE, TRUE, NULL, &pParentISF, &pRelativePIDL)) {
		return hsiMaybe;
	}

	HRESULT hr = S_OK;
	if(pRelativePIDL == reinterpret_cast<PCUITEMID_CHILD>(-1)) {
		// IsSlowItem did not retrieve pParentISF and pRelativePIDL
		pParentISF = NULL;
		pRelativePIDL = NULL;
		hr = SHBindToParent(pAbsolutePIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	} else if(!pRelativePIDL) {
		// IsSlowItem did retrieve pParentISF and pRelativePIDL, but an error occured
		hr = E_FAIL;
	}
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		return HasAtLeastOneSubItem(pParentISF, const_cast<PITEMID_CHILD>(pRelativePIDL), pAbsolutePIDL, pCachedEnumSettings, timeCritical, TRUE);
	}
	return hsiMaybe;
}

HasSubItems HasAtLeastOneSubItem(LPSHELLFOLDER pParentISF, PITEMID_CHILD pRelativePIDL, PIDLIST_ABSOLUTE pAbsolutePIDL, CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical, BOOL skipSlowItemCheck/* = FALSE*/)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pCachedEnumSettings, CachedEnumSettings);

	if(!skipSlowItemCheck) {
		if(IsSlowItem(pParentISF, pRelativePIDL, pAbsolutePIDL, TRUE, TRUE)) {
			return hsiMaybe;
		}
	}

	CComPtr<IShellFolder> pNamespaceISF = NULL;
	if(ILIsEmpty(pRelativePIDL)) {
		SHGetDesktopFolder(&pNamespaceISF);
	} else {
		pParentISF->BindToObject(pRelativePIDL, NULL, IID_PPV_ARGS(&pNamespaceISF));
	}
	if(!pNamespaceISF) {
		return hsiNo;
	}

	// get the enumerator
	CComPtr<IEnumIDList> pEnum = NULL;
	pNamespaceISF->EnumObjects(NULL, pCachedEnumSettings->enumerationFlags, &pEnum);
	// will be NULL if access was denied
	if(!pEnum) {
		return hsiNo;
	}

	ULONG fetched = 0;
	pRelativePIDL = NULL;
	DWORD enumerationBegin = GetTickCount();
	while((pEnum->Next(1, &pRelativePIDL, &fetched) == S_OK) && (fetched > 0)) {
		BOOL b = ShouldShowItem(pNamespaceISF, pRelativePIDL, pCachedEnumSettings);
		ILFree(pRelativePIDL);
		if(b) {
			return hsiYes;
		}

		if(timeCritical && (GetTickCount() - enumerationBegin >= 100)) {
			// abort
			return hsiMaybe;
		}
	}
	return hsiNo;
}

HasSubItems HasAtLeastOneSubItem(IShellItem* pShellItem, CachedEnumSettings* pCachedEnumSettings, BOOL timeCritical, BOOL skipSlowItemCheck/* = FALSE*/)
{
	ATLASSUME(pShellItem);
	ATLASSERT_POINTER(pCachedEnumSettings, CachedEnumSettings);

	if(!skipSlowItemCheck) {
		if(IsSlowItem(pShellItem, TRUE, TRUE)) {
			return hsiMaybe;
		}
	}

	SFGAOF attributesMask = SFGAO_FOLDER;
	SFGAOF attributes = 0;
	pShellItem->GetAttributes(attributesMask, &attributes);
	if(!(attributes & attributesMask)) {
		return hsiNo;
	}

	// get the enumerator
	CComPtr<IEnumShellItems> pEnum = NULL;
	CComPtr<IBindCtx> pBindContext;
	// Note that STR_ENUM_ITEMS_FLAGS is ignored on Windows Vista and 7! That's why we use this code path only on Windows 8 and newer.
	if(SUCCEEDED(CreateDwordBindCtx(STR_ENUM_ITEMS_FLAGS, pCachedEnumSettings->enumerationFlags, &pBindContext))) {
		pShellItem->BindToHandler(pBindContext, BHID_EnumItems, IID_PPV_ARGS(&pEnum));
	}
	// will be NULL if access was denied
	if(!pEnum) {
		return hsiNo;
	}

	ULONG fetched = 0;
	IShellItem* pChildItem = NULL;
	DWORD enumerationBegin = GetTickCount();
	while((pEnum->Next(1, &pChildItem, &fetched) == S_OK) && (fetched > 0)) {
		BOOL b = ShouldShowItem(pChildItem, pCachedEnumSettings);
		pChildItem->Release();
		pChildItem = NULL;
		if(b) {
			return hsiYes;
		}

		if(timeCritical && (GetTickCount() - enumerationBegin >= 100)) {
			// abort
			return hsiMaybe;
		}
	}
	return hsiNo;
}


int CALLBACK FreeDPAColumnElement(LPVOID pColumnData, LPVOID /*pParam*/)
{
	if(pColumnData) {
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			if(reinterpret_cast<LPSHELLLISTVIEWCOLUMNDATA>(pColumnData)->pPropertyDescription) {
				reinterpret_cast<LPSHELLLISTVIEWCOLUMNDATA>(pColumnData)->pPropertyDescription->Release();
			}
		#endif
		delete pColumnData;
	}
	return 1;
}

int CALLBACK FreeDPAEnumeratedItemElement(LPVOID pEnumeratedItem, LPVOID /*pParam*/)
{
	if(pEnumeratedItem) {
		if(reinterpret_cast<LPENUMERATEDITEM>(pEnumeratedItem)->pIDL) {
			ILFree(reinterpret_cast<LPENUMERATEDITEM>(pEnumeratedItem)->pIDL);
		}
		delete pEnumeratedItem;
	}
	return 1;
}

int CALLBACK FreeDPAPIDLElement(LPVOID pIDL, LPVOID /*pParam*/)
{
	if(pIDL) {
		ILFree(reinterpret_cast<PIDLIST_RELATIVE>(pIDL));
	}
	return 1;
}


SHGDNF DisplayNameTypeToSHGDNFlags(DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop)
{
	SHGDNF flags;
	if(displayNameType == dntAddressBarNameFollowSysSettings) {
		BOOL fullAddresses = FALSE;
		DWORD dataSize = sizeof(fullAddresses);
		SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CabinetState"), TEXT("FullPathAddress"), NULL, &fullAddresses, &dataSize);
		if(fullAddresses) {
			flags = SHGDN_NORMAL | SHGDN_FORADDRESSBAR | SHGDN_FORPARSING;
		} else {
			flags = SHGDN_INFOLDER | SHGDN_FORADDRESSBAR;
		}
	} else if(displayNameType == dntURL) {
		flags = SHGDN_NORMAL | SHGDN_FORPARSING;
	} else {
		flags = displayNameType;
		if(relativeToDesktop) {
			flags = flags | SHGDN_NORMAL;
		} else {
			flags = flags | SHGDN_INFOLDER;
		}
	}
	return flags;
}

SIGDN DisplayNameTypeToSIGDNFlags(DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop)
{
	switch(displayNameType) {
		case dntDisplayName:
			return (relativeToDesktop ? SIGDN_NORMALDISPLAY : SIGDN_PARENTRELATIVE);
			break;
		case dntEditingName:
			return (relativeToDesktop ? SIGDN_DESKTOPABSOLUTEEDITING : SIGDN_PARENTRELATIVEEDITING);
			break;
		case dntAddressBarNameFollowSysSettings: {
			BOOL fullAddresses = FALSE;
			DWORD dataSize = sizeof(fullAddresses);
			SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CabinetState"), TEXT("FullPathAddress"), NULL, &fullAddresses, &dataSize);
			if(fullAddresses) {
				return static_cast<SIGDN>(SIGDN_PARENTRELATIVEFORADDRESSBAR & ~0x0001);
			} else {
				return SIGDN_PARENTRELATIVE;
			}
			break;
		}
		case dntAddressBarName:
			return (relativeToDesktop ? static_cast<SIGDN>(SIGDN_PARENTRELATIVEFORADDRESSBAR & ~0x0001) : SIGDN_PARENTRELATIVEFORADDRESSBAR);
			break;
		case dntParsingName:
			return (relativeToDesktop ? SIGDN_DESKTOPABSOLUTEPARSING : SIGDN_PARENTRELATIVEPARSING);
			break;
		case dntFileSystemPath:
			return SIGDN_FILESYSPATH;
			break;
		case dntURL:
			return SIGDN_URL;
			break;
	}
	return SIGDN_NORMALDISPLAY;
}

HRESULT GetDisplayName(PCIDLIST_ABSOLUTE pIDL, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, DisplayNameTypeConstants displayNameType, VARIANT_BOOL relativeToDesktop, BSTR* pDisplayName)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pDisplayName, BSTR);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	if(!pDisplayName) {
		return E_POINTER;
	}

	LPWSTR pBuffer = NULL;
	HRESULT hr = GetDisplayName(pIDL, pParentISF, pRelativePIDL, displayNameType, VARIANTBOOL2BOOL(relativeToDesktop), &pBuffer);
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		*pDisplayName = _bstr_t(pBuffer).Detach();
	} else {
		*pDisplayName = NULL;
	}
	// NOTE: CoTaskMemFree can handle NULL
	CoTaskMemFree(pBuffer);
	return hr;
}

HRESULT GetDisplayName(PCIDLIST_ABSOLUTE pIDL, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, DisplayNameTypeConstants displayNameType, BOOL relativeToDesktop, LPWSTR* ppDisplayName)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	//ATLASSUME(pParentISF);
	//ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(ppDisplayName, LPWSTR);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	if(!ppDisplayName) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	*ppDisplayName = NULL;
	BOOL useFallBack = TRUE;

	CComPtr<IShellItem> pShellItem;
	APIWrapper::SHCreateShellItem(NULL, NULL, reinterpret_cast<PCUITEMID_CHILD>(pIDL), &pShellItem, &hr);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pShellItem);
		SIGDN flags = DisplayNameTypeToSIGDNFlags(displayNameType, relativeToDesktop);

		hr = pShellItem->GetDisplayName(flags, ppDisplayName);
		if(FAILED(hr) && (flags == SIGDN_FILESYSPATH)) {
			hr = S_OK;
			useFallBack = FALSE;
		} else {
			useFallBack = (*ppDisplayName == NULL);
		}
	}

	if(useFallBack) {
		CoTaskMemFree(*ppDisplayName);
		hr = E_FAIL;
		if(displayNameType == dntFileSystemPath) {
			*ppDisplayName = static_cast<LPWSTR>(CoTaskMemAlloc(2048 * sizeof(WCHAR)));
			if(*ppDisplayName) {
				ZeroMemory(*ppDisplayName, 2048 * sizeof(WCHAR));
				SHGetPathFromIDListW(pIDL, *ppDisplayName);
				hr = S_OK;
			} else {
				hr = E_OUTOFMEMORY;
			}
		} else if(pParentISF && pRelativePIDL) {
			SHGDNF flags = DisplayNameTypeToSHGDNFlags(displayNameType, relativeToDesktop);

			STRRET displayName = {0};
			displayName.uType = STRRET_OFFSET;
			hr = pParentISF->GetDisplayNameOf(pRelativePIDL, flags, &displayName);
			ATLASSERT(SUCCEEDED(hr));
			hr = StrRetToStrW(&displayName, pRelativePIDL, ppDisplayName);
			if(*ppDisplayName) {
				if(displayNameType == dntURL) {
					WCHAR pURLBuffer[INTERNET_MAX_URL_LENGTH];
					ZeroMemory(pURLBuffer, INTERNET_MAX_URL_LENGTH * sizeof(WCHAR));
					DWORD bufferSize = INTERNET_MAX_URL_LENGTH;
					hr = UrlCreateFromPathW(*ppDisplayName, pURLBuffer, &bufferSize, NULL);
					CoTaskMemFree(*ppDisplayName);
					*ppDisplayName = static_cast<LPWSTR>(CoTaskMemAlloc(INTERNET_MAX_URL_LENGTH * sizeof(WCHAR)));
					if(*ppDisplayName) {
						lstrcpynW(*ppDisplayName, pURLBuffer, INTERNET_MAX_URL_LENGTH);
						hr = S_OK;
					} else {
						hr = E_OUTOFMEMORY;
					}
				}
			}
		}
	}
	return hr;
}

HRESULT GetFileAttributes(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, ItemFileAttributesConstants mask, ItemFileAttributesConstants* pAttributes)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT(pAttributes);
	if(!pParentISF || !pRelativePIDL) {
		return E_INVALIDARG;
	}
	if(!pAttributes) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	*pAttributes = static_cast<ItemFileAttributesConstants>(0);

	LPWSTR pPath = NULL;
	ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
	if(pPath) {
		DWORD fileAttributes = GetFileAttributesW(pPath);
		CoTaskMemFree(pPath);

		if(fileAttributes == INVALID_FILE_ATTRIBUTES) {
			hr = HRESULT_FROM_WIN32(GetLastError());
		} else {
			*pAttributes = static_cast<ItemFileAttributesConstants>(fileAttributes & mask);
			hr = S_OK;
		}
	}
	return hr;
}

HRESULT GetFileSystemPath(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, LPWSTR* ppPath)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(ppPath, LPWSTR);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	if(!ppPath) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	BOOL freePIDL = FALSE;
	if(resolveShortcuts) {
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
			SFGAOF attributes = SFGAO_LINK;
			if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes)) && (attributes & SFGAO_LINK)) {
				PIDLIST_ABSOLUTE pIDLTarget = NULL;
				hr = GetLinkTarget(hWndShellUIParentWindow, pIDL, &pIDLTarget);
				if(SUCCEEDED(hr)) {
					pIDL = pIDLTarget;
					freePIDL = TRUE;
				} else {
					return hr;
				}
			}
		}
	}

	*ppPath = NULL;
	BOOL useFallBack = TRUE;

	CComPtr<IShellItem> pShellItem;
	APIWrapper::SHCreateShellItem(NULL, NULL, reinterpret_cast<PCUITEMID_CHILD>(pIDL), &pShellItem, &hr);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pShellItem);
		hr = pShellItem->GetDisplayName(SIGDN_FILESYSPATH, ppPath);
		if(FAILED(hr)) {
			hr = S_OK;
			useFallBack = FALSE;
		} else {
			useFallBack = (*ppPath == NULL);
		}
	}

	if(useFallBack) {
		CoTaskMemFree(*ppPath);
		*ppPath = static_cast<LPWSTR>(CoTaskMemAlloc(2048 * sizeof(WCHAR)));
		if(*ppPath) {
			ZeroMemory(*ppPath, 2048 * sizeof(WCHAR));
			SHGetPathFromIDListW(pIDL, *ppPath);
			hr = S_OK;
		} else {
			hr = E_OUTOFMEMORY;
		}
	}
	if(freePIDL) {
		CoTaskMemFree(pIDL);
	}
	return hr;
}

HRESULT GetFileSystemPath(HWND hWndShellUIParentWindow, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, BOOL resolveShortcuts, LPWSTR* ppPath)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(ppPath, LPWSTR);
	if(!pParentISF || !pRelativePIDL) {
		return E_INVALIDARG;
	}
	if(!ppPath) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IShellFolder> pParentISFToUse = NULL;
	if(resolveShortcuts) {
		SFGAOF attributes = SFGAO_LINK;
		if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes)) && (attributes & SFGAO_LINK)) {
			PIDLIST_ABSOLUTE pIDLTarget = NULL;
			PIDLIST_ABSOLUTE pIDL = MakeAbsolutePIDL(pParentISF, pRelativePIDL);
			if(pIDL) {
				if(SUCCEEDED(GetLinkTarget(hWndShellUIParentWindow, pIDL, &pIDLTarget))) {
					hr = SHBindToParent(pIDLTarget, IID_PPV_ARGS(&pParentISFToUse), &pRelativePIDL);
					if(FAILED(hr)) {
						ILFree(pIDL);
						CoTaskMemFree(pIDLTarget);
						return hr;
					}
					CoTaskMemFree(pIDLTarget);
				}
				ILFree(pIDL);
			}
		} else {
			// no shortcut
			pParentISF->QueryInterface(IID_PPV_ARGS(&pParentISFToUse));
		}
	} else {
		pParentISF->QueryInterface(IID_PPV_ARGS(&pParentISFToUse));
	}
	if(!pParentISFToUse) {
		return E_INVALIDARG;
	}

	*ppPath = NULL;
	BOOL useFallBack = TRUE;

	CComPtr<IShellItem> pShellItem;
	APIWrapper::SHCreateShellItem(NULL, pParentISFToUse, pRelativePIDL, &pShellItem, &hr);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pShellItem);
		hr = pShellItem->GetDisplayName(SIGDN_FILESYSPATH, ppPath);
		if(FAILED(hr)) {
			hr = S_OK;
			useFallBack = FALSE;
		} else {
			useFallBack = (*ppPath == NULL);
		}
	}

	if(useFallBack) {
		CoTaskMemFree(*ppPath);

		hr = E_FAIL;
		PIDLIST_ABSOLUTE pIDL = MakeAbsolutePIDL(pParentISFToUse, pRelativePIDL);
		if(pIDL) {
			*ppPath = static_cast<LPWSTR>(CoTaskMemAlloc(2048 * sizeof(WCHAR)));
			if(*ppPath) {
				ZeroMemory(*ppPath, 2048 * sizeof(WCHAR));
				SHGetPathFromIDListW(pIDL, *ppPath);
				hr = S_OK;
			} else {
				hr = E_OUTOFMEMORY;
			}
			ILFree(pIDL);
		}
	}
	return hr;
}

HRESULT GetInfoTip(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, InfoTipFlagsConstants infoTipFlags, LPWSTR* ppInfoTip)
{
	ATLASSERT(IsWindow(hWndShellUIParentWindow));
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(ppInfoTip, LPWSTR);

	*ppInfoTip = NULL;
	if(infoTipFlags & itfNoInfoTipFollowSystemSettings) {
		infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags & ~itfNoInfoTipFollowSystemSettings);
		SHELLSTATE shellState = {0};
		SHGetSetSettings(&shellState, SSF_SHOWINFOTIP, FALSE);
		if(shellState.fShowInfoTip) {
			infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags & ~itfNoInfoTip);
		} else {
			infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags | itfNoInfoTip);
		}
	}
	if(infoTipFlags & itfNoInfoTip) {
		return S_OK;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		CComPtr<IQueryInfo> pQueryInfo = NULL;
		hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IQueryInfo, 0, reinterpret_cast<LPVOID*>(&pQueryInfo));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pQueryInfo);

			if(infoTipFlags & itfAllowSlowInfoTipFollowSysSettings) {
				infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags & ~itfAllowSlowInfoTipFollowSysSettings);
				BOOL slowInfoTips = FALSE;
				DWORD dataSize = sizeof(slowInfoTips);
				SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"), TEXT("FolderContentsInfoTip"), NULL, &slowInfoTips, &dataSize);
				if(slowInfoTips) {
					infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags | itfAllowSlowInfoTip);
				} else {
					infoTipFlags = static_cast<InfoTipFlagsConstants>(infoTipFlags & ~itfAllowSlowInfoTip);
				}
			}

			DWORD flags = static_cast<DWORD>(infoTipFlags);
			hr = pQueryInfo->GetInfoTip(flags, ppInfoTip);
		}
	}
	return hr;
}

HRESULT GetLinkTarget(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, LPTSTR pBuffer, int bufferSize)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pBuffer, TCHAR);
	if(!pBuffer) {
		return E_POINTER;
	}

	pBuffer[0] = TEXT('\0');

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		CComPtr<IShellLink> pLink = NULL;
		hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IShellLink, 0, reinterpret_cast<LPVOID*>(&pLink));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pLink);
			WIN32_FIND_DATA findData = {0};
			hr = pLink->GetPath(pBuffer, bufferSize, &findData, SLGP_UNCPRIORITY);
			if(hr != S_OK) {
				// try to get the pIDL
				PIDLIST_ABSOLUTE targetPIDL = NULL;
				hr = pLink->GetIDList(&targetPIDL);
				if(SUCCEEDED(hr)) {
					// get the parsing name
					ATLASSERT(targetPIDL);
					pParentISF = NULL;
					pRelativePIDL = NULL;
					hr = SHBindToParent(targetPIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
					if(SUCCEEDED(hr)) {
						ATLASSUME(pParentISF);
						ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
						LPWSTR pTarget = NULL;
						hr = GetDisplayName(targetPIDL, pParentISF, pRelativePIDL, dntParsingName, TRUE, &pTarget);
						if(SUCCEEDED(hr)) {
							ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, bufferSize, CW2T(pTarget))));
						}
						// NOTE: CoTaskMemFree can handle NULL
						CoTaskMemFree(pTarget);
					}
					ILFree(targetPIDL);
				}
			}
		}
	}

	if(FAILED(hr)) {
		LPWSTR pPath = NULL;
		hr = GetFileSystemPath(hWndShellUIParentWindow, pParentISF, pRelativePIDL, FALSE, &pPath);
		ATLASSERT(SUCCEEDED(hr));
		if(pPath) {
			HANDLE hReparsePoint = CreateFileW(pPath, 0, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT, NULL);
			CoTaskMemFree(pPath);

			if(hReparsePoint != INVALID_HANDLE_VALUE) {
				DWORD writtenBytes;
				LPBYTE pReparseBuffer = static_cast<LPBYTE>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, MAXIMUM_REPARSE_DATA_BUFFER_SIZE));
				if(pReparseBuffer) {
					BOOL b = DeviceIoControl(hReparsePoint, FSCTL_GET_REPARSE_POINT, NULL, 0, pReparseBuffer, MAXIMUM_REPARSE_DATA_BUFFER_SIZE, &writtenBytes, NULL);
					if(b) {
						typedef struct _REPARSE_DATA_BUFFER
						{
							DWORD ReparseTag;
							WORD ReparseDataLength;
							WORD Reserved;
							union
							{
								struct
								{
									WORD SubstituteNameOffset;
									WORD SubstituteNameLength;
									WORD PrintNameOffset;
									WORD PrintNameLength;
									WCHAR PathBuffer[1];
								} SymbolicLinkReparseBuffer;
								struct
								{
									WORD SubstituteNameOffset;
									WORD SubstituteNameLength;
									WORD PrintNameOffset;
									WORD PrintNameLength;
									WCHAR PathBuffer[1];
								} MountPointReparseBuffer;
								struct
								{
									BYTE DataBuffer[1];
								} GenericReparseBuffer;
							};
						} REPARSE_DATA_BUFFER, *PREPARSE_DATA_BUFFER;

						PREPARSE_DATA_BUFFER pReparseDetails = reinterpret_cast<PREPARSE_DATA_BUFFER>(pReparseBuffer);
						int size = min(bufferSize, pReparseDetails->MountPointReparseBuffer.SubstituteNameLength + 1);
						#ifdef UNICODE
							lstrcpynW(pBuffer, reinterpret_cast<LPWSTR>(reinterpret_cast<LPBYTE>(pReparseDetails->MountPointReparseBuffer.PathBuffer) + pReparseDetails->MountPointReparseBuffer.SubstituteNameOffset), size);
						#else
							LPWSTR pUnicodeBuffer = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, size + 1));
							if(pUnicodeBuffer) {
								lstrcpynW(pUnicodeBuffer, reinterpret_cast<LPWSTR>(reinterpret_cast<LPBYTE>(pReparseDetails->MountPointReparseBuffer.PathBuffer) + pReparseDetails->MountPointReparseBuffer.SubstituteNameOffset), size);
								CW2T converter(pUnicodeBuffer);
								lstrcpyn(pBuffer, converter, size);
								HeapFree(GetProcessHeap(), 0, pUnicodeBuffer);
							}
						#endif
					}
					HeapFree(GetProcessHeap(), 0, pReparseBuffer);
				}
				CloseHandle(hReparsePoint);
			}
		}
	}

	return hr;
}

HRESULT GetLinkTarget(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, PIDLIST_ABSOLUTE* ppIDLTarget)
{
	ATLASSERT_POINTER(ppIDLTarget, PCIDLIST_ABSOLUTE);
	if(!ppIDLTarget) {
		return E_POINTER;
	}

	*ppIDLTarget = NULL;

	// try to get the pIDL directly
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		CComPtr<IShellLink> pLink = NULL;
		hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IShellLink, 0, reinterpret_cast<LPVOID*>(&pLink));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pLink);
			hr = pLink->GetIDList(ppIDLTarget);
		}
	}

	if(FAILED(hr)) {     // shouldn't happen
		// get the target's path and create a pIDL from it
		TCHAR pBuffer[MAX_PATH];
		hr = GetLinkTarget(hWndShellUIParentWindow, pIDL, pBuffer, MAX_PATH);
		if(SUCCEEDED(hr)) {
			*ppIDLTarget = APIWrapper::ILCreateFromPath_LFN(pBuffer);
		}
	}
	return hr;
}

HRESULT GetDropTarget(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, LPDROPTARGET* ppDropTarget, BOOL forceCreateViewObject)
{
	ATLASSERT(IsWindow(hWndShellUIParentWindow));
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(ppDropTarget, LPDROPTARGET);
	if(!ppDropTarget) {
		return E_POINTER;
	}

	HRESULT hr;
	*ppDropTarget = NULL;

	if(ILIsEmpty(pIDL)) {
		// special handling for the Desktop
		CComPtr<IShellFolder> pDesktopISF = NULL;
		hr = SHGetDesktopFolder(&pDesktopISF);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pDesktopISF);
			hr = pDesktopISF->CreateViewObject(hWndShellUIParentWindow, IID_PPV_ARGS(ppDropTarget));
		}
	} else {
		if(forceCreateViewObject) {
			CComPtr<IShellFolder> pNamespaceISF;
			hr = BindToPIDL(pIDL, IID_PPV_ARGS(&pNamespaceISF));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pNamespaceISF);
				hr = pNamespaceISF->CreateViewObject(hWndShellUIParentWindow, IID_PPV_ARGS(ppDropTarget));
			}
		} else {
			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
			if(SUCCEEDED(hr)) {
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
				hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IDropTarget, NULL, reinterpret_cast<LPVOID*>(ppDropTarget));
			}
		}
	}

	return hr;
}

#ifdef ACTIVATE_SECZONE_FEATURES
	HRESULT GetZone(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, URLZONE* pZone)
	{
		ATLASSERT_POINTER(pZone, URLZONE);
		if(!pZone) {
			return E_POINTER;
		}

		if(ILIsEmpty(pIDL)) {
			*pZone = URLZONE_INVALID;
			return S_OK;
		}

		CComPtr<IInternetSecurityManager> pSecurityManager;
		HRESULT hr = pSecurityManager.CoCreateInstance(CLSID_InternetSecurityManager);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pSecurityManager);
			LPWSTR pUrl = NULL;

			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
			ATLASSERT(SUCCEEDED(hr));

			/*CComPtr<IShellView> pView = NULL;
			hr = pParentISF->CreateViewObject(hWndShellUIParentWindow, IID_PPV_ARGS(&pView));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pView);

				CComObject<DummyShellBrowser>* pBrowser = NULL;
				CComObject<DummyShellBrowser>::CreateInstance(&pBrowser);
				pBrowser->hWnd = hWndShellUIParentWindow;
				pBrowser->AddRef();

				FOLDERSETTINGS settings = {0};
				settings.ViewMode = FVM_LIST;
				RECT rc = {0};
				HWND hWndView = NULL;
				hr = pView->CreateViewWindow(NULL, &settings, pBrowser, &rc, &hWndView);
				ATLASSERT(SUCCEEDED(hr));
				hr = pView->UIActivate(SVUIA_ACTIVATE_FOCUS);
				ATLASSERT(SUCCEEDED(hr));
				hr = pView->UIActivate(SVUIA_DEACTIVATE);
				ATLASSERT(SUCCEEDED(hr));
				hr = pView->DestroyViewWindow();
				ATLASSERT(SUCCEEDED(hr));

				*pZone = static_cast<URLZONE>(HIWORD(pBrowser->zone));
				pBrowser->Release();
			}*/

			if(SUCCEEDED(hr)) {
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

				BOOL isFolderLink = FALSE;
				SFGAOF attributes = SFGAO_FOLDER | SFGAO_LINK;
				hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				if(SUCCEEDED(hr)) {
					isFolderLink = ((attributes & (SFGAO_FOLDER | SFGAO_LINK)) == (SFGAO_FOLDER | SFGAO_LINK));
				}

				if(!isFolderLink) {
					CComPtr<IUri> pUri;
					hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IUri, NULL, reinterpret_cast<LPVOID*>(&pUri));
					if(SUCCEEDED(hr) && pUri) {
						BSTR uri;
						hr = pUri->GetAbsoluteUri(&uri);
						if(SUCCEEDED(hr)) {
							UINT l = SysStringByteLen(uri) + sizeof(WCHAR);
							pUrl = static_cast<LPWSTR>(CoTaskMemAlloc(l));
							if(pUrl) {
								lstrcpynW(pUrl, OLE2W(uri), l);
							}
							SysFreeString(uri);
						}
					}
				}

				if(!pUrl && isFolderLink && APIWrapper::IsSupported_PSGetPropertyDescription() && APIWrapper::IsSupported_PSGetPropertyValue() && APIWrapper::IsSupported_PropVariantToString()) {
					CComPtr<IPropertyStore> pPropertyStore;
					hr = pParentISF->BindToObject(pRelativePIDL, NULL, IID_PPV_ARGS(&pPropertyStore));

					if(SUCCEEDED(hr) && pPropertyStore) {
						CComPtr<IPropertyDescription> pPropertyDescription;
						APIWrapper::PSGetPropertyDescription(PKEY_Link_TargetParsingPath, IID_PPV_ARGS(&pPropertyDescription), &hr);
						if(SUCCEEDED(hr)) {
							PROPVARIANT propertyValue;
							PropVariantInit(&propertyValue);
							APIWrapper::PSGetPropertyValue(pPropertyStore, pPropertyDescription, &propertyValue, &hr);
							if(SUCCEEDED(hr)) {
								pUrl = static_cast<LPWSTR>(CoTaskMemAlloc(2048 * sizeof(WCHAR)));
								APIWrapper::PropVariantToString(propertyValue, pUrl, 2048, &hr);
								if(FAILED(hr)) {
									CoTaskMemFree(pUrl);
									pUrl = NULL;
								}
								PropVariantClear(&propertyValue);
							}
						}
					}
				}

				if(!pUrl) {
					hr = GetDisplayName(pIDL, pParentISF, pRelativePIDL, dntParsingName, TRUE, &pUrl);
					ATLASSERT(SUCCEEDED(hr));
				}
			}

			if(SUCCEEDED(hr) && pUrl) {
				hr = pSecurityManager->MapUrlToZone(pUrl, reinterpret_cast<LPDWORD>(pZone), MUTZ_ISFILE | MUTZ_DONT_UNESCAPE);
				if(SUCCEEDED(hr)) {
					/*if(*pZone == URLZONE_INTERNET) {
						// HACK: the system reports Internet zone for virtual items like My Computer
						SFGAOF attributes = SFGAO_FILESYSTEM;
						if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes))) {
							if(!(attributes & SFGAO_FILESYSTEM)) {
								*pZone = URLZONE_LOCAL_MACHINE;
							}
						}
					} else */if(*pZone < URLZONE_INVALID || *pZone > URLZONE_USER_MAX) {
						*pZone = URLZONE_INVALID;
					}
				}
				CoTaskMemFree(pUrl);
			}
		}
		return hr;
	}
#endif

HRESULT GetMIMEType(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, LPWSTR* ppMIMEType)
{
	*ppMIMEType = NULL;

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		CComPtr<IQueryAssociations> pIQueryAssociations = NULL;
		if(resolveShortcuts) {
			SFGAOF attributes = SFGAO_LINK;
			if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes)) && (attributes & SFGAO_LINK)) {
				PIDLIST_ABSOLUTE pIDLTarget = NULL;
				if(SUCCEEDED(GetLinkTarget(hWndShellUIParentWindow, pIDL, &pIDLTarget))) {
					pParentISF = NULL;
					pRelativePIDL = NULL;
					if(SUCCEEDED(SHBindToParent(pIDLTarget, IID_PPV_ARGS(&pParentISF), &pRelativePIDL))) {
						ATLASSUME(pParentISF);
						ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
						hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
					}
					CoTaskMemFree(pIDLTarget);
				}
			} else {
				// no shortcut
				hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
			}
		} else {
			hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
		}

		if(SUCCEEDED(hr)) {
			ATLASSUME(pIQueryAssociations);
			DWORD bufferSize = 0;
			hr = pIQueryAssociations->GetString(ASSOCF_INIT_DEFAULTTOSTAR, ASSOCSTR_CONTENTTYPE, NULL, NULL, &bufferSize);
			if(SUCCEEDED(hr)) {
				*ppMIMEType = static_cast<LPWSTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
				hr = pIQueryAssociations->GetString(ASSOCF_INIT_DEFAULTTOSTAR, ASSOCSTR_CONTENTTYPE, NULL, *ppMIMEType, &bufferSize);
				ATLASSERT(SUCCEEDED(hr));
			}
		}
	}
	return hr;
}

HRESULT GetPerceivedType(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, PERCEIVED* pPerceivedType, LPWSTR* ppPerceivedTypeString)
{
	*pPerceivedType = PERCEIVED_TYPE_UNSPECIFIED;
	if(ppPerceivedTypeString) {
		*ppPerceivedTypeString = NULL;
	}

	/* NOTE: The perceived type could also be retrieved using the property system (PKEY_PerceivedType), but it
	         does not seem to provide the perceived type string. */

	LPWSTR pPath = NULL;
	HRESULT hr = GetFileSystemPath(hWndShellUIParentWindow, pIDL, resolveShortcuts, &pPath);
	if(SUCCEEDED(hr)) {
		WCHAR pStar[2] = {L'*', L'\0'};
		LPWSTR pExtension = PathFindExtensionW(pPath);
		if(!pExtension || lstrlenW(pExtension) == 0) {
			pExtension = pStar;
		}
		PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
		hr = AssocGetPerceivedType(pExtension, pPerceivedType, &flags, ppPerceivedTypeString);
		CoTaskMemFree(pPath);
	}
	return hr;
}

HRESULT GetThumbnailAdornment(LPWSTR pPerceivedType, LPDWORD pAdornment)
{
	*pAdornment = 1;			// drop shadow is the default

	DWORD dataType = REG_DWORD;
	DWORD dataSize = sizeof(*pAdornment);
	//LSTATUS status = RegGetValueW(HKEY_CLASSES_ROOT, pPerceivedType, L"Treatment", RRF_RT_DWORD, &dataType, pAdornment, &dataSize);
	HKEY hKey = NULL;
	LSTATUS status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pPerceivedType, 0, KEY_QUERY_VALUE, &hKey);
	if(status == ERROR_SUCCESS) {
		status = RegQueryValueExW(hKey, L"Treatment", NULL, &dataType, reinterpret_cast<LPBYTE>(pAdornment), &dataSize);
		RegCloseKey(hKey);
	}
	if(status != ERROR_SUCCESS) {
		WCHAR pBuffer[256];
		lstrcpynW(pBuffer, L"SystemFileAssociations\\", 256);
		ATLVERIFY(SUCCEEDED(StringCchCatW(pBuffer, 256, pPerceivedType)));
		status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pBuffer, 0, KEY_QUERY_VALUE, &hKey);
		if(status == ERROR_SUCCESS) {
			status = RegQueryValueExW(hKey, L"Treatment", NULL, &dataType, reinterpret_cast<LPBYTE>(pAdornment), &dataSize);
			RegCloseKey(hKey);
		}
	}
	return HRESULT_FROM_WIN32(status);
}

HRESULT GetThumbnailAdornment(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, LPDWORD pAdornment, LPWSTR* ppTypeOverlay)
{
	/* TODO: Currently we test .ext, then SystemFileAssociations\.ext, then <type>, then
	 *       SystemFileAssociations\<type>. This might be wrong and .ext, <type>, SystemFileAssociations\.ext,
	 *       SystemFileAssociations\<type> might be correct.
	 */

	*pAdornment = 1;			// drop shadow is the default

	LPWSTR pPath = NULL;
	HRESULT hr = GetFileSystemPath(hWndShellUIParentWindow, pIDL, resolveShortcuts, &pPath);
	if(FAILED(hr)) {
		return hr;
	} else if(!pPath) {
		// this happens e.g. in the Fonts folder
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
			CComQIPtr<IShellFolder2> pParentISF2 = pParentISF;
			if(pParentISF2) {
				VARIANT v;
				hr = pParentISF2->GetDetailsEx(pRelativePIDL, &PKEY_FileName, &v);
				if(v.vt == VT_BSTR) {
					UINT l = SysStringByteLen(v.bstrVal) + sizeof(WCHAR);
					pPath = static_cast<LPWSTR>(CoTaskMemAlloc(l));
					if(pPath) {
						lstrcpynW(pPath, OLE2W(v.bstrVal), l);
					}
				}
				VariantClear(&v);
			}
		}
		if(FAILED(hr)) {
			return E_FAIL;
		}
	}

	hr = E_FAIL;
	LPWSTR pExtension = PathFindExtensionW(pPath);
	if(pExtension) {
		hr = GetThumbnailAdornment(pExtension, pAdornment);
	}
	if(FAILED(hr)) {
		DWORD dataType = REG_SZ;
		DWORD dataSize = 0;
		//LSTATUS status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, NULL, &dataSize);
		HKEY hKey = NULL;
		LSTATUS status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pExtension, 0, KEY_QUERY_VALUE, &hKey);
		if(status == ERROR_SUCCESS) {
			status = RegQueryValueExW(hKey, NULL, NULL, &dataType, NULL, &dataSize);
			if(status == ERROR_SUCCESS) {
				LPWSTR pFileTypeClass = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, dataSize));
				if(pFileTypeClass) {
					//status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, pFileTypeClass, &dataSize);
					status = RegQueryValueExW(hKey, NULL, NULL, &dataType, reinterpret_cast<LPBYTE>(pFileTypeClass), &dataSize);
					if(status == ERROR_SUCCESS) {
						hr = GetThumbnailAdornment(pFileTypeClass, pAdornment);
					}
					HeapFree(GetProcessHeap(), 0, pFileTypeClass);
				}
			}
			RegCloseKey(hKey);
		}
	}
	if(FAILED(hr)) {
		WCHAR pStar[2] = {L'*', L'\0'};
		LPWSTR pExtension2 = PathFindExtensionW(pPath);
		if(!pExtension2 || lstrlenW(pExtension2) == 0) {
			pExtension2 = pStar;
		}
		PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
		PERCEIVED perceivedType = PERCEIVED_TYPE_UNSPECIFIED;
		LPWSTR pPerceivedTypeString = NULL;
		hr = AssocGetPerceivedType(pExtension2, &perceivedType, &flags, &pPerceivedTypeString);
		if(SUCCEEDED(hr)) {
			hr = GetThumbnailAdornment(pPerceivedTypeString, pAdornment);
			CoTaskMemFree(pPerceivedTypeString);
		}
	}

	if(ppTypeOverlay) {
		hr = E_FAIL;
		if(pExtension) {
			hr = GetThumbnailTypeOverlay(pExtension, ppTypeOverlay);
		}
		if(FAILED(hr)) {
			DWORD dataType = REG_SZ;
			DWORD dataSize = 0;
			//LSTATUS status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, NULL, &dataSize);
			HKEY hKey = NULL;
			LSTATUS status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pExtension, 0, KEY_QUERY_VALUE, &hKey);
			if(status == ERROR_SUCCESS) {
				status = RegQueryValueExW(hKey, NULL, NULL, &dataType, NULL, &dataSize);
				if(status == ERROR_SUCCESS) {
					LPWSTR pFileTypeClass = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, dataSize));
					if(pFileTypeClass) {
						//status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, pFileTypeClass, &dataSize);
						status = RegQueryValueExW(hKey, NULL, NULL, &dataType, reinterpret_cast<LPBYTE>(pFileTypeClass), &dataSize);
						if(status == ERROR_SUCCESS) {
							hr = GetThumbnailTypeOverlay(pFileTypeClass, ppTypeOverlay);
						}
						HeapFree(GetProcessHeap(), 0, pFileTypeClass);
					}
				}
				RegCloseKey(hKey);
			}
		}
		if(FAILED(hr)) {
			WCHAR pStar[2] = {L'*', L'\0'};
			LPWSTR pExtension2 = PathFindExtensionW(pPath);
			if(!pExtension2 || lstrlenW(pExtension2) == 0) {
				pExtension2 = pStar;
			}
			PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
			PERCEIVED perceivedType = PERCEIVED_TYPE_UNSPECIFIED;
			LPWSTR pPerceivedTypeString = NULL;
			hr = AssocGetPerceivedType(pExtension2, &perceivedType, &flags, &pPerceivedTypeString);
			if(SUCCEEDED(hr)) {
				hr = GetThumbnailTypeOverlay(pPerceivedTypeString, ppTypeOverlay);
				CoTaskMemFree(pPerceivedTypeString);
			}
		}
	}

	CoTaskMemFree(pPath);
	return S_OK;
}

HRESULT GetThumbnailTypeOverlay(LPWSTR pPerceivedType, LPWSTR* ppTypeOverlay)
{
	*ppTypeOverlay = NULL;			// use the executable's icon by default

	DWORD dataType = REG_SZ;
	DWORD dataSize = 0;
	//LSTATUS status = RegGetValueW(HKEY_CLASSES_ROOT, pPerceivedType, L"TypeOverlay", RRF_RT_REG_SZ, &dataType, NULL, &dataSize);
	HKEY hKey = NULL;
	LSTATUS status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pPerceivedType, 0, KEY_QUERY_VALUE, &hKey);
	if(status == ERROR_SUCCESS) {
		status = RegQueryValueExW(hKey, L"TypeOverlay", NULL, &dataType, NULL, &dataSize);
		if(status == ERROR_SUCCESS) {
			*ppTypeOverlay = static_cast<LPWSTR>(CoTaskMemAlloc(dataSize));
			if(*ppTypeOverlay) {
				//status = RegGetValueW(HKEY_CLASSES_ROOT, pPerceivedType, L"TypeOverlay", RRF_RT_REG_SZ, &dataType, *ppTypeOverlay, &dataSize);
				status = RegQueryValueExW(hKey, L"TypeOverlay", NULL, &dataType, reinterpret_cast<LPBYTE>(*ppTypeOverlay), &dataSize);
			} else {
				return E_OUTOFMEMORY;
			}
		}
		RegCloseKey(hKey);
	}
	if(status != ERROR_SUCCESS) {
		if(*ppTypeOverlay) {
			CoTaskMemFree(*ppTypeOverlay);
			*ppTypeOverlay = NULL;
		}

		WCHAR pBuffer[256];
		lstrcpynW(pBuffer, L"SystemFileAssociations\\", 256);
		ATLVERIFY(SUCCEEDED(StringCchCatW(pBuffer, 256, pPerceivedType)));
		status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pBuffer, 0, KEY_QUERY_VALUE, &hKey);
		if(status == ERROR_SUCCESS) {
			//status = RegGetValueW(HKEY_CLASSES_ROOT, pBuffer, L"TypeOverlay", RRF_RT_REG_SZ, &dataType, NULL, &dataSize);
			status = RegQueryValueExW(hKey, L"TypeOverlay", NULL, &dataType, NULL, &dataSize);
			if(status == ERROR_SUCCESS) {
				*ppTypeOverlay = static_cast<LPWSTR>(CoTaskMemAlloc(dataSize));
				if(*ppTypeOverlay) {
					//status = RegGetValueW(HKEY_CLASSES_ROOT, pBuffer, L"TypeOverlay", RRF_RT_REG_SZ, &dataType, *ppTypeOverlay, &dataSize);
					status = RegQueryValueExW(hKey, L"TypeOverlay", NULL, &dataType, reinterpret_cast<LPBYTE>(*ppTypeOverlay), &dataSize);
					if(status != ERROR_SUCCESS) {
						CoTaskMemFree(*ppTypeOverlay);
						*ppTypeOverlay = NULL;
					}
				} else {
					return E_OUTOFMEMORY;
				}
			}
			RegCloseKey(hKey);
		}
	}
	return HRESULT_FROM_WIN32(status);
}

HRESULT GetThumbnailTypeOverlay(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE pIDL, BOOL resolveShortcuts, LPWSTR* ppTypeOverlay)
{
	/* TODO: Currently we test .ext, then SystemFileAssociations\.ext, then <type>, then
	 *       SystemFileAssociations\<type>. This might be wrong and .ext, <type>, SystemFileAssociations\.ext,
	 *       SystemFileAssociations\<type> might be correct.
	 */

	*ppTypeOverlay = NULL;			// use the executable's icon by default

	LPWSTR pPath = NULL;
	HRESULT hr = GetFileSystemPath(hWndShellUIParentWindow, pIDL, resolveShortcuts, &pPath);
	if(FAILED(hr)) {
		return hr;
	}

	hr = E_FAIL;
	LPWSTR pExtension = PathFindExtensionW(pPath);
	if(pExtension) {
		hr = GetThumbnailTypeOverlay(pExtension, ppTypeOverlay);
	}
	if(FAILED(hr)) {
		DWORD dataType = REG_SZ;
		DWORD dataSize = 0;
		//LSTATUS status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, NULL, &dataSize);
		HKEY hKey = NULL;
		LSTATUS status = RegOpenKeyExW(HKEY_CLASSES_ROOT, pExtension, 0, KEY_QUERY_VALUE, &hKey);
		if(status == ERROR_SUCCESS) {
			status = RegQueryValueExW(hKey, NULL, NULL, &dataType, NULL, &dataSize);
			if(status == ERROR_SUCCESS) {
				LPWSTR pFileTypeClass = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, dataSize));
				if(pFileTypeClass) {
					//status = RegGetValueW(HKEY_CLASSES_ROOT, pExtension, NULL, RRF_RT_REG_SZ, &dataType, pFileTypeClass, &dataSize);
					status = RegQueryValueExW(hKey, NULL, NULL, &dataType, reinterpret_cast<LPBYTE>(pFileTypeClass), &dataSize);
					if(status == ERROR_SUCCESS) {
						hr = GetThumbnailTypeOverlay(pFileTypeClass, ppTypeOverlay);
					}
					HeapFree(GetProcessHeap(), 0, pFileTypeClass);
				}
			}
			RegCloseKey(hKey);
		}
	}
	if(FAILED(hr)) {
		WCHAR pStar[2] = {L'*', L'\0'};
		LPWSTR pExtension2 = PathFindExtensionW(pPath);
		if(!pExtension2 || lstrlenW(pExtension2) == 0) {
			pExtension2 = pStar;
		}
		PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
		PERCEIVED perceivedType = PERCEIVED_TYPE_UNSPECIFIED;
		LPWSTR pPerceivedTypeString = NULL;
		hr = AssocGetPerceivedType(pExtension2, &perceivedType, &flags, &pPerceivedTypeString);
		if(SUCCEEDED(hr)) {
			hr = GetThumbnailTypeOverlay(pPerceivedTypeString, ppTypeOverlay);
			CoTaskMemFree(pPerceivedTypeString);
		}
	}
	CoTaskMemFree(pPath);
	return S_OK;
}

/*HICON LoadIconResource(LPWSTR pResourceReference, int desiredWidth, int desiredHeight)
{
	LPWSTR pSeparator = StrStrW(pResourceReference, L"@,");
	ZeroMemory(pSeparator, 2 * sizeof(WCHAR));
	UINT resourceID = static_cast<UINT>(_wtoi(pSeparator + 2));
	HICON hIcon = ExtractIconW(ModuleHelper::GetResourceInstance(), pResourceReference, resourceID);
	return hIcon;
}*/

void FlushSystemImageList(void)
{
	static DWORD timeOfLastCacheFlush = 0;
	DWORD currentTime = GetTickCount();
	if(currentTime - timeOfLastCacheFlush > 10000) {
		HKEY hKey = NULL;
		RegCreateKeyEx(HKEY_CURRENT_USER, TEXT("Control Panel\\Desktop\\WindowMetrics"), 0, NULL, 0, KEY_QUERY_VALUE | KEY_SET_VALUE, NULL, &hKey, NULL);
		if(hKey) {
			TCHAR pIconSize[10];
			LONG dataSize = 10;
			LSTATUS result = RegQueryValue(hKey, TEXT("Shell Icon Size"), pIconSize, &dataSize);
			BOOL valueExisted = (result != ERROR_FILE_NOT_FOUND);
			if(!valueExisted) {
				lstrcpyn(pIconSize, TEXT("32"), 10);
				result = ERROR_SUCCESS;
			}
			if(result == ERROR_SUCCESS) {
				int iconSize = _ttoi(pIconSize);
				if(SUCCEEDED(StringCchPrintf(pIconSize, 10, TEXT("%i"), iconSize + 1))) {
					timeOfLastCacheFlush = currentTime;
					RegSetValueEx(hKey, TEXT("Shell Icon Size"), 0, REG_SZ, reinterpret_cast<LPBYTE>(pIconSize), dataSize);
					APIWrapper::FileIconInit(TRUE, NULL);
					if(valueExisted) {
						if(SUCCEEDED(StringCchPrintf(pIconSize, 10, TEXT("%i"), iconSize))) {
							RegSetValueEx(hKey, TEXT("Shell Icon Size"), 0, REG_SZ, reinterpret_cast<LPBYTE>(pIconSize), dataSize);
							APIWrapper::FileIconInit(TRUE, NULL);
						}
					} else {
						RegDeleteValue(hKey, TEXT("Shell Icon Size"));
						APIWrapper::FileIconInit(TRUE, NULL);
					}
				}
			}
			RegCloseKey(hKey);
		}
	}
}

int GetOverlayIndex(PCIDLIST_ABSOLUTE pIDL)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	ATLVERIFY(SUCCEEDED(SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
	return GetOverlayIndex(pParentISF, pRelativePIDL);
}

int GetOverlayIndex(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	CComQIPtr<IShellIconOverlay> pParentISIO = pParentISF;
	if(pParentISIO) {
		// the folder implements IShellIconOverlay
		int overlayIndex = 0;
		if(pParentISIO->GetOverlayIndex(pRelativePIDL, &overlayIndex) == S_OK) {
			return overlayIndex;
		}
	} else {
		// fallback solution
		SFGAOF attributes = SFGAO_LINK | SFGAO_SHARE;
		pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
		if(attributes & SFGAO_LINK) {
			return 2;
		} else if(attributes & SFGAO_SHARE) {
			return 1;
		}
	}
	return 0;
}


HRESULT BindToPIDL(PCIDLIST_ABSOLUTE pIDL, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	if(!pIDL) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(ppInterfaceImpl, LPVOID);
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	*ppInterfaceImpl = NULL;

	HRESULT hr;
	if(ILIsEmpty(pIDL)) {
		CComPtr<IShellFolder> pDesktopISF = NULL;
		hr = SHGetDesktopFolder(&pDesktopISF);
		if(pDesktopISF) {
			if(requiredInterface == IID_IShellFolder) {
				hr = pDesktopISF->QueryInterface(requiredInterface, ppInterfaceImpl);
			} else {
				hr = pDesktopISF->BindToObject(pIDL, NULL, requiredInterface, ppInterfaceImpl);
			}
		}
	} else {
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
		if(pParentISF && pRelativePIDL) {
			hr = pParentISF->BindToObject(pRelativePIDL, NULL, requiredInterface, ppInterfaceImpl);
		}
	}
	return hr;
}

BOOL CanBeCustomized(PCIDLIST_ABSOLUTE pIDL)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	const int nonCustomizableFolders[] = {
		CSIDL_WINDOWS,
		CSIDL_SYSTEM,
		CSIDL_SYSTEMX86,
		CSIDL_PROGRAM_FILES,
		CSIDL_PROGRAM_FILESX86,
		-1
	};

	if(SHRestricted(REST_NOCUSTOMIZEWEBVIEW) || SHRestricted(REST_NOCUSTOMIZETHISFOLDER) || ILIsEmpty(pIDL)) {
		return FALSE;
	}

	TCHAR pPath[2048];
	ZeroMemory(pPath, 2048);
	if(!SHGetPathFromIDList(pIDL, pPath)) {
		return FALSE;
	}

	TCHAR pBuffer[MAX_PATH];
	for(int i = 0; nonCustomizableFolders[i] != -1; ++i) {
		SHGetFolderPath(NULL, nonCustomizableFolders[i] | CSIDL_FLAG_DONT_VERIFY, NULL, SHGFP_TYPE_CURRENT, pBuffer);
		if(!lstrcmpi(pBuffer, pPath)) {
			return FALSE;
		}
	}

	PathAppend(pPath, TEXT("Desktop.ini"));
	BOOL fileExists = PathFileExists(pPath);
	if(fileExists) {
		DWORD attributes = GetFileAttributes(pPath);
		if(attributes != INVALID_FILE_ATTRIBUTES) {
			return ((attributes & FILE_ATTRIBUTE_READONLY) == 0);
		}
	}

	HANDLE hFile = CreateFile(pPath, GENERIC_WRITE, FILE_SHARE_READ, NULL, (fileExists ? OPEN_EXISTING : CREATE_NEW), FILE_ATTRIBUTE_NORMAL, NULL);
	if(hFile != INVALID_HANDLE_VALUE) {
		CloseHandle(hFile);
		if(!fileExists) {
			DeleteFile(pPath);
		}
		return TRUE;
	}
	return FALSE;
}

HRESULT CustomizeFolder(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	TCHAR pPath[2048];
	ZeroMemory(pPath, 2048);
	if(!SHGetPathFromIDList(pIDL, pPath)) {
		return E_INVALIDARG;
	}

	BOOL isWindowsXPOrNewer = TRUE;/*FALSE;
	OSVERSIONINFOEX osVersion;
	osVersion.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	if(!GetVersionEx(reinterpret_cast<LPOSVERSIONINFO>(&osVersion))) {
		osVersion.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		GetVersionEx(reinterpret_cast<LPOSVERSIONINFO>(&osVersion));
	}
	if(osVersion.dwPlatformId == VER_PLATFORM_WIN32_NT) {
		if(osVersion.dwMajorVersion == 5) {
			isWindowsXPOrNewer = (osVersion.dwMinorVersion >= 1);
		} else {
			isWindowsXPOrNewer = (osVersion.dwMajorVersion > 5);
		}
	}*/

	BOOL succeeded = FALSE;
	if(isWindowsXPOrNewer) {
		LPWSTR pCustomizeTabTitle = NULL;
		HINSTANCE hShell32 = LoadLibraryEx(TEXT("shell32.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
		if(hShell32) {
			pCustomizeTabTitle = MLLoadDialogTitle(hShell32, TEXT("#1124"));
			FreeLibrary(hShell32);
		}
		ATLASSERT(pCustomizeTabTitle);
		CT2W converter(pPath);
		LPWSTR pBuffer = converter;
		succeeded = SHObjectProperties(hWndShellUIParentWindow, SHOP_FILEPATH, pBuffer, pCustomizeTabTitle);
		if(pCustomizeTabTitle) {
			delete[] pCustomizeTabTitle;
		}
	} else {
		TCHAR pParam[MAX_PATH + 20];
		// the command line is of the form <dirname>;<Parent-hWnd>
		wnsprintf(pParam, sizeof(pParam), TEXT("%s;%d"), pPath, reinterpret_cast<INT_PTR>(hWndShellUIParentWindow));

		SHELLEXECUTEINFO executionInfo = {0};
		executionInfo.cbSize = sizeof(SHELLEXECUTEINFO);
		executionInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
		executionInfo.hwnd = hWndShellUIParentWindow;
		executionInfo.lpFile = TEXT("IESHWIZ.EXE");
		executionInfo.lpParameters = pParam;
		executionInfo.lpDirectory = pPath;
		executionInfo.nShow = SW_SHOWNORMAL;
		succeeded = ShellExecuteEx(&executionInfo);
	}
	return (succeeded ? S_OK : E_FAIL);
}

/*int CompareFullIDLists(LPARAM lParam, PCUIDLIST_RELATIVE pIDL1, PCUIDLIST_RELATIVE pIDL2, BOOL checkingForEquality)
{
	// TODO: on Vista, try to use SHCompareIDsFull
	ATLASSERT_POINTER(pIDL1, ITEMIDLIST_RELATIVE);
	ATLASSERT_POINTER(pIDL2, ITEMIDLIST_RELATIVE);

	if(checkingForEquality) {
		// compare the item counts
		PCUIDLIST_RELATIVE p = pIDL1;
		UINT c1 = 0;
		while(!ILIsEmpty(p)) {
			p = ILNext(p);
			++c1;
		}
		p = pIDL2;
		UINT c2 = 0;
		while(!ILIsEmpty(p)) {
			p = ILNext(p);
			++c2;
		}
		if(c1 != c2) {
			// items can't be equal
			return -1;
		}
	}

	LPSHELLFOLDER pParentISF = NULL;
	ATLVERIFY(SUCCEEDED(SHGetDesktopFolder(&pParentISF)));
	ATLASSUME(pParentISF);
	LPSHELLFOLDER pBuffer = NULL;

	int result = 0;
	while(TRUE) {
		BOOL isEmpty1 = ILIsEmpty(pIDL1);
		BOOL isEmpty2 = ILIsEmpty(pIDL2);
		if(isEmpty1 && isEmpty2) {
			result = 0;
			break;
		} else if(isEmpty1) {
			result = -1;
			break;
		} else if(isEmpty2) {
			result = 1;
			break;
		}

		PITEMID_CHILD pIDLFirst1 = ILCloneFirst(pIDL1);
		PITEMID_CHILD pIDLFirst2 = ILCloneFirst(pIDL2);
		result = static_cast<short>(HRESULT_CODE(pParentISF->CompareIDs(lParam, pIDLFirst1, pIDLFirst2)));
		if(result != 0) {
			ILFree(pIDLFirst1);
			ILFree(pIDLFirst2);
			break;
		}
		ATLVERIFY(SUCCEEDED(pParentISF->BindToObject(pIDLFirst1, NULL, IID_PPV_ARGS(&pBuffer))));
		pParentISF->Release();
		pParentISF = pBuffer;
		ATLASSUME(pParentISF);
		ILFree(pIDLFirst1);
		ILFree(pIDLFirst2);

		pIDL1 = ILNext(pIDL1);
		pIDL2 = ILNext(pIDL2);
	}
	pParentISF->Release();
	return result;
}*/

HRESULT CALLBACK ContextMenuCallback(LPSHELLFOLDER /*pShellFolder*/, HWND /*hWnd*/, LPDATAOBJECT /*pDataObject*/, UINT message, WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	HRESULT hr = E_NOTIMPL;
	switch(message) {
		case DFM_MERGECONTEXTMENU:
			hr = S_OK;
			break;
		case DFM_INVOKECOMMAND:
		case DFM_INVOKECOMMANDEX:
		case DFM_GETDEFSTATICID:
			hr = S_FALSE;
			break;
	}
	return hr;
}

HRESULT GetCrossNamespaceContextMenu(HWND hWndShellUIParentWindow, PIDLIST_ABSOLUTE* ppIDLs, UINT pIDLCount, IContextMenuCB* pCallbackObject, LPSHELLFOLDER* ppParentISF, LPCONTEXTMENU* ppContextMenu)
{
	ATLASSERT(IsWindow(hWndShellUIParentWindow));
	ATLASSERT(pIDLCount > 0);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PIDLIST_ABSOLUTE, pIDLCount);
	ATLASSUME(pCallbackObject);
	ATLASSERT_POINTER(ppParentISF, LPSHELLFOLDER);
	ATLASSERT_POINTER(ppContextMenu, LPCONTEXTMENU);
	if(!ppIDLs || !ppParentISF || !ppContextMenu) {
		return E_INVALIDARG;
	}

	LPSHELLFOLDER pParentISF = NULL;
	ATLVERIFY(SUCCEEDED(DesktopShellFolderHook::CreateInstance(&pParentISF)));
	ATLASSUME(pParentISF);
	*ppParentISF = pParentISF;

	HRESULT hr = E_FAIL;
	if(APIWrapper::IsSupported_SHCreateDefaultContextMenu()) {
		DEFCONTEXTMENU details = {0};
		details.hwnd = hWndShellUIParentWindow;
		details.pcmcb = pCallbackObject;
		//details.pidlFolder = NULL;
		details.psf = pParentISF;
		details.cidl = pIDLCount;
		details.apidl = reinterpret_cast<PCUITEMID_CHILD_ARRAY>(ppIDLs);
		/*details.punkAssociationInfo = NULL;
		details.cKeys = 0;
		details.aKeys = NULL;*/
		APIWrapper::SHCreateDefaultContextMenu(&details, IID_IContextMenu, reinterpret_cast<LPVOID*>(ppContextMenu), &hr);
	} else if(APIWrapper::IsSupported_CDefFolderMenu_Create2()) {
		PIDLIST_ABSOLUTE pIDLDesktop = GetDesktopPIDL();
		ATLASSERT_POINTER(pIDLDesktop, ITEMIDLIST_ABSOLUTE);
		if(pIDLDesktop) {
			TCHAR pPath[2048];
			LPTSTR pExtension = NULL;
			BOOL isFolder = FALSE;

			LPSHELLFOLDER pFocusedItemParentISF;
			PCUITEMID_CHILD pIDLFocusedItem = NULL;
			hr = SHBindToParent(ppIDLs[0], IID_PPV_ARGS(&pFocusedItemParentISF), &pIDLFocusedItem);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pFocusedItemParentISF);
				ATLASSERT_POINTER(pIDLFocusedItem, ITEMID_CHILD);

				SFGAOF attributes = SFGAO_FOLDER | SFGAO_FILESYSTEM;
				if(SUCCEEDED(pFocusedItemParentISF->GetAttributesOf(1, &pIDLFocusedItem, &attributes))) {
					if(attributes & SFGAO_FOLDER) {
						isFolder = TRUE;
					} else {
						if(attributes & SFGAO_FILESYSTEM) {
							if(SHGetPathFromIDList(ppIDLs[0], pPath)) {
								pExtension = PathFindExtension(pPath);
							}
						} else {
							// TODO: Should we do anything here?
						}
					}
				}
			}

			// CDefFolderMenu_Create2 is limited to 16 keys
			HKEY pKeys[16];
			ZeroMemory(pKeys, 16 * sizeof(HKEY));
			UINT keyCount = 0;

			if(isFolder) {
				// use the Folder file type
				HKEY hKeyExtension = NULL;
				if(RegOpenKeyEx(HKEY_CLASSES_ROOT, TEXT("Folder"), 0, KEY_READ, &hKeyExtension) == ERROR_SUCCESS) {
					pKeys[keyCount++] = hKeyExtension;
				}
			} else {
				BOOL addUnknownKey = TRUE;

				if(pExtension && lstrlen(pExtension) > 0) {
					HKEY hKeyExtension = NULL;
					if(RegOpenKeyEx(HKEY_CLASSES_ROOT, pExtension, 0, KEY_READ, &hKeyExtension) == ERROR_SUCCESS) {
						addUnknownKey = FALSE;
						// check for a file class
						DWORD size = 0;
						RegQueryValueEx(hKeyExtension, NULL, NULL, NULL, NULL, &size);
						if(size > 0) {
							// it has one - open it
							LPBYTE pBuffer = new BYTE[size];
							if(RegQueryValueEx(hKeyExtension, NULL, NULL, NULL, pBuffer, &size) == ERROR_SUCCESS) {
								LPTSTR pClass = reinterpret_cast<LPTSTR>(pBuffer);
								HKEY hKeyFileClass = NULL;
								if(RegOpenKeyEx(HKEY_CLASSES_ROOT, pClass, 0, KEY_READ, &hKeyFileClass) == ERROR_SUCCESS) {
									// check for a version
									HKEY hKeyVersion = NULL;
									if(RegOpenKeyEx(hKeyFileClass, TEXT("CurVer"), 0, KEY_READ, &hKeyVersion) == ERROR_SUCCESS) {
										// it has one - open it
										size = 0;
										RegQueryValueEx(hKeyVersion, NULL, NULL, NULL, NULL, &size);
										if(size > 0) {
											LPBYTE pBuffer2 = new BYTE[size];
											if(RegQueryValueEx(hKeyVersion, NULL, NULL, NULL, pBuffer2, &size) == ERROR_SUCCESS) {
												LPTSTR pVersion = reinterpret_cast<LPTSTR>(pBuffer2);
												HKEY hKeyVersionedFileClass = NULL;
												if(RegOpenKeyEx(HKEY_CLASSES_ROOT, pVersion, 0, KEY_READ, &hKeyVersionedFileClass) == ERROR_SUCCESS) {
													pKeys[keyCount++] = hKeyVersionedFileClass;
												}
											}
											delete[] pBuffer2;
										}
										RegCloseKey(hKeyVersion);
									} else {
										pKeys[keyCount++] = hKeyFileClass;
									}
								}
							}
							delete[] pBuffer;
						} else {
							// no file class
							pKeys[keyCount++] = hKeyExtension;
						}

						HKEY hKeySysFileAssociations = NULL;
						if(RegOpenKeyEx(HKEY_CLASSES_ROOT, TEXT("SystemFileAssociations"), 0, KEY_READ, &hKeySysFileAssociations) == ERROR_SUCCESS) {
							HKEY hKeySysFileAssociation = NULL;
							if(RegOpenKeyEx(hKeySysFileAssociations, pExtension, 0, KEY_READ, &hKeySysFileAssociation) == ERROR_SUCCESS) {
								pKeys[keyCount++] = hKeySysFileAssociation;
							}
							size = 0;
							RegQueryValueEx(hKeyExtension, TEXT("PerceivedType"), NULL, NULL, NULL, &size);
							if(size > 0) {
								LPBYTE pBuffer = new BYTE[size];
								if(RegQueryValueEx(hKeyExtension, TEXT("PerceivedType"), NULL, NULL, pBuffer, &size) == ERROR_SUCCESS) {
									LPTSTR pPerceivedType = reinterpret_cast<LPTSTR>(pBuffer);
									hKeySysFileAssociation = NULL;
									if(RegOpenKeyEx(hKeySysFileAssociations, pPerceivedType, 0, KEY_READ, &hKeySysFileAssociation) == ERROR_SUCCESS) {
										pKeys[keyCount++] = hKeySysFileAssociation;
									}
								}
								delete[] pBuffer;
							}
							RegCloseKey(hKeySysFileAssociations);
						}
					}
					if(pKeys[0] != hKeyExtension) {
						RegCloseKey(hKeyExtension);
					}
				}

				if(addUnknownKey) {
					// use the Unknown file type
					HKEY hKeyExtension = NULL;
					if(RegOpenKeyEx(HKEY_CLASSES_ROOT, TEXT("Unknown"), 0, KEY_READ, &hKeyExtension) == ERROR_SUCCESS) {
						pKeys[keyCount++] = hKeyExtension;
					}
				}
				ATLASSERT(keyCount <= 14);

				HKEY hKey = NULL;
				if(RegOpenKeyEx(HKEY_CLASSES_ROOT, TEXT("*"), 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
					pKeys[keyCount++] = hKey;
				}
				hKey = NULL;
				if(RegOpenKeyEx(HKEY_CLASSES_ROOT, TEXT("AllFilesystemObjects"), 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
					pKeys[keyCount++] = hKey;
				}
			}

			// TODO: https://groups.google.de/group/microsoft.public.platformsdk.shell/browse_thread/thread/32ae490ae4ea681d/3fb841aea3811256
			APIWrapper::CDefFolderMenu_Create2(pIDLDesktop, hWndShellUIParentWindow, pIDLCount, reinterpret_cast<PCUITEMID_CHILD_ARRAY>(ppIDLs), pParentISF, ContextMenuCallback, keyCount, pKeys, ppContextMenu, &hr);

			for(UINT i = 0; i < keyCount; ++i) {
				RegCloseKey(pKeys[i]);
			}
			ILFree(pIDLDesktop);
		}
	}

	return hr;
}

PIDLIST_ABSOLUTE GetDesktopPIDL(void)
{
	PIDLIST_ABSOLUTE pIDLDesktop = NULL;
	HRESULT hr = E_FAIL;
	APIWrapper::SHGetKnownFolderIDList(FOLDERID_Desktop, KF_FLAG_DONT_VERIFY, NULL, &pIDLDesktop, &hr);
	if(FAILED(hr)) {
		ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_DESKTOP, NULL, 0, &pIDLDesktop)));
	}
	return pIDLDesktop;
}

PIDLIST_ABSOLUTE GetGlobalRecycleBinPIDL(void)
{
	PIDLIST_ABSOLUTE pIDLRecycler = NULL;
	HRESULT hr = E_FAIL;
	APIWrapper::SHGetKnownFolderIDList(FOLDERID_RecycleBinFolder, KF_FLAG_DONT_VERIFY, NULL, &pIDLRecycler, &hr);
	if(FAILED(hr)) {
		ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_BITBUCKET, NULL, 0, &pIDLRecycler)));
	}
	return pIDLRecycler;
}

PIDLIST_ABSOLUTE GetMyComputerPIDL(void)
{
	PIDLIST_ABSOLUTE pIDLMyComputer = NULL;
	HRESULT hr = E_FAIL;
	APIWrapper::SHGetKnownFolderIDList(FOLDERID_ComputerFolder, KF_FLAG_DONT_VERIFY, NULL, &pIDLMyComputer, &hr);
	if(FAILED(hr)) {
		ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_DRIVES, NULL, 0, &pIDLMyComputer)));
	}
	return pIDLMyComputer;
}

PIDLIST_ABSOLUTE GetMyDocsPIDL(void)
{
	PIDLIST_ABSOLUTE pIDLMyDocs = NULL;
	HRESULT hr = E_FAIL;
	APIWrapper::SHGetKnownFolderIDList(FOLDERID_Documents, KF_FLAG_DONT_VERIFY, NULL, &pIDLMyDocs, &hr);
	if(FAILED(hr)) {
		ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_MYDOCUMENTS, NULL, 0, &pIDLMyDocs)));
	}
	return pIDLMyDocs;
}

PIDLIST_ABSOLUTE GetPrintersPIDL(void)
{
	PIDLIST_ABSOLUTE pIDLPrinters = NULL;
	HRESULT hr = E_FAIL;
	APIWrapper::SHGetKnownFolderIDList(FOLDERID_PrintersFolder, KF_FLAG_DONT_VERIFY, NULL, &pIDLPrinters, &hr);
	if(FAILED(hr)) {
		ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_PRINTERS, NULL, 0, &pIDLPrinters)));
	}
	return pIDLPrinters;
}

BOOL HaveSameParent(PCIDLIST_ABSOLUTE_ARRAY ppIDLs, UINT pIDLCount)
{
	ATLASSERT_ARRAYPOINTER(ppIDLs, PIDLIST_ABSOLUTE, pIDLCount);
	if(!ppIDLs) {
		return FALSE;
	}
	PIDLIST_ABSOLUTE pIDLParent = ILCloneFull(ppIDLs[0]);
	ATLASSERT_POINTER(pIDLParent, ITEMIDLIST_ABSOLUTE);
	if(!pIDLParent) {
		return FALSE;
	}
	ILRemoveLastID(pIDLParent);
	ATLASSERT_POINTER(pIDLParent, ITEMIDLIST_ABSOLUTE);

	BOOL ret = TRUE;
	for(UINT i = 1; i < pIDLCount; ++i) {
		if(!ILIsParent(pIDLParent, ppIDLs[i], TRUE)) {
			ret = FALSE;
			break;
		}
	}
	ILFree(pIDLParent);
	return ret;
}

UINT ILCount(PCUIDLIST_RELATIVE pIDL)
{
	UINT c = 0;
	while(!ILIsEmpty(pIDL)) {
		++c;
		pIDL = ILNext(pIDL);
	}
	return c;
}

BOOL IsElevationRequired(HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDL)
{
	ATLASSERT(IsWindow(hWndShellUIParentWindow));
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		PCITEMID_CHILD pIDLs[1] = {pRelativePIDL};
		CComPtr<IExtractIcon> pIExtractIcon = NULL;
		hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, pIDLs, IID_IExtractIcon, NULL, reinterpret_cast<LPVOID*>(&pIExtractIcon));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pIExtractIcon);
			TCHAR pBuffer[1024];
			int index;
			UINT flags = 0;
			pIExtractIcon->GetIconLocation(GIL_CHECKSHIELD, pBuffer, 1024, &index, &flags);
			return (flags & GIL_SHIELD);
		}
	}

	BOOL ret = FALSE;

	if(APIWrapper::IsSupported_IsElevationRequired()) {
		LPWSTR pPath = NULL;
		hr = GetFileSystemPath(hWndShellUIParentWindow, pParentISF, pRelativePIDL, TRUE, &pPath);
		ATLASSERT(SUCCEEDED(hr));
		if(pPath) {
			APIWrapper::IsElevationRequired(pPath, &ret);
			CoTaskMemFree(pPath);
		}
	}
	return ret;
}

BOOL IsSlowItem(PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive/* = NULL*/, LPSHELLFOLDER* ppParentISF/* = NULL*/, PCUITEMID_CHILD* ppRelativePIDL/* = NULL*/)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	if(ppParentISF) {
		*ppParentISF = NULL;
	}
	if(ppRelativePIDL) {
		*ppRelativePIDL = reinterpret_cast<PCUITEMID_CHILD>(-1);
	}
	if(!pAbsolutePIDL) {
		return FALSE;
	}
	if(pIsUnreachableNetDrive) {
		*pIsUnreachableNetDrive = FALSE;
	}

	TCHAR pPath[1024];
	ZeroMemory(pPath, 1024);

	// Treat any item as slow that is not on an ordinary drive.
	BOOL gotPath = SHGetPathFromIDList(pAbsolutePIDL, pPath);
	if(treatAnyVirtualItemsAsSlow && (!gotPath || PathGetDriveNumber(pPath) == -1)) {
		return TRUE;
	} else if(!treatAnyVirtualItemsAsSlow) {
		// recognize DLNA devices
		BOOL isDLNA = FALSE;
		LPWSTR pParsingName = NULL;
		if(SUCCEEDED(GetDisplayName(pAbsolutePIDL, NULL, NULL, dntParsingName, FALSE, &pParsingName)) && pParsingName) {
			int parsingNameLength = lstrlenW(pParsingName);
			LPCWSTR pDLNAPrefixWin7 = L"Provider\\Microsoft.Networking.SSDP//uuid:";
			int prefixLength = lstrlenW(pDLNAPrefixWin7);
			if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin7, prefixLength) == CSTR_EQUAL) {
				// Windows 7 style DLNA
				isDLNA = TRUE;
			}
			if(!isDLNA && RunTimeHelperEx::IsWin8()) {
				LPCWSTR pDLNAPrefixWin8 = L"uuid:";
				prefixLength = lstrlenW(pDLNAPrefixWin8);
				if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin8, prefixLength) == CSTR_EQUAL) {
					// Windows 8/8.1 style DLNA
					isDLNA = TRUE;
				}
			}
		}
		if(pParsingName) {
			CoTaskMemFree(pParsingName);
		}
		if(isDLNA) {
			return TRUE;
		}
	}

	// Is the item located on a slow drive?
	if(gotPath && PathStripToRoot(pPath)) {
		switch(GetDriveType(pPath)) {
			case DRIVE_REMOVABLE:
			case DRIVE_REMOTE:
			case DRIVE_CDROM:
				return TRUE;
				break;
			case DRIVE_NO_ROOT_DIR:					// NOTE: unreachable network drives are reported as DRIVE_NO_ROOT_DIR
				if(pIsUnreachableNetDrive) {
					*pIsUnreachableNetDrive = TRUE;
				}
				return TRUE;
				break;
		}
	}

	// Is it network stuff?
	if(IsSubItemOfNethood(pAbsolutePIDL, FALSE)) {
		return TRUE;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pAbsolutePIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	if(ppRelativePIDL) {
		*ppRelativePIDL = NULL;
		if(SUCCEEDED(hr)) {
			*ppRelativePIDL = pRelativePIDL;
		}
	}
	if(ppParentISF) {
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			pParentISF->QueryInterface(IID_PPV_ARGS(ppParentISF));
		}
	}
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		// Is it removable media or is it marked as being slow?
		SFGAOF attributes = SFGAO_ISSLOW | SFGAO_REMOVABLE;
		pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
		if(attributes & (SFGAO_ISSLOW | SFGAO_REMOVABLE)) {
			return TRUE;
		}

		if(checkForRemoteFolderLinks && IsRemoteFolderLink(pParentISF, pRelativePIDL)) {
			return TRUE;
		}
	}
	return FALSE;
}

BOOL IsSlowItem(LPSHELLFOLDER pParentISF, PITEMID_CHILD pRelativePIDL, PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive/* = NULL*/)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	if(!pParentISF || !pRelativePIDL || !pAbsolutePIDL) {
		return FALSE;
	}
	if(pIsUnreachableNetDrive) {
		*pIsUnreachableNetDrive = FALSE;
	}

	TCHAR pPath[1024];
	ZeroMemory(pPath, 1024);

	// Treat any item as slow that is not on an ordinary drive.
	BOOL gotPath = SHGetPathFromIDList(pAbsolutePIDL, pPath);
	if(treatAnyVirtualItemsAsSlow && (!gotPath || PathGetDriveNumber(pPath) == -1)) {
		return TRUE;
	} else if(!treatAnyVirtualItemsAsSlow) {
		// recognize DLNA devices
		BOOL isDLNA = FALSE;
		LPWSTR pParsingName = NULL;
		if(SUCCEEDED(GetDisplayName(pAbsolutePIDL, NULL, NULL, dntParsingName, FALSE, &pParsingName)) && pParsingName) {
			int parsingNameLength = lstrlenW(pParsingName);
			LPCWSTR pDLNAPrefixWin7 = L"Provider\\Microsoft.Networking.SSDP//uuid:";
			int prefixLength = lstrlenW(pDLNAPrefixWin7);
			if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin7, prefixLength) == CSTR_EQUAL) {
				// Windows 7 style DLNA
				isDLNA = TRUE;
			}
			if(!isDLNA && RunTimeHelperEx::IsWin8()) {
				LPCWSTR pDLNAPrefixWin8 = L"uuid:";
				prefixLength = lstrlenW(pDLNAPrefixWin8);
				if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin8, prefixLength) == CSTR_EQUAL) {
					// Windows 8/8.1 style DLNA
					isDLNA = TRUE;
				}
			}
		}
		if(pParsingName) {
			CoTaskMemFree(pParsingName);
		}
		if(isDLNA) {
			return TRUE;
		}
	}

	// Is the item located on a slow drive?
	if(PathStripToRoot(pPath)) {
		switch(GetDriveType(pPath)) {
			case DRIVE_REMOVABLE:
			case DRIVE_REMOTE:
			case DRIVE_CDROM:
				return TRUE;
				break;
			case DRIVE_NO_ROOT_DIR:					// NOTE: unreachable network drives are reported as DRIVE_NO_ROOT_DIR
				if(pIsUnreachableNetDrive) {
					*pIsUnreachableNetDrive = TRUE;
				}
				return TRUE;
				break;
		}
	}

	// Is it network stuff?
	if(IsSubItemOfNethood(pAbsolutePIDL, FALSE)) {
		return TRUE;
	}

	// Is it removable media or is it marked as being slow?
	SFGAOF attributes = SFGAO_ISSLOW | SFGAO_REMOVABLE;
	pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
	if(attributes & (SFGAO_ISSLOW | SFGAO_REMOVABLE)) {
		return TRUE;
	}

	if(checkForRemoteFolderLinks && IsRemoteFolderLink(pParentISF, pRelativePIDL)) {
		return TRUE;
	}
	return FALSE;
}

BOOL IsSlowItem(IShellItem* pShellItem, BOOL treatAnyVirtualItemsAsSlow, BOOL checkForRemoteFolderLinks, LPBOOL pIsUnreachableNetDrive/* = NULL*/)
{
	ATLASSUME(pShellItem);
	if(!pShellItem) {
		return FALSE;
	}
	PIDLIST_ABSOLUTE pAbsolutePIDL = NULL;
	CComQIPtr<IPersistIDList>(pShellItem)->GetIDList(&pAbsolutePIDL);
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	if(!pAbsolutePIDL) {
		return FALSE;
	}
	if(pIsUnreachableNetDrive) {
		*pIsUnreachableNetDrive = FALSE;
	}

	TCHAR pPath[1024];
	ZeroMemory(pPath, 1024);

	// Treat any item as slow that is not on an ordinary drive.
	BOOL gotPath = SHGetPathFromIDList(pAbsolutePIDL, pPath);
	if(treatAnyVirtualItemsAsSlow && (!gotPath || PathGetDriveNumber(pPath) == -1)) {
		return TRUE;
	} else if(!treatAnyVirtualItemsAsSlow) {
		// recognize DLNA devices
		BOOL isDLNA = FALSE;
		LPWSTR pParsingName = NULL;
		if(SUCCEEDED(GetDisplayName(pAbsolutePIDL, NULL, NULL, dntParsingName, FALSE, &pParsingName)) && pParsingName) {
			int parsingNameLength = lstrlenW(pParsingName);
			LPCWSTR pDLNAPrefixWin7 = L"Provider\\Microsoft.Networking.SSDP//uuid:";
			int prefixLength = lstrlenW(pDLNAPrefixWin7);
			if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin7, prefixLength) == CSTR_EQUAL) {
				// Windows 7 style DLNA
				isDLNA = TRUE;
			}
			if(!isDLNA && RunTimeHelperEx::IsWin8()) {
				LPCWSTR pDLNAPrefixWin8 = L"uuid:";
				prefixLength = lstrlenW(pDLNAPrefixWin8);
				if(CompareStringW(GetThreadLocale(), NORM_IGNORECASE, pParsingName, min(parsingNameLength, prefixLength), pDLNAPrefixWin8, prefixLength) == CSTR_EQUAL) {
					// Windows 8/8.1 style DLNA
					isDLNA = TRUE;
				}
			}
		}
		if(pParsingName) {
			CoTaskMemFree(pParsingName);
		}
		if(isDLNA) {
			return TRUE;
		}
	}

	// Is the item located on a slow drive?
	if(PathStripToRoot(pPath)) {
		switch(GetDriveType(pPath)) {
			case DRIVE_REMOVABLE:
			case DRIVE_REMOTE:
			case DRIVE_CDROM:
				return TRUE;
				break;
			case DRIVE_NO_ROOT_DIR:					// NOTE: unreachable network drives are reported as DRIVE_NO_ROOT_DIR
				if(pIsUnreachableNetDrive) {
					*pIsUnreachableNetDrive = TRUE;
				}
				return TRUE;
				break;
		}
	}

	// Is it network stuff?
	if(IsSubItemOfNethood(pAbsolutePIDL, FALSE)) {
		return TRUE;
	}

	// Is it removable media or is it marked as being slow?
	SFGAOF attributesMask = SFGAO_ISSLOW | SFGAO_REMOVABLE;
	SFGAOF attributes = 0;
	pShellItem->GetAttributes(attributesMask, &attributes);
	if(attributes & attributesMask) {
		return TRUE;
	}

	if(checkForRemoteFolderLinks && IsRemoteFolderLink(pShellItem)) {
		return TRUE;
	}
	return FALSE;
}

BOOL IsSubItemOfDriveRecycleBin(PCIDLIST_ABSOLUTE pAbsolutePIDL)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);

	BOOL ret = FALSE;

	LPTSTR pPath = new TCHAR[2048];
	ATLASSERT_ARRAYPOINTER(pPath, TCHAR, 2048);
	if(!pPath) {
		return FALSE;
	}
	ZeroMemory(pPath, 2048);
	SHGetPathFromIDList(pAbsolutePIDL, pPath);

	int pathLength = lstrlen(pPath);
	if(pathLength > 0) {
		// Are we on a hard disk?
		LPTSTR pRoot = new TCHAR[pathLength + 1];
		ATLASSERT_ARRAYPOINTER(pRoot, TCHAR, pathLength + 1);
		if(pRoot) {
			lstrcpyn(pRoot, pPath, pathLength + 1);
			if(PathStripToRoot(pRoot)) {
				if(GetDriveType(pRoot) == DRIVE_FIXED) {
					// get the name of the recycler folder
					TCHAR pRecyclerName[MAX_PATH];
					ZeroMemory(pRecyclerName, MAX_PATH);
					if(RunTimeHelper::IsVista()) {
						lstrcpyn(pRecyclerName, TEXT("$Recycle.Bin\\"), MAX_PATH);
					} else {
						TCHAR pFilesystem[MAX_PATH + 1];
						ZeroMemory(pFilesystem, MAX_PATH + 1);
						if(GetVolumeInformation(pRoot, NULL, 0, NULL, NULL, NULL, pFilesystem, MAX_PATH + 1)) {
							// TODO: What's up with drives converted from FAT to NTFS?
							if(lstrcmpi(pFilesystem, TEXT("NTFS")) == 0) {
								lstrcpyn(pRecyclerName, TEXT("RECYCLER\\"), MAX_PATH);
							} else {
								lstrcpyn(pRecyclerName, TEXT("RECYCLED\\"), MAX_PATH);
							}
						}
					}
					ATLASSERT(lstrlen(pRecyclerName) > 0);

					// build the recycler's path
					LPTSTR pRecyclerPath = new TCHAR[2048];
					ATLASSERT_ARRAYPOINTER(pRecyclerPath, TCHAR, 2048);
					if(pRecyclerPath) {
						ZeroMemory(pRecyclerPath, 2048);
						lstrcpyn(pRecyclerPath, pRoot, 2048);
						PathAddBackslash(pRecyclerPath);
						ATLVERIFY(SUCCEEDED(StringCchCat(pRecyclerPath, 2048, pRecyclerName)));

						// Does the passed pIDL's path start with the recycler's path?
						int recyclerPathLength = lstrlen(pRecyclerPath);
						if(pathLength > recyclerPathLength) {
							// truncate the path
							pPath[recyclerPathLength] = TEXT('\0');
						}
						// now compare
						ret = (lstrcmpi(pPath, pRecyclerPath) == 0);

						delete[] pRecyclerPath;
					}
				}
			}
			delete[] pRoot;
		}
	}
	delete[] pPath;
	return ret;
}

BOOL IsSubItemOfGlobalRecycleBin(PCIDLIST_ABSOLUTE pAbsolutePIDL)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);

	BOOL ret = FALSE;

	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();
	if(pIDLRecycler) {
		ret = ILIsParent(pIDLRecycler, pAbsolutePIDL, FALSE);
		ILFree(pIDLRecycler);
	}
	return ret;
}

BOOL IsSubItemOfMyComputer(PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL immediateSubItem)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);

	BOOL ret = FALSE;
	if(pAbsolutePIDL) {
		PIDLIST_ABSOLUTE pIDLComputer = NULL;
		HRESULT hr = E_FAIL;
		APIWrapper::SHGetKnownFolderIDList(FOLDERID_ComputerFolder, KF_FLAG_DONT_VERIFY, NULL, &pIDLComputer, &hr);
		if(FAILED(hr)) {
			ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_DRIVES, NULL, 0, &pIDLComputer)));
		}
		if(pIDLComputer) {
			ret = ILIsParent(pIDLComputer, pAbsolutePIDL, immediateSubItem);
			ILFree(pIDLComputer);
		}
	}
	return ret;
}

BOOL IsSubItemOfNethood(PIDLIST_ABSOLUTE pAbsolutePIDL, BOOL immediateSubItem)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);

	BOOL ret = FALSE;
	if(pAbsolutePIDL) {
		PIDLIST_ABSOLUTE pIDLNetwork = NULL;
		HRESULT hr = E_FAIL;
		APIWrapper::SHGetKnownFolderIDList(FOLDERID_NetworkFolder, KF_FLAG_DONT_VERIFY, NULL, &pIDLNetwork, &hr);
		if(FAILED(hr)) {
			ATLVERIFY(SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_NETWORK, NULL, 0, &pIDLNetwork)));
		}
		if(pIDLNetwork) {
			ret = ILIsParent(pIDLNetwork, pAbsolutePIDL, immediateSubItem);
			ILFree(pIDLNetwork);
		}
	}
	return ret;
}

BOOL IsRemoteFolderLink(PIDLIST_ABSOLUTE pAbsolutePIDL)
{
	ATLASSERT_POINTER(pAbsolutePIDL, ITEMIDLIST_ABSOLUTE);
	if(!pAbsolutePIDL) {
		return FALSE;
	}
	
	BOOL ret = FALSE;
	CComPtr<IShellItem> pShellItem = NULL;
	HRESULT hr = E_FAIL;
	if(APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		APIWrapper::SHCreateItemFromIDList(pAbsolutePIDL, IID_PPV_ARGS(&pShellItem), &hr);
	}
	if(SUCCEEDED(hr)) {
		ret = IsRemoteFolderLink(pShellItem);
	} else if(APIWrapper::IsSupported_PSGetPropertyDescription() && APIWrapper::IsSupported_PSGetPropertyValue()) {
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		hr = SHBindToParent(pAbsolutePIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

			ret = IsRemoteFolderLink(pParentISF, pRelativePIDL);
		}
	}
	return ret;
}

BOOL IsRemoteFolderLink(LPSHELLFOLDER pParentISF, PCITEMID_CHILD pRelativePIDL)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	if(!pParentISF || !pRelativePIDL) {
		return FALSE;
	}

	BOOL ret = FALSE;
	if(APIWrapper::IsSupported_PSGetPropertyDescription() && APIWrapper::IsSupported_PSGetPropertyValue() && APIWrapper::IsSupported_PropVariantToString()) {
		SFGAOF attributes = SFGAO_FOLDER | SFGAO_LINK;
		pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
		if((attributes & (SFGAO_FOLDER | SFGAO_LINK)) == (SFGAO_FOLDER | SFGAO_LINK)) {
			CComPtr<IPropertyStore> pPropertyStore;
			if(SUCCEEDED(pParentISF->BindToObject(pRelativePIDL, NULL, IID_PPV_ARGS(&pPropertyStore)))) {
				HRESULT hr;
				CComPtr<IPropertyDescription> pPropertyDescription;
				APIWrapper::PSGetPropertyDescription(PKEY_Link_TargetParsingPath, IID_PPV_ARGS(&pPropertyDescription), &hr);
				if(SUCCEEDED(hr)) {
					PROPVARIANT propertyValue;
					PropVariantInit(&propertyValue);
					APIWrapper::PSGetPropertyValue(pPropertyStore, pPropertyDescription, &propertyValue, &hr);
					if(SUCCEEDED(hr)) {
						WCHAR pBuffer[1024];
						APIWrapper::PropVariantToString(propertyValue, pBuffer, 1024, &hr);
						if(SUCCEEDED(hr)) {
							if(lstrlenW(pBuffer) > 0 && (PathIsUNCW(pBuffer) || PathIsURLW(pBuffer))) {
								ret = TRUE;
							}
						}
						PropVariantClear(&propertyValue);
					}
				}
			}
		}
	}
	return ret;
}

BOOL IsRemoteFolderLink(IShellItem* pShellItem)
{
	ATLASSUME(pShellItem);
	if(!pShellItem) {
		return FALSE;
	}

	BOOL ret = FALSE;
	if(APIWrapper::IsSupported_PSGetPropertyDescription() && APIWrapper::IsSupported_PSGetPropertyValue() && APIWrapper::IsSupported_PropVariantToString()) {
		SFGAOF attributesMask = SFGAO_FOLDER | SFGAO_LINK;
		SFGAOF attributes = 0;
		pShellItem->GetAttributes(attributesMask, &attributes);
		if((attributes & attributesMask) == attributesMask) {
			CComPtr<IPropertyStore> pPropertyStore;
			if(SUCCEEDED(pShellItem->BindToHandler(NULL, BHID_PropertyStore, IID_PPV_ARGS(&pPropertyStore)))) {
				HRESULT hr;
				CComPtr<IPropertyDescription> pPropertyDescription;
				APIWrapper::PSGetPropertyDescription(PKEY_Link_TargetParsingPath, IID_PPV_ARGS(&pPropertyDescription), &hr);
				if(SUCCEEDED(hr)) {
					PROPVARIANT propertyValue;
					PropVariantInit(&propertyValue);
					APIWrapper::PSGetPropertyValue(pPropertyStore, pPropertyDescription, &propertyValue, &hr);
					if(SUCCEEDED(hr)) {
						WCHAR pBuffer[1024];
						APIWrapper::PropVariantToString(propertyValue, pBuffer, 1024, &hr);
						if(SUCCEEDED(hr)) {
							if(lstrlenW(pBuffer) > 0 && (PathIsUNCW(pBuffer) || PathIsURLW(pBuffer))) {
								ret = TRUE;
							}
						}
						PropVariantClear(&propertyValue);
					}
				}
			}
		}
	}
	return ret;
}

PIDLIST_ABSOLUTE MakeAbsolutePIDL(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	PIDLIST_ABSOLUTE pIDL = NULL;

	CComQIPtr<IPersistIDList> pParentPersistIDList = pParentISF;
	if(pParentPersistIDList) {
		ATLVERIFY(SUCCEEDED(pParentPersistIDList->GetIDList(&pIDL)));
	}
	if(!pIDL) {
		CComQIPtr<IPersistFolder2> pParentPersistFolder2 = pParentISF;
		if(pParentPersistFolder2) {
			ATLVERIFY(SUCCEEDED(pParentPersistFolder2->GetCurFolder(&pIDL)));
		}
	}
	if(pIDL) {
		// append the relative pIDL
		pIDL = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(pIDL, reinterpret_cast<LPCSHITEMID>(pRelativePIDL), TRUE));
	} else {
		// okay, this really sucks
		STRRET path = {0};
		path.uType = STRRET_OFFSET;
		ATLVERIFY(SUCCEEDED(pParentISF->GetDisplayNameOf(pRelativePIDL, SHGDN_FORPARSING, &path)));
		LPTSTR pPath = NULL;
		ATLVERIFY(SUCCEEDED(StrRetToStr(&path, pRelativePIDL, &pPath)));
		pIDL = APIWrapper::ILCreateFromPath_LFN(pPath);
		CoTaskMemFree(pPath);
	}
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	return pIDL;
}

HRESULT SetFileAttributes(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pRelativePIDL, ItemFileAttributesConstants mask, ItemFileAttributesConstants attributes)
{
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	HRESULT hr = E_FAIL;

	LPWSTR pPath = NULL;
	ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
	if(pPath) {
		if(PathIsRootW(pPath) && lstrlenW(pPath) == 3 && pPath[1] == L':' && pPath[2] == L'\\') {
			pPath[2] = L'\0';
		}
		BOOL succeeded = TRUE;

		DWORD oldAttributes = GetFileAttributesW(pPath) & static_cast<DWORD>(mask);
		oldAttributes &= ~(FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT);
		DWORD newAttributes = static_cast<DWORD>(attributes) & static_cast<DWORD>(mask);
		newAttributes &= ~(FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT);

		if((oldAttributes & FILE_ATTRIBUTE_COMPRESSED) != (newAttributes & FILE_ATTRIBUTE_COMPRESSED)) {
			// change compression state using DeviceIoControl
			HANDLE hFile = CreateFileW(pPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_WRITE | FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
			if(hFile == INVALID_HANDLE_VALUE) {
				hr = HRESULT_FROM_WIN32(GetLastError());
				CoTaskMemFree(pPath);
				return hr;
			} else {
				USHORT newCompressionState;
				if(newAttributes & FILE_ATTRIBUTE_COMPRESSED) {
					newCompressionState = COMPRESSION_FORMAT_DEFAULT;
				} else {
					newCompressionState = COMPRESSION_FORMAT_NONE;
				}
				DWORD bytesReturned = 0;
				succeeded = DeviceIoControl(hFile, FSCTL_SET_COMPRESSION, &newCompressionState, sizeof(newCompressionState), NULL, 0, &bytesReturned, NULL);
				CloseHandle(hFile);
				if(!succeeded) {
					hr = HRESULT_FROM_WIN32(GetLastError());
					CoTaskMemFree(pPath);
					return hr;
				}
			}
		}

		if((oldAttributes & FILE_ATTRIBUTE_ENCRYPTED) != (newAttributes & FILE_ATTRIBUTE_ENCRYPTED)) {
			// change encryption state using EncryptFile/DecryptFile
			if(newAttributes & FILE_ATTRIBUTE_ENCRYPTED) {
				succeeded = EncryptFileW(pPath);
			} else {
				succeeded = DecryptFileW(pPath, 0);
			}
			if(!succeeded) {
				hr = HRESULT_FROM_WIN32(GetLastError());
				CoTaskMemFree(pPath);
				return hr;
			}
		}

		if((oldAttributes & FILE_ATTRIBUTE_SPARSE_FILE) != (newAttributes & FILE_ATTRIBUTE_SPARSE_FILE)) {
			if(newAttributes & FILE_ATTRIBUTE_SPARSE_FILE) {
				// make the file a sparse file using DeviceIoControl
				HANDLE hFile = CreateFileW(pPath, GENERIC_WRITE, FILE_SHARE_WRITE | FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
				if(hFile == INVALID_HANDLE_VALUE) {
					hr = HRESULT_FROM_WIN32(GetLastError());
					CoTaskMemFree(pPath);
					return hr;
				} else {
					DWORD bytesReturned = 0;
					succeeded = DeviceIoControl(hFile, FSCTL_SET_SPARSE, NULL, 0, NULL, 0, &bytesReturned, NULL);
					CloseHandle(hFile);
					if(!succeeded) {
						hr = HRESULT_FROM_WIN32(GetLastError());
						CoTaskMemFree(pPath);
						return hr;
					}
				}
			} else {
				CoTaskMemFree(pPath);
				return E_FAIL;
			}
		}

		newAttributes &= ~(FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_ENCRYPTED | FILE_ATTRIBUTE_SPARSE_FILE);

		succeeded = SetFileAttributesW(pPath, newAttributes);
		if(succeeded) {
			hr = S_OK;
		} else {
			hr = HRESULT_FROM_WIN32(GetLastError());
		}
		CoTaskMemFree(pPath);
	}
	return hr;
}

PIDLIST_ABSOLUTE SimpleToRealPIDL(PCIDLIST_ABSOLUTE pSimplePIDL)
{
	ATLASSERT_POINTER(pSimplePIDL, ITEMIDLIST_ABSOLUTE);

	PIDLIST_ABSOLUTE pRealPIDL = NULL;

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pIDLRelative = NULL;
	HRESULT hr = SHBindToParent(pSimplePIDL, IID_PPV_ARGS(&pParentISF), &pIDLRelative);
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		PUITEMID_CHILD pIDL = NULL;
		if(SUCCEEDED(SHGetRealIDL(pParentISF, pIDLRelative, &pIDL))) {
			ATLASSERT_POINTER(pIDL, ITEMID_CHILD);

			pRealPIDL = MakeAbsolutePIDL(pParentISF, pIDL);
			ILFree(pIDL);
		} else {
			// well, it already was a real pIDL...
			pRealPIDL = ILCloneFull(pSimplePIDL);
		}
	}
	return pRealPIDL;
}

BOOL SupportsNewFolders(PCIDLIST_ABSOLUTE pIDL)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	if(APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		CComPtr<IShellItem> pShellItem = NULL;
		HRESULT hr = E_FAIL;
		APIWrapper::SHCreateItemFromIDList(pIDL, IID_PPV_ARGS(&pShellItem), &hr);
		if(SUCCEEDED(hr)) {
			SFGAOF mask = SFGAO_FILESYSTEM | SFGAO_FOLDER | SFGAO_FILESYSANCESTOR;
			SFGAOF attributes = 0;
			if(SUCCEEDED(pShellItem->GetAttributes(mask, &attributes))) {
				return ((attributes & mask) == mask);
			}
		}
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	if(SUCCEEDED(SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL))) {
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		SFGAOF mask = SFGAO_FILESYSTEM | SFGAO_FOLDER | SFGAO_FILESYSANCESTOR;
		SFGAOF attributes = mask;
		if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes))) {
			return ((attributes & mask) == mask);
		}
	}
	return FALSE;
}

BOOL ValidatePIDL(PCIDLIST_ABSOLUTE pAbsolutePIDL)
{
	BOOL ret = FALSE;

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(pAbsolutePIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		SFGAOF attributes = SFGAO_VALIDATE;
		ret = SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes));
	}
	return ret;
}

HRESULT CreateDwordBindCtx(LPCWSTR pOptionName, DWORD optionValue, IBindCtx** ppBindContext)
{
	CComPtr<IBindCtx> bindContext;
	HRESULT hr = CreateBindCtx(0, &bindContext);
	if(SUCCEEDED(hr)) {
		hr = AddBindCtxDWORD(bindContext, pOptionName, optionValue);
	}
	*ppBindContext = SUCCEEDED(hr) ? bindContext.Detach() : NULL;
	return hr;
}

HRESULT AddBindCtxDWORD(IBindCtx* pBindContext, LPCWSTR pOptionName, DWORD optionValue)
{
	CComPtr<IPropertyBag> pPropertyBag;
	HRESULT hr = EnsureBindCtxPropertyBag(pBindContext, IID_PPV_ARGS(&pPropertyBag));
	if(SUCCEEDED(hr)) {
		APIWrapper::PSPropertyBag_WriteDWORD(pPropertyBag, pOptionName, optionValue, &hr);
	}
	return hr;
}

HRESULT EnsureBindCtxPropertyBag(IBindCtx* pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	*ppInterfaceImpl = NULL;
	CComPtr<IUnknown> pUnk;
	HRESULT hr = pBindContext->GetObjectParam(STR_PROPERTYBAG_PARAM, &pUnk);
	if(FAILED(hr)) {
		APIWrapper::PSCreateMemoryPropertyStore(IID_PPV_ARGS(&pUnk), &hr);
		if(SUCCEEDED(hr)) {
			hr = pBindContext->RegisterObjectParam(STR_PROPERTYBAG_PARAM, pUnk);
		}
	}
	if(SUCCEEDED(hr)) {
		hr = pUnk->QueryInterface(requiredInterface, ppInterfaceImpl);
	}
	return hr;
}