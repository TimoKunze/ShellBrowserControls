// APIWrapper.cpp: A wrapper class for API functions not available on all supported systems

#include "stdafx.h"
#include "APIWrapper.h"


BOOL APIWrapper::checkedSupport_AssocGetDetailsOfPropKey = FALSE;
#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
	BOOL APIWrapper::checkedSupport_CancelSynchronousIo = FALSE;
#endif
BOOL APIWrapper::checkedSupport_CDefFolderMenu_Create2 = FALSE;
BOOL APIWrapper::checkedSupport_FileIconInit = FALSE;
BOOL APIWrapper::checkedSupport_HIMAGELIST_QueryInterface = FALSE;
BOOL APIWrapper::checkedSupport_ImageList_CoCreateInstance = FALSE;
BOOL APIWrapper::checkedSupport_IsElevationRequired = FALSE;
BOOL APIWrapper::checkedSupport_IsPathShared = FALSE;
BOOL APIWrapper::checkedSupport_PropVariantToString = FALSE;
BOOL APIWrapper::checkedSupport_PSCreateMemoryPropertyStore = FALSE;
BOOL APIWrapper::checkedSupport_PSFormatForDisplay = FALSE;
BOOL APIWrapper::checkedSupport_PSGetItemPropertyHandler = FALSE;
BOOL APIWrapper::checkedSupport_PSGetPropertyDescription = FALSE;
BOOL APIWrapper::checkedSupport_PSGetPropertyDescriptionListFromString = FALSE;
BOOL APIWrapper::checkedSupport_PSGetPropertyKeyFromName = FALSE;
BOOL APIWrapper::checkedSupport_PSGetPropertyValue = FALSE;
BOOL APIWrapper::checkedSupport_PSPropertyBag_WriteDWORD = FALSE;
BOOL APIWrapper::checkedSupport_SHCreateDefaultContextMenu = FALSE;
BOOL APIWrapper::checkedSupport_SHCreateItemFromIDList = FALSE;
BOOL APIWrapper::checkedSupport_SHCreateShellItem = FALSE;
BOOL APIWrapper::checkedSupport_SHCreateThreadRef = FALSE;
BOOL APIWrapper::checkedSupport_SHGetIDListFromObject = FALSE;
BOOL APIWrapper::checkedSupport_SHGetImageList = FALSE;
BOOL APIWrapper::checkedSupport_SHGetKnownFolderIDList = FALSE;
BOOL APIWrapper::checkedSupport_SHGetStockIconInfo = FALSE;
BOOL APIWrapper::checkedSupport_SHGetThreadRef = FALSE;
BOOL APIWrapper::checkedSupport_SHLimitInputEdit = FALSE;
BOOL APIWrapper::checkedSupport_SHSetThreadRef = FALSE;
BOOL APIWrapper::checkedSupport_StampIconForElevation = FALSE;
BOOL APIWrapper::checkedSupport_VariantToPropVariant = FALSE;
AssocGetDetailsOfPropKeyFn* APIWrapper::pfnAssocGetDetailsOfPropKey = NULL;
#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
	CancelSynchronousIoFn* APIWrapper::pfnCancelSynchronousIo = NULL;
#endif
CDefFolderMenu_Create2Fn* APIWrapper::pfnCDefFolderMenu_Create2 = NULL;
FileIconInitFn* APIWrapper::pfnFileIconInit = NULL;
HIMAGELIST_QueryInterfaceFn* APIWrapper::pfnHIMAGELIST_QueryInterface = NULL;
ImageList_CoCreateInstanceFn* APIWrapper::pfnImageList_CoCreateInstance = NULL;
IsElevationRequiredFn* APIWrapper::pfnIsElevationRequired = NULL;
IsPathSharedFn* APIWrapper::pfnIsPathShared = NULL;
PropVariantToStringFn* APIWrapper::pfnPropVariantToString = NULL;
PSCreateMemoryPropertyStoreFn* APIWrapper::pfnPSCreateMemoryPropertyStore = NULL;
PSFormatForDisplayFn* APIWrapper::pfnPSFormatForDisplay = NULL;
PSGetItemPropertyHandlerFn* APIWrapper::pfnPSGetItemPropertyHandler = NULL;
PSGetPropertyDescriptionFn* APIWrapper::pfnPSGetPropertyDescription = NULL;
PSGetPropertyDescriptionListFromStringFn* APIWrapper::pfnPSGetPropertyDescriptionListFromString = NULL;
PSGetPropertyKeyFromNameFn* APIWrapper::pfnPSGetPropertyKeyFromName = NULL;
PSGetPropertyValueFn* APIWrapper::pfnPSGetPropertyValue = NULL;
PSPropertyBag_WriteDWORDFn* APIWrapper::pfnPSPropertyBag_WriteDWORD = NULL;
SHCreateDefaultContextMenuFn* APIWrapper::pfnSHCreateDefaultContextMenu = NULL;
SHCreateItemFromIDListFn* APIWrapper::pfnSHCreateItemFromIDList = NULL;
SHCreateShellItemFn* APIWrapper::pfnSHCreateShellItem = NULL;
SHCreateThreadRefFn* APIWrapper::pfnSHCreateThreadRef = NULL;
SHGetIDListFromObjectFn* APIWrapper::pfnSHGetIDListFromObject = NULL;
SHGetImageListFn* APIWrapper::pfnSHGetImageList = NULL;
SHGetKnownFolderIDListFn* APIWrapper::pfnSHGetKnownFolderIDList = NULL;
SHGetStockIconInfoFn* APIWrapper::pfnSHGetStockIconInfo = NULL;
SHGetThreadRefFn* APIWrapper::pfnSHGetThreadRef = NULL;
SHLimitInputEditFn* APIWrapper::pfnSHLimitInputEdit = NULL;
SHSetThreadRefFn* APIWrapper::pfnSHSetThreadRef = NULL;
StampIconForElevationFn* APIWrapper::pfnStampIconForElevation = NULL;
VariantToPropVariantFn* APIWrapper::pfnVariantToPropVariant = NULL;


APIWrapper::APIWrapper(void)
{
}


BOOL APIWrapper::IsSupported_AssocGetDetailsOfPropKey(void)
{
	if(!checkedSupport_AssocGetDetailsOfPropKey) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnAssocGetDetailsOfPropKey = reinterpret_cast<AssocGetDetailsOfPropKeyFn*>(GetProcAddress(hShell32, "AssocGetDetailsOfPropKey"));
		}
		checkedSupport_AssocGetDetailsOfPropKey = TRUE;
	}

	return (pfnAssocGetDetailsOfPropKey != NULL);
}

#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
	BOOL APIWrapper::IsSupported_CancelSynchronousIo(void)
	{
		if(!checkedSupport_CancelSynchronousIo) {
			//if(!RunTimeHelper::IsWin8()) {
				HMODULE hKernel32 = GetModuleHandle(TEXT("kernel32.dll"));
				if(hKernel32) {
					pfnCancelSynchronousIo = reinterpret_cast<CancelSynchronousIoFn*>(GetProcAddress(hKernel32, "CancelSynchronousIo"));
				}
			//}
			checkedSupport_CancelSynchronousIo = TRUE;
		}

		return (pfnCancelSynchronousIo != NULL);
	}
#endif

BOOL APIWrapper::IsSupported_CDefFolderMenu_Create2(void)
{
	if(!checkedSupport_CDefFolderMenu_Create2) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnCDefFolderMenu_Create2 = reinterpret_cast<CDefFolderMenu_Create2Fn*>(GetProcAddress(hShell32, "CDefFolderMenu_Create2"));
			if(!pfnCDefFolderMenu_Create2) {
				pfnCDefFolderMenu_Create2 = reinterpret_cast<CDefFolderMenu_Create2Fn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(701)));
			}
		}
		checkedSupport_CDefFolderMenu_Create2 = TRUE;
	}

	return (pfnCDefFolderMenu_Create2 != NULL);
}

BOOL APIWrapper::IsSupported_FileIconInit(void)
{
	if(!checkedSupport_FileIconInit) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnFileIconInit = reinterpret_cast<FileIconInitFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(660)));
		}
		checkedSupport_FileIconInit = TRUE;
	}

	return (pfnFileIconInit != NULL);
}

BOOL APIWrapper::IsSupported_HIMAGELIST_QueryInterface(void)
{
	if(!checkedSupport_HIMAGELIST_QueryInterface) {
		HMODULE hComctl32 = GetModuleHandle(TEXT("comctl32.dll"));
		if(hComctl32) {
			pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hComctl32, "HIMAGELIST_QueryInterface"));
		}
		checkedSupport_HIMAGELIST_QueryInterface = TRUE;
	}

	return (pfnHIMAGELIST_QueryInterface != NULL);
}

BOOL APIWrapper::IsSupported_ImageList_CoCreateInstance(void)
{
	if(!checkedSupport_ImageList_CoCreateInstance) {
		HMODULE hComctl32 = GetModuleHandle(TEXT("comctl32.dll"));
		if(hComctl32) {
			pfnImageList_CoCreateInstance = reinterpret_cast<ImageList_CoCreateInstanceFn*>(GetProcAddress(hComctl32, "ImageList_CoCreateInstance"));
		}
		checkedSupport_ImageList_CoCreateInstance = TRUE;
	}

	return (pfnImageList_CoCreateInstance != NULL);
}

BOOL APIWrapper::IsSupported_IsElevationRequired(void)
{
	if(!checkedSupport_IsElevationRequired) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnIsElevationRequired = reinterpret_cast<IsElevationRequiredFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(865)));
		}
		checkedSupport_IsElevationRequired = TRUE;
	}

	return (pfnIsElevationRequired != NULL);
}

BOOL APIWrapper::IsSupported_IsPathShared(void)
{
	if(!checkedSupport_IsPathShared) {
		HKEY hKey = NULL;
		RegCreateKeyEx(HKEY_CLASSES_ROOT, TEXT("Network\\SharingHandler"), 0, NULL, 0, KEY_QUERY_VALUE, NULL, &hKey, NULL);
		if(hKey) {
			TCHAR pSharingHandler[MAX_PATH];
			LONG dataSize = MAX_PATH;
			LSTATUS result = RegQueryValue(hKey, NULL, pSharingHandler, &dataSize);
			if(result == ERROR_FILE_NOT_FOUND) {
				lstrcpyn(pSharingHandler, TEXT("ntshrui.dll"), MAX_PATH);
				result = ERROR_SUCCESS;
			}
			if(result == ERROR_SUCCESS) {
				HMODULE hSharingHandler = GetModuleHandle(pSharingHandler);
				if(hSharingHandler) {
					#ifdef UNICODE
						pfnIsPathShared = reinterpret_cast<IsPathSharedWFn*>(GetProcAddress(hSharingHandler, "IsPathSharedW"));
					#else
						pfnIsPathShared = reinterpret_cast<IsPathSharedAFn*>(GetProcAddress(hSharingHandler, "IsPathSharedA"));
					#endif
				}
			}
			RegCloseKey(hKey);
		}
		checkedSupport_IsPathShared = TRUE;
	}

	return (pfnIsPathShared != NULL);
}

BOOL APIWrapper::IsSupported_PropVariantToString(void)
{
	if(!checkedSupport_PropVariantToString) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPropVariantToString = reinterpret_cast<PropVariantToStringFn*>(GetProcAddress(hPropSys, "PropVariantToString"));
		}
		checkedSupport_PropVariantToString = TRUE;
	}

	return (pfnPropVariantToString != NULL);
}

BOOL APIWrapper::IsSupported_PSCreateMemoryPropertyStore(void)
{
	if(!checkedSupport_PSCreateMemoryPropertyStore) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSCreateMemoryPropertyStore = reinterpret_cast<PSCreateMemoryPropertyStoreFn*>(GetProcAddress(hPropSys, "PSCreateMemoryPropertyStore"));
		}
		checkedSupport_PSCreateMemoryPropertyStore = TRUE;
	}

	return (pfnPSCreateMemoryPropertyStore != NULL);
}

BOOL APIWrapper::IsSupported_PSFormatForDisplay(void)
{
	if(!checkedSupport_PSFormatForDisplay) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSFormatForDisplay = reinterpret_cast<PSFormatForDisplayFn*>(GetProcAddress(hPropSys, "PSFormatForDisplay"));
		}
		checkedSupport_PSFormatForDisplay = TRUE;
	}

	return (pfnPSFormatForDisplay != NULL);
}

BOOL APIWrapper::IsSupported_PSGetItemPropertyHandler(void)
{
	if(!checkedSupport_PSGetItemPropertyHandler) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSGetItemPropertyHandler = reinterpret_cast<PSGetItemPropertyHandlerFn*>(GetProcAddress(hPropSys, "PSGetItemPropertyHandler"));
		}
		checkedSupport_PSGetItemPropertyHandler = TRUE;
	}

	return (pfnPSGetItemPropertyHandler != NULL);
}

BOOL APIWrapper::IsSupported_PSGetPropertyDescription(void)
{
	if(!checkedSupport_PSGetPropertyDescription) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSGetPropertyDescription = reinterpret_cast<PSGetPropertyDescriptionFn*>(GetProcAddress(hPropSys, "PSGetPropertyDescription"));
		}
		checkedSupport_PSGetPropertyDescription = TRUE;
	}

	return (pfnPSGetPropertyDescription != NULL);
}

BOOL APIWrapper::IsSupported_PSGetPropertyDescriptionListFromString(void)
{
	if(!checkedSupport_PSGetPropertyDescriptionListFromString) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSGetPropertyDescriptionListFromString = reinterpret_cast<PSGetPropertyDescriptionListFromStringFn*>(GetProcAddress(hPropSys, "PSGetPropertyDescriptionListFromString"));
		}
		checkedSupport_PSGetPropertyDescriptionListFromString = TRUE;
	}

	return (pfnPSGetPropertyDescriptionListFromString != NULL);
}

BOOL APIWrapper::IsSupported_PSGetPropertyKeyFromName(void)
{
	if(!checkedSupport_PSGetPropertyKeyFromName) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSGetPropertyKeyFromName = reinterpret_cast<PSGetPropertyKeyFromNameFn*>(GetProcAddress(hPropSys, "PSGetPropertyKeyFromName"));
		}
		checkedSupport_PSGetPropertyKeyFromName = TRUE;
	}

	return (pfnPSGetPropertyKeyFromName != NULL);
}

BOOL APIWrapper::IsSupported_PSGetPropertyValue(void)
{
	if(!checkedSupport_PSGetPropertyValue) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSGetPropertyValue = reinterpret_cast<PSGetPropertyValueFn*>(GetProcAddress(hPropSys, "PSGetPropertyValue"));
		}
		checkedSupport_PSGetPropertyValue = TRUE;
	}

	return (pfnPSGetPropertyValue != NULL);
}

BOOL APIWrapper::IsSupported_PSPropertyBag_WriteDWORD(void)
{
	if(!checkedSupport_PSPropertyBag_WriteDWORD) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnPSPropertyBag_WriteDWORD = reinterpret_cast<PSPropertyBag_WriteDWORDFn*>(GetProcAddress(hPropSys, "PSPropertyBag_WriteDWORD"));
		}
		checkedSupport_PSPropertyBag_WriteDWORD = TRUE;
	}

	return (pfnPSPropertyBag_WriteDWORD != NULL);
}

BOOL APIWrapper::IsSupported_SHCreateDefaultContextMenu(void)
{
	if(!checkedSupport_SHCreateDefaultContextMenu) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHCreateDefaultContextMenu = reinterpret_cast<SHCreateDefaultContextMenuFn*>(GetProcAddress(hShell32, "SHCreateDefaultContextMenu"));
		}
		checkedSupport_SHCreateDefaultContextMenu = TRUE;
	}

	return (pfnSHCreateDefaultContextMenu != NULL);
}

BOOL APIWrapper::IsSupported_SHCreateItemFromIDList(void)
{
	if(!checkedSupport_SHCreateItemFromIDList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHCreateItemFromIDList = reinterpret_cast<SHCreateItemFromIDListFn*>(GetProcAddress(hShell32, "SHCreateItemFromIDList"));
		}
		checkedSupport_SHCreateItemFromIDList = TRUE;
	}

	return (pfnSHCreateItemFromIDList != NULL);
}

BOOL APIWrapper::IsSupported_SHCreateShellItem(void)
{
	if(!checkedSupport_SHCreateShellItem) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHCreateShellItem = reinterpret_cast<SHCreateShellItemFn*>(GetProcAddress(hShell32, "SHCreateShellItem"));
		}
		checkedSupport_SHCreateShellItem = TRUE;
	}

	return (pfnSHCreateShellItem != NULL);
}

BOOL APIWrapper::IsSupported_SHCreateThreadRef(void)
{
	if(!checkedSupport_SHCreateThreadRef) {
		HMODULE hShlwapi = GetModuleHandle(TEXT("shlwapi.dll"));
		if(hShlwapi) {
			pfnSHCreateThreadRef = reinterpret_cast<SHCreateThreadRefFn*>(GetProcAddress(hShlwapi, "SHCreateThreadRef"));
		}
		checkedSupport_SHCreateThreadRef = TRUE;
	}

	return (pfnSHCreateThreadRef != NULL);
}

BOOL APIWrapper::IsSupported_SHGetIDListFromObject(void)
{
	if(!checkedSupport_SHGetIDListFromObject) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetIDListFromObject = reinterpret_cast<SHGetIDListFromObjectFn*>(GetProcAddress(hShell32, "SHGetIDListFromObject"));
		}
		checkedSupport_SHGetIDListFromObject = TRUE;
	}

	return (pfnSHGetIDListFromObject != NULL);
}

BOOL APIWrapper::IsSupported_SHGetImageList(void)
{
	if(!checkedSupport_SHGetImageList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, "SHGetImageList"));
			if(!pfnSHGetImageList) {
				pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(727)));
			}
		}
		checkedSupport_SHGetImageList = TRUE;
	}

	return (pfnSHGetImageList != NULL);
}

BOOL APIWrapper::IsSupported_SHGetKnownFolderIDList(void)
{
	if(!checkedSupport_SHGetKnownFolderIDList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetKnownFolderIDList = reinterpret_cast<SHGetKnownFolderIDListFn*>(GetProcAddress(hShell32, "SHGetKnownFolderIDList"));
		}
		checkedSupport_SHGetKnownFolderIDList = TRUE;
	}

	return (pfnSHGetKnownFolderIDList != NULL);
}

BOOL APIWrapper::IsSupported_SHGetStockIconInfo(void)
{
	if(!checkedSupport_SHGetStockIconInfo) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetStockIconInfo = reinterpret_cast<SHGetStockIconInfoFn*>(GetProcAddress(hShell32, "SHGetStockIconInfo"));
		}
		checkedSupport_SHGetStockIconInfo = TRUE;
	}

	return (pfnSHGetStockIconInfo != NULL);
}

BOOL APIWrapper::IsSupported_SHGetThreadRef(void)
{
	if(!checkedSupport_SHGetThreadRef) {
		HMODULE hShlwapi = GetModuleHandle(TEXT("shlwapi.dll"));
		if(hShlwapi) {
			pfnSHGetThreadRef = reinterpret_cast<SHGetThreadRefFn*>(GetProcAddress(hShlwapi, "SHGetThreadRef"));
		}
		checkedSupport_SHGetThreadRef = TRUE;
	}

	return (pfnSHGetThreadRef != NULL);
}

BOOL APIWrapper::IsSupported_SHLimitInputEdit(void)
{
	if(!checkedSupport_SHLimitInputEdit) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHLimitInputEdit = reinterpret_cast<SHLimitInputEditFn*>(GetProcAddress(hShell32, "SHLimitInputEdit"));
			if(!pfnSHLimitInputEdit) {
				pfnSHLimitInputEdit = reinterpret_cast<SHLimitInputEditFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(747)));
			}
		}
		checkedSupport_SHLimitInputEdit = TRUE;
	}

	return (pfnSHLimitInputEdit != NULL);
}

BOOL APIWrapper::IsSupported_SHSetThreadRef(void)
{
	if(!checkedSupport_SHSetThreadRef) {
		HMODULE hShlwapi = GetModuleHandle(TEXT("shlwapi.dll"));
		if(hShlwapi) {
			pfnSHSetThreadRef = reinterpret_cast<SHSetThreadRefFn*>(GetProcAddress(hShlwapi, "SHSetThreadRef"));
		}
		checkedSupport_SHSetThreadRef = TRUE;
	}

	return (pfnSHSetThreadRef != NULL);
}

BOOL APIWrapper::IsSupported_StampIconForElevation(void)
{
	if(!checkedSupport_StampIconForElevation) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnStampIconForElevation = reinterpret_cast<StampIconForElevationFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(864)));
		}
		checkedSupport_StampIconForElevation = TRUE;
	}

	return (pfnStampIconForElevation != NULL);
}

BOOL APIWrapper::IsSupported_VariantToPropVariant(void)
{
	if(!checkedSupport_VariantToPropVariant) {
		HMODULE hPropSys = GetModuleHandle(TEXT("propsys.dll"));
		if(hPropSys) {
			pfnVariantToPropVariant = reinterpret_cast<VariantToPropVariantFn*>(GetProcAddress(hPropSys, "VariantToPropVariant"));
		}
		checkedSupport_VariantToPropVariant = TRUE;
	}

	return (pfnVariantToPropVariant != NULL);
}

HRESULT APIWrapper::AssocGetDetailsOfPropKey(LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pIDLChild, const PROPERTYKEY* propertyKey, VARIANT* pValue, BOOL* pFoundKey, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_AssocGetDetailsOfPropKey()) {
		*pReturnValue = pfnAssocGetDetailsOfPropKey(pParentISF, pIDLChild, propertyKey, pValue, pFoundKey);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
	HRESULT APIWrapper::CancelSynchronousIo(HANDLE hThread, BOOL* pReturnValue)
	{
		BOOL dummy;
		if(!pReturnValue) {
			pReturnValue = &dummy;
		}

		if(IsSupported_CancelSynchronousIo()) {
			//ATLTRACE(TEXT("2 - Calling CancelSynchronousIo: 0x%X\n"), hThread);
			*pReturnValue = pfnCancelSynchronousIo(hThread);
			return S_OK;
		} else {
			return E_NOTIMPL;
		}
	}
#endif

HRESULT APIWrapper::CDefFolderMenu_Create2(PCIDLIST_ABSOLUTE pIDLParentFolder, HWND hWnd, UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, LPSHELLFOLDER pParentISF, LPFNDFMCALLBACK pCallback, UINT keyCount, const HKEY* pKeys, LPCONTEXTMENU* ppContextMenu, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_CDefFolderMenu_Create2()) {
		*pReturnValue = pfnCDefFolderMenu_Create2(pIDLParentFolder, hWnd, pIDLCount, pIDLs, pParentISF, pCallback, keyCount, pKeys, ppContextMenu);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::FileIconInit(BOOL restoreCache, BOOL* pReturnValue)
{
	BOOL dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_FileIconInit()) {
		*pReturnValue = pfnFileIconInit(restoreCache);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::HIMAGELIST_QueryInterface(HIMAGELIST hImageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_HIMAGELIST_QueryInterface()) {
		*pReturnValue = pfnHIMAGELIST_QueryInterface(hImageList, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

PIDLIST_ABSOLUTE APIWrapper::ILCreateFromPathW_LFN(LPCWSTR pPath)
{
	if(lstrlenW(pPath) > MAX_PATH - 10) {
		int bufferSize = lstrlenW(pPath) + 12;
		LPWSTR pBuffer = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, sizeof(WCHAR) * bufferSize));
		if(pBuffer) {
			lstrcpynW(pBuffer, L"file://\\\\?\\", bufferSize - 1);
			StringCchCatW(pBuffer, bufferSize, pPath);
			PIDLIST_ABSOLUTE pIDL = ::ILCreateFromPathW(pBuffer);
			HeapFree(GetProcessHeap(), 0, pBuffer);
			return pIDL;
		} else {
			return NULL;
		}
	}
	return ::ILCreateFromPathW(pPath);
}

PIDLIST_ABSOLUTE APIWrapper::ILCreateFromPath_LFN(LPCTSTR pPath)
{
	if(lstrlen(pPath) > MAX_PATH - 10) {
		int bufferSize = lstrlen(pPath) + 12;
		LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, sizeof(TCHAR) * bufferSize));
		if(pBuffer) {
			lstrcpyn(pBuffer, TEXT("file://\\\\?\\"), bufferSize - 1);
			StringCchCat(pBuffer, bufferSize, pPath);
			PIDLIST_ABSOLUTE pIDL = ::ILCreateFromPath(pBuffer);
			HeapFree(GetProcessHeap(), 0, pBuffer);
			return pIDL;
		} else {
			return NULL;
		}
	}
	return ::ILCreateFromPath(pPath);
}

HRESULT APIWrapper::ImageList_CoCreateInstance(REFCLSID classID, const LPUNKNOWN pUnknownOuter, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_ImageList_CoCreateInstance()) {
		*pReturnValue = pfnImageList_CoCreateInstance(classID, pUnknownOuter, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::IsElevationRequired(LPWSTR pPath, BOOL* pReturnValue)
{
	BOOL dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_IsElevationRequired()) {
		*pReturnValue = pfnIsElevationRequired(pPath);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::IsPathShared(LPCTSTR pPath, BOOL flushCache, BOOL* pReturnValue)
{
	BOOL dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_IsPathShared()) {
		#ifndef UNICODE
			if(!pPath) {
				// ANSI version doesn't support NULL paths
				*pReturnValue = pfnIsPathShared(TEXT(""), flushCache);
				return S_OK;
			}
		#endif
		*pReturnValue = pfnIsPathShared(pPath, flushCache);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PropVariantToString(REFPROPVARIANT propertyValue, LPWSTR pBuffer, DWORD bufferSize, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PropVariantToString()) {
		*pReturnValue = pfnPropVariantToString(propertyValue, pBuffer, bufferSize);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSCreateMemoryPropertyStore(REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSCreateMemoryPropertyStore()) {
		*pReturnValue = pfnPSCreateMemoryPropertyStore(requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSFormatForDisplay(REFPROPERTYKEY propertyKey, REFPROPVARIANT propertyValue, PROPDESC_FORMAT_FLAGS flags, LPWSTR pBuffer, DWORD bufferSize, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSFormatForDisplay()) {
		*pReturnValue = pfnPSFormatForDisplay(propertyKey, propertyValue, flags, pBuffer, bufferSize);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSGetItemPropertyHandler(LPUNKNOWN pShellItem, BOOL writable, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSGetItemPropertyHandler()) {
		*pReturnValue = pfnPSGetItemPropertyHandler(pShellItem, writable, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSGetPropertyDescription(REFPROPERTYKEY propertyKey, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSGetPropertyDescription()) {
		*pReturnValue = pfnPSGetPropertyDescription(propertyKey, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSGetPropertyDescriptionListFromString(LPCWSTR pPropertyList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSGetPropertyDescriptionListFromString()) {
		*pReturnValue = pfnPSGetPropertyDescriptionListFromString(pPropertyList, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSGetPropertyKeyFromName(PCWSTR pCanonicalName, PROPERTYKEY* pPropertyKey, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSGetPropertyKeyFromName()) {
		*pReturnValue = pfnPSGetPropertyKeyFromName(pCanonicalName, pPropertyKey);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSGetPropertyValue(LPPROPERTYSTORE pPropertyStore, IPropertyDescription* pPropertyDescription, LPPROPVARIANT pPropertyValue, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSGetPropertyValue()) {
		*pReturnValue = pfnPSGetPropertyValue(pPropertyStore, pPropertyDescription, pPropertyValue);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::PSPropertyBag_WriteDWORD(IPropertyBag* pPropertyBag, LPCWSTR pPropertyName, DWORD propertyValue, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_PSPropertyBag_WriteDWORD()) {
		*pReturnValue = pfnPSPropertyBag_WriteDWORD(pPropertyBag, pPropertyName, propertyValue);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHCreateDefaultContextMenu(const DEFCONTEXTMENU* pDetails, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHCreateDefaultContextMenu()) {
		*pReturnValue = pfnSHCreateDefaultContextMenu(pDetails, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHCreateItemFromIDList(PCIDLIST_ABSOLUTE pIDL, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHCreateItemFromIDList()) {
		*pReturnValue = pfnSHCreateItemFromIDList(pIDL, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHCreateShellItem(PCIDLIST_ABSOLUTE pIDLParent, LPSHELLFOLDER pParentISF, PCUITEMID_CHILD pIDL, IShellItem** ppShellItem, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHCreateShellItem()) {
		*pReturnValue = pfnSHCreateShellItem(pIDLParent, pParentISF, pIDL, ppShellItem);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHCreateThreadRef(LONG* pReferenceCounter, LPUNKNOWN* ppThreadObject, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHCreateThreadRef()) {
		*pReturnValue = pfnSHCreateThreadRef(pReferenceCounter, ppThreadObject);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetIDListFromObject(LPUNKNOWN pObject, PIDLIST_ABSOLUTE* ppIDL, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetIDListFromObject()) {
		*pReturnValue = pfnSHGetIDListFromObject(pObject, ppIDL);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetImageList()) {
		*pReturnValue = pfnSHGetImageList(imageList, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetKnownFolderIDList(REFKNOWNFOLDERID folderID, DWORD flags, HANDLE hToken, PIDLIST_ABSOLUTE* ppIDL, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetKnownFolderIDList()) {
		*pReturnValue = pfnSHGetKnownFolderIDList(folderID, flags, hToken, ppIDL);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetStockIconInfo(SHSTOCKICONID stockIconID, UINT flags, SHSTOCKICONINFO* pStockIconInfo, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetStockIconInfo()) {
		*pReturnValue = pfnSHGetStockIconInfo(stockIconID, flags, pStockIconInfo);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetThreadRef(LPUNKNOWN* ppThreadObject, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetThreadRef()) {
		*pReturnValue = pfnSHGetThreadRef(ppThreadObject);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHLimitInputEdit(HWND hWndEdit, LPSHELLFOLDER pParentISF, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHLimitInputEdit()) {
		*pReturnValue = pfnSHLimitInputEdit(hWndEdit, pParentISF);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHSetThreadRef(LPUNKNOWN pThreadObject, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHSetThreadRef()) {
		*pReturnValue = pfnSHSetThreadRef(pThreadObject);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::StampIconForElevation(HICON hSourceIcon, BOOL iconIsNonSquare, HICON* pStampedIcon, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_StampIconForElevation()) {
		*pReturnValue = pfnStampIconForElevation(hSourceIcon, iconIsNonSquare, pStampedIcon);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::VariantToPropVariant(const VARIANT* pVariant, PROPVARIANT* pPropVariant, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_VariantToPropVariant()) {
		*pReturnValue = pfnVariantToPropVariant(pVariant, pPropVariant);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}