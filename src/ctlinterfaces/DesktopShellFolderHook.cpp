// DesktopShellFolderHook.cpp: Hooks the desktop's IShellFolder implementation

#include "stdafx.h"
#include "DesktopShellFolderHook.h"


DesktopShellFolderHook::DesktopShellFolderHook(HRESULT* pHResult)
{
	referenceCount = 1;
	pDataObject = NULL;
	pDesktopISF = NULL;
	SHGetDesktopFolder(&pDesktopISF);
	ATLASSUME(pDesktopISF);
	*pHResult = (pDesktopISF ? S_OK : E_FAIL);
}

DesktopShellFolderHook::~DesktopShellFolderHook()
{
	if(pDesktopISF) {
		pDesktopISF->Release();
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
ULONG STDMETHODCALLTYPE DesktopShellFolderHook::AddRef()
{
	return InterlockedIncrement(&referenceCount);
}

ULONG STDMETHODCALLTYPE DesktopShellFolderHook::Release()
{
	ULONG ret = InterlockedDecrement(&referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		delete this;
	}
	return ret;
}

HRESULT STDMETHODCALLTYPE DesktopShellFolderHook::QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = this;
		AddRef();
	} else if(requiredInterface == IID_IShellFolder) {
		*ppInterfaceImpl = this;
		AddRef();
	} else {
		ATLASSUME(pDesktopISF);
		return pDesktopISF->QueryInterface(requiredInterface, ppInterfaceImpl);
	}

	return S_OK;
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IShellFolder
STDMETHODIMP DesktopShellFolderHook::ParseDisplayName(HWND hWnd, LPBINDCTX pBindContext, LPWSTR pDisplayName, PULONG pEatenCharacters, PIDLIST_RELATIVE* ppIDL, PULONG pAttributes)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->ParseDisplayName(hWnd, pBindContext, pDisplayName, pEatenCharacters, ppIDL, pAttributes);
}

STDMETHODIMP DesktopShellFolderHook::EnumObjects(HWND hWnd, SHCONTF flags, LPENUMIDLIST* ppEnumerator)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->EnumObjects(hWnd, flags, ppEnumerator);
}

STDMETHODIMP DesktopShellFolderHook::BindToObject(PCUIDLIST_RELATIVE pIDL, LPBINDCTX pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->BindToObject(pIDL, pBindContext, requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP DesktopShellFolderHook::BindToStorage(PCUIDLIST_RELATIVE pIDL, LPBINDCTX pBindContext, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->BindToStorage(pIDL, pBindContext, requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP DesktopShellFolderHook::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pIDL1, PCUIDLIST_RELATIVE pIDL2)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->CompareIDs(lParam, pIDL1, pIDL2);
}

STDMETHODIMP DesktopShellFolderHook::CreateViewObject(HWND hWndOwner, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->CreateViewObject(hWndOwner, requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP DesktopShellFolderHook::GetAttributesOf(UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, SFGAOF* pAttributes)
{
	// the desktop's GetAttributesOf implementation doesn't handle fully qualified pIDLs very well
	ATLASSERT_POINTER(pAttributes, SFGAOF);

	SFGAOF combinedAttributes = *pAttributes;
	for(UINT i = 0; i < pIDLCount; ++i) {
		CComPtr<IShellFolder> pParentISF;
		PCUITEMID_CHILD pIDLRelative = NULL;
		HRESULT hr = SHBindToParent(reinterpret_cast<PCIDLIST_ABSOLUTE>(pIDLs[i]), IID_PPV_ARGS(&pParentISF), &pIDLRelative);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pIDLRelative, ITEMID_CHILD);

			DWORD attributes = *pAttributes;
			if(SUCCEEDED(pParentISF->GetAttributesOf(1, &pIDLRelative, &attributes))) {
				combinedAttributes &= attributes;
			}
		}
	}
	*pAttributes = combinedAttributes;
	return S_OK;
}

STDMETHODIMP DesktopShellFolderHook::GetUIObjectOf(HWND hWndOwner, UINT pIDLCount, PCUITEMID_CHILD_ARRAY pIDLs, REFIID requiredInterface, PUINT pReserved, LPVOID* ppInterfaceImpl)
{
	if(requiredInterface == IID_IDataObject) {
		// the desktop's data object doesn't provide CF_HDROP if the files are from different drives
		if(!pDataObject) {
			ATLVERIFY(SUCCEEDED(DataObject::CreateInstance(&pDataObject)));
			*ppInterfaceImpl = pDataObject;
		} else {
			ATLVERIFY(SUCCEEDED(pDataObject->QueryInterface(requiredInterface, ppInterfaceImpl)));
		}
		ATLASSUME(*ppInterfaceImpl);
		if(!*ppInterfaceImpl) {
			return E_NOINTERFACE;
		}

		// add CF_HDROP
		TCHAR pPath[2048];
		FORMATETC format = {0};
		format.cfFormat = CF_HDROP;
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.ptd = NULL;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM storageMedium = {0};
		storageMedium.tymed = TYMED_HGLOBAL;

		// calculate the total size of memory we'll have to allocate
		// the list is null-terminated
		size_t requiredSize = sizeof(DROPFILES) + sizeof(TCHAR);
		for(UINT i = 0; i < pIDLCount; ++i) {
			// entries are null-separated
			ZeroMemory(pPath, 2048 * sizeof(TCHAR));
			if(SHGetPathFromIDList(reinterpret_cast<PCIDLIST_ABSOLUTE>(pIDLs[i]), pPath)) {
				requiredSize += (lstrlen(pPath) + 1) * sizeof(TCHAR);
			}
		}

		// store the data
		storageMedium.hGlobal = GlobalAlloc(GPTR, requiredSize);
		if(storageMedium.hGlobal) {
			LPDROPFILES pFiles = static_cast<LPDROPFILES>(GlobalLock(storageMedium.hGlobal));
			if(pFiles) {
				// Windows Explorer doesn't set the pt member, too
				#ifdef UNICODE
					pFiles->fWide = TRUE;
				#else
					pFiles->fWide = FALSE;
				#endif
				pFiles->pFiles = sizeof(DROPFILES);

				++pFiles;
				for(UINT i = 0; i < pIDLCount; ++i) {
					int textSize = 0;
					ZeroMemory(pPath, 2048 * sizeof(TCHAR));
					if(SHGetPathFromIDList(reinterpret_cast<PCIDLIST_ABSOLUTE>(pIDLs[i]), pPath)) {
						textSize = lstrlen(pPath) * sizeof(TCHAR);
						CopyMemory(pFiles, pPath, textSize);
					}
					pFiles = reinterpret_cast<LPDROPFILES>(reinterpret_cast<LPSTR>(pFiles) + textSize + sizeof(TCHAR));
				}
				GlobalUnlock(storageMedium.hGlobal);
			}
		}
		ATLVERIFY(SUCCEEDED(pDataObject->SetData(&format, &storageMedium, TRUE)));


		// add CFSTR_SHELLIDLIST
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_SHELLIDLIST));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.ptd = NULL;
		format.tymed = TYMED_HGLOBAL;
		storageMedium.tymed = TYMED_HGLOBAL;

		// calculate the total size of memory we'll have to allocate
		PIDLIST_ABSOLUTE pIDLDesktop = GetDesktopPIDL();
		ATLASSERT_POINTER(pIDLDesktop, ITEMIDLIST_ABSOLUTE);
		if(pIDLDesktop) {
			requiredSize = sizeof(CIDA) + ILGetSize(pIDLDesktop);
			for(UINT i = 0; i < pIDLCount; ++i) {
				// entries are null-separated
				requiredSize += sizeof(UINT) + ILGetSize(pIDLs[i]);
			}

			// store the data
			storageMedium.hGlobal = GlobalAlloc(GPTR, requiredSize);
			if(storageMedium.hGlobal) {
				LPIDA pIDA = static_cast<LPIDA>(GlobalLock(storageMedium.hGlobal));
				if(pIDA) {
					pIDA->cidl = pIDLCount;
					pIDA->aoffset[0] = sizeof(pIDA->cidl) + (pIDA->cidl + 1) * sizeof(pIDA->aoffset[0]);

					LPBYTE ppIDLs = reinterpret_cast<LPBYTE>(pIDA) + pIDA->aoffset[0];
					UINT pIDLSize = ILGetSize(pIDLDesktop);
					CopyMemory(ppIDLs, pIDLDesktop, pIDLSize);
					ppIDLs += pIDLSize;

					for(UINT i = 0; i < pIDLCount; ++i) {
						pIDA->aoffset[i + 1] = pIDA->aoffset[i] + pIDLSize;
						pIDLSize = ILGetSize(pIDLs[i]);
						CopyMemory(ppIDLs, pIDLs[i], pIDLSize);
						ppIDLs += pIDLSize;
					}
					GlobalUnlock(storageMedium.hGlobal);
				}
			}
			ATLVERIFY(SUCCEEDED(pDataObject->SetData(&format, &storageMedium, TRUE)));
			ILFree(pIDLDesktop);
		}

		return S_OK;
	} else {
		ATLASSUME(pDesktopISF);
		return pDesktopISF->GetUIObjectOf(hWndOwner, pIDLCount, pIDLs, requiredInterface, pReserved, ppInterfaceImpl);
	}
}

STDMETHODIMP DesktopShellFolderHook::GetDisplayNameOf(PCUITEMID_CHILD pIDL, SHGDNF flags, LPSTRRET pDisplayName)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->GetDisplayNameOf(pIDL, flags, pDisplayName);
}

STDMETHODIMP DesktopShellFolderHook::SetNameOf(HWND hWnd, PCUITEMID_CHILD pIDL, LPCWSTR pName, SHGDNF flags, PITEMID_CHILD* pNewPIDL)
{
	ATLASSUME(pDesktopISF);
	return pDesktopISF->SetNameOf(hWnd, pIDL, pName, flags, pNewPIDL);
}
// implementation of IShellFolder
//////////////////////////////////////////////////////////////////////


HRESULT DesktopShellFolderHook::CreateInstance(LPSHELLFOLDER* ppShellFolder)
{
	ATLASSERT_POINTER(ppShellFolder, LPSHELLFOLDER);
	if(!ppShellFolder) {
		return E_POINTER;
	}

	*ppShellFolder = NULL;

	HRESULT hr = NOERROR;
	DesktopShellFolderHook* pNewShellFolder = new DesktopShellFolderHook(&hr);
	if(pNewShellFolder) {
		if(SUCCEEDED(hr)) {
			*ppShellFolder = pNewShellFolder;
		} else {
			delete pNewShellFolder;
			hr = E_FAIL;
		}
	} else {
		hr = E_OUTOFMEMORY;
	}
	return hr;
}