// LegacyTileViewTask.cpp: A specialization of RunnableTask for retrieving shell item's tile view sub-items

#include "stdafx.h"
#include "LegacyTileViewTask.h"


ShLvwLegacyTileViewTask::ShLvwLegacyTileViewTask(void)
    : RunnableTask(TRUE)
{
	pBackgroundTileViewQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwLegacyTileViewTask::FinalRelease()
{
	if(pParentISF2) {
		pParentISF2->Release();
		pParentISF2 = NULL;
	}
	if(pIQueryAssociations) {
		pIQueryAssociations->Release();
		pIQueryAssociations = NULL;
	}
	if(pIPropertiesUI) {
		pIPropertiesUI->Release();
		pIPropertiesUI = NULL;
	}
	#ifdef USE_STL
		for(std::vector<PROPERTYKEY*>::const_iterator it = propertyKeys.cbegin(); it != propertyKeys.cend(); ++it) {
			delete *it;
		}
		propertyKeys.clear();
	#else
		for(size_t i = 0; i < propertyKeys.GetCount(); ++i) {
			delete propertyKeys[i];
		}
		propertyKeys.RemoveAll();
	#endif
	if(pTileInfoString) {
		HeapFree(GetProcessHeap(), 0, pTileInfoString);
		pTileInfoString = NULL;
	}
	if(pIDL) {
		ILFree(pIDL);
		pIDL = NULL;
	}
	if(pIDLParentNamespace) {
		ILFree(pIDLParentNamespace);
		pIDLParentNamespace = NULL;
	}
	if(pResult) {
		if(pResult->hColumnBuffer) {
			DPA_Destroy(pResult->hColumnBuffer);
		}
		delete pResult;
		pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount)
#else
	HRESULT ShLvwLegacyTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pBackgroundTileViewQueue = pBackgroundTileViewQueue;
	this->pCriticalSection = pCriticalSection;
	this->hColumnsReadyEvent = hColumnsReadyEvent;
	this->pIDL = ILCloneFull(pIDL);
	if(pIDLParentNamespace) {
		this->pIDLParentNamespace = ILCloneFull(pIDLParentNamespace);
	} else {
		this->pIDLParentNamespace = NULL;
	}
	this->maxColumnCount = maxColumnCount;

	pParentISF2 = NULL;
	pIQueryAssociations = NULL;
	pIPropertiesUI = NULL;
	pTileInfoString = NULL;
	pResult = new SHLVWBACKGROUNDTILEVIEWINFO;
	if(pResult) {
		ZeroMemory(pResult, sizeof(SHLVWBACKGROUNDTILEVIEWINFO));
		pResult->itemID = itemID;
		pResult->hColumnBuffer = DPA_Create(maxColumnCount);
	} else {
		return E_OUTOFMEMORY;
	}
	status = TVTS_NOTHINGDONE;
	return S_OK;
}

#ifdef USE_STL
	HRESULT ShLvwLegacyTileViewTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyTileViewTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pBackgroundTileViewQueue || !pCriticalSection || !pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwLegacyTileViewTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyTileViewTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, hWndShellUIParentWindow, pBackgroundTileViewQueue, pCriticalSection, hColumnsReadyEvent, pIDL, itemID, pIDLParentNamespace, maxColumnCount);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyTileViewTask::DoInternalResume(void)
{
	if(!(status == TVTS_DONE && !pResult)) {
		ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDTILEVIEWINFO);
		ATLASSUME(pResult->hColumnBuffer);
		ATLASSUME(pParentISF2);

		HRESULT hr;

		if(status != TVTS_DONE) {
			ATLASSERT(status == TVTS_COLLECTINGPROPERTYKEYS || status == TVTS_FINDINGCOLUMNS);
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}

			if(status == TVTS_COLLECTINGPROPERTYKEYS) {
				if(!pTileInfoString) {
					// retrieve a string defining which properties to display
					ATLASSUME(pIQueryAssociations);

					DWORD bufferSize = 0;
					hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE, ASSOCSTR_TILEINFO, NULL, NULL, &bufferSize);
					if(hr == S_FALSE) {
						pTileInfoString = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (bufferSize + 1) * sizeof(WCHAR)));
						ATLASSERT(pTileInfoString);
						if(!pTileInfoString) {
							return E_OUTOFMEMORY;
						}
						hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE, ASSOCSTR_TILEINFO, NULL, pTileInfoString, &bufferSize);
						ATLASSERT(SUCCEEDED(hr));
					}
					pIQueryAssociations->Release();
					pIQueryAssociations = NULL;

					hr = CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pIPropertiesUI));
					ATLASSERT(SUCCEEDED(hr));
					if(!pTileInfoString || !pIPropertiesUI) {
						return E_FAIL;
					}
					pNextTileInfoProperty = pTileInfoString;
					if((tolower(pNextTileInfoProperty[0]) == L'p') && (tolower(pNextTileInfoProperty[1]) == L'r') && (tolower(pNextTileInfoProperty[2]) == L'o') && (tolower(pNextTileInfoProperty[3]) == L'p') && (tolower(pNextTileInfoProperty[4]) == L':')) {
						pNextTileInfoProperty += 5;
					}

					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}

				// parse the retrieved string and translate each single property string into a property key
				ATLASSUME(pIPropertiesUI);
				ATLASSERT_POINTER(pNextTileInfoProperty, WCHAR);
				while(pNextTileInfoProperty[0] != L'\0') {
					LPWSTR pProperty = pNextTileInfoProperty;
					while(pNextTileInfoProperty[0] != L'\0') {
						if(pNextTileInfoProperty[0] == L';') {
							pNextTileInfoProperty[0] = L'\0';
							++pNextTileInfoProperty;
							break;
						}
						++pNextTileInfoProperty;
					}

					PROPERTYKEY* pPropertyKey = new PROPERTYKEY;
					if(!pPropertyKey) {
						return E_OUTOFMEMORY;
					}
					ZeroMemory(pPropertyKey, sizeof(PROPERTYKEY));
					ULONG dummy = 0;
					hr = pIPropertiesUI->ParsePropertyName(pProperty, &pPropertyKey->fmtid, &pPropertyKey->pid, &dummy);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						#ifdef USE_STL
							propertyKeys.push_back(pPropertyKey);
						#else
							propertyKeys.Add(pPropertyKey);
						#endif
						if(DPA_AppendPtr(pResult->hColumnBuffer, reinterpret_cast<LPVOID>(-1)) == -1) {
							return E_OUTOFMEMORY;
						}
					}

					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}
				status = TVTS_FINDINGCOLUMNS;
				HeapFree(GetProcessHeap(), 0, pTileInfoString);
				pTileInfoString = NULL;
				pNextTileInfoProperty = NULL;
				pIPropertiesUI->Release();
				pIPropertiesUI = NULL;
			}

			if(status == TVTS_FINDINGCOLUMNS) {
				SHELLDETAILS shellDetails = {0};
				while(pParentISF2->GetDetailsOf(NULL, currentRealColumnIndex, &shellDetails) == S_OK) {
					SHCOLUMNID propertyKey = {0};
					hr = pParentISF2->MapColumnToSCID(currentRealColumnIndex, &propertyKey);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						#ifdef USE_STL
							size_t columnCount = propertyKeys.size();
						#else
							size_t columnCount = propertyKeys.GetCount();
						#endif
						for(size_t i = 0; i < columnCount; ++i) {
							if(propertyKey.fmtid == propertyKeys[i]->fmtid && propertyKey.pid == propertyKeys[i]->pid) {
								DPA_SetPtr(pResult->hColumnBuffer, i, reinterpret_cast<LPVOID>(currentRealColumnIndex));
								break;
							}
						}
					}

					currentRealColumnIndex++;
					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}

				// now remove any columns that are still -1
				UINT i = 0;
				while(i < static_cast<UINT>(DPA_GetPtrCount(pResult->hColumnBuffer))) {
					if(reinterpret_cast<int>(DPA_GetPtr(pResult->hColumnBuffer, i)) == -1) {
						DPA_DeletePtr(pResult->hColumnBuffer, i);
					} else {
						i++;
					}
				}
				// "remove" any columns that are too many
				if(static_cast<UINT>(DPA_GetPtrCount(pResult->hColumnBuffer)) > maxColumnCount) {
					DPA_SetPtrCount(pResult->hColumnBuffer, maxColumnCount);
				}
				status = TVTS_DONE;
			}
		}
		ATLASSERT(status == TVTS_DONE);

		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pBackgroundTileViewQueue->push(pResult);
		#else
			pBackgroundTileViewQueue->AddTail(pResult);
		#endif
		pResult = NULL;
		LeaveCriticalSection(pCriticalSection);
	}

	if(IsWindow(hWndToNotify)) {
		HANDLE handles[2];
		handles[0] = hDoneEvent;
		handles[1] = hColumnsReadyEvent;
		DWORD result = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
		if(result == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		} else if(result == WAIT_OBJECT_0 + 1) {
			PostMessage(hWndToNotify, WM_TRIGGER_UPDATETILEVIEWCOLUMNS, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyTileViewTask::DoRun(void)
{
	currentRealColumnIndex = 0;
	pNextTileInfoProperty = NULL;

	CComPtr<IShellFolder> pParentISF = NULL;
	if(pIDLParentNamespace) {
		BindToPIDL(pIDLParentNamespace, IID_PPV_ARGS(&pParentISF));
	}
	if(!pParentISF) {
		return E_FAIL;
	}
	pParentISF->QueryInterface(IID_PPV_ARGS(&pParentISF2));
	if(!pParentISF2) {
		return E_FAIL;
	}

	PCITEMID_CHILD pIDLChild = ILFindLastID(pIDL);
	HRESULT hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pIDLChild, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
	if(FAILED(hr)) {
		return E_FAIL;
	}
	SFGAOF attributes = SFGAO_FILESYSTEM | SFGAO_FOLDER | SFGAO_STREAM;
	hr = pParentISF->GetAttributesOf(1, &pIDLChild, &attributes);
	ATLASSERT(SUCCEEDED(hr));

	BOOL initialized = FALSE;
	if(attributes & SFGAO_FILESYSTEM) {
		if((attributes & SFGAO_FOLDER) && ((attributes & SFGAO_STREAM) == 0)) {
		//if(attributes & SFGAO_FOLDER) {
			hr = pIQueryAssociations->Init(ASSOCF_INIT_DEFAULTTOFOLDER, L"Folder", NULL, NULL);
			initialized = SUCCEEDED(hr);
		} else {
			STRRET buffer;
			hr = pParentISF->GetDisplayNameOf(pIDLChild, SHGDN_FORPARSING, &buffer);
			if(SUCCEEDED(hr)) {
				LPWSTR pParsingName = NULL;
				hr = StrRetToStrW(&buffer, pIDLChild, &pParsingName);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr) && pParsingName && !PathIsDirectoryW(pParsingName)) {
					LPWSTR pExtension = PathFindExtensionW(pParsingName);
					if(pExtension[0] == L'\0') {
						hr = pIQueryAssociations->Init(attributes & SFGAO_FOLDER ? ASSOCF_INIT_DEFAULTTOFOLDER : ASSOCF_INIT_DEFAULTTOSTAR, L"*", NULL, NULL);
					} else {
						hr = pIQueryAssociations->Init(attributes & SFGAO_FOLDER ? ASSOCF_INIT_DEFAULTTOFOLDER : ASSOCF_INIT_DEFAULTTOSTAR, pExtension, NULL, NULL);
					}
					initialized = SUCCEEDED(hr);
				}
				CoTaskMemFree(pParsingName);
			}
		}
	} else {
		STRRET buffer;
		hr = pParentISF->GetDisplayNameOf(pIDLChild, SHGDN_FORPARSING | SHGDN_INFOLDER, &buffer);
		if(SUCCEEDED(hr)) {
			LPWSTR pParsingName = NULL;
			hr = StrRetToStrW(&buffer, pIDLChild, &pParsingName);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr) && pParsingName) {
				hr = pIQueryAssociations->Init(0, pParsingName + 2/*"::"*/, NULL, NULL);
				initialized = SUCCEEDED(hr);
			}
			CoTaskMemFree(pParsingName);
		}
	}
	if(!initialized) {
		return E_FAIL;
	}
	status = TVTS_COLLECTINGPROPERTYKEYS;
	return E_PENDING;
}