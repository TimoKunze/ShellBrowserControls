// LegacyTileViewTask.cpp: A specialization of RunnableTask for retrieving shell item's tile view sub-items

#include "stdafx.h"
#include "LegacyTileViewTask.h"


ShLvwLegacyTileViewTask::ShLvwLegacyTileViewTask(void)
    : RunnableTask(TRUE)
{
	properties.pBackgroundTileViewQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwLegacyTileViewTask::FinalRelease()
{
	if(properties.pParentISF2) {
		properties.pParentISF2->Release();
		properties.pParentISF2 = NULL;
	}
	if(properties.pIQueryAssociations) {
		properties.pIQueryAssociations->Release();
		properties.pIQueryAssociations = NULL;
	}
	if(properties.pIPropertiesUI) {
		properties.pIPropertiesUI->Release();
		properties.pIPropertiesUI = NULL;
	}
	#ifdef USE_STL
		for(std::vector<PROPERTYKEY*>::const_iterator it = properties.propertyKeys.cbegin(); it != properties.propertyKeys.cend(); ++it) {
			delete *it;
		}
		properties.propertyKeys.clear();
	#else
		for(size_t i = 0; i < properties.propertyKeys.GetCount(); ++i) {
			delete properties.propertyKeys[i];
		}
		properties.propertyKeys.RemoveAll();
	#endif
	if(properties.pTileInfoString) {
		HeapFree(GetProcessHeap(), 0, properties.pTileInfoString);
		properties.pTileInfoString = NULL;
	}
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
	if(properties.pIDLParentNamespace) {
		ILFree(properties.pIDLParentNamespace);
		properties.pIDLParentNamespace = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->hColumnBuffer) {
			DPA_Destroy(properties.pResult->hColumnBuffer);
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount)
#else
	HRESULT ShLvwLegacyTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pBackgroundTileViewQueue = pBackgroundTileViewQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.hColumnsReadyEvent = hColumnsReadyEvent;
	this->properties.pIDL = ILCloneFull(pIDL);
	if(pIDLParentNamespace) {
		this->properties.pIDLParentNamespace = ILCloneFull(pIDLParentNamespace);
	} else {
		this->properties.pIDLParentNamespace = NULL;
	}
	this->properties.maxColumnCount = maxColumnCount;

	properties.pParentISF2 = NULL;
	properties.pIQueryAssociations = NULL;
	properties.pIPropertiesUI = NULL;
	properties.pTileInfoString = NULL;
	properties.pResult = new SHLVWBACKGROUNDTILEVIEWINFO;
	if(properties.pResult) {
		ZeroMemory(properties.pResult, sizeof(SHLVWBACKGROUNDTILEVIEWINFO));
		properties.pResult->itemID = itemID;
		properties.pResult->hColumnBuffer = DPA_Create(maxColumnCount);
	} else {
		return E_OUTOFMEMORY;
	}
	properties.status = TVTS_NOTHINGDONE;
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
	if(!(properties.status == TVTS_DONE && !properties.pResult)) {
		ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTILEVIEWINFO);
		ATLASSUME(properties.pResult->hColumnBuffer);
		ATLASSUME(properties.pParentISF2);

		HRESULT hr;

		if(properties.status != TVTS_DONE) {
			ATLASSERT(properties.status == TVTS_COLLECTINGPROPERTYKEYS || properties.status == TVTS_FINDINGCOLUMNS);
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}

			if(properties.status == TVTS_COLLECTINGPROPERTYKEYS) {
				if(!properties.pTileInfoString) {
					// retrieve a string defining which properties to display
					ATLASSUME(properties.pIQueryAssociations);

					DWORD bufferSize = 0;
					hr = properties.pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE, ASSOCSTR_TILEINFO, NULL, NULL, &bufferSize);
					if(hr == S_FALSE) {
						properties.pTileInfoString = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (bufferSize + 1) * sizeof(WCHAR)));
						ATLASSERT(properties.pTileInfoString);
						if(!properties.pTileInfoString) {
							return E_OUTOFMEMORY;
						}
						hr = properties.pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE, ASSOCSTR_TILEINFO, NULL, properties.pTileInfoString, &bufferSize);
						ATLASSERT(SUCCEEDED(hr));
					}
					properties.pIQueryAssociations->Release();
					properties.pIQueryAssociations = NULL;

					hr = CoCreateInstance(CLSID_PropertiesUI, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&properties.pIPropertiesUI));
					ATLASSERT(SUCCEEDED(hr));
					if(!properties.pTileInfoString || !properties.pIPropertiesUI) {
						return E_FAIL;
					}
					properties.pNextTileInfoProperty = properties.pTileInfoString;
					if((tolower(properties.pNextTileInfoProperty[0]) == L'p') && (tolower(properties.pNextTileInfoProperty[1]) == L'r') && (tolower(properties.pNextTileInfoProperty[2]) == L'o') && (tolower(properties.pNextTileInfoProperty[3]) == L'p') && (tolower(properties.pNextTileInfoProperty[4]) == L':')) {
						properties.pNextTileInfoProperty += 5;
					}

					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}

				// parse the retrieved string and translate each single property string into a property key
				ATLASSUME(properties.pIPropertiesUI);
				ATLASSERT_POINTER(properties.pNextTileInfoProperty, WCHAR);
				while(properties.pNextTileInfoProperty[0] != L'\0') {
					LPWSTR pProperty = properties.pNextTileInfoProperty;
					while(properties.pNextTileInfoProperty[0] != L'\0') {
						if(properties.pNextTileInfoProperty[0] == L';') {
							properties.pNextTileInfoProperty[0] = L'\0';
							++properties.pNextTileInfoProperty;
							break;
						}
						++properties.pNextTileInfoProperty;
					}

					PROPERTYKEY* pPropertyKey = new PROPERTYKEY;
					if(!pPropertyKey) {
						return E_OUTOFMEMORY;
					}
					ZeroMemory(pPropertyKey, sizeof(PROPERTYKEY));
					ULONG dummy = 0;
					hr = properties.pIPropertiesUI->ParsePropertyName(pProperty, &pPropertyKey->fmtid, &pPropertyKey->pid, &dummy);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						#ifdef USE_STL
							properties.propertyKeys.push_back(pPropertyKey);
						#else
							properties.propertyKeys.Add(pPropertyKey);
						#endif
						if(DPA_AppendPtr(properties.pResult->hColumnBuffer, reinterpret_cast<LPVOID>(-1)) == -1) {
							return E_OUTOFMEMORY;
						}
					}

					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}
				properties.status = TVTS_FINDINGCOLUMNS;
				HeapFree(GetProcessHeap(), 0, properties.pTileInfoString);
				properties.pTileInfoString = NULL;
				properties.pNextTileInfoProperty = NULL;
				properties.pIPropertiesUI->Release();
				properties.pIPropertiesUI = NULL;
			}

			if(properties.status == TVTS_FINDINGCOLUMNS) {
				SHELLDETAILS shellDetails = {0};
				while(properties.pParentISF2->GetDetailsOf(NULL, properties.currentRealColumnIndex, &shellDetails) == S_OK) {
					SHCOLUMNID propertyKey = {0};
					hr = properties.pParentISF2->MapColumnToSCID(properties.currentRealColumnIndex, &propertyKey);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						#ifdef USE_STL
							size_t columnCount = properties.propertyKeys.size();
						#else
							size_t columnCount = properties.propertyKeys.GetCount();
						#endif
						for(size_t i = 0; i < columnCount; ++i) {
							if(propertyKey.fmtid == properties.propertyKeys[i]->fmtid && propertyKey.pid == properties.propertyKeys[i]->pid) {
								DPA_SetPtr(properties.pResult->hColumnBuffer, i, reinterpret_cast<LPVOID>(properties.currentRealColumnIndex));
								break;
							}
						}
					}

					properties.currentRealColumnIndex++;
					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}

				// now remove any columns that are still -1
				UINT i = 0;
				while(i < static_cast<UINT>(DPA_GetPtrCount(properties.pResult->hColumnBuffer))) {
					if(reinterpret_cast<int>(DPA_GetPtr(properties.pResult->hColumnBuffer, i)) == -1) {
						DPA_DeletePtr(properties.pResult->hColumnBuffer, i);
					} else {
						i++;
					}
				}
				// "remove" any columns that are too many
				if(static_cast<UINT>(DPA_GetPtrCount(properties.pResult->hColumnBuffer)) > properties.maxColumnCount) {
					DPA_SetPtrCount(properties.pResult->hColumnBuffer, properties.maxColumnCount);
				}
				properties.status = TVTS_DONE;
			}
		}
		ATLASSERT(properties.status == TVTS_DONE);

		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pBackgroundTileViewQueue->push(properties.pResult);
		#else
			properties.pBackgroundTileViewQueue->AddTail(properties.pResult);
		#endif
		properties.pResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);
	}

	if(IsWindow(properties.hWndToNotify)) {
		HANDLE handles[2];
		handles[0] = hDoneEvent;
		handles[1] = properties.hColumnsReadyEvent;
		DWORD result = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
		if(result == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		} else if(result == WAIT_OBJECT_0 + 1) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATETILEVIEWCOLUMNS, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyTileViewTask::DoRun(void)
{
	properties.currentRealColumnIndex = 0;
	properties.pNextTileInfoProperty = NULL;

	CComPtr<IShellFolder> pParentISF = NULL;
	if(properties.pIDLParentNamespace) {
		BindToPIDL(properties.pIDLParentNamespace, IID_PPV_ARGS(&pParentISF));
	}
	if(!pParentISF) {
		return E_FAIL;
	}
	pParentISF->QueryInterface(IID_PPV_ARGS(&properties.pParentISF2));
	if(!properties.pParentISF2) {
		return E_FAIL;
	}

	PCITEMID_CHILD pIDLChild = ILFindLastID(properties.pIDL);
	HRESULT hr = pParentISF->GetUIObjectOf(properties.hWndShellUIParentWindow, 1, &pIDLChild, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&properties.pIQueryAssociations));
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
			hr = properties.pIQueryAssociations->Init(ASSOCF_INIT_DEFAULTTOFOLDER, L"Folder", NULL, NULL);
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
						hr = properties.pIQueryAssociations->Init(attributes & SFGAO_FOLDER ? ASSOCF_INIT_DEFAULTTOFOLDER : ASSOCF_INIT_DEFAULTTOSTAR, L"*", NULL, NULL);
					} else {
						hr = properties.pIQueryAssociations->Init(attributes & SFGAO_FOLDER ? ASSOCF_INIT_DEFAULTTOFOLDER : ASSOCF_INIT_DEFAULTTOSTAR, pExtension, NULL, NULL);
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
				hr = properties.pIQueryAssociations->Init(0, pParsingName + 2/*"::"*/, NULL, NULL);
				initialized = SUCCEEDED(hr);
			}
			CoTaskMemFree(pParsingName);
		}
	}
	if(!initialized) {
		return E_FAIL;
	}
	properties.status = TVTS_COLLECTINGPROPERTYKEYS;
	return E_PENDING;
}