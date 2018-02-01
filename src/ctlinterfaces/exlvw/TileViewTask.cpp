// TileViewTask.cpp: A specialization of RunnableTask for retrieving shell item's tile view sub-items

#include "stdafx.h"
#include "TileViewTask.h"


ShLvwTileViewTask::ShLvwTileViewTask(void)
    : RunnableTask(TRUE)
{
	properties.pBackgroundTileViewQueue = NULL;
	properties.pParentISF2 = NULL;
	properties.pPropertyDescriptionList = NULL;
	properties.pPropertyKeys = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwTileViewTask::FinalRelease()
{
	if(properties.pParentISF2) {
		properties.pParentISF2->Release();
		properties.pParentISF2 = NULL;
	}
	if(properties.pPropertyDescriptionList) {
		properties.pPropertyDescriptionList->Release();
		properties.pPropertyDescriptionList = NULL;
	}
	if(properties.pPropertyKeys) {
		HeapFree(GetProcessHeap(), 0, properties.pPropertyKeys);
		properties.pPropertyKeys = NULL;
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
	HRESULT ShLvwTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad)
#else
	HRESULT ShLvwTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad)
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
	this->properties.propertyListToLoad = propertyListToLoad;

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
	HRESULT ShLvwTileViewTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad, IRunnableTask** ppTask)
#else
	HRESULT ShLvwTileViewTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad, IRunnableTask** ppTask)
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

	CComObject<ShLvwTileViewTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwTileViewTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, hWndShellUIParentWindow, pBackgroundTileViewQueue, pCriticalSection, hColumnsReadyEvent, pIDL, itemID, pIDLParentNamespace, maxColumnCount, propertyListToLoad);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwTileViewTask::DoInternalResume(void)
{
	if(!(properties.status == TVTS_DONE && !properties.pResult)) {
		ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTILEVIEWINFO);
		ATLASSUME(properties.pResult->hColumnBuffer);
		ATLASSUME(properties.pParentISF2);

		HRESULT hr;

		if(properties.numberOfPropertyKeys <= 0) {
			return NOERROR;
		}
		if(properties.status != TVTS_DONE) {
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}

			if(properties.status == TVTS_COLLECTINGPROPERTYKEYS) {
				ATLASSUME(properties.pPropertyDescriptionList);
				while(properties.nextPropertyKeyToFetch < properties.numberOfPropertyKeys) {
					if(DPA_AppendPtr(properties.pResult->hColumnBuffer, reinterpret_cast<LPVOID>(-1)) == -1) {
						return E_OUTOFMEMORY;
					}

					CComPtr<IPropertyDescription> pPropertyDescription = NULL;
					hr = properties.pPropertyDescriptionList->GetAt(properties.nextPropertyKeyToFetch, IID_PPV_ARGS(&pPropertyDescription));
					ATLASSERT(SUCCEEDED(hr));
					if(FAILED(hr)) {
						break;
					}
					ATLASSUME(pPropertyDescription);

					hr = pPropertyDescription->GetPropertyKey(&properties.pPropertyKeys[properties.nextPropertyKeyToFetch]);
					ATLASSERT(SUCCEEDED(hr));
					if(FAILED(hr)) {
						break;
					}

					++properties.nextPropertyKeyToFetch;
					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}
				properties.status = TVTS_FINDINGCOLUMNS;
				properties.pPropertyDescriptionList->Release();
				properties.pPropertyDescriptionList = NULL;
			}

			if(properties.status == TVTS_FINDINGCOLUMNS) {
				SHELLDETAILS shellDetails = {0};
				while(properties.pParentISF2->GetDetailsOf(NULL, properties.currentRealColumnIndex, &shellDetails) == S_OK) {
					SHCOLUMNID propertyKey = {0};
					hr = properties.pParentISF2->MapColumnToSCID(properties.currentRealColumnIndex, &propertyKey);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						for(UINT i = 0; i < properties.numberOfPropertyKeys; ++i) {
							if(propertyKey.fmtid == properties.pPropertyKeys[i].fmtid && propertyKey.pid == properties.pPropertyKeys[i].pid) {
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
						++i;
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

STDMETHODIMP ShLvwTileViewTask::DoRun(void)
{
	properties.nextPropertyKeyToFetch = 0;
	properties.numberOfPropertyKeys = 0;
	properties.currentRealColumnIndex = 0;

	/* NOTE: We could also use IShellItem2::GetPropertyDescriptionList, but this list doesn't seem to be
	         always up-to-date. */

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

	HRESULT hr, hr2;

	/* TODO: According to an older comment, IShellItem2::GetPropertyDescriptionList isn't always reliable and
	         tends to retrieve a list that is out of date. Is this still correct? */
	VARIANT v;
	hr = APIWrapper::AssocGetDetailsOfPropKey(pParentISF, ILFindLastID(properties.pIDL), &properties.propertyListToLoad, &v, NULL, &hr2);
	if(SUCCEEDED(hr) && SUCCEEDED(hr2)) {
		hr = APIWrapper::PSGetPropertyDescriptionListFromString(OLE2W_EX_DEF(v.bstrVal), IID_PPV_ARGS(&properties.pPropertyDescriptionList), &hr2);
		VariantClear(&v);
	}
	if(!properties.pPropertyDescriptionList) {
		// try IShellItem2::GetPropertyDescritpionList
		CComPtr<IShellItem2> pShellItem;
		hr = APIWrapper::SHCreateItemFromIDList(properties.pIDL, IID_PPV_ARGS(&pShellItem), &hr2);
		if(pShellItem) {
			hr = pShellItem->GetPropertyDescriptionList(properties.propertyListToLoad, IID_PPV_ARGS(&properties.pPropertyDescriptionList));
		}
		if(FAILED(hr) || FAILED(hr2)) {
			return E_FAIL;
		}
	}
	ATLASSUME(properties.pPropertyDescriptionList);

	hr = properties.pPropertyDescriptionList->GetCount(&properties.numberOfPropertyKeys);
	ATLASSERT(SUCCEEDED(hr));
	if(FAILED(hr)) {
		return E_FAIL;
	}
	ATLASSERT(properties.numberOfPropertyKeys > 0);
	if(properties.numberOfPropertyKeys > 0) {
		properties.pPropertyKeys = static_cast<PROPERTYKEY*>(HeapAlloc(GetProcessHeap(), 0, properties.numberOfPropertyKeys * sizeof(PROPERTYKEY)));
		if(!properties.pPropertyKeys) {
			return E_OUTOFMEMORY;
		}
	}
	properties.status = TVTS_COLLECTINGPROPERTYKEYS;
	return E_PENDING;
}