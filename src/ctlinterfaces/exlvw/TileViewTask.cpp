// TileViewTask.cpp: A specialization of RunnableTask for retrieving shell item's tile view sub-items

#include "stdafx.h"
#include "TileViewTask.h"


ShLvwTileViewTask::ShLvwTileViewTask(void)
    : RunnableTask(TRUE)
{
	pBackgroundTileViewQueue = NULL;
	pParentISF2 = NULL;
	pPropertyDescriptionList = NULL;
	pPropertyKeys = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwTileViewTask::FinalRelease()
{
	if(pParentISF2) {
		pParentISF2->Release();
		pParentISF2 = NULL;
	}
	if(pPropertyDescriptionList) {
		pPropertyDescriptionList->Release();
		pPropertyDescriptionList = NULL;
	}
	if(pPropertyKeys) {
		HeapFree(GetProcessHeap(), 0, pPropertyKeys);
		pPropertyKeys = NULL;
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
	HRESULT ShLvwTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad)
#else
	HRESULT ShLvwTileViewTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO>* pBackgroundTileViewQueue, LPCRITICAL_SECTION pCriticalSection, HANDLE hColumnsReadyEvent, PCIDLIST_ABSOLUTE pIDL, LONG itemID, PCIDLIST_ABSOLUTE pIDLParentNamespace, UINT maxColumnCount, const PROPERTYKEY& propertyListToLoad)
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
	this->propertyListToLoad = propertyListToLoad;

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
	if(!(status == TVTS_DONE && !pResult)) {
		ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDTILEVIEWINFO);
		ATLASSUME(pResult->hColumnBuffer);
		ATLASSUME(pParentISF2);

		HRESULT hr;

		if(numberOfPropertyKeys <= 0) {
			return NOERROR;
		}
		if(status != TVTS_DONE) {
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}

			if(status == TVTS_COLLECTINGPROPERTYKEYS) {
				ATLASSUME(pPropertyDescriptionList);
				while(nextPropertyKeyToFetch < numberOfPropertyKeys) {
					if(DPA_AppendPtr(pResult->hColumnBuffer, reinterpret_cast<LPVOID>(-1)) == -1) {
						return E_OUTOFMEMORY;
					}

					CComPtr<IPropertyDescription> pPropertyDescription = NULL;
					hr = pPropertyDescriptionList->GetAt(nextPropertyKeyToFetch, IID_PPV_ARGS(&pPropertyDescription));
					ATLASSERT(SUCCEEDED(hr));
					if(FAILED(hr)) {
						break;
					}
					ATLASSUME(pPropertyDescription);

					hr = pPropertyDescription->GetPropertyKey(&pPropertyKeys[nextPropertyKeyToFetch]);
					ATLASSERT(SUCCEEDED(hr));
					if(FAILED(hr)) {
						break;
					}

					++nextPropertyKeyToFetch;
					if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
						return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
					}
				}
				status = TVTS_FINDINGCOLUMNS;
				pPropertyDescriptionList->Release();
				pPropertyDescriptionList = NULL;
			}

			if(status == TVTS_FINDINGCOLUMNS) {
				SHELLDETAILS shellDetails = {0};
				while(pParentISF2->GetDetailsOf(NULL, currentRealColumnIndex, &shellDetails) == S_OK) {
					SHCOLUMNID propertyKey = {0};
					hr = pParentISF2->MapColumnToSCID(currentRealColumnIndex, &propertyKey);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						for(UINT i = 0; i < numberOfPropertyKeys; ++i) {
							if(propertyKey.fmtid == pPropertyKeys[i].fmtid && propertyKey.pid == pPropertyKeys[i].pid) {
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
						++i;
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

STDMETHODIMP ShLvwTileViewTask::DoRun(void)
{
	nextPropertyKeyToFetch = 0;
	numberOfPropertyKeys = 0;
	currentRealColumnIndex = 0;

	/* NOTE: We could also use IShellItem2::GetPropertyDescriptionList, but this list doesn't seem to be
	         always up-to-date. */

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

	HRESULT hr, hr2;

	/* TODO: According to an older comment, IShellItem2::GetPropertyDescriptionList isn't always reliable and
	         tends to retrieve a list that is out of date. Is this still correct? */
	VARIANT v;
	hr = APIWrapper::AssocGetDetailsOfPropKey(pParentISF, ILFindLastID(pIDL), &propertyListToLoad, &v, NULL, &hr2);
	if(SUCCEEDED(hr) && SUCCEEDED(hr2)) {
		hr = APIWrapper::PSGetPropertyDescriptionListFromString(OLE2W_EX_DEF(v.bstrVal), IID_PPV_ARGS(&pPropertyDescriptionList), &hr2);
		VariantClear(&v);
	}
	if(!pPropertyDescriptionList) {
		// try IShellItem2::GetPropertyDescritpionList
		CComPtr<IShellItem2> pShellItem;
		hr = APIWrapper::SHCreateItemFromIDList(pIDL, IID_PPV_ARGS(&pShellItem), &hr2);
		if(pShellItem) {
			hr = pShellItem->GetPropertyDescriptionList(propertyListToLoad, IID_PPV_ARGS(&pPropertyDescriptionList));
		}
		if(FAILED(hr) || FAILED(hr2)) {
			return E_FAIL;
		}
	}
	ATLASSUME(pPropertyDescriptionList);

	hr = pPropertyDescriptionList->GetCount(&numberOfPropertyKeys);
	ATLASSERT(SUCCEEDED(hr));
	if(FAILED(hr)) {
		return E_FAIL;
	}
	ATLASSERT(numberOfPropertyKeys > 0);
	if(numberOfPropertyKeys > 0) {
		pPropertyKeys = static_cast<PROPERTYKEY*>(HeapAlloc(GetProcessHeap(), 0, numberOfPropertyKeys * sizeof(PROPERTYKEY)));
		if(!pPropertyKeys) {
			return E_OUTOFMEMORY;
		}
	}
	status = TVTS_COLLECTINGPROPERTYKEYS;
	return E_PENDING;
}