// BackgroundItemEnumTask.cpp: A specialization of RunnableTask for enumerating listview items

#include "stdafx.h"
#include "BackgroundItemEnumTask.h"


ShLvwBackgroundItemEnumTask::ShLvwBackgroundItemEnumTask(void)
    : RunnableTask(TRUE)
{
	pNamespaceObject = NULL;
	pEnumratedItemsQueue = NULL;
	pEnumResult = NULL;
	pCriticalSection = NULL;
	taskID = GetNewTaskID();
}

void ShLvwBackgroundItemEnumTask::FinalRelease()
{
	if(pMarshaledParentISI) {
		pMarshaledParentISI->Release();
		pMarshaledParentISI = NULL;
	}
	if(pParentISI) {
		pParentISI->Release();
		pParentISI = NULL;
	}
	if(pEnumerator) {
		pEnumerator->Release();
		pEnumerator = NULL;
	}
	if(pEnumSettings) {
		pEnumSettings->Release();
		pEnumSettings = NULL;
	}
	if(pNamespaceObject) {
		pNamespaceObject->Release();
		pNamespaceObject = NULL;
	}
	if(pIDLParent) {
		ILFree(pIDLParent);
		pIDLParent = NULL;
	}
	if(pEnumResult) {
		if(pEnumResult->hPIDLBuffer) {
			DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
		}
		delete pEnumResult;
		pEnumResult = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShLvwBackgroundItemEnumTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = taskID;
	return S_OK;
}

STDMETHODIMP ShLvwBackgroundItemEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSERT_POINTER(ppInterfaceImpl, LPVOID);
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(!pNamespaceObject) {
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	return pNamespaceObject->QueryInterface(requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP ShLvwBackgroundItemEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
{
	if(!pNamespaceObject) {
		return E_INVALIDARG;
	}

	if(this->pNamespaceObject) {
		this->pNamespaceObject->Release();
		this->pNamespaceObject = NULL;
	}

	if(pNamespaceObject) {
		return pNamespaceObject->QueryInterface(IID_PPV_ARGS(&this->pNamespaceObject));
	}
	return S_OK;
}
// implementation of INamespaceEnumTask
//////////////////////////////////////////////////////////////////////


#ifdef USE_STL
	HRESULT ShLvwBackgroundItemEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
#else
	HRESULT ShLvwBackgroundItemEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->pCriticalSection = pCriticalSection;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pIDLParent = ILCloneFull(pIDLParent);
	this->pMarshaledParentISI = pMarshaledParentISI;
	if(this->pMarshaledParentISI) {
		this->pMarshaledParentISI->AddRef();
	}
	this->pEnumerator = pEnumerator;
	if(this->pEnumerator) {
		this->pEnumerator->AddRef();
	}
	this->pEnumSettings = pEnumSettings;
	this->pEnumSettings->AddRef();

	pParentISI = NULL;
	pEnumResult = new SHLVWBACKGROUNDITEMENUMINFO;
	if(pEnumResult) {
		ZeroMemory(pEnumResult, sizeof(SHLVWBACKGROUNDITEMENUMINFO));
		pEnumResult->taskID = taskID;
		pEnumResult->insertAt = insertAt;
		pEnumResult->checkForDuplicates = checkForDuplicates;
		pEnumResult->forceAutoSort = forceAutoSort;
		pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		pEnumResult->hPIDLBuffer = DPA_Create(16);
	} else {
		return E_OUTOFMEMORY;
	}
	enumSettings = CacheEnumSettings(pEnumSettings);
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShLvwBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, IRunnableTask** ppTask)
#else
	HRESULT ShLvwBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	ATLASSUME(pEnumSettings);
	if(!pEnumratedItemsQueue || !pCriticalSection || !pIDLParent || !pEnumSettings) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwBackgroundItemEnumTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwBackgroundItemEnumTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, hWndShellUIParentWindow, insertAt, pIDLParent, pMarshaledParentISI, pEnumerator, pEnumSettings, namespacePIDLToSet, checkForDuplicates, forceAutoSort);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwBackgroundItemEnumTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pEnumResult, SHLVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	if(!pEnumerator) {
		CComPtr<IBindCtx> pBindContext;
		// Note that STR_ENUM_ITEMS_FLAGS is ignored on Windows Vista and 7! That's why we use this code path only on Windows 8 and newer.
		if(SUCCEEDED(CreateDwordBindCtx(STR_ENUM_ITEMS_FLAGS, enumSettings.enumerationFlags, &pBindContext))) {
			pParentISI->BindToHandler(pBindContext, BHID_EnumItems, IID_PPV_ARGS(&pEnumerator));
		}
		if(!pEnumerator) {
			goto Done;
		}

		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	IShellItem* pChildItem = NULL;
	while(pEnumerator->Next(1, &pChildItem, NULL) == S_OK) {
		if(isSlowNamespace) {
			for(int i = 0; TRUE; ++i) {
				PIDLIST_ABSOLUTE pIDLItem = reinterpret_cast<PIDLIST_ABSOLUTE>(DPA_GetPtr(pEnumResult->hPIDLBuffer, i));
				if(!pIDLItem) {
					break;
				}
				CComPtr<IShellItem> pShellItem;
				HRESULT hr2 = E_FAIL;
				if(SUCCEEDED(APIWrapper::SHCreateItemFromIDList(pIDLItem, IID_PPV_ARGS(&pShellItem), &hr2)) && SUCCEEDED(hr2) && pShellItem) {
					int order = 0;
					if(SUCCEEDED(pShellItem->Compare(pChildItem, static_cast<SICHINTF>(SICHINT_ALLFIELDS), &order)) && order == 0) {
						pChildItem->Release();
						pChildItem = NULL;
						break;
					}
				}
			}
		}
		if(!pChildItem) {
			continue;
		}

		if(!ShouldShowItem(pChildItem, &enumSettings)) {
			pChildItem->Release();
			pChildItem = NULL;
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}
			continue;
		}

		PIDLIST_ABSOLUTE pIDLSubItem = NULL;
		if(FAILED(CComQIPtr<IPersistIDList>(pChildItem)->GetIDList(&pIDLSubItem)) || !pIDLSubItem) {
			pChildItem->Release();
			pChildItem = NULL;
			pIDLSubItem = NULL;
			hr = E_OUTOFMEMORY;
			break;	//goto Done;
		}
		ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);

		pChildItem->Release();
		pChildItem = NULL;
		if(DPA_AppendPtr(pEnumResult->hPIDLBuffer, pIDLSubItem) == -1) {
			ILFree(pIDLSubItem);
			hr = E_OUTOFMEMORY;
			break;	//goto Done;
		}
		
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

Done:
	EnterCriticalSection(pCriticalSection);
	#ifdef USE_STL
		pEnumratedItemsQueue->push(pEnumResult);
	#else
		pEnumratedItemsQueue->AddTail(pEnumResult);
	#endif
	pEnumResult = NULL;
	LeaveCriticalSection(pCriticalSection);

	if(IsWindow(hWndToNotify)) {
		PostMessage(hWndToNotify, WM_TRIGGER_ITEMENUMCOMPLETE, 0, 0);
	}
	return hr;
}

STDMETHODIMP ShLvwBackgroundItemEnumTask::DoRun(void)
{
	if(pMarshaledParentISI) {
		CoGetInterfaceAndReleaseStream(pMarshaledParentISI, IID_PPV_ARGS(&pParentISI));
		pMarshaledParentISI = NULL;
	} else if(APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		APIWrapper::SHCreateItemFromIDList(pIDLParent, IID_PPV_ARGS(&pParentISI), NULL);
	}
	if(!pParentISI) {
		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pEnumratedItemsQueue->push(pEnumResult);
		#else
			pEnumratedItemsQueue->AddTail(pEnumResult);
		#endif
		pEnumResult = NULL;
		LeaveCriticalSection(pCriticalSection);

		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_ITEMENUMCOMPLETE, 0, 0);
		}
		return E_FAIL;
	}
	isSlowNamespace = IsSlowItem(pParentISI, FALSE, TRUE);

	return E_PENDING;
}