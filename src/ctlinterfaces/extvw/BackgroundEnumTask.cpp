// BackgroundEnumTask.cpp: A specialization of RunnableTask for enumerating treeview items

#include "stdafx.h"
#include "BackgroundEnumTask.h"


ShTvwBackgroundItemEnumTask::ShTvwBackgroundItemEnumTask(void)
    : RunnableTask(TRUE)
{
	properties.pNamespaceObject = NULL;
	properties.pEnumratedItemsQueue = NULL;
	properties.pEnumResult = NULL;
	properties.pCriticalSection = NULL;
	properties.taskID = GetNewTaskID();
}

void ShTvwBackgroundItemEnumTask::FinalRelease()
{
	if(properties.pMarshaledParentISI) {
		properties.pMarshaledParentISI->Release();
		properties.pMarshaledParentISI = NULL;
	}
	if(properties.pParentISI) {
		properties.pParentISI->Release();
		properties.pParentISI = NULL;
	}
	if(properties.pEnumerator) {
		properties.pEnumerator->Release();
		properties.pEnumerator = NULL;
	}
	if(properties.pEnumSettings) {
		properties.pEnumSettings->Release();
		properties.pEnumSettings = NULL;
	}
	if(properties.pNamespaceObject) {
		properties.pNamespaceObject->Release();
		properties.pNamespaceObject = NULL;
	}
	if(properties.pIDLParent) {
		ILFree(properties.pIDLParent);
		properties.pIDLParent = NULL;
	}
	if(properties.pEnumResult) {
		if(properties.pEnumResult->hPIDLBuffer) {
			DPA_DestroyCallback(properties.pEnumResult->hPIDLBuffer, FreeDPAEnumeratedItemElement, NULL);
		}
		delete properties.pEnumResult;
		properties.pEnumResult = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShTvwBackgroundItemEnumTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = properties.taskID;
	return S_OK;
}

STDMETHODIMP ShTvwBackgroundItemEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSERT_POINTER(ppInterfaceImpl, LPVOID);
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(!properties.pNamespaceObject) {
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	return properties.pNamespaceObject->QueryInterface(requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP ShTvwBackgroundItemEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
{
	if(!pNamespaceObject) {
		return E_INVALIDARG;
	}

	if(this->properties.pNamespaceObject) {
		this->properties.pNamespaceObject->Release();
		this->properties.pNamespaceObject = NULL;
	}

	if(pNamespaceObject) {
		return pNamespaceObject->QueryInterface(IID_PPV_ARGS(&this->properties.pNamespaceObject));
	}
	return S_OK;
}
// implementation of INamespaceEnumTask
//////////////////////////////////////////////////////////////////////


#ifdef USE_STL
	HRESULT ShTvwBackgroundItemEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates)
#else
	HRESULT ShTvwBackgroundItemEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pIDLParent = ILCloneFull(pIDLParent);
	this->properties.pMarshaledParentISI = pMarshaledParentISI;
	if(this->properties.pMarshaledParentISI) {
		this->properties.pMarshaledParentISI->AddRef();
	}
	this->properties.pEnumerator = pEnumerator;
	if(this->properties.pEnumerator) {
		this->properties.pEnumerator->AddRef();
	}
	this->properties.pEnumSettings = pEnumSettings;
	this->properties.pEnumSettings->AddRef();

	properties.pParentISI = NULL;
	properties.pEnumResult = new SHTVWBACKGROUNDITEMENUMINFO;
	if(properties.pEnumResult) {
		ZeroMemory(properties.pEnumResult, sizeof(SHTVWBACKGROUNDITEMENUMINFO));
		properties.pEnumResult->taskID = properties.taskID;
		properties.pEnumResult->hParentItem = hParentItem;
		properties.pEnumResult->hInsertAfter = (hInsertAfter ? hInsertAfter : TVI_LAST);
		properties.pEnumResult->checkForDuplicates = checkForDuplicates;
		properties.pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		properties.pEnumResult->hPIDLBuffer = DPA_Create(16);
	} else {
		return E_OUTOFMEMORY;
	}

	properties.enumSettings = CacheEnumSettings(pEnumSettings);
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShTvwBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, IRunnableTask** ppTask)
#else
	HRESULT ShTvwBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISI, IEnumShellItems* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, IRunnableTask** ppTask)
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

	CComObject<ShTvwBackgroundItemEnumTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShTvwBackgroundItemEnumTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, hWndShellUIParentWindow, hParentItem, hInsertAfter, pIDLParent, pMarshaledParentISI, pEnumerator, pEnumSettings, namespacePIDLToSet, checkForDuplicates);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShTvwBackgroundItemEnumTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(properties.pEnumResult, SHTVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(properties.pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	if(!properties.pEnumerator) {
		CComPtr<IBindCtx> pBindContext;
		// Note that STR_ENUM_ITEMS_FLAGS is ignored on Windows Vista and 7! That's why we use this code path only on Windows 8 and newer.
		if(SUCCEEDED(CreateDwordBindCtx(STR_ENUM_ITEMS_FLAGS, properties.enumSettings.enumerationFlags, &pBindContext))) {
			properties.pParentISI->BindToHandler(pBindContext, BHID_EnumItems, IID_PPV_ARGS(&properties.pEnumerator));
		}
		if(!properties.pEnumerator) {
			goto Done;
		}

		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	IShellItem* pChildItem = NULL;
	while(properties.pEnumerator->Next(1, &pChildItem, NULL) == S_OK) {
		if(properties.isSlowNamespace) {
			for(int i = 0; TRUE; ++i) {
				LPENUMERATEDITEM pItem = reinterpret_cast<LPENUMERATEDITEM>(DPA_GetPtr(properties.pEnumResult->hPIDLBuffer, i));
				if(!pItem) {
					break;
				}
				CComPtr<IShellItem> pShellItem;
				HRESULT hr2 = E_FAIL;
				if(SUCCEEDED(APIWrapper::SHCreateItemFromIDList(pItem->pIDL, IID_PPV_ARGS(&pShellItem), &hr2)) && SUCCEEDED(hr2) && pShellItem) {
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

		if(!ShouldShowItem(pChildItem, &properties.enumSettings)) {
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
		HasSubItems hasSubItems = HasAtLeastOneSubItem(pChildItem, &properties.enumSettings, TRUE);

		pChildItem->Release();
		pChildItem = NULL;

		LPENUMERATEDITEM pItem = new ENUMERATEDITEM(pIDLSubItem, hasSubItems);
		if(!pItem || (DPA_AppendPtr(properties.pEnumResult->hPIDLBuffer, pItem) == -1)) {
			ILFree(pIDLSubItem);
			hr = E_OUTOFMEMORY;
			break;	//goto Done;
		}

		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

Done:
	EnterCriticalSection(properties.pCriticalSection);
	#ifdef USE_STL
		properties.pEnumratedItemsQueue->push(properties.pEnumResult);
	#else
		properties.pEnumratedItemsQueue->AddTail(properties.pEnumResult);
	#endif
	properties.pEnumResult = NULL;
	LeaveCriticalSection(properties.pCriticalSection);

	if(IsWindow(properties.hWndToNotify)) {
		PostMessage(properties.hWndToNotify, WM_TRIGGER_ITEMENUMCOMPLETE, 0, 0);
	}
	return hr;
}

STDMETHODIMP ShTvwBackgroundItemEnumTask::DoRun(void)
{
	if(properties.pMarshaledParentISI) {
		CoGetInterfaceAndReleaseStream(properties.pMarshaledParentISI, IID_PPV_ARGS(&properties.pParentISI));
		properties.pMarshaledParentISI = NULL;
	} else if(APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		APIWrapper::SHCreateItemFromIDList(properties.pIDLParent, IID_PPV_ARGS(&properties.pParentISI), NULL);
	}
	if(!properties.pParentISI) {
		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pEnumratedItemsQueue->push(properties.pEnumResult);
		#else
			properties.pEnumratedItemsQueue->AddTail(properties.pEnumResult);
		#endif
		properties.pEnumResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);

		if(IsWindow(properties.hWndToNotify)) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_ITEMENUMCOMPLETE, 0, 0);
		}
		return E_FAIL;
	}
	properties.isSlowNamespace = IsSlowItem(properties.pParentISI, FALSE, TRUE);

	return E_PENDING;
}