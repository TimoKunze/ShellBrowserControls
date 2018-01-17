// LegacyBackgroundEnumTask.cpp: A specialization of RunnableTask for enumerating treeview items

#include "stdafx.h"
#include "LegacyBackgroundEnumTask.h"


ShTvwLegacyBackgroundItemEnumTask::ShTvwLegacyBackgroundItemEnumTask(void)
    : RunnableTask(TRUE)
{
	pNamespaceObject = NULL;
	pEnumratedItemsQueue = NULL;
	pEnumResult = NULL;
	pCriticalSection = NULL;
	taskID = GetNewTaskID();
}

void ShTvwLegacyBackgroundItemEnumTask::FinalRelease()
{
	if(pMarshaledParentISF) {
		pMarshaledParentISF->Release();
		pMarshaledParentISF = NULL;
	}
	if(pParentISF) {
		pParentISF->Release();
		pParentISF = NULL;
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
			DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAEnumeratedItemElement, NULL);
		}
		delete pEnumResult;
		pEnumResult = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShTvwLegacyBackgroundItemEnumTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = taskID;
	return S_OK;
}

STDMETHODIMP ShTvwLegacyBackgroundItemEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
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

STDMETHODIMP ShTvwLegacyBackgroundItemEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
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
	HRESULT ShTvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates)
#else
	HRESULT ShTvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->pCriticalSection = pCriticalSection;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pIDLParent = ILCloneFull(pIDLParent);
	this->pMarshaledParentISF = pMarshaledParentISF;
	if(this->pMarshaledParentISF) {
		this->pMarshaledParentISF->AddRef();
	}
	this->pEnumerator = pEnumerator;
	if(this->pEnumerator) {
		this->pEnumerator->AddRef();
	}
	this->pEnumSettings = pEnumSettings;
	this->pEnumSettings->AddRef();

	pParentISF = NULL;
	pEnumResult = new SHTVWBACKGROUNDITEMENUMINFO;
	if(pEnumResult) {
		ZeroMemory(pEnumResult, sizeof(SHTVWBACKGROUNDITEMENUMINFO));
		pEnumResult->taskID = taskID;
		pEnumResult->hParentItem = hParentItem;
		pEnumResult->hInsertAfter = (hInsertAfter ? hInsertAfter : TVI_LAST);
		pEnumResult->checkForDuplicates = checkForDuplicates;
		pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		pEnumResult->hPIDLBuffer = DPA_Create(16);
	} else {
		return E_OUTOFMEMORY;
	}

	enumSettings = CacheEnumSettings(pEnumSettings);
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShTvwLegacyBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, IRunnableTask** ppTask)
#else
	HRESULT ShTvwLegacyBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, IRunnableTask** ppTask)
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

	CComObject<ShTvwLegacyBackgroundItemEnumTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShTvwLegacyBackgroundItemEnumTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, hWndShellUIParentWindow, hParentItem, hInsertAfter, pIDLParent, pMarshaledParentISF, pEnumerator, pEnumSettings, namespacePIDLToSet, checkForDuplicates);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShTvwLegacyBackgroundItemEnumTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pEnumResult, SHTVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	if(!pEnumerator) {
		pParentISF->EnumObjects(hWndShellUIParentWindow, enumSettings.enumerationFlags, &pEnumerator);
		if(!pEnumerator) {
			goto Done;
		}

		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	PUITEMID_CHILD pIDLChild = NULL;
	while(pEnumerator->Next(1, &pIDLChild, NULL) == S_OK) {
		if(isSlowNamespace) {
			for(int i = 0; TRUE; ++i) {
				LPENUMERATEDITEM pItem = reinterpret_cast<LPENUMERATEDITEM>(DPA_GetPtr(pEnumResult->hPIDLBuffer, i));
				if(!pItem) {
					break;
				}
				if(static_cast<short>(HRESULT_CODE(pParentISF->CompareIDs(SHCIDS_ALLFIELDS, ILFindLastID(pItem->pIDL), pIDLChild))) == 0) {
					ILFree(pIDLChild);
					pIDLChild = NULL;
					break;
				}
			}
		}
		if(!pIDLChild) {
			continue;
		}

		if(!ShouldShowItem(pParentISF, pIDLChild, &enumSettings)) {
			ILFree(pIDLChild);
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}
			continue;
		}

		// build a fully qualified pIDL
		PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(pIDLParent);
		ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
		PIDLIST_ABSOLUTE pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pIDLChild), TRUE));
		ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);
		if(!pIDLSubItem) {
			ILFree(pIDLChild);
			pIDLChild = NULL;
			hr = E_OUTOFMEMORY;
			goto Done;
		}

		HasSubItems hasSubItems = HasAtLeastOneSubItem(pParentISF, pIDLChild, pIDLSubItem, &enumSettings, TRUE);
		ILFree(pIDLChild);
		pIDLChild = NULL;

		LPENUMERATEDITEM pItem = new ENUMERATEDITEM(pIDLSubItem, hasSubItems);
		if(!pItem || (DPA_AppendPtr(pEnumResult->hPIDLBuffer, pItem) == -1)) {
			ILFree(pIDLSubItem);
			hr = E_OUTOFMEMORY;
			goto Done;
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

STDMETHODIMP ShTvwLegacyBackgroundItemEnumTask::DoRun(void)
{
	if(pMarshaledParentISF) {
		CoGetInterfaceAndReleaseStream(pMarshaledParentISF, IID_PPV_ARGS(&pParentISF));
		pMarshaledParentISF = NULL;
	} else {
		BindToPIDL(pIDLParent, IID_PPV_ARGS(&pParentISF));
	}
	if(!pParentISF) {
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
	isSlowNamespace = IsSlowItem(pIDLParent, FALSE, TRUE);

	return E_PENDING;
}