// LegacyBackgroundItemEnumTask.cpp: A specialization of RunnableTask for enumerating listview items

#include "stdafx.h"
#include "LegacyBackgroundItemEnumTask.h"


ShLvwLegacyBackgroundItemEnumTask::ShLvwLegacyBackgroundItemEnumTask(void)
    : RunnableTask(TRUE)
{
	pNamespaceObject = NULL;
	pEnumratedItemsQueue = NULL;
	pEnumResult = NULL;
	pCriticalSection = NULL;
	taskID = GetNewTaskID();
}

void ShLvwLegacyBackgroundItemEnumTask::FinalRelease()
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
			DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
		}
		delete pEnumResult;
		pEnumResult = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = taskID;
	return S_OK;
}

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
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

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
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
	HRESULT ShLvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
#else
	HRESULT ShLvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
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
	HRESULT ShLvwLegacyBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyBackgroundItemEnumTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort, IRunnableTask** ppTask)
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

	CComObject<ShLvwLegacyBackgroundItemEnumTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyBackgroundItemEnumTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, hWndShellUIParentWindow, insertAt, pIDLParent, pMarshaledParentISF, pEnumerator, pEnumSettings, namespacePIDLToSet, checkForDuplicates, forceAutoSort);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pEnumResult, SHLVWBACKGROUNDITEMENUMINFO);
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
				PIDLIST_ABSOLUTE pIDLItem = reinterpret_cast<PIDLIST_ABSOLUTE>(DPA_GetPtr(pEnumResult->hPIDLBuffer, i));
				if(!pIDLItem) {
					break;
				}
				if(static_cast<short>(HRESULT_CODE(pParentISF->CompareIDs(SHCIDS_ALLFIELDS, ILFindLastID(pIDLItem), pIDLChild))) == 0) {
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

		if(DPA_AppendPtr(pEnumResult->hPIDLBuffer, pIDLSubItem) == -1) {
			ILFree(pIDLSubItem);
			hr = E_OUTOFMEMORY;
			goto Done;
		}
		ILFree(pIDLChild);
		pIDLChild = NULL;
		
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

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::DoRun(void)
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