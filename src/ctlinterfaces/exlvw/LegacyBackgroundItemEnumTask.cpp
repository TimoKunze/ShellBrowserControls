// LegacyBackgroundItemEnumTask.cpp: A specialization of RunnableTask for enumerating listview items

#include "stdafx.h"
#include "LegacyBackgroundItemEnumTask.h"


ShLvwLegacyBackgroundItemEnumTask::ShLvwLegacyBackgroundItemEnumTask(void)
    : RunnableTask(TRUE)
{
	properties.pNamespaceObject = NULL;
	properties.pEnumratedItemsQueue = NULL;
	properties.pEnumResult = NULL;
	properties.pCriticalSection = NULL;
	properties.taskID = GetNewTaskID();
}

void ShLvwLegacyBackgroundItemEnumTask::FinalRelease()
{
	if(properties.pMarshaledParentISF) {
		properties.pMarshaledParentISF->Release();
		properties.pMarshaledParentISF = NULL;
	}
	if(properties.pParentISF) {
		properties.pParentISF->Release();
		properties.pParentISF = NULL;
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
			DPA_DestroyCallback(properties.pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
		}
		delete properties.pEnumResult;
		properties.pEnumResult = NULL;
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
	*pTaskID = properties.taskID;
	return S_OK;
}

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
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

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
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
	HRESULT ShLvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
#else
	HRESULT ShLvwLegacyBackgroundItemEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, LONG insertAt, PCIDLIST_ABSOLUTE pIDLParent, LPSTREAM pMarshaledParentISF, IEnumIDList* pEnumerator, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL forceAutoSort)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pIDLParent = ILCloneFull(pIDLParent);
	this->properties.pMarshaledParentISF = pMarshaledParentISF;
	if(this->properties.pMarshaledParentISF) {
		this->properties.pMarshaledParentISF->AddRef();
	}
	this->properties.pEnumerator = pEnumerator;
	if(this->properties.pEnumerator) {
		this->properties.pEnumerator->AddRef();
	}
	this->properties.pEnumSettings = pEnumSettings;
	this->properties.pEnumSettings->AddRef();

	properties.pParentISF = NULL;
	properties.pEnumResult = new SHLVWBACKGROUNDITEMENUMINFO;
	if(properties.pEnumResult) {
		ZeroMemory(properties.pEnumResult, sizeof(SHLVWBACKGROUNDITEMENUMINFO));
		properties.pEnumResult->taskID = properties.taskID;
		properties.pEnumResult->insertAt = insertAt;
		properties.pEnumResult->checkForDuplicates = checkForDuplicates;
		properties.pEnumResult->forceAutoSort = forceAutoSort;
		properties.pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		properties.pEnumResult->hPIDLBuffer = DPA_Create(16);
	} else {
		return E_OUTOFMEMORY;
	}
	properties.enumSettings = CacheEnumSettings(pEnumSettings);
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
	ATLASSERT_POINTER(properties.pEnumResult, SHLVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(properties.pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	if(!properties.pEnumerator) {
		properties.pParentISF->EnumObjects(properties.hWndShellUIParentWindow, properties.enumSettings.enumerationFlags, &properties.pEnumerator);
		if(!properties.pEnumerator) {
			goto Done;
		}

		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	PUITEMID_CHILD pIDLChild = NULL;
	while(properties.pEnumerator->Next(1, &pIDLChild, NULL) == S_OK) {
		if(properties.isSlowNamespace) {
			for(int i = 0; TRUE; ++i) {
				PIDLIST_ABSOLUTE pIDLItem = reinterpret_cast<PIDLIST_ABSOLUTE>(DPA_GetPtr(properties.pEnumResult->hPIDLBuffer, i));
				if(!pIDLItem) {
					break;
				}
				if(static_cast<short>(HRESULT_CODE(properties.pParentISF->CompareIDs(SHCIDS_ALLFIELDS, ILFindLastID(pIDLItem), pIDLChild))) == 0) {
					ILFree(pIDLChild);
					pIDLChild = NULL;
					break;
				}
			}
		}
		if(!pIDLChild) {
			continue;
		}

		if(!ShouldShowItem(properties.pParentISF, pIDLChild, &properties.enumSettings)) {
			ILFree(pIDLChild);
			if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
				return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
			}
			continue;
		}

		// build a fully qualified pIDL
		PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(properties.pIDLParent);
		ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
		PIDLIST_ABSOLUTE pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pIDLChild), TRUE));
		ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);
		if(!pIDLSubItem) {
			ILFree(pIDLChild);
			pIDLChild = NULL;
			hr = E_OUTOFMEMORY;
			goto Done;
		}

		if(DPA_AppendPtr(properties.pEnumResult->hPIDLBuffer, pIDLSubItem) == -1) {
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

STDMETHODIMP ShLvwLegacyBackgroundItemEnumTask::DoRun(void)
{
	if(properties.pMarshaledParentISF) {
		CoGetInterfaceAndReleaseStream(properties.pMarshaledParentISF, IID_PPV_ARGS(&properties.pParentISF));
		properties.pMarshaledParentISF = NULL;
	} else {
		BindToPIDL(properties.pIDLParent, IID_PPV_ARGS(&properties.pParentISF));
	}
	if(!properties.pParentISF) {
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
	properties.isSlowNamespace = IsSlowItem(properties.pIDLParent, FALSE, TRUE);

	return E_PENDING;
}