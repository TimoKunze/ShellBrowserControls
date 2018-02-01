// InsertSingleItemTask.cpp: A specialization of RunnableTask for inserting a single listview item

#include "stdafx.h"
#include "InsertSingleItemTask.h"


ShLvwInsertSingleItemTask::ShLvwInsertSingleItemTask(void)
    : RunnableTask(FALSE)
{
	properties.pNamespaceObject = NULL;
	properties.pEnumratedItemsQueue = NULL;
	properties.pEnumResult = NULL;
	properties.pCriticalSection = NULL;
	properties.taskID = GetNewTaskID();
}

void ShLvwInsertSingleItemTask::FinalRelease()
{
	if(properties.pEnumSettings) {
		properties.pEnumSettings->Release();
		properties.pEnumSettings = NULL;
	}
	if(properties.pNamespaceObject) {
		properties.pNamespaceObject->Release();
		properties.pNamespaceObject = NULL;
	}
	if(properties.pIDLToAdd) {
		ILFree(properties.pIDLToAdd);
		properties.pIDLToAdd = NULL;
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
STDMETHODIMP ShLvwInsertSingleItemTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = properties.taskID;
	return S_OK;
}

STDMETHODIMP ShLvwInsertSingleItemTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
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

STDMETHODIMP ShLvwInsertSingleItemTask::SetTarget(LPDISPATCH pNamespaceObject)
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
	HRESULT ShLvwInsertSingleItemTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, LONG insertAt, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit)
#else
	HRESULT ShLvwInsertSingleItemTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, LONG insertAt, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.pIDLToAdd = ILCloneFull(pIDLToAdd);
	this->properties.isSimplePIDL = isSimplePIDL;
	this->properties.pIDLParent = ILCloneFull(pIDLParent);
	this->properties.pEnumSettings = pEnumSettings;
	this->properties.pEnumSettings->AddRef();
	this->properties.autoLabelEdit = autoLabelEdit;

	properties.pEnumResult = new SHLVWBACKGROUNDITEMENUMINFO;
	if(properties.pEnumResult) {
		ZeroMemory(properties.pEnumResult, sizeof(SHLVWBACKGROUNDITEMENUMINFO));
		properties.pEnumResult->taskID = properties.taskID;
		properties.pEnumResult->insertAt = insertAt;
		properties.pEnumResult->checkForDuplicates = checkForDuplicates;
		properties.pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		properties.pEnumResult->hPIDLBuffer = DPA_Create(1);
	} else {
		return E_OUTOFMEMORY;
	}

	properties.enumSettings = CacheEnumSettings(pEnumSettings);
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShLvwInsertSingleItemTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, LONG insertAt, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, IRunnableTask** ppTask)
#else
	HRESULT ShLvwInsertSingleItemTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, LONG insertAt, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	ATLASSUME(pEnumSettings);
	if(!pEnumratedItemsQueue || !pCriticalSection || !pIDLToAdd || !pIDLParent || !pEnumSettings) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwInsertSingleItemTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwInsertSingleItemTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, insertAt, pIDLToAdd, isSimplePIDL, pIDLParent, pEnumSettings, namespacePIDLToSet, checkForDuplicates, autoLabelEdit);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwInsertSingleItemTask::DoRun(void)
{
	ATLASSERT_POINTER(properties.pEnumResult, SHLVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(properties.pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	CComPtr<IShellFolder> pParentISF = NULL;
	PUITEMID_CHILD pIDLToAddRelative = NULL;
	HRESULT hr2 = SHBindToParent(properties.pIDLToAdd, IID_PPV_ARGS(&pParentISF), const_cast<PCUITEMID_CHILD*>(&pIDLToAddRelative));
	ATLASSERT(SUCCEEDED(hr2));
	if(FAILED(hr2)) {
		goto Done;
	}

	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pIDLToAddRelative, ITEMID_CHILD);

	PIDLIST_ABSOLUTE pIDLSubItem = NULL;
	if(properties.isSimplePIDL) {
		// we need a real pIDL
		PUITEMID_CHILD pIDL = NULL;
		if(SUCCEEDED(SHGetRealIDL(pParentISF, pIDLToAddRelative, &pIDL))) {
			ATLASSERT_POINTER(pIDL, ITEMID_CHILD);
			ATLASSERT(pIDL != pIDLToAddRelative);
			pIDLToAddRelative = pIDL;

			PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(properties.pIDLParent);
			ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
			pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pIDLToAddRelative), TRUE));
		} else {
			// well, it already was a real pIDL...
			pIDLSubItem = properties.pIDLToAdd;
			properties.pIDLToAdd = NULL;
			properties.isSimplePIDL = FALSE;
		}
	} else {
		pIDLSubItem = properties.pIDLToAdd;
		properties.pIDLToAdd = NULL;
	}
	ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);

	if(!ShouldShowItem(pParentISF, pIDLToAddRelative, &properties.enumSettings)) {
		if(properties.isSimplePIDL) {
			ILFree(pIDLToAddRelative);
			pIDLToAddRelative = NULL;
		}
		ILFree(pIDLSubItem);
		goto Done;
	}
	if(properties.isSimplePIDL) {
		ILFree(pIDLToAddRelative);
		pIDLToAddRelative = NULL;
	}

	if(DPA_AppendPtr(properties.pEnumResult->hPIDLBuffer, pIDLSubItem) == -1) {
		ILFree(pIDLSubItem);
		hr = E_OUTOFMEMORY;
		goto Done;
	}

	if(properties.autoLabelEdit) {
		properties.pEnumResult->pIDLToLabelEdit = pIDLSubItem;
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