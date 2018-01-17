// InsertSingleItemTask.cpp: A specialization of RunnableTask for inserting a single treeview item

#include "stdafx.h"
#include "InsertSingleItemTask.h"


ShTvwInsertSingleItemTask::ShTvwInsertSingleItemTask(void)
    : RunnableTask(FALSE)
{
	pNamespaceObject = NULL;
	pEnumratedItemsQueue = NULL;
	pEnumResult = NULL;
	pCriticalSection = NULL;
	taskID = GetNewTaskID();
}

void ShTvwInsertSingleItemTask::FinalRelease()
{
	if(pEnumSettings) {
		pEnumSettings->Release();
		pEnumSettings = NULL;
	}
	if(pNamespaceObject) {
		pNamespaceObject->Release();
		pNamespaceObject = NULL;
	}
	if(pIDLToAdd) {
		ILFree(pIDLToAdd);
		pIDLToAdd = NULL;
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
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShTvwInsertSingleItemTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = taskID;
	return S_OK;
}

STDMETHODIMP ShTvwInsertSingleItemTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
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

STDMETHODIMP ShTvwInsertSingleItemTask::SetTarget(LPDISPATCH pNamespaceObject)
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
	HRESULT ShTvwInsertSingleItemTask::Attach(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit)
#else
	HRESULT ShTvwInsertSingleItemTask::Attach(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pEnumratedItemsQueue = pEnumratedItemsQueue;
	this->pCriticalSection = pCriticalSection;
	this->pIDLToAdd = ILCloneFull(pIDLToAdd);
	this->isSimplePIDL = isSimplePIDL;
	this->pIDLParent = ILCloneFull(pIDLParent);
	this->pEnumSettings = pEnumSettings;
	this->pEnumSettings->AddRef();
	this->autoLabelEdit = autoLabelEdit;

	pEnumResult = new SHTVWBACKGROUNDITEMENUMINFO;
	if(pEnumResult) {
		ZeroMemory(pEnumResult, sizeof(SHTVWBACKGROUNDITEMENUMINFO));
		pEnumResult->taskID = taskID;
		pEnumResult->hParentItem = hParentItem;
		pEnumResult->hInsertAfter = (hInsertAfter ? hInsertAfter : TVI_LAST);
		pEnumResult->checkForDuplicates = checkForDuplicates;
		pEnumResult->namespacePIDLToSet = namespacePIDLToSet;
		pEnumResult->hPIDLBuffer = DPA_Create(1);
	} else {
		return E_OUTOFMEMORY;
	}

	enumSettings = CacheEnumSettings(pEnumSettings);
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShTvwInsertSingleItemTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, IRunnableTask** ppTask)
#else
	HRESULT ShTvwInsertSingleItemTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO>* pEnumratedItemsQueue, LPCRITICAL_SECTION pCriticalSection, HTREEITEM hParentItem, HTREEITEM hInsertAfter, PCIDLIST_ABSOLUTE pIDLToAdd, BOOL isSimplePIDL, PCIDLIST_ABSOLUTE pIDLParent, INamespaceEnumSettings* pEnumSettings, PCIDLIST_ABSOLUTE namespacePIDLToSet, BOOL checkForDuplicates, BOOL autoLabelEdit, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	ATLASSUME(pEnumSettings);
	if(!pEnumratedItemsQueue || !pCriticalSection || !pEnumSettings || !pIDLToAdd || !pIDLParent) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShTvwInsertSingleItemTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShTvwInsertSingleItemTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedItemsQueue, pCriticalSection, hParentItem, hInsertAfter, pIDLToAdd, isSimplePIDL, pIDLParent, pEnumSettings, namespacePIDLToSet, checkForDuplicates, autoLabelEdit);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShTvwInsertSingleItemTask::DoRun(void)
{
	ATLASSERT_POINTER(pEnumResult, SHTVWBACKGROUNDITEMENUMINFO);
	ATLASSUME(pEnumResult->hPIDLBuffer);

	HRESULT hr = NOERROR;
	CComPtr<IShellFolder> pParentISF = NULL;
	PUITEMID_CHILD pIDLToAddRelative = NULL;
	HRESULT hr2 = SHBindToParent(pIDLToAdd, IID_PPV_ARGS(&pParentISF), const_cast<PCUITEMID_CHILD*>(&pIDLToAddRelative));
	ATLASSERT(SUCCEEDED(hr2));
	if(FAILED(hr2)) {
		goto Done;
	}

	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pIDLToAddRelative, ITEMID_CHILD);

	PIDLIST_ABSOLUTE pIDLSubItem = NULL;
	if(isSimplePIDL) {
		// we need a real pIDL
		PUITEMID_CHILD pIDL = NULL;
		if(SUCCEEDED(SHGetRealIDL(pParentISF, pIDLToAddRelative, &pIDL))) {
			ATLASSERT_POINTER(pIDL, ITEMID_CHILD);
			ATLASSERT(pIDL != pIDLToAddRelative);
			pIDLToAddRelative = pIDL;

			PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(pIDLParent);
			ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
			pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pIDLToAddRelative), TRUE));
		} else {
			// well, it already was a real pIDL...
			pIDLSubItem = pIDLToAdd;
			pIDLToAdd = NULL;
			isSimplePIDL = FALSE;
		}
	} else {
		pIDLSubItem = pIDLToAdd;
		pIDLToAdd = NULL;
	}
	ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);
	if(!pIDLSubItem) {
		hr = E_OUTOFMEMORY;
		goto Done;
	}

	if(!ShouldShowItem(pParentISF, pIDLToAddRelative, &enumSettings)) {
		if(isSimplePIDL) {
			ILFree(pIDLToAddRelative);
			pIDLToAddRelative = NULL;
		}
		ILFree(pIDLSubItem);
		goto Done;
	}

	HasSubItems hasSubItems = HasAtLeastOneSubItem(pParentISF, pIDLToAddRelative, pIDLSubItem, &enumSettings, TRUE);
	if(isSimplePIDL) {
		ILFree(pIDLToAddRelative);
		pIDLToAddRelative = NULL;
	}

	LPENUMERATEDITEM pItem = new ENUMERATEDITEM(pIDLSubItem, hasSubItems);
	if(!pItem || (DPA_AppendPtr(pEnumResult->hPIDLBuffer, pItem) == -1)) {
		ILFree(pIDLSubItem);
		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_ITEMENUMCOMPLETE, 0, 0);
		}
		hr = E_OUTOFMEMORY;
		goto Done;
	}

	if(autoLabelEdit) {
		pEnumResult->pIDLToLabelEdit = pIDLSubItem;
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