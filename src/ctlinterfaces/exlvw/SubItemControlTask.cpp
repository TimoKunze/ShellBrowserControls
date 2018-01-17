// SubItemControlTask.cpp: A specialization of RunnableTask for retrieving a sub-item control's current value

#include "stdafx.h"
#include "SubItemControlTask.h"


ShLvwSubItemControlTask::ShLvwSubItemControlTask(void)
    : RunnableTask(TRUE)
{
	pIDL = NULL;
	pIDLNamespace = NULL;
	pParentISF = NULL;
	pParentISF2 = NULL;
	pSubItemControlQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwSubItemControlTask::FinalRelease()
{
	if(pParentISF) {
		pParentISF->Release();
		pParentISF = NULL;
	}
	if(pParentISF2) {
		pParentISF2->Release();
		pParentISF2 = NULL;
	}
	if(pIDL) {
		ILFree(pIDL);
		pIDL = NULL;
	}
	if(pIDLNamespace) {
		ILFree(pIDLNamespace);
		pIDLNamespace = NULL;
	}
	if(pResult) {
		if(pResult->pPropertyValue) {
			PropVariantClear(pResult->pPropertyValue);
			delete pResult->pPropertyValue;
		}
		delete pResult;
		pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwSubItemControlTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#else
	HRESULT ShLvwSubItemControlTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pSubItemControlQueue = pSubItemControlQueue;
	this->pCriticalSection = pCriticalSection;
	this->pIDL = ILCloneFull(pIDL);
	this->realColumnIndex = realColumnIndex;
	this->pIDLNamespace = ILCloneFull(pIDLNamespace);

	pResult = new SHLVWBACKGROUNDCOLUMNINFO;
	if(pResult) {
		ZeroMemory(pResult, sizeof(SHLVWBACKGROUNDCOLUMNINFO));
		pResult->itemID = itemID;
		pResult->columnID = columnID;
	} else {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}

#ifdef USE_STL
	HRESULT ShLvwSubItemControlTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, IRunnableTask** ppTask)
#else
	HRESULT ShLvwSubItemControlTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pSubItemControlQueue || !pCriticalSection || !pIDL || !pIDLNamespace) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwSubItemControlTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwSubItemControlTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, hWndShellUIParentWindow, pSubItemControlQueue, pCriticalSection, pIDL, itemID, columnID, realColumnIndex, pIDLNamespace);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwSubItemControlTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDCOLUMNINFO);

	HRESULT hr = E_FAIL;

	// at first try IShellItem2::GetPropertyStore -> IPropertyStore::GetValue
	/* NOTE: We could use IShellFolder::BindToObject to get the property store directly, but then we cannot
	         specify any GPS_* flags. */
	BOOL hasPropertyKey = FALSE;
	BOOL hasValue = FALSE;
	PROPERTYKEY propertyKey;
	LPPROPVARIANT pPv = new PROPVARIANT;
	PropVariantInit(pPv);
	CComPtr<IShellItem2> pShellItem;
	if(pParentISF2) {
		hasPropertyKey = SUCCEEDED(pParentISF2->MapColumnToSCID(realColumnIndex, &propertyKey));
	}
	if(hasPropertyKey && APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		ATLVERIFY(SUCCEEDED(APIWrapper::SHCreateItemFromIDList(pIDL, IID_PPV_ARGS(&pShellItem), &hr)));
	}
	if(pShellItem) {
		CComPtr<IPropertyStore> pPropertyStore;
		hr = pShellItem->GetPropertyStore(GPS_DEFAULT, IID_PPV_ARGS(&pPropertyStore));
		if(SUCCEEDED(hr) && pPropertyStore) {
			hr = pPropertyStore->GetValue(propertyKey, pPv);
			hasValue = SUCCEEDED(hr);
			if(hasValue) {
				pResult->pPropertyValue = pPv;
			}
		}
	}

	if(!hasValue) {
		// we did not succeed, so try IShellFolder2::GetDetailsEx
		if(hasPropertyKey) {
			VARIANT v;
			VariantInit(&v);
			hr = pParentISF2->GetDetailsEx(ILFindLastID(pIDL), &propertyKey, &v);
			if(SUCCEEDED(hr)) {
				hasValue = SUCCEEDED(APIWrapper::VariantToPropVariant(&v, pPv, &hr));
				if(hasValue) {
					pResult->pPropertyValue = pPv;
				}
			}
			VariantClear(&v);
		}
	}

	if(hasValue) {
		hr = S_OK;
	} else {
		hr = E_FAIL;
	}
	if(pResult->pPropertyValue != pPv && pPv) {
		delete pPv;
		pPv = NULL;
	}

	if(SUCCEEDED(hr)) {
		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pSubItemControlQueue->push(pResult);
		#else
			pSubItemControlQueue->AddTail(pResult);
		#endif
		pResult = NULL;
		LeaveCriticalSection(pCriticalSection);

		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_UPDATESUBITEMCONTROL, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwSubItemControlTask::DoRun(void)
{
	if(pIDLNamespace) {
		BindToPIDL(pIDLNamespace, IID_PPV_ARGS(&pParentISF));
	}
	if(!pParentISF) {
		return E_FAIL;
	}
	pParentISF->QueryInterface(IID_PPV_ARGS(&pParentISF2));

	return E_PENDING;
}