// SubItemControlTask.cpp: A specialization of RunnableTask for retrieving a sub-item control's current value

#include "stdafx.h"
#include "SubItemControlTask.h"


ShLvwSubItemControlTask::ShLvwSubItemControlTask(void)
    : RunnableTask(TRUE)
{
	properties.pIDL = NULL;
	properties.pIDLNamespace = NULL;
	properties.pParentISF = NULL;
	properties.pParentISF2 = NULL;
	properties.pSubItemControlQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwSubItemControlTask::FinalRelease()
{
	if(properties.pParentISF) {
		properties.pParentISF->Release();
		properties.pParentISF = NULL;
	}
	if(properties.pParentISF2) {
		properties.pParentISF2->Release();
		properties.pParentISF2 = NULL;
	}
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
	if(properties.pIDLNamespace) {
		ILFree(properties.pIDLNamespace);
		properties.pIDLNamespace = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->pPropertyValue) {
			PropVariantClear(properties.pResult->pPropertyValue);
			delete properties.pResult->pPropertyValue;
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwSubItemControlTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#else
	HRESULT ShLvwSubItemControlTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSubItemControlQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pSubItemControlQueue = pSubItemControlQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.realColumnIndex = realColumnIndex;
	this->properties.pIDLNamespace = ILCloneFull(pIDLNamespace);

	properties.pResult = new SHLVWBACKGROUNDCOLUMNINFO;
	if(properties.pResult) {
		ZeroMemory(properties.pResult, sizeof(SHLVWBACKGROUNDCOLUMNINFO));
		properties.pResult->itemID = itemID;
		properties.pResult->columnID = columnID;
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
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDCOLUMNINFO);

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
	if(properties.pParentISF2) {
		hasPropertyKey = SUCCEEDED(properties.pParentISF2->MapColumnToSCID(properties.realColumnIndex, &propertyKey));
	}
	if(hasPropertyKey && APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		ATLVERIFY(SUCCEEDED(APIWrapper::SHCreateItemFromIDList(properties.pIDL, IID_PPV_ARGS(&pShellItem), &hr)));
	}
	if(pShellItem) {
		CComPtr<IPropertyStore> pPropertyStore;
		hr = pShellItem->GetPropertyStore(GPS_DEFAULT, IID_PPV_ARGS(&pPropertyStore));
		if(SUCCEEDED(hr) && pPropertyStore) {
			hr = pPropertyStore->GetValue(propertyKey, pPv);
			hasValue = SUCCEEDED(hr);
			if(hasValue) {
				properties.pResult->pPropertyValue = pPv;
			}
		}
	}

	if(!hasValue) {
		// we did not succeed, so try IShellFolder2::GetDetailsEx
		if(hasPropertyKey) {
			VARIANT v;
			VariantInit(&v);
			hr = properties.pParentISF2->GetDetailsEx(ILFindLastID(properties.pIDL), &propertyKey, &v);
			if(SUCCEEDED(hr)) {
				hasValue = SUCCEEDED(APIWrapper::VariantToPropVariant(&v, pPv, &hr));
				if(hasValue) {
					properties.pResult->pPropertyValue = pPv;
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
	if(properties.pResult->pPropertyValue != pPv && pPv) {
		delete pPv;
		pPv = NULL;
	}

	if(SUCCEEDED(hr)) {
		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pSubItemControlQueue->push(properties.pResult);
		#else
			properties.pSubItemControlQueue->AddTail(properties.pResult);
		#endif
		properties.pResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);

		if(IsWindow(properties.hWndToNotify)) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATESUBITEMCONTROL, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwSubItemControlTask::DoRun(void)
{
	if(properties.pIDLNamespace) {
		BindToPIDL(properties.pIDLNamespace, IID_PPV_ARGS(&properties.pParentISF));
	}
	if(!properties.pParentISF) {
		return E_FAIL;
	}
	properties.pParentISF->QueryInterface(IID_PPV_ARGS(&properties.pParentISF2));

	return E_PENDING;
}