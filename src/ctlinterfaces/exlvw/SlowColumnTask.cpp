// SlowColumnTask.cpp: A specialization of RunnableTask for retrieving listview sub-item's text

#include "stdafx.h"
#include "SlowColumnTask.h"


ShLvwSlowColumnTask::ShLvwSlowColumnTask(void)
    : RunnableTask(TRUE)
{
	pIDL = NULL;
	pIDLNamespace = NULL;
	pParentISF = NULL;
	pParentISF2 = NULL;
	pSlowColumnQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwSlowColumnTask::FinalRelease()
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
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			if(pResult->pPropertyValue) {
				PropVariantClear(pResult->pPropertyValue);
				delete pResult->pPropertyValue;
			}
		#endif
		delete pResult;
		pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwSlowColumnTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#else
	HRESULT ShLvwSlowColumnTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pSlowColumnQueue = pSlowColumnQueue;
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
	HRESULT ShLvwSlowColumnTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, IRunnableTask** ppTask)
#else
	HRESULT ShLvwSlowColumnTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pSlowColumnQueue || !pCriticalSection || !pIDL || !pIDLNamespace) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwSlowColumnTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwSlowColumnTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, hWndShellUIParentWindow, pSlowColumnQueue, pCriticalSection, pIDL, itemID, columnID, realColumnIndex, pIDLNamespace);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwSlowColumnTask::DoInternalResume(void)
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
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				if(hasValue) {
					pResult->pPropertyValue = pPv;
				}
			#endif
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
				#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
					if(hasValue) {
						pResult->pPropertyValue = pPv;
					}
				#endif
			}
			VariantClear(&v);
		}
	}

	if(hasValue) {
		APIWrapper::PSFormatForDisplay(propertyKey, *pPv, PDFF_DEFAULT, pResult->pText, MAX_LVITEMTEXTLENGTH, NULL);
		hr = S_OK;
	} else if(pParentISF2) {
		// still no success, so try IShellFolder2::GetDetailsOf
		SHELLDETAILS column = {0};
		hr = pParentISF2->GetDetailsOf(ILFindLastID(pIDL), realColumnIndex, &column);
		if(SUCCEEDED(hr)) {
			ATLVERIFY(SUCCEEDED(StrRetToBufW(&column.str, ILFindLastID(pIDL), pResult->pText, MAX_LVITEMTEXTLENGTH)));
		}
	} else {
		hr = E_FAIL;
	}
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		if(pResult->pPropertyValue != pPv && pPv) {
	#else
		if(pPv) {
	#endif
		delete pPv;
		pPv = NULL;
	}

	if(FAILED(hr) && pParentISF) {
		// still no success, so try IShellDetails::GetDetailsOf
		CComPtr<IShellDetails> pParentISD;
		hr = pParentISF->CreateViewObject(hWndShellUIParentWindow, IID_PPV_ARGS(&pParentISD));
		if(pParentISD) {
			SHELLDETAILS column = {0};
			hr = pParentISD->GetDetailsOf(ILFindLastID(pIDL), realColumnIndex, &column);
			if(SUCCEEDED(hr)) {
				ATLVERIFY(SUCCEEDED(StrRetToBufW(&column.str, ILFindLastID(pIDL), pResult->pText, MAX_LVITEMTEXTLENGTH)));
			}
		} else {
			// give up
			return NOERROR;
		}
	}

	if(SUCCEEDED(hr)) {
		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pSlowColumnQueue->push(pResult);
		#else
			pSlowColumnQueue->AddTail(pResult);
		#endif
		pResult = NULL;
		LeaveCriticalSection(pCriticalSection);

		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_UPDATETEXT, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwSlowColumnTask::DoRun(void)
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