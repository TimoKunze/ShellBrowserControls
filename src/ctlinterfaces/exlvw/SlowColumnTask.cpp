// SlowColumnTask.cpp: A specialization of RunnableTask for retrieving listview sub-item's text

#include "stdafx.h"
#include "SlowColumnTask.h"


ShLvwSlowColumnTask::ShLvwSlowColumnTask(void)
    : RunnableTask(TRUE)
{
	properties.pIDL = NULL;
	properties.pIDLNamespace = NULL;
	properties.pParentISF = NULL;
	properties.pParentISF2 = NULL;
	properties.pSlowColumnQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwSlowColumnTask::FinalRelease()
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
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			if(properties.pResult->pPropertyValue) {
				PropVariantClear(properties.pResult->pPropertyValue);
				delete properties.pResult->pPropertyValue;
			}
		#endif
		delete properties.pResult;
		properties.pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwSlowColumnTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#else
	HRESULT ShLvwSlowColumnTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO>* pSlowColumnQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, LONG columnID, int realColumnIndex, PCIDLIST_ABSOLUTE pIDLNamespace)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pSlowColumnQueue = pSlowColumnQueue;
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
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				if(hasValue) {
					properties.pResult->pPropertyValue = pPv;
				}
			#endif
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
				#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
					if(hasValue) {
						properties.pResult->pPropertyValue = pPv;
					}
				#endif
			}
			VariantClear(&v);
		}
	}

	if(hasValue) {
		APIWrapper::PSFormatForDisplay(propertyKey, *pPv, PDFF_DEFAULT, properties.pResult->pText, MAX_LVITEMTEXTLENGTH, NULL);
		hr = S_OK;
	} else if(properties.pParentISF2) {
		// still no success, so try IShellFolder2::GetDetailsOf
		SHELLDETAILS column = {0};
		hr = properties.pParentISF2->GetDetailsOf(ILFindLastID(properties.pIDL), properties.realColumnIndex, &column);
		if(SUCCEEDED(hr)) {
			ATLVERIFY(SUCCEEDED(StrRetToBufW(&column.str, ILFindLastID(properties.pIDL), properties.pResult->pText, MAX_LVITEMTEXTLENGTH)));
		}
	} else {
		hr = E_FAIL;
	}
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		if(properties.pResult->pPropertyValue != pPv && pPv) {
	#else
		if(pPv) {
	#endif
		delete pPv;
		pPv = NULL;
	}

	if(FAILED(hr) && properties.pParentISF) {
		// still no success, so try IShellDetails::GetDetailsOf
		CComPtr<IShellDetails> pParentISD;
		hr = properties.pParentISF->CreateViewObject(properties.hWndShellUIParentWindow, IID_PPV_ARGS(&pParentISD));
		if(pParentISD) {
			SHELLDETAILS column = {0};
			hr = pParentISD->GetDetailsOf(ILFindLastID(properties.pIDL), properties.realColumnIndex, &column);
			if(SUCCEEDED(hr)) {
				ATLVERIFY(SUCCEEDED(StrRetToBufW(&column.str, ILFindLastID(properties.pIDL), properties.pResult->pText, MAX_LVITEMTEXTLENGTH)));
			}
		} else {
			// give up
			return NOERROR;
		}
	}

	if(SUCCEEDED(hr)) {
		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pSlowColumnQueue->push(properties.pResult);
		#else
			properties.pSlowColumnQueue->AddTail(properties.pResult);
		#endif
		properties.pResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);

		if(IsWindow(properties.hWndToNotify)) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATETEXT, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwSlowColumnTask::DoRun(void)
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