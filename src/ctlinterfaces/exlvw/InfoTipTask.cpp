// InfoTipTask.cpp: A specialization of RunnableTask for retrieving listview item info tip

#include "stdafx.h"
#include "InfoTipTask.h"


ShLvwInfoTipTask::ShLvwInfoTipTask(void)
    : RunnableTask(FALSE)
{
	properties.pIDL = NULL;
	properties.pInfoTipQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwInfoTipTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->pInfoTipText) {
			CoTaskMemFree(properties.pResult->pInfoTipText);
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwInfoTipTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize)
#else
	HRESULT ShLvwInfoTipTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.pInfoTipQueue = pInfoTipQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.infoTipFlags = infoTipFlags;

	properties.pResult = new SHLVWBACKGROUNDINFOTIPINFO;
	if(properties.pResult) {
		ZeroMemory(properties.pResult, sizeof(SHLVWBACKGROUNDINFOTIPINFO));
		properties.pResult->itemID = itemID;
		if(textToPrependSize > 0) {
			properties.pResult->pInfoTipText = static_cast<LPWSTR>(CoTaskMemAlloc((textToPrependSize + 1) * sizeof(WCHAR)));
			if(properties.pResult->pInfoTipText) {
				properties.pResult->infoTipTextSize = textToPrependSize;
				lstrcpynW(properties.pResult->pInfoTipText, pTextToPrepend, textToPrependSize + 1);
			}
		}
	} else {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}

#ifdef USE_STL
	HRESULT ShLvwInfoTipTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize, IRunnableTask** ppTask)
#else
	HRESULT ShLvwInfoTipTask::CreateInstance(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pInfoTipQueue || !pCriticalSection || !pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwInfoTipTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwInfoTipTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, hWndShellUIParentWindow, pInfoTipQueue, pCriticalSection, pIDL, itemID, infoTipFlags, pTextToPrepend, textToPrependSize);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwInfoTipTask::DoRun(void)
{
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDINFOTIPINFO);

	LPWSTR pInfoTipText = NULL;
	HRESULT hr = GetInfoTip(properties.hWndShellUIParentWindow, properties.pIDL, properties.infoTipFlags, &pInfoTipText);
	if(SUCCEEDED(hr)) {
		UINT infoTipTextSize = lstrlenW(pInfoTipText);
		if(properties.pResult->pInfoTipText) {
			// append pInfoTipText to pResult->pInfoTipText
			if(infoTipTextSize > 0) {
				// append a line break and pInfoTipText
				LPWSTR pCombinedInfoTip = static_cast<LPWSTR>(CoTaskMemAlloc((properties.pResult->infoTipTextSize + 2 + infoTipTextSize + 1) * sizeof(WCHAR)));
				if(pCombinedInfoTip) {
					lstrcpynW(pCombinedInfoTip, properties.pResult->pInfoTipText, properties.pResult->infoTipTextSize + 1);
					lstrcpynW(pCombinedInfoTip + properties.pResult->infoTipTextSize, L"\r\n", 3);
					lstrcpynW(pCombinedInfoTip + properties.pResult->infoTipTextSize + 2, pInfoTipText, infoTipTextSize + 1);
					pInfoTipText = NULL;
					CoTaskMemFree(properties.pResult->pInfoTipText);
					properties.pResult->pInfoTipText = pCombinedInfoTip;
					properties.pResult->infoTipTextSize += 2 + infoTipTextSize;
				}
			} else {
				// actually there's nothing to append
			}
			if(pInfoTipText) {
				CoTaskMemFree(pInfoTipText);
				pInfoTipText = NULL;
			}
		} else {
			// nothing to prepend
			properties.pResult->pInfoTipText = pInfoTipText;
			properties.pResult->infoTipTextSize = infoTipTextSize;
		}

		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pInfoTipQueue->push(properties.pResult);
		#else
			properties.pInfoTipQueue->AddTail(properties.pResult);
		#endif
		properties.pResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);

		if(IsWindow(properties.hWndToNotify)) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_SETINFOTIP, 0, 0);
		}
	} else if(pInfoTipText) {
		CoTaskMemFree(pInfoTipText);
	}
	return NOERROR;
}