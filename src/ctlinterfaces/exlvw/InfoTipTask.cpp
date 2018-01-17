// InfoTipTask.cpp: A specialization of RunnableTask for retrieving listview item info tip

#include "stdafx.h"
#include "InfoTipTask.h"


ShLvwInfoTipTask::ShLvwInfoTipTask(void)
    : RunnableTask(FALSE)
{
	pIDL = NULL;
	pInfoTipQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwInfoTipTask::FinalRelease()
{
	if(pIDL) {
		ILFree(pIDL);
		pIDL = NULL;
	}
	if(pResult) {
		if(pResult->pInfoTipText) {
			CoTaskMemFree(pResult->pInfoTipText);
		}
		delete pResult;
		pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwInfoTipTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, std::queue<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize)
#else
	HRESULT ShLvwInfoTipTask::Attach(HWND hWndToNotify, HWND hWndShellUIParentWindow, CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO>* pInfoTipQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, InfoTipFlagsConstants infoTipFlags, LPWSTR pTextToPrepend, UINT textToPrependSize)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->pInfoTipQueue = pInfoTipQueue;
	this->pCriticalSection = pCriticalSection;
	this->pIDL = ILCloneFull(pIDL);
	this->infoTipFlags = infoTipFlags;

	pResult = new SHLVWBACKGROUNDINFOTIPINFO;
	if(pResult) {
		ZeroMemory(pResult, sizeof(SHLVWBACKGROUNDINFOTIPINFO));
		pResult->itemID = itemID;
		if(textToPrependSize > 0) {
			pResult->pInfoTipText = static_cast<LPWSTR>(CoTaskMemAlloc((textToPrependSize + 1) * sizeof(WCHAR)));
			if(pResult->pInfoTipText) {
				pResult->infoTipTextSize = textToPrependSize;
				lstrcpynW(pResult->pInfoTipText, pTextToPrepend, textToPrependSize + 1);
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
	ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDINFOTIPINFO);

	LPWSTR pInfoTipText = NULL;
	HRESULT hr = GetInfoTip(hWndShellUIParentWindow, pIDL, infoTipFlags, &pInfoTipText);
	if(SUCCEEDED(hr)) {
		UINT infoTipTextSize = lstrlenW(pInfoTipText);
		if(pResult->pInfoTipText) {
			// append pInfoTipText to pResult->pInfoTipText
			if(infoTipTextSize > 0) {
				// append a line break and pInfoTipText
				LPWSTR pCombinedInfoTip = static_cast<LPWSTR>(CoTaskMemAlloc((pResult->infoTipTextSize + 2 + infoTipTextSize + 1) * sizeof(WCHAR)));
				if(pCombinedInfoTip) {
					lstrcpynW(pCombinedInfoTip, pResult->pInfoTipText, pResult->infoTipTextSize + 1);
					lstrcpynW(pCombinedInfoTip + pResult->infoTipTextSize, L"\r\n", 3);
					lstrcpynW(pCombinedInfoTip + pResult->infoTipTextSize + 2, pInfoTipText, infoTipTextSize + 1);
					pInfoTipText = NULL;
					CoTaskMemFree(pResult->pInfoTipText);
					pResult->pInfoTipText = pCombinedInfoTip;
					pResult->infoTipTextSize += 2 + infoTipTextSize;
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
			pResult->pInfoTipText = pInfoTipText;
			pResult->infoTipTextSize = infoTipTextSize;
		}

		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pInfoTipQueue->push(pResult);
		#else
			pInfoTipQueue->AddTail(pResult);
		#endif
		pResult = NULL;
		LeaveCriticalSection(pCriticalSection);

		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_SETINFOTIP, 0, 0);
		}
	} else if(pInfoTipText) {
		CoTaskMemFree(pInfoTipText);
	}
	return NOERROR;
}