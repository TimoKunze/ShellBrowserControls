// OverlayTask.cpp: A specialization of RunnableTask for retrieving listview item's overlay icon index

#include "stdafx.h"
#include "OverlayTask.h"


ShLvwOverlayTask::ShLvwOverlayTask(void)
    : RunnableTask(FALSE)
{
}

void ShLvwOverlayTask::FinalRelease()
{
	if(pIDL) {
		ILFree(pIDL);
		pIDL = NULL;
	}
	if(pParentISIO) {
		pParentISIO->Release();
		pParentISIO = NULL;
	}
}


HRESULT ShLvwOverlayTask::Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIconOverlay* pParentISIO)
{
	this->hWndToNotify = hWndToNotify;
	this->pIDL = ILCloneFull(pIDL);
	this->itemID = itemID;
	this->pParentISIO = pParentISIO;
	this->pParentISIO->AddRef();
	return S_OK;
}


HRESULT ShLvwOverlayTask::CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIconOverlay* pParentISIO, IRunnableTask** ppTask)
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	ATLASSUME(pParentISIO);
	if(!pIDL || !pParentISIO) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwOverlayTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwOverlayTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pIDL, itemID, pParentISIO);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwOverlayTask::DoRun(void)
{
	int overlayIndex = 0;
	pParentISIO->GetOverlayIndex(ILFindLastID(pIDL), &overlayIndex);
	if(overlayIndex > 0) {
		PostMessage(hWndToNotify, WM_TRIGGER_UPDATEOVERLAY, itemID, overlayIndex);
	}
	return NOERROR;
}