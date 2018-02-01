// OverlayTask.cpp: A specialization of RunnableTask for retrieving listview item's overlay icon index

#include "stdafx.h"
#include "OverlayTask.h"


ShLvwOverlayTask::ShLvwOverlayTask(void)
    : RunnableTask(FALSE)
{
}

void ShLvwOverlayTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
	if(properties.pParentISIO) {
		properties.pParentISIO->Release();
		properties.pParentISIO = NULL;
	}
}


HRESULT ShLvwOverlayTask::Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIconOverlay* pParentISIO)
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.itemID = itemID;
	this->properties.pParentISIO = pParentISIO;
	this->properties.pParentISIO->AddRef();
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
	properties.pParentISIO->GetOverlayIndex(ILFindLastID(properties.pIDL), &overlayIndex);
	if(overlayIndex > 0) {
		PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATEOVERLAY, properties.itemID, overlayIndex);
	}
	return NOERROR;
}