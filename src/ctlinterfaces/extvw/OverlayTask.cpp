// OverlayTask.cpp: A specialization of RunnableTask for retrieving treeview item's overlay icon index

#include "stdafx.h"
#include "OverlayTask.h"


ShTvwOverlayTask::ShTvwOverlayTask(void)
    : RunnableTask(FALSE)
{
}

void ShTvwOverlayTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
}


HRESULT ShTvwOverlayTask::Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle)
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.itemHandle = itemHandle;
	return S_OK;
}


HRESULT ShTvwOverlayTask::CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, IRunnableTask** ppTask)
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShTvwOverlayTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShTvwOverlayTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pIDL, itemHandle);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShTvwOverlayTask::DoRun(void)
{
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = SHBindToParent(properties.pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	if(SUCCEEDED(hr)) {
		int overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
		if(overlayIndex > 0) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATEOVERLAY, reinterpret_cast<WPARAM>(properties.itemHandle), overlayIndex);
		}
		hr = NOERROR;
	}
	return hr;
}