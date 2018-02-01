// ElevationShieldTask.cpp: A specialization of RunnableTask for retrieving shell item elevation requirements

#include "stdafx.h"
#include "ElevationShieldTask.h"


ElevationShieldTask::ElevationShieldTask(void)
    : RunnableTask(FALSE)
{
}

void ElevationShieldTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
}


HRESULT ElevationShieldTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID)
{
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.itemID = itemID;
	this->properties.pIDL = ILCloneFull(pIDL);
	return S_OK;
}

HRESULT ElevationShieldTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IRunnableTask** ppTask)
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ElevationShieldTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ElevationShieldTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndShellUIParentWindow, hWndToNotify, pIDL, itemID);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ElevationShieldTask::DoRun(void)
{
	if(IsElevationRequired(properties.hWndShellUIParentWindow, properties.pIDL)) {
		// NOTE: As a small optimization we don't send a message if no elevation is required.
		PostMessage(properties.hWndToNotify, WM_TRIGGER_SETELEVATIONSHIELD, properties.itemID, TRUE);
	}
	return NOERROR;
}