// BackgroundColumnEnumTask.cpp: A specialization of RunnableTask for enumerating shell columns

#include "stdafx.h"
#include "BackgroundColumnEnumTask.h"


ShLvwBackgroundColumnEnumTask::ShLvwBackgroundColumnEnumTask(void)
    : RunnableTask(TRUE)
{
	pNamespaceObject = NULL;
	pEnumratedColumnsQueue = NULL;
	pEnumResult = NULL;
	pEnumratedColumnsQueue = NULL;
	taskID = GetNewTaskID();
}

void ShLvwBackgroundColumnEnumTask::FinalRelease()
{
	if(pParentISF2) {
		pParentISF2->Release();
		pParentISF2 = NULL;
	}
	if(pParentISD) {
		pParentISD->Release();
		pParentISD = NULL;
	}
	if(pNamespaceObject) {
		pNamespaceObject->Release();
		pNamespaceObject = NULL;
	}
	if(pIDLParent) {
		ILFree(pIDLParent);
		pIDLParent = NULL;
	}
	if(pEnumResult) {
		if(pEnumResult->hColumnBuffer) {
			DPA_DestroyCallback(pEnumResult->hColumnBuffer, FreeDPAColumnElement, NULL);
		}
		delete pEnumResult;
		pEnumResult = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of INamespaceEnumTask
STDMETHODIMP ShLvwBackgroundColumnEnumTask::GetTaskID(PULONGLONG pTaskID)
{
	ATLASSERT_POINTER(pTaskID, ULONGLONG);
	if(!pTaskID) {
		return E_POINTER;
	}
	*pTaskID = taskID;
	return S_OK;
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSERT_POINTER(ppInterfaceImpl, LPVOID);
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(!pNamespaceObject) {
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	return pNamespaceObject->QueryInterface(requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
{
	if(!pNamespaceObject) {
		return E_INVALIDARG;
	}

	if(this->pNamespaceObject) {
		this->pNamespaceObject->Release();
		this->pNamespaceObject = NULL;
	}

	if(pNamespaceObject) {
		return pNamespaceObject->QueryInterface(IID_PPV_ARGS(&this->pNamespaceObject));
	}
	return S_OK;
}
// implementation of INamespaceEnumTask
//////////////////////////////////////////////////////////////////////


#ifdef USE_STL
	HRESULT ShLvwBackgroundColumnEnumTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth)
#else
	HRESULT ShLvwBackgroundColumnEnumTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pEnumratedColumnsQueue = pEnumratedColumnsQueue;
	this->pCriticalSection = pCriticalSection;
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	currentRealColumnIndex = realColumnIndex;
	this->averageCharWidth = averageCharWidth;
	this->pIDLParent = ILCloneFull(pIDLParent);

	pParentISF2 = NULL;
	pParentISD = NULL;
	pEnumResult = new SHLVWBACKGROUNDCOLUMNENUMINFO;
	if(pEnumResult) {
		ZeroMemory(pEnumResult, sizeof(SHLVWBACKGROUNDCOLUMNENUMINFO));
		pEnumResult->taskID = taskID;
		pEnumResult->baseRealColumnIndex = realColumnIndex;
		pEnumResult->hColumnBuffer = DPA_Create(10);
		#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
			pEnumResult->firstNewColumn = 0;
			pEnumResult->lastNewColumn = -1;
		#endif
	} else {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShLvwBackgroundColumnEnumTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth, IRunnableTask** ppTask)
#else
	HRESULT ShLvwBackgroundColumnEnumTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO>* pEnumratedColumnsQueue, LPCRITICAL_SECTION pCriticalSection, HWND hWndShellUIParentWindow, PCIDLIST_ABSOLUTE pIDLParent, int realColumnIndex, int averageCharWidth, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pEnumratedColumnsQueue || !pCriticalSection || !pIDLParent) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwBackgroundColumnEnumTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwBackgroundColumnEnumTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pEnumratedColumnsQueue, pCriticalSection, hWndShellUIParentWindow, pIDLParent, realColumnIndex, averageCharWidth);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwBackgroundColumnEnumTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO);
	ATLASSUME(pEnumResult->hColumnBuffer);

	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		const int CHUNKSIZE = 10;
	#endif

	HRESULT hr = E_FAIL;
	SHELLDETAILS column;
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		pEnumResult->moreColumnsToCome = TRUE;
	#endif
	do {
		ZeroMemory(&column, sizeof(column));
		SHCOLSTATEF columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
		if(pParentISF2) {
			hr = pParentISF2->GetDetailsOf(NULL, currentRealColumnIndex, &column);
			if(FAILED(pParentISF2->GetDefaultColumnState(currentRealColumnIndex, &columnState))) {
				columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
			}
		} else {
			hr = E_FAIL;
		}
		if(FAILED(hr) && pParentISD) {
			hr = pParentISD->GetDetailsOf(NULL, currentRealColumnIndex, &column);
		}

		if(SUCCEEDED(hr)) {
			// store the column
			LPSHELLLISTVIEWCOLUMNDATA pColumnData = new SHELLLISTVIEWCOLUMNDATA;
			ATLASSERT_POINTER(pColumnData, SHELLLISTVIEWCOLUMNDATA);
			pColumnData->columnID = -1;
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				pColumnData->pPropertyDescription = NULL;
				pColumnData->displayType = PDDT_STRING;
				pColumnData->typeFlags = PDTF_DEFAULT;
			#endif
			pColumnData->visibleInDetailsView = FALSE;
			pColumnData->visibilityHasBeenChangedExplicitly = FALSE;
			pColumnData->lastIndex_DetailsView = -1;
			pColumnData->lastIndex_TilesView = -1;
			ATLVERIFY(SUCCEEDED(StrRetToBuf(&column.str, NULL, pColumnData->pText, MAX_LVCOLUMNTEXTLENGTH)));
			if(column.cxChar <= 0) {
				pColumnData->width = 96;
			} else {
				pColumnData->width = column.cxChar * averageCharWidth;
			}
			pColumnData->format = column.fmt;
			if(pParentISF2 && APIWrapper::IsSupported_PSGetPropertyDescription()) {
				PROPERTYKEY propertyKey;
				if(SUCCEEDED(pParentISF2->MapColumnToSCID(currentRealColumnIndex, &propertyKey))) {
					CComPtr<IPropertyDescription> pPropertyDescription = NULL;
					APIWrapper::PSGetPropertyDescription(propertyKey, IID_PPV_ARGS(&pPropertyDescription), NULL);
					if(pPropertyDescription) {
						UINT width = 0;
						if(SUCCEEDED(pPropertyDescription->GetDefaultColumnWidth(&width))) {
							pColumnData->width = static_cast<int>(width) * averageCharWidth;
						}

						PROPDESC_VIEW_FLAGS viewFlags;
						if(SUCCEEDED(pPropertyDescription->GetViewFlags(&viewFlags))) {
							// TODO: Is this correct?
							if(viewFlags & PDVF_BEGINNEWGROUP) {
								pColumnData->format |= LVCFMT_LINE_BREAK;
							}
							if(viewFlags & PDVF_FILLAREA) {
								pColumnData->format |= LVCFMT_FILL;
							}
							// TODO: Is this correct?
							if(viewFlags & PDVF_HIDELABEL) {
								pColumnData->format |= LVCFMT_NO_TITLE;
							}
							if(viewFlags & PDVF_CANWRAP) {
								pColumnData->format |= LVCFMT_WRAP;
							}
						}

						#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
							if(RunTimeHelper::IsVista()) {
								pPropertyDescription->GetDisplayType(&pColumnData->displayType);
								pPropertyDescription->GetTypeFlags(PDTF_MASK_ALL, &pColumnData->typeFlags);
							}
						#endif
					}
				}
			}
			pColumnData->state = columnState;

			if(DPA_AppendPtr(pEnumResult->hColumnBuffer, pColumnData) == -1) {
				delete pColumnData;
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				} else {
					enumeratedColumns++;
					pEnumResult->lastNewColumn++;
			#endif
			}

			currentRealColumnIndex++;
		}

		#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
			if(enumeratedColumns % CHUNKSIZE == 0) {
				// causes problems on shutdown
				EnterCriticalSection(pCriticalSection);
				#ifdef USE_STL
					pEnumratedColumnsQueue->push(pEnumResult);
				#else
					pEnumratedColumnsQueue->AddTail(pEnumResult);
				#endif
				LPSHLVWBACKGROUNDCOLUMNENUMINFO pNewEnumResult = new SHLVWBACKGROUNDCOLUMNENUMINFO;
				if(pNewEnumResult) {
					*pNewEnumResult = *pEnumResult;
					pNewEnumResult->hColumnBuffer = DPA_Create(CHUNKSIZE);
					pNewEnumResult->firstNewColumn += CHUNKSIZE;
					pNewEnumResult->baseRealColumnIndex += CHUNKSIZE;
				}
				pEnumResult = pNewEnumResult;
				LeaveCriticalSection(pCriticalSection);
				ATLASSERT_POINTER(pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO);
				ATLASSUME(pEnumResult->hColumnBuffer);

				if(IsWindow(hWndToNotify)) {
					PostMessage(hWndToNotify, WM_TRIGGER_COLUMNENUMCOMPLETE, 0, 0);
				}
			}
		#endif
		if(hr == S_OK && WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	} while(hr == S_OK);

	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		pEnumResult->moreColumnsToCome = FALSE;
	#endif

	EnterCriticalSection(pCriticalSection);
	#ifdef USE_STL
		pEnumratedColumnsQueue->push(pEnumResult);
	#else
		pEnumratedColumnsQueue->AddTail(pEnumResult);
	#endif
	pEnumResult = NULL;
	LeaveCriticalSection(pCriticalSection);

	if(IsWindow(hWndToNotify)) {
		PostMessage(hWndToNotify, WM_TRIGGER_COLUMNENUMCOMPLETE, 0, 0);
	}
	return NOERROR;
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::DoRun(void)
{
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		enumeratedColumns = 0;
	#endif

	CComPtr<IShellFolder> pParentISF;
	HRESULT hr = BindToPIDL(pIDLParent, IID_PPV_ARGS(&pParentISF));
	if(FAILED(hr)) {
		return E_FAIL;
	}
	ATLASSUME(pParentISF);
	pParentISF->QueryInterface(IID_PPV_ARGS(&pParentISF2));
	pParentISF->CreateViewObject(hWndShellUIParentWindow, IID_PPV_ARGS(&pParentISD));
	if(!pParentISF2 && !pParentISD) {
		return E_FAIL;
	}
	return E_PENDING;
}