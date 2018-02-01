// BackgroundColumnEnumTask.cpp: A specialization of RunnableTask for enumerating shell columns

#include "stdafx.h"
#include "BackgroundColumnEnumTask.h"


ShLvwBackgroundColumnEnumTask::ShLvwBackgroundColumnEnumTask(void)
    : RunnableTask(TRUE)
{
	properties.pNamespaceObject = NULL;
	properties.pEnumratedColumnsQueue = NULL;
	properties.pEnumResult = NULL;
	properties.pEnumratedColumnsQueue = NULL;
	properties.taskID = GetNewTaskID();
}

void ShLvwBackgroundColumnEnumTask::FinalRelease()
{
	if(properties.pParentISF2) {
		properties.pParentISF2->Release();
		properties.pParentISF2 = NULL;
	}
	if(properties.pParentISD) {
		properties.pParentISD->Release();
		properties.pParentISD = NULL;
	}
	if(properties.pNamespaceObject) {
		properties.pNamespaceObject->Release();
		properties.pNamespaceObject = NULL;
	}
	if(properties.pIDLParent) {
		ILFree(properties.pIDLParent);
		properties.pIDLParent = NULL;
	}
	if(properties.pEnumResult) {
		if(properties.pEnumResult->hColumnBuffer) {
			DPA_DestroyCallback(properties.pEnumResult->hColumnBuffer, FreeDPAColumnElement, NULL);
		}
		delete properties.pEnumResult;
		properties.pEnumResult = NULL;
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
	*pTaskID = properties.taskID;
	return S_OK;
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::GetTarget(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	ATLASSERT_POINTER(ppInterfaceImpl, LPVOID);
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(!properties.pNamespaceObject) {
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	return properties.pNamespaceObject->QueryInterface(requiredInterface, ppInterfaceImpl);
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::SetTarget(LPDISPATCH pNamespaceObject)
{
	if(!pNamespaceObject) {
		return E_INVALIDARG;
	}

	if(this->properties.pNamespaceObject) {
		this->properties.pNamespaceObject->Release();
		this->properties.pNamespaceObject = NULL;
	}

	if(pNamespaceObject) {
		return pNamespaceObject->QueryInterface(IID_PPV_ARGS(&this->properties.pNamespaceObject));
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
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pEnumratedColumnsQueue = pEnumratedColumnsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	properties.currentRealColumnIndex = realColumnIndex;
	this->properties.averageCharWidth = averageCharWidth;
	this->properties.pIDLParent = ILCloneFull(pIDLParent);

	properties.pParentISF2 = NULL;
	properties.pParentISD = NULL;
	properties.pEnumResult = new SHLVWBACKGROUNDCOLUMNENUMINFO;
	if(properties.pEnumResult) {
		ZeroMemory(properties.pEnumResult, sizeof(SHLVWBACKGROUNDCOLUMNENUMINFO));
		properties.pEnumResult->taskID = properties.taskID;
		properties.pEnumResult->baseRealColumnIndex = realColumnIndex;
		properties.pEnumResult->hColumnBuffer = DPA_Create(10);
		#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
			properties.pEnumResult->firstNewColumn = 0;
			properties.pEnumResult->lastNewColumn = -1;
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
	ATLASSERT_POINTER(properties.pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO);
	ATLASSUME(properties.pEnumResult->hColumnBuffer);

	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		const int CHUNKSIZE = 10;
	#endif

	HRESULT hr = E_FAIL;
	SHELLDETAILS column;
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		properties.pEnumResult->moreColumnsToCome = TRUE;
	#endif
	do {
		ZeroMemory(&column, sizeof(column));
		SHCOLSTATEF columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
		if(properties.pParentISF2) {
			hr = properties.pParentISF2->GetDetailsOf(NULL, properties.currentRealColumnIndex, &column);
			if(FAILED(properties.pParentISF2->GetDefaultColumnState(properties.currentRealColumnIndex, &columnState))) {
				columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
			}
		} else {
			hr = E_FAIL;
		}
		if(FAILED(hr) && properties.pParentISD) {
			hr = properties.pParentISD->GetDetailsOf(NULL, properties.currentRealColumnIndex, &column);
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
				pColumnData->width = column.cxChar * properties.averageCharWidth;
			}
			pColumnData->format = column.fmt;
			if(properties.pParentISF2 && APIWrapper::IsSupported_PSGetPropertyDescription()) {
				PROPERTYKEY propertyKey;
				if(SUCCEEDED(properties.pParentISF2->MapColumnToSCID(properties.currentRealColumnIndex, &propertyKey))) {
					CComPtr<IPropertyDescription> pPropertyDescription = NULL;
					APIWrapper::PSGetPropertyDescription(propertyKey, IID_PPV_ARGS(&pPropertyDescription), NULL);
					if(pPropertyDescription) {
						UINT width = 0;
						if(SUCCEEDED(pPropertyDescription->GetDefaultColumnWidth(&width))) {
							pColumnData->width = static_cast<int>(width) * properties.averageCharWidth;
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

			if(DPA_AppendPtr(properties.pEnumResult->hColumnBuffer, pColumnData) == -1) {
				delete pColumnData;
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				} else {
					properties.enumeratedColumns++;
					properties.pEnumResult->lastNewColumn++;
			#endif
			}

			properties.currentRealColumnIndex++;
		}

		#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
			if(properties.enumeratedColumns % CHUNKSIZE == 0) {
				// causes problems on shutdown
				EnterCriticalSection(properties.pCriticalSection);
				#ifdef USE_STL
					properties.pEnumratedColumnsQueue->push(properties.pEnumResult);
				#else
					properties.pEnumratedColumnsQueue->AddTail(properties.pEnumResult);
				#endif
				LPSHLVWBACKGROUNDCOLUMNENUMINFO pNewEnumResult = new SHLVWBACKGROUNDCOLUMNENUMINFO;
				if(pNewEnumResult) {
					*pNewEnumResult = *properties.pEnumResult;
					pNewEnumResult->hColumnBuffer = DPA_Create(CHUNKSIZE);
					pNewEnumResult->firstNewColumn += CHUNKSIZE;
					pNewEnumResult->baseRealColumnIndex += CHUNKSIZE;
				}
				properties.pEnumResult = pNewEnumResult;
				LeaveCriticalSection(properties.pCriticalSection);
				ATLASSERT_POINTER(properties.pEnumResult, SHLVWBACKGROUNDCOLUMNENUMINFO);
				ATLASSUME(properties.pEnumResult->hColumnBuffer);

				if(IsWindow(properties.hWndToNotify)) {
					PostMessage(properties.hWndToNotify, WM_TRIGGER_COLUMNENUMCOMPLETE, 0, 0);
				}
			}
		#endif
		if(hr == S_OK && WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	} while(hr == S_OK);

	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		properties.pEnumResult->moreColumnsToCome = FALSE;
	#endif

	EnterCriticalSection(properties.pCriticalSection);
	#ifdef USE_STL
		properties.pEnumratedColumnsQueue->push(properties.pEnumResult);
	#else
		properties.pEnumratedColumnsQueue->AddTail(properties.pEnumResult);
	#endif
	properties.pEnumResult = NULL;
	LeaveCriticalSection(properties.pCriticalSection);

	if(IsWindow(properties.hWndToNotify)) {
		PostMessage(properties.hWndToNotify, WM_TRIGGER_COLUMNENUMCOMPLETE, 0, 0);
	}
	return NOERROR;
}

STDMETHODIMP ShLvwBackgroundColumnEnumTask::DoRun(void)
{
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		properties.enumeratedColumns = 0;
	#endif

	CComPtr<IShellFolder> pParentISF;
	HRESULT hr = BindToPIDL(properties.pIDLParent, IID_PPV_ARGS(&pParentISF));
	if(FAILED(hr)) {
		return E_FAIL;
	}
	ATLASSUME(pParentISF);
	pParentISF->QueryInterface(IID_PPV_ARGS(&properties.pParentISF2));
	pParentISF->CreateViewObject(properties.hWndShellUIParentWindow, IID_PPV_ARGS(&properties.pParentISD));
	if(!properties.pParentISF2 && !properties.pParentISD) {
		return E_FAIL;
	}
	return E_PENDING;
}