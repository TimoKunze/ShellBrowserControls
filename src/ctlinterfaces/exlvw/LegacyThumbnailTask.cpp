// LegacyThumbnailTask.cpp: A specialization of RunnableTask for retrieving listview item thumbnail details

#include "stdafx.h"
#include "LegacyThumbnailTask.h"
#include "../../ClassFactory.h"


ShLvwLegacyThumbnailTask::ShLvwLegacyThumbnailTask(void)
    : RunnableTask(TRUE)
{
	pParentISF = NULL;
	pRelativePIDL = NULL;
	pExtractImage = NULL;
	pExtractImage2 = NULL;
	pThumbnailDiskCache = NULL;
	pExecutableString = NULL;
	ZeroMemory(&dateStamp, sizeof(dateStamp));
	pBackgroundThumbnailsQueue = NULL;
	pResult = NULL;
	pBackgroundThumbnailsCritSection = NULL;
	pTasksToEnqueue = NULL;
	pTaskToEnqueue = NULL;
	pTasksToEnqueueCritSection = NULL;
}

void ShLvwLegacyThumbnailTask::FinalRelease()
{
	if(pParentISF) {
		pParentISF->Release();
		pParentISF = NULL;
	}
	if(pExtractImage) {
		pExtractImage->Release();
		pExtractImage = NULL;
	}
	if(pExtractImage2) {
		pExtractImage2->Release();
		pExtractImage2 = NULL;
	}
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->Release();
		pThumbnailDiskCache = NULL;
	}
	if(pExecutableString) {
		CoTaskMemFree(pExecutableString);
		pExecutableString = NULL;
	}
	if(pResult) {
		if(pResult->hThumbnailOrIcon) {
			DeleteObject(pResult->hThumbnailOrIcon);
		}
		delete pResult;
		pResult = NULL;
	}
	if(pTaskToEnqueue) {
		if(pTaskToEnqueue->pTask) {
			pTaskToEnqueue->pTask->Release();
		}
		delete pTaskToEnqueue;
		pTaskToEnqueue = NULL;
	}
	if(pIDL) {
		ILFree(pIDL);
		pIDL = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays)
#else
	HRESULT ShLvwLegacyThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays)
#endif
{
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->hWndToNotify = hWndToNotify;
	this->pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->pBackgroundThumbnailsCritSection = pBackgroundThumbnailsCritSection;
	this->pTasksToEnqueue = pTasksToEnqueue;
	this->pTasksToEnqueueCritSection = pTasksToEnqueueCritSection;
	this->pIDL = ILCloneFull(pIDL);
	this->useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->pThumbnailDiskCache));
	}
	this->displayFileTypeOverlays = displayFileTypeOverlays;

	pResult = new SHLVWBACKGROUNDTHUMBNAILINFO;
	if(pResult) {
		ZeroMemory(pResult, sizeof(SHLVWBACKGROUNDTHUMBNAILINFO));
		pResult->itemID = itemID;
		pResult->targetThumbnailSize = imageSize;
		pResult->executableIconIndex = -1;
		pResult->pOverlayImageResource = NULL;
	} else {
		return E_OUTOFMEMORY;
	}

	status = LTTS_NOTHINGDONE;
	return S_OK;
}

#ifdef USE_STL
	HRESULT ShLvwLegacyThumbnailTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyThumbnailTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pBackgroundThumbnailsQueue || !pBackgroundThumbnailsCritSection || !pTasksToEnqueue || !pTasksToEnqueueCritSection || !pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwLegacyThumbnailTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyThumbnailTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndShellUIParentWindow, hWndToNotify, pBackgroundThumbnailsQueue, pBackgroundThumbnailsCritSection, pTasksToEnqueue, pTasksToEnqueueCritSection, pIDL, itemID, useThumbnailDiskCache, pThumbnailDiskCache, imageSize, displayFileTypeOverlays);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyThumbnailTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDTHUMBNAILINFO);

	if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
		return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
	}

	HRESULT hr;

	if(status == LTTS_RETRIEVINGDATESTAMP) {
		if(pResult->targetThumbnailSize.cx >= 32) {
			ATLASSUME(pParentISF);
			pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IExtractImage, NULL, reinterpret_cast<LPVOID*>(&pExtractImage));
			if(pExtractImage) {
				pResult->flags |= AII_HASTHUMBNAIL;
				pExtractImage->QueryInterface(IID_PPV_ARGS(&pExtractImage2));
				BOOL hasDateStamp = FALSE;
				if(pExtractImage2) {
					hasDateStamp = SUCCEEDED(pExtractImage2->GetDateStamp(&dateStamp));
				}
				if(!hasDateStamp) {
					// use the last write time
					// TODO: The times written by Windows Explorer to the disk cache are slightly different (1-3 seconds).
					WCHAR pItemPath[1024];
					pItemPath[0] = L'\0';
					STRRET path = {0};
					hr = pParentISF->GetDisplayNameOf(pRelativePIDL, SHGDN_FORPARSING, &path);
					if(SUCCEEDED(hr)) {
						StrRetToBufW(&path, pRelativePIDL, pItemPath, 1024);
						WIN32_FIND_DATAW fileDetails;
						HANDLE hFind = FindFirstFileW(pItemPath, &fileDetails);
						if(hFind != INVALID_HANDLE_VALUE) {
							dateStamp = fileDetails.ftLastWriteTime;
							FindClose(hFind);
						}
					}
				}
			}
			if(pResult->flags & AII_HASTHUMBNAIL) {
				SFGAOF attributes = SFGAO_FOLDER;
				hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				if(SUCCEEDED(hr) && (attributes & SFGAO_FOLDER) == SFGAO_FOLDER) {
					itemIsFile = FALSE;
					pResult->flags |= AII_NOFILETYPEOVERLAY;
				} else {
					itemIsFile = TRUE;
					// TODO: Does Windows XP support this registry value?
					pResult->pOverlayImageResource = NULL;
					if(displayFileTypeOverlays == sldftoFollowSystemSettings) {
						SHELLSTATE shellState = {0};
						SHGetSetSettings(&shellState, SSF_SHOWTYPEOVERLAY, FALSE);
						if(shellState.fShowTypeOverlay) {
							if(SUCCEEDED(GetThumbnailTypeOverlay(hWndShellUIParentWindow, pIDL, TRUE, &pResult->pOverlayImageResource))) {
								if(pResult->pOverlayImageResource) {
									if(lstrlenW(pResult->pOverlayImageResource) > 0) {
										pResult->flags |= AII_IMAGERESOURCEOVERLAY;
									} else {
										pResult->flags |= AII_NOFILETYPEOVERLAY;
									}
								} else {
									pResult->flags |= AII_EXECUTABLEICONOVERLAY;
								}
							} else {
								pResult->flags |= AII_EXECUTABLEICONOVERLAY;
							}
						}
					} else if(displayFileTypeOverlays & sldftoExecutableIcon) {
						pResult->flags |= AII_EXECUTABLEICONOVERLAY;
					}
				}
			}
		}
		status = ((pExtractImage && pResult->targetThumbnailSize.cx >= 32) ? LTTS_CHECKINGCACHE : LTTS_DONE);
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(status == LTTS_CHECKINGCACHE) {
		ATLASSUME(pExtractImage);
		DWORD flags = IEIFLAG_ASYNC | IEIFLAG_ORIGSIZE | IEIFLAG_QUALITY | IEIFLAG_SCREEN;
		WCHAR pThumbnailPath[1024];
		DWORD priority = TASK_PRIORITY_LV_GET_ICON;
		SIZE size = pResult->targetThumbnailSize;
		hr = pExtractImage->GetLocation(pThumbnailPath, 1024, &priority, &size, ILC_COLOR32, &flags);
		BOOL asynchronousExtraction = (hr == E_PENDING);
		if(!(SUCCEEDED(hr) || asynchronousExtraction)) {
			pThumbnailPath[0] = L'\0';
			STRRET path = {0};
			hr = pParentISF->GetDisplayNameOf(pRelativePIDL, SHGDN_FORPARSING, &path);
			if(SUCCEEDED(hr)) {
				StrRetToBufW(&path, pRelativePIDL, pThumbnailPath, 1024);
			}
			hr = S_OK;
		}
		// TODO: use the date stamp and the path to lookup the thumbnail in our own cache

		//if(!foundInCache) {
			WCHAR pItemPath[1024];
			pItemPath[0] = L'\0';
			STRRET path = {0};
			hr = pParentISF->GetDisplayNameOf(pRelativePIDL, SHGDN_FORPARSING | SHGDN_INFOLDER, &path);
			if(SUCCEEDED(hr)) {
				StrRetToBufW(&path, pRelativePIDL, pItemPath, 1024);
			}

			CComPtr<IRunnableTask> pTestCacheTask;
			priority = max(priority, 1);
			hr = ShLvwLegacyTestThumbnailCacheTask::CreateInstance(hWndShellUIParentWindow, hWndToNotify, pTasksToEnqueue, pTasksToEnqueueCritSection, pBackgroundThumbnailsQueue, pBackgroundThumbnailsCritSection, pIDL, pResult->itemID, itemIsFile, pExtractImage, useThumbnailDiskCache, pThumbnailDiskCache, pResult->targetThumbnailSize, pThumbnailPath, pItemPath, dateStamp, flags, priority, asynchronousExtraction, &pTestCacheTask);
			if(SUCCEEDED(hr)) {
				if(!asynchronousExtraction) {
					// run in foreground
					hr = pTestCacheTask->Run();
				} else {
					pTaskToEnqueue = new SHCTLSBACKGROUNDTASKINFO;
					if(pTaskToEnqueue) {
						pTestCacheTask->QueryInterface(IID_PPV_ARGS(&pTaskToEnqueue->pTask));
						pTaskToEnqueue->taskID = TASKID_ShLvwThumbnailTestCache;
						pTaskToEnqueue->taskPriority = priority;
					} else {
						return E_OUTOFMEMORY;
					}
					EnterCriticalSection(pTasksToEnqueueCritSection);
					#ifdef USE_STL
						pTasksToEnqueue->push(pTaskToEnqueue);
					#else
						pTasksToEnqueue->AddTail(pTaskToEnqueue);
					#endif
					pTaskToEnqueue = NULL;
					LeaveCriticalSection(pTasksToEnqueueCritSection);

					if(IsWindow(hWndToNotify)) {
						PostMessage(hWndToNotify, WM_TRIGGER_ENQUEUETASK, 0, 0);
					}
				}
			}
		//}
		pExtractImage->Release();
		pExtractImage = NULL;

		status = LTTS_RETRIEVINGEXECUTABLEICON;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(status == LTTS_RETRIEVINGEXECUTABLEICON) {
		if(!pExecutableString) {
			ATLASSUME(pParentISF);
			CComPtr<IQueryAssociations> pIQueryAssociations = NULL;
			SFGAOF attributes = SFGAO_LINK;
			hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
			if(SUCCEEDED(hr) && (attributes & SFGAO_LINK) == SFGAO_LINK) {
				PIDLIST_ABSOLUTE pIDLTarget = NULL;
				GetLinkTarget(hWndShellUIParentWindow, pIDL, &pIDLTarget);
				CComPtr<IShellFolder> pTargetItem = NULL;
				PCUITEMID_CHILD pChildPIDL = NULL;
				if(pIDLTarget) {
					::SHBindToParent(pIDLTarget, IID_PPV_ARGS(&pTargetItem), &pChildPIDL);
					ATLASSUME(pTargetItem);
					ATLASSUME(pChildPIDL);
				}
				if(pTargetItem && pChildPIDL) {
					hr = pTargetItem->GetUIObjectOf(hWndShellUIParentWindow, 1, &pChildPIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
				}
				ILFree(pIDLTarget);
			} else {
				hr = pParentISF->GetUIObjectOf(hWndShellUIParentWindow, 1, &pRelativePIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
			}
			if(SUCCEEDED(hr)) {
				ATLASSUME(pIQueryAssociations);
				DWORD bufferSize = 0;
				hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE | ASSOCF_REMAPRUNDLL, ASSOCSTR_EXECUTABLE, NULL, NULL, &bufferSize);
				if(hr == S_FALSE) {
					pExecutableString = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (bufferSize + 1) * sizeof(WCHAR)));
					ATLASSERT(pExecutableString);
					if(!pExecutableString) {
						return E_OUTOFMEMORY;
					}
					hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE | ASSOCF_REMAPRUNDLL, ASSOCSTR_EXECUTABLE, NULL, pExecutableString, &bufferSize);
					ATLASSERT(SUCCEEDED(hr));
				}
			}
		}
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
		if(pExecutableString) {
			SHFILEINFOW fileInfoData = {0};
			SHGetFileInfoW(pExecutableString, 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
			ATLASSERT(fileInfoData.iIcon >= 0);
			pResult->executableIconIndex = fileInfoData.iIcon;

			status = LTTS_DONE;
			HeapFree(GetProcessHeap(), 0, pExecutableString);
			pExecutableString = NULL;
		} else {
			status = LTTS_DONE;
		}
		pParentISF->Release();
		pParentISF = NULL;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}
	ATLASSERT(status == LTTS_DONE);

	pResult->mask = SIIF_OVERLAYIMAGE | SIIF_FLAGS;
	if(pResult->hThumbnailOrIcon) {
		pResult->mask |= SIIF_IMAGE;
	}
	EnterCriticalSection(pBackgroundThumbnailsCritSection);
	#ifdef USE_STL
		pBackgroundThumbnailsQueue->push(pResult);
	#else
		pBackgroundThumbnailsQueue->AddTail(pResult);
	#endif
	pResult = NULL;
	LeaveCriticalSection(pBackgroundThumbnailsCritSection);

	if(IsWindow(hWndToNotify)) {
		PostMessage(hWndToNotify, WM_TRIGGER_UPDATETHUMBNAIL, 0, 0);
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyThumbnailTask::DoRun(void)
{
	if(FAILED(::SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL))) {
		return E_FAIL;
	}
	ATLASSUME(pParentISF);
	ATLASSUME(pRelativePIDL);

	status = LTTS_RETRIEVINGDATESTAMP;
	return E_PENDING;
}