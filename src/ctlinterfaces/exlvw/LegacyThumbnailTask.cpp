// LegacyThumbnailTask.cpp: A specialization of RunnableTask for retrieving listview item thumbnail details

#include "stdafx.h"
#include "LegacyThumbnailTask.h"
#include "../../ClassFactory.h"


ShLvwLegacyThumbnailTask::ShLvwLegacyThumbnailTask(void)
    : RunnableTask(TRUE)
{
	properties.pParentISF = NULL;
	properties.pRelativePIDL = NULL;
	properties.pExtractImage = NULL;
	properties.pExtractImage2 = NULL;
	properties.pThumbnailDiskCache = NULL;
	properties.pExecutableString = NULL;
	ZeroMemory(&properties.dateStamp, sizeof(properties.dateStamp));
	properties.pBackgroundThumbnailsQueue = NULL;
	properties.pResult = NULL;
	properties.pBackgroundThumbnailsCritSection = NULL;
	properties.pTasksToEnqueue = NULL;
	properties.pTaskToEnqueue = NULL;
	properties.pTasksToEnqueueCritSection = NULL;
}

void ShLvwLegacyThumbnailTask::FinalRelease()
{
	if(properties.pParentISF) {
		properties.pParentISF->Release();
		properties.pParentISF = NULL;
	}
	if(properties.pExtractImage) {
		properties.pExtractImage->Release();
		properties.pExtractImage = NULL;
	}
	if(properties.pExtractImage2) {
		properties.pExtractImage2->Release();
		properties.pExtractImage2 = NULL;
	}
	if(properties.pThumbnailDiskCache) {
		properties.pThumbnailDiskCache->Release();
		properties.pThumbnailDiskCache = NULL;
	}
	if(properties.pExecutableString) {
		CoTaskMemFree(properties.pExecutableString);
		properties.pExecutableString = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->hThumbnailOrIcon) {
			DeleteObject(properties.pResult->hThumbnailOrIcon);
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
	if(properties.pTaskToEnqueue) {
		if(properties.pTaskToEnqueue->pTask) {
			properties.pTaskToEnqueue->pTask->Release();
		}
		delete properties.pTaskToEnqueue;
		properties.pTaskToEnqueue = NULL;
	}
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays)
#else
	HRESULT ShLvwLegacyThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays)
#endif
{
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->properties.pBackgroundThumbnailsCritSection = pBackgroundThumbnailsCritSection;
	this->properties.pTasksToEnqueue = pTasksToEnqueue;
	this->properties.pTasksToEnqueueCritSection = pTasksToEnqueueCritSection;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->properties.pThumbnailDiskCache));
	}
	this->properties.displayFileTypeOverlays = displayFileTypeOverlays;

	properties.pResult = new SHLVWBACKGROUNDTHUMBNAILINFO;
	if(properties.pResult) {
		ZeroMemory(properties.pResult, sizeof(SHLVWBACKGROUNDTHUMBNAILINFO));
		properties.pResult->itemID = itemID;
		properties.pResult->targetThumbnailSize = imageSize;
		properties.pResult->executableIconIndex = -1;
		properties.pResult->pOverlayImageResource = NULL;
	} else {
		return E_OUTOFMEMORY;
	}

	properties.status = LTTS_NOTHINGDONE;
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
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTHUMBNAILINFO);

	if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
		return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
	}

	HRESULT hr;

	if(properties.status == LTTS_RETRIEVINGDATESTAMP) {
		if(properties.pResult->targetThumbnailSize.cx >= 32) {
			ATLASSUME(properties.pParentISF);
			properties.pParentISF->GetUIObjectOf(properties.hWndShellUIParentWindow, 1, &properties.pRelativePIDL, IID_IExtractImage, NULL, reinterpret_cast<LPVOID*>(&properties.pExtractImage));
			if(properties.pExtractImage) {
				properties.pResult->flags |= AII_HASTHUMBNAIL;
				properties.pExtractImage->QueryInterface(IID_PPV_ARGS(&properties.pExtractImage2));
				BOOL hasDateStamp = FALSE;
				if(properties.pExtractImage2) {
					hasDateStamp = SUCCEEDED(properties.pExtractImage2->GetDateStamp(&properties.dateStamp));
				}
				if(!hasDateStamp) {
					// use the last write time
					// TODO: The times written by Windows Explorer to the disk cache are slightly different (1-3 seconds).
					WCHAR pItemPath[1024];
					pItemPath[0] = L'\0';
					STRRET path = {0};
					hr = properties.pParentISF->GetDisplayNameOf(properties.pRelativePIDL, SHGDN_FORPARSING, &path);
					if(SUCCEEDED(hr)) {
						StrRetToBufW(&path, properties.pRelativePIDL, pItemPath, 1024);
						WIN32_FIND_DATAW fileDetails;
						HANDLE hFind = FindFirstFileW(pItemPath, &fileDetails);
						if(hFind != INVALID_HANDLE_VALUE) {
							properties.dateStamp = fileDetails.ftLastWriteTime;
							FindClose(hFind);
						}
					}
				}
			}
			if(properties.pResult->flags & AII_HASTHUMBNAIL) {
				SFGAOF attributes = SFGAO_FOLDER;
				hr = properties.pParentISF->GetAttributesOf(1, &properties.pRelativePIDL, &attributes);
				if(SUCCEEDED(hr) && (attributes & SFGAO_FOLDER) == SFGAO_FOLDER) {
					properties.itemIsFile = FALSE;
					properties.pResult->flags |= AII_NOFILETYPEOVERLAY;
				} else {
					properties.itemIsFile = TRUE;
					// TODO: Does Windows XP support this registry value?
					properties.pResult->pOverlayImageResource = NULL;
					if(properties.displayFileTypeOverlays == sldftoFollowSystemSettings) {
						SHELLSTATE shellState = {0};
						SHGetSetSettings(&shellState, SSF_SHOWTYPEOVERLAY, FALSE);
						if(shellState.fShowTypeOverlay) {
							if(SUCCEEDED(GetThumbnailTypeOverlay(properties.hWndShellUIParentWindow, properties.pIDL, TRUE, &properties.pResult->pOverlayImageResource))) {
								if(properties.pResult->pOverlayImageResource) {
									if(lstrlenW(properties.pResult->pOverlayImageResource) > 0) {
										properties.pResult->flags |= AII_IMAGERESOURCEOVERLAY;
									} else {
										properties.pResult->flags |= AII_NOFILETYPEOVERLAY;
									}
								} else {
									properties.pResult->flags |= AII_EXECUTABLEICONOVERLAY;
								}
							} else {
								properties.pResult->flags |= AII_EXECUTABLEICONOVERLAY;
							}
						}
					} else if(properties.displayFileTypeOverlays & sldftoExecutableIcon) {
						properties.pResult->flags |= AII_EXECUTABLEICONOVERLAY;
					}
				}
			}
		}
		properties.status = ((properties.pExtractImage && properties.pResult->targetThumbnailSize.cx >= 32) ? LTTS_CHECKINGCACHE : LTTS_DONE);
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(properties.status == LTTS_CHECKINGCACHE) {
		ATLASSUME(properties.pExtractImage);
		DWORD flags = IEIFLAG_ASYNC | IEIFLAG_ORIGSIZE | IEIFLAG_QUALITY | IEIFLAG_SCREEN;
		WCHAR pThumbnailPath[1024];
		DWORD priority = TASK_PRIORITY_LV_GET_ICON;
		SIZE size = properties.pResult->targetThumbnailSize;
		hr = properties.pExtractImage->GetLocation(pThumbnailPath, 1024, &priority, &size, ILC_COLOR32, &flags);
		BOOL asynchronousExtraction = (hr == E_PENDING);
		if(!(SUCCEEDED(hr) || asynchronousExtraction)) {
			pThumbnailPath[0] = L'\0';
			STRRET path = {0};
			hr = properties.pParentISF->GetDisplayNameOf(properties.pRelativePIDL, SHGDN_FORPARSING, &path);
			if(SUCCEEDED(hr)) {
				StrRetToBufW(&path, properties.pRelativePIDL, pThumbnailPath, 1024);
			}
			hr = S_OK;
		}
		// TODO: use the date stamp and the path to lookup the thumbnail in our own cache

		//if(!foundInCache) {
			WCHAR pItemPath[1024];
			pItemPath[0] = L'\0';
			STRRET path = {0};
			hr = properties.pParentISF->GetDisplayNameOf(properties.pRelativePIDL, SHGDN_FORPARSING | SHGDN_INFOLDER, &path);
			if(SUCCEEDED(hr)) {
				StrRetToBufW(&path, properties.pRelativePIDL, pItemPath, 1024);
			}

			CComPtr<IRunnableTask> pTestCacheTask;
			priority = max(priority, 1);
			hr = ShLvwLegacyTestThumbnailCacheTask::CreateInstance(properties.hWndShellUIParentWindow, properties.hWndToNotify, properties.pTasksToEnqueue, properties.pTasksToEnqueueCritSection, properties.pBackgroundThumbnailsQueue, properties.pBackgroundThumbnailsCritSection, properties.pIDL, properties.pResult->itemID, properties.itemIsFile, properties.pExtractImage, properties.useThumbnailDiskCache, properties.pThumbnailDiskCache, properties.pResult->targetThumbnailSize, pThumbnailPath, pItemPath, properties.dateStamp, flags, priority, asynchronousExtraction, &pTestCacheTask);
			if(SUCCEEDED(hr)) {
				if(!asynchronousExtraction) {
					// run in foreground
					hr = pTestCacheTask->Run();
				} else {
					properties.pTaskToEnqueue = new SHCTLSBACKGROUNDTASKINFO;
					if(properties.pTaskToEnqueue) {
						pTestCacheTask->QueryInterface(IID_PPV_ARGS(&properties.pTaskToEnqueue->pTask));
						properties.pTaskToEnqueue->taskID = TASKID_ShLvwThumbnailTestCache;
						properties.pTaskToEnqueue->taskPriority = priority;
					} else {
						return E_OUTOFMEMORY;
					}
					EnterCriticalSection(properties.pTasksToEnqueueCritSection);
					#ifdef USE_STL
						properties.pTasksToEnqueue->push(properties.pTaskToEnqueue);
					#else
						properties.pTasksToEnqueue->AddTail(properties.pTaskToEnqueue);
					#endif
						properties.pTaskToEnqueue = NULL;
					LeaveCriticalSection(properties.pTasksToEnqueueCritSection);

					if(IsWindow(properties.hWndToNotify)) {
						PostMessage(properties.hWndToNotify, WM_TRIGGER_ENQUEUETASK, 0, 0);
					}
				}
			}
		//}
		properties.pExtractImage->Release();
		properties.pExtractImage = NULL;

		properties.status = LTTS_RETRIEVINGEXECUTABLEICON;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(properties.status == LTTS_RETRIEVINGEXECUTABLEICON) {
		if(!properties.pExecutableString) {
			ATLASSUME(properties.pParentISF);
			CComPtr<IQueryAssociations> pIQueryAssociations = NULL;
			SFGAOF attributes = SFGAO_LINK;
			hr = properties.pParentISF->GetAttributesOf(1, &properties.pRelativePIDL, &attributes);
			if(SUCCEEDED(hr) && (attributes & SFGAO_LINK) == SFGAO_LINK) {
				PIDLIST_ABSOLUTE pIDLTarget = NULL;
				GetLinkTarget(properties.hWndShellUIParentWindow, properties.pIDL, &pIDLTarget);
				CComPtr<IShellFolder> pTargetItem = NULL;
				PCUITEMID_CHILD pChildPIDL = NULL;
				if(pIDLTarget) {
					::SHBindToParent(pIDLTarget, IID_PPV_ARGS(&pTargetItem), &pChildPIDL);
					ATLASSUME(pTargetItem);
					ATLASSUME(pChildPIDL);
				}
				if(pTargetItem && pChildPIDL) {
					hr = pTargetItem->GetUIObjectOf(properties.hWndShellUIParentWindow, 1, &pChildPIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
				}
				ILFree(pIDLTarget);
			} else {
				hr = properties.pParentISF->GetUIObjectOf(properties.hWndShellUIParentWindow, 1, &properties.pRelativePIDL, IID_IQueryAssociations, NULL, reinterpret_cast<LPVOID*>(&pIQueryAssociations));
			}
			if(SUCCEEDED(hr)) {
				ATLASSUME(pIQueryAssociations);
				DWORD bufferSize = 0;
				hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE | ASSOCF_REMAPRUNDLL, ASSOCSTR_EXECUTABLE, NULL, NULL, &bufferSize);
				if(hr == S_FALSE) {
					properties.pExecutableString = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (bufferSize + 1) * sizeof(WCHAR)));
					ATLASSERT(properties.pExecutableString);
					if(!properties.pExecutableString) {
						return E_OUTOFMEMORY;
					}
					hr = pIQueryAssociations->GetString(ASSOCF_NOTRUNCATE | ASSOCF_REMAPRUNDLL, ASSOCSTR_EXECUTABLE, NULL, properties.pExecutableString, &bufferSize);
					ATLASSERT(SUCCEEDED(hr));
				}
			}
		}
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
		if(properties.pExecutableString) {
			SHFILEINFOW fileInfoData = {0};
			SHGetFileInfoW(properties.pExecutableString, 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
			ATLASSERT(fileInfoData.iIcon >= 0);
			properties.pResult->executableIconIndex = fileInfoData.iIcon;

			properties.status = LTTS_DONE;
			HeapFree(GetProcessHeap(), 0, properties.pExecutableString);
			properties.pExecutableString = NULL;
		} else {
			properties.status = LTTS_DONE;
		}
		properties.pParentISF->Release();
		properties.pParentISF = NULL;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}
	ATLASSERT(properties.status == LTTS_DONE);

	properties.pResult->mask = SIIF_OVERLAYIMAGE | SIIF_FLAGS;
	if(properties.pResult->hThumbnailOrIcon) {
		properties.pResult->mask |= SIIF_IMAGE;
	}
	EnterCriticalSection(properties.pBackgroundThumbnailsCritSection);
	#ifdef USE_STL
		properties.pBackgroundThumbnailsQueue->push(pResult);
	#else
		properties.pBackgroundThumbnailsQueue->AddTail(properties.pResult);
	#endif
	properties.pResult = NULL;
	LeaveCriticalSection(properties.pBackgroundThumbnailsCritSection);

	if(IsWindow(properties.hWndToNotify)) {
		PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATETHUMBNAIL, 0, 0);
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyThumbnailTask::DoRun(void)
{
	if(FAILED(::SHBindToParent(properties.pIDL, IID_PPV_ARGS(&properties.pParentISF), &properties.pRelativePIDL))) {
		return E_FAIL;
	}
	ATLASSUME(properties.pParentISF);
	ATLASSUME(properties.pRelativePIDL);

	properties.status = LTTS_RETRIEVINGDATESTAMP;
	return E_PENDING;
}