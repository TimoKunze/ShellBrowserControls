// LegacyTestThumbnailCacheTask.cpp: A specialization of RunnableTask for extracting a thumbnail

#include "stdafx.h"
#include "LegacyTestThumbnailCacheTask.h"


ShLvwLegacyTestThumbnailCacheTask::ShLvwLegacyTestThumbnailCacheTask(void)
    : RunnableTask(FALSE)
{
	pExtractImage = NULL;
	pThumbnailDiskCache = NULL;
	pThumbnailPath[0] = TEXT('\0');
	pItemPath[0] = L'\0';
	pTasksToEnqueue = NULL;
	pTaskToEnqueue = NULL;
	pTasksToEnqueueCritSection = NULL;
	pBackgroundThumbnailsQueue = NULL;
	pBackgroundThumbnailsCritSection = NULL;
}

void ShLvwLegacyTestThumbnailCacheTask::FinalRelease()
{
	if(pExtractImage) {
		pExtractImage->Release();
		pExtractImage = NULL;
	}
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->Release();
		pThumbnailDiskCache = NULL;
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
	HRESULT ShLvwLegacyTestThumbnailCacheTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction)
#else
	HRESULT ShLvwLegacyTestThumbnailCacheTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction)
#endif
{
	this->hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->hWndToNotify = hWndToNotify;
	this->pTasksToEnqueue = pTasksToEnqueue;
	this->pTasksToEnqueueCritSection = pTasksToEnqueueCritSection;
	this->pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->pBackgroundThumbnailsCritSection = pBackgroundThumbnailsCritSection;
	this->itemIsFile = itemIsFile;
	pExtractImage->QueryInterface(IID_PPV_ARGS(&this->pExtractImage));
	this->useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->pThumbnailDiskCache));
	}
	this->pIDL = ILCloneFull(pIDL);
	this->itemID = itemID;
	this->imageSize = imageSize;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->pThumbnailPath, 1024, pThumbnailPath, lstrlenW(pThumbnailPath))));
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));
	this->dateStamp = dateStamp;
	this->flags = flags;
	this->priority = priority;
	this->asynchronousExtraction = asynchronousExtraction;

	pTaskToEnqueue = new SHCTLSBACKGROUNDTASKINFO;
	if(pTaskToEnqueue) {
		ZeroMemory(pTaskToEnqueue, sizeof(SHCTLSBACKGROUNDTASKINFO));
	} else {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShLvwLegacyTestThumbnailCacheTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyTestThumbnailCacheTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pTasksToEnqueue || !pTasksToEnqueueCritSection || !pBackgroundThumbnailsQueue || !pBackgroundThumbnailsCritSection || !pIDL || !pExtractImage || !pThumbnailPath) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwLegacyTestThumbnailCacheTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyTestThumbnailCacheTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndShellUIParentWindow, hWndToNotify, pTasksToEnqueue, pTasksToEnqueueCritSection, pBackgroundThumbnailsQueue, pBackgroundThumbnailsCritSection, pIDL, itemID, itemIsFile, pExtractImage, useThumbnailDiskCache, pThumbnailDiskCache, imageSize, pThumbnailPath, pItemPath, dateStamp, flags, priority, asynchronousExtraction);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyTestThumbnailCacheTask::DoRun(void)
{
	ATLASSERT_POINTER(pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO);

	CComPtr<IRunnableTask> pExtractionTask;
	HRESULT hr = E_FAIL;
	BOOL schedule = TRUE;

	BOOL closeDiskCacheImmediately = (pThumbnailDiskCache == NULL);
	if(useThumbnailDiskCache) {
		if(!pThumbnailDiskCache) {
			// no cache was provided
			hr = CoCreateInstance(CLSID_ShellThumbnailDiskCache, NULL, CLSCTX_INPROC, IID_IShellImageStore, reinterpret_cast<LPVOID*>(&pThumbnailDiskCache));
			if(SUCCEEDED(hr)) {
				CComQIPtr<IPersistFolder, &IID_IPersistFolder> pPersistFolder = pThumbnailDiskCache;
				if(pPersistFolder) {
					PIDLIST_ABSOLUTE pIDLParent = ILCloneFull(pIDL);
					ATLASSERT_POINTER(pIDLParent, ITEMIDLIST_ABSOLUTE);
					if(pIDLParent) {
						ILRemoveLastID(pIDLParent);
						if(FAILED(pPersistFolder->Initialize(pIDLParent))) {
							pThumbnailDiskCache->Release();
							pThumbnailDiskCache = NULL;
						}
						ILFree(pIDLParent);
					} else {
						pThumbnailDiskCache->Release();
						pThumbnailDiskCache = NULL;
					}
				} else {
					pThumbnailDiskCache->Release();
					pThumbnailDiskCache = NULL;
				}
			}
		}

		if(itemIsFile) {
			// check disk cache
			ATLASSUME(pThumbnailDiskCache);
			if(pThumbnailDiskCache) {
				DWORD lock = 0;
				hr = pThumbnailDiskCache->Open(STGM_READ, &lock);
				if(SUCCEEDED(hr)) {
					if(!closeDiskCacheImmediately) {
						if(IsWindow(hWndToNotify)) {
							closeDiskCacheImmediately = !SendMessage(hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, itemID, GetTickCount());
						}
					}

					FILETIME cacheTimeStamp;
					hr = pThumbnailDiskCache->IsEntryInStore(pItemPath, &cacheTimeStamp);
					if(hr == S_OK && CompareFileTime(&cacheTimeStamp, &dateStamp) == 0) {
						hr = ShLvwLegacyExtractThumbnailFromDiskCacheTask::CreateInstance(hWndToNotify, pBackgroundThumbnailsQueue, pBackgroundThumbnailsCritSection, itemID, pItemPath, pThumbnailDiskCache, closeDiskCacheImmediately, imageSize, &pExtractionTask);
						if(SUCCEEDED(hr)) {
							pTaskToEnqueue->taskID = TASKID_ShLvwExtractThumbnailFromDiskCache;
						}
					} else {
						hr = E_FAIL;
					}
					pThumbnailDiskCache->ReleaseLock(&lock);
				}
				if(closeDiskCacheImmediately) {
					pThumbnailDiskCache->Close(NULL);
				}
			}
		}
	}

	if(!pExtractionTask) {
		// extract a fresh thumbnail
		hr = ShLvwLegacyExtractThumbnailTask::CreateInstance(hWndToNotify, pBackgroundThumbnailsQueue, pBackgroundThumbnailsCritSection, itemID, itemIsFile, pExtractImage, useThumbnailDiskCache, pThumbnailDiskCache, closeDiskCacheImmediately, pItemPath, dateStamp, imageSize, &pExtractionTask);
		if(SUCCEEDED(hr)) {
			pTaskToEnqueue->taskID = TASKID_ShLvwExtractThumbnail;
			if(!asynchronousExtraction) {
				// run in foreground
				hr = pExtractionTask->Run();
				schedule = FALSE;
			}
		}
	}

	if(SUCCEEDED(hr)) {
		// a task has been created
		if(schedule) {
			pExtractionTask->QueryInterface(IID_PPV_ARGS(&pTaskToEnqueue->pTask));
			pTaskToEnqueue->taskPriority = priority;

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
			return NOERROR;
		}
	}
	return hr;
}