// LegacyTestThumbnailCacheTask.cpp: A specialization of RunnableTask for extracting a thumbnail

#include "stdafx.h"
#include "LegacyTestThumbnailCacheTask.h"


ShLvwLegacyTestThumbnailCacheTask::ShLvwLegacyTestThumbnailCacheTask(void)
    : RunnableTask(FALSE)
{
	properties.pExtractImage = NULL;
	properties.pThumbnailDiskCache = NULL;
	properties.pThumbnailPath[0] = TEXT('\0');
	properties.pItemPath[0] = L'\0';
	properties.pTasksToEnqueue = NULL;
	properties.pTaskToEnqueue = NULL;
	properties.pTasksToEnqueueCritSection = NULL;
	properties.pBackgroundThumbnailsQueue = NULL;
	properties.pBackgroundThumbnailsCritSection = NULL;
}

void ShLvwLegacyTestThumbnailCacheTask::FinalRelease()
{
	if(properties.pExtractImage) {
		properties.pExtractImage->Release();
		properties.pExtractImage = NULL;
	}
	if(properties.pThumbnailDiskCache) {
		properties.pThumbnailDiskCache->Release();
		properties.pThumbnailDiskCache = NULL;
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
	HRESULT ShLvwLegacyTestThumbnailCacheTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction)
#else
	HRESULT ShLvwLegacyTestThumbnailCacheTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHCTLSBACKGROUNDTASKINFO>* pTasksToEnqueue, LPCRITICAL_SECTION pTasksToEnqueueCritSection, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pBackgroundThumbnailsCritSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, SIZE imageSize, LPWSTR pThumbnailPath, LPWSTR pItemPath, FILETIME dateStamp, DWORD flags, DWORD priority, BOOL asynchronousExtraction)
#endif
{
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pTasksToEnqueue = pTasksToEnqueue;
	this->properties.pTasksToEnqueueCritSection = pTasksToEnqueueCritSection;
	this->properties.pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->properties.pBackgroundThumbnailsCritSection = pBackgroundThumbnailsCritSection;
	this->properties.itemIsFile = itemIsFile;
	pExtractImage->QueryInterface(IID_PPV_ARGS(&this->properties.pExtractImage));
	this->properties.useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->properties.pThumbnailDiskCache));
	}
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.itemID = itemID;
	this->properties.imageSize = imageSize;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->properties.pThumbnailPath, 1024, pThumbnailPath, lstrlenW(pThumbnailPath))));
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->properties.pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));
	this->properties.dateStamp = dateStamp;
	this->properties.flags = flags;
	this->properties.priority = priority;
	this->properties.asynchronousExtraction = asynchronousExtraction;

	properties.pTaskToEnqueue = new SHCTLSBACKGROUNDTASKINFO;
	if(properties.pTaskToEnqueue) {
		ZeroMemory(properties.pTaskToEnqueue, sizeof(SHCTLSBACKGROUNDTASKINFO));
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
	ATLASSERT_POINTER(properties.pTaskToEnqueue, SHCTLSBACKGROUNDTASKINFO);

	CComPtr<IRunnableTask> pExtractionTask;
	HRESULT hr = E_FAIL;
	BOOL schedule = TRUE;

	BOOL closeDiskCacheImmediately = (properties.pThumbnailDiskCache == NULL);
	if(properties.useThumbnailDiskCache) {
		if(!properties.pThumbnailDiskCache) {
			// no cache was provided
			hr = CoCreateInstance(CLSID_ShellThumbnailDiskCache, NULL, CLSCTX_INPROC, IID_IShellImageStore, reinterpret_cast<LPVOID*>(&properties.pThumbnailDiskCache));
			if(SUCCEEDED(hr)) {
				CComQIPtr<IPersistFolder, &IID_IPersistFolder> pPersistFolder = properties.pThumbnailDiskCache;
				if(pPersistFolder) {
					PIDLIST_ABSOLUTE pIDLParent = ILCloneFull(properties.pIDL);
					ATLASSERT_POINTER(pIDLParent, ITEMIDLIST_ABSOLUTE);
					if(pIDLParent) {
						ILRemoveLastID(pIDLParent);
						if(FAILED(pPersistFolder->Initialize(pIDLParent))) {
							properties.pThumbnailDiskCache->Release();
							properties.pThumbnailDiskCache = NULL;
						}
						ILFree(pIDLParent);
					} else {
						properties.pThumbnailDiskCache->Release();
						properties.pThumbnailDiskCache = NULL;
					}
				} else {
					properties.pThumbnailDiskCache->Release();
					properties.pThumbnailDiskCache = NULL;
				}
			}
		}

		if(properties.itemIsFile) {
			// check disk cache
			ATLASSUME(properties.pThumbnailDiskCache);
			if(properties.pThumbnailDiskCache) {
				DWORD lock = 0;
				hr = properties.pThumbnailDiskCache->Open(STGM_READ, &lock);
				if(SUCCEEDED(hr)) {
					if(!closeDiskCacheImmediately) {
						if(IsWindow(properties.hWndToNotify)) {
							closeDiskCacheImmediately = !SendMessage(properties.hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, properties.itemID, GetTickCount());
						}
					}

					FILETIME cacheTimeStamp;
					hr = properties.pThumbnailDiskCache->IsEntryInStore(properties.pItemPath, &cacheTimeStamp);
					if(hr == S_OK && CompareFileTime(&cacheTimeStamp, &properties.dateStamp) == 0) {
						hr = ShLvwLegacyExtractThumbnailFromDiskCacheTask::CreateInstance(properties.hWndToNotify, properties.pBackgroundThumbnailsQueue, properties.pBackgroundThumbnailsCritSection, properties.itemID, properties.pItemPath, properties.pThumbnailDiskCache, closeDiskCacheImmediately, properties.imageSize, &pExtractionTask);
						if(SUCCEEDED(hr)) {
							properties.pTaskToEnqueue->taskID = TASKID_ShLvwExtractThumbnailFromDiskCache;
						}
					} else {
						hr = E_FAIL;
					}
					properties.pThumbnailDiskCache->ReleaseLock(&lock);
				}
				if(closeDiskCacheImmediately) {
					properties.pThumbnailDiskCache->Close(NULL);
				}
			}
		}
	}

	if(!pExtractionTask) {
		// extract a fresh thumbnail
		hr = ShLvwLegacyExtractThumbnailTask::CreateInstance(properties.hWndToNotify, properties.pBackgroundThumbnailsQueue, properties.pBackgroundThumbnailsCritSection, properties.itemID, properties.itemIsFile, properties.pExtractImage, properties.useThumbnailDiskCache, properties.pThumbnailDiskCache, closeDiskCacheImmediately, properties.pItemPath, properties.dateStamp, properties.imageSize, &pExtractionTask);
		if(SUCCEEDED(hr)) {
			properties.pTaskToEnqueue->taskID = TASKID_ShLvwExtractThumbnail;
			if(!properties.asynchronousExtraction) {
				// run in foreground
				hr = pExtractionTask->Run();
				schedule = FALSE;
			}
		}
	}

	if(SUCCEEDED(hr)) {
		// a task has been created
		if(schedule) {
			pExtractionTask->QueryInterface(IID_PPV_ARGS(&properties.pTaskToEnqueue->pTask));
			properties.pTaskToEnqueue->taskPriority = properties.priority;

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
			return NOERROR;
		}
	}
	return hr;
}