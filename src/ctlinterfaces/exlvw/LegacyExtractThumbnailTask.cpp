// LegacyExtractThumbnailTask.cpp: A specialization of RunnableTask for extracting a thumbnail

#include "stdafx.h"
#include "LegacyExtractThumbnailTask.h"


ShLvwLegacyExtractThumbnailTask::ShLvwLegacyExtractThumbnailTask(void)
    : RunnableTask(TRUE)
{
	properties.pExtractImage = NULL;
	properties.pThumbnailDiskCache = NULL;
	properties.pItemPath[0] = L'\0';
	properties.pExtractionTask = NULL;
	properties.pBackgroundThumbnailsQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwLegacyExtractThumbnailTask::FinalRelease()
{
	if(properties.pExtractImage) {
		properties.pExtractImage->Release();
		properties.pExtractImage = NULL;
	}
	if(properties.pThumbnailDiskCache) {
		properties.pThumbnailDiskCache->Release();
		properties.pThumbnailDiskCache = NULL;
	}
	if(properties.pExtractionTask) {
		properties.pExtractionTask->Release();
		properties.pExtractionTask = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->hThumbnailOrIcon) {
			DeleteObject(properties.pResult->hThumbnailOrIcon);
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyExtractThumbnailTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize)
#else
	HRESULT ShLvwLegacyExtractThumbnailTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.itemIsFile = itemIsFile;
	pExtractImage->QueryInterface(IID_PPV_ARGS(&this->properties.pExtractImage));
	this->properties.useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->properties.pThumbnailDiskCache));
	}
	this->properties.closeDiskCacheImmediately = closeDiskCacheImmediately;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->properties.pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));
	this->properties.dateStamp = dateStamp;

	properties.pResult = new SHLVWBACKGROUNDTHUMBNAILINFO;
	if(properties.pResult) {
		ZeroMemory(properties.pResult, sizeof(SHLVWBACKGROUNDTHUMBNAILINFO));
		properties.pResult->itemID = itemID;
		properties.pResult->targetThumbnailSize = imageSize;
		properties.pResult->executableIconIndex = -1;
		properties.pResult->pOverlayImageResource = NULL;
		properties.pResult->mask = SIIF_IMAGE;
	} else {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}


#ifdef USE_STL
	HRESULT ShLvwLegacyExtractThumbnailTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyExtractThumbnailTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pBackgroundThumbnailsQueue || !pCriticalSection || !pExtractImage) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwLegacyExtractThumbnailTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyExtractThumbnailTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pBackgroundThumbnailsQueue, pCriticalSection, itemID, itemIsFile, pExtractImage, useThumbnailDiskCache, pThumbnailDiskCache, closeDiskCacheImmediately, pItemPath, dateStamp, imageSize);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoRun(void)
{
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTHUMBNAILINFO);

	ATLASSUME(properties.pExtractImage);
	properties.pExtractImage->QueryInterface(IID_PPV_ARGS(&properties.pExtractionTask));

	HRESULT hr = properties.pExtractImage->Extract(&properties.pResult->hThumbnailOrIcon);
	if(properties.pExtractionTask) {
		properties.pExtractionTask->Release();
		properties.pExtractionTask = NULL;
	}

	// TODO: Maybe allow an interruption here?

	if(SUCCEEDED(hr)) {
		if(properties.useThumbnailDiskCache && properties.pThumbnailDiskCache) {
			// if the item is a file, add the image to the thumbnail disk cache
			if(properties.itemIsFile && properties.pResult->hThumbnailOrIcon) {
				DWORD lock;
				BOOL lockedDiskCache = FALSE;
				hr = properties.pThumbnailDiskCache->Open(STGM_WRITE, &lock);
				if(hr == STG_E_FILENOTFOUND) {
					hr = properties.pThumbnailDiskCache->Create(STGM_WRITE, &lock);
				}
				lockedDiskCache = SUCCEEDED(hr);
				if(SUCCEEDED(hr)) {
					hr = properties.pThumbnailDiskCache->AddEntry(properties.pItemPath, &properties.dateStamp, STGM_WRITE, properties.pResult->hThumbnailOrIcon);
					if(!properties.closeDiskCacheImmediately) {
						if(IsWindow(properties.hWndToNotify)) {
							properties.closeDiskCacheImmediately = !SendMessage(properties.hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, properties.pResult->itemID, GetTickCount());
						}
					}
					if(lockedDiskCache) {
						hr = properties.pThumbnailDiskCache->ReleaseLock(&lock);
					}
					if(properties.closeDiskCacheImmediately) {
						properties.pThumbnailDiskCache->Close(NULL);
					}
				}
			}
		}

		BITMAP bitmap = {0};
		if(GetObject(properties.pResult->hThumbnailOrIcon, sizeof(bitmap), reinterpret_cast<LPVOID*>(&bitmap))) {
			if(bitmap.bmWidth != properties.pResult->targetThumbnailSize.cx || bitmap.bmHeight != properties.pResult->targetThumbnailSize.cy) {
				CDC memoryDC;
				memoryDC.CreateCompatibleDC();
				if(!memoryDC.IsNull()) {
					LPBITMAPINFO pBitmapInfo = static_cast<LPBITMAPINFO>(HeapAlloc(GetProcessHeap(), 0, sizeof(BITMAPINFO) + sizeof(RGBQUAD) * 256));
					if(pBitmapInfo) {
						ZeroMemory(pBitmapInfo, sizeof(BITMAPINFO) + sizeof(RGBQUAD) * 256);
						pBitmapInfo->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
						if(GetDIBits(memoryDC, properties.pResult->hThumbnailOrIcon, 0, 0, NULL, pBitmapInfo, DIB_RGB_COLORS)) {
							// we have the header, now get the data
							LPRGBQUAD pBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pBitmapInfo->bmiHeader.biSizeImage));
							if(pBits) {
								if(GetDIBits(memoryDC, properties.pResult->hThumbnailOrIcon, 0, pBitmapInfo->bmiHeader.biHeight, pBits, pBitmapInfo, DIB_RGB_COLORS)) {
									RECT boundingRectangle = {0, 0, bitmap.bmWidth, bitmap.bmHeight};
									CalculateAspectRatio(&properties.pResult->targetThumbnailSize, &boundingRectangle);
									HBITMAP h = NULL;
									if(DrawOntoWhiteBackground(pBitmapInfo, pBits, &properties.pResult->targetThumbnailSize, &boundingRectangle, &h)) {
										DeleteObject(properties.pResult->hThumbnailOrIcon);
										properties.pResult->hThumbnailOrIcon = h;
									}
								}
								HeapFree(GetProcessHeap(), 0, pBits);
							}
						}
						HeapFree(GetProcessHeap(), 0, pBitmapInfo);
					}
				}
			} else {
				// the original bitmap is fine
			}
		}

		EnterCriticalSection(properties.pCriticalSection);
		#ifdef USE_STL
			properties.pBackgroundThumbnailsQueue->push(properties.pResult);
		#else
			properties.pBackgroundThumbnailsQueue->AddTail(properties.pResult);
		#endif
		properties.pResult = NULL;
		LeaveCriticalSection(properties.pCriticalSection);

		if(IsWindow(properties.hWndToNotify)) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATETHUMBNAIL, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoKill(BOOL /*unused*/)
{
	if(properties.pExtractionTask) {
		return properties.pExtractionTask->Kill(FALSE);
	}
	return E_NOTIMPL;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoSuspend(void)
{
	if(properties.pExtractionTask) {
		return properties.pExtractionTask->Suspend();
	}
	return E_NOTIMPL;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoResume(void)
{
	if(properties.pExtractionTask) {
		return properties.pExtractionTask->Resume();
	}
	return DoInternalResume();
}