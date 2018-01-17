// LegacyExtractThumbnailTask.cpp: A specialization of RunnableTask for extracting a thumbnail

#include "stdafx.h"
#include "LegacyExtractThumbnailTask.h"


ShLvwLegacyExtractThumbnailTask::ShLvwLegacyExtractThumbnailTask(void)
    : RunnableTask(TRUE)
{
	pExtractImage = NULL;
	pThumbnailDiskCache = NULL;
	pItemPath[0] = L'\0';
	pExtractionTask = NULL;
	pBackgroundThumbnailsQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwLegacyExtractThumbnailTask::FinalRelease()
{
	if(pExtractImage) {
		pExtractImage->Release();
		pExtractImage = NULL;
	}
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->Release();
		pThumbnailDiskCache = NULL;
	}
	if(pExtractionTask) {
		pExtractionTask->Release();
		pExtractionTask = NULL;
	}
	if(pResult) {
		if(pResult->hThumbnailOrIcon) {
			DeleteObject(pResult->hThumbnailOrIcon);
		}
		delete pResult;
		pResult = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwLegacyExtractThumbnailTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize)
#else
	HRESULT ShLvwLegacyExtractThumbnailTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, BOOL itemIsFile, LPEXTRACTIMAGE pExtractImage, BOOL useThumbnailDiskCache, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, LPWSTR pItemPath, FILETIME dateStamp, SIZE imageSize)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->pCriticalSection = pCriticalSection;
	this->itemIsFile = itemIsFile;
	pExtractImage->QueryInterface(IID_PPV_ARGS(&this->pExtractImage));
	this->useThumbnailDiskCache = useThumbnailDiskCache;
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->pThumbnailDiskCache));
	}
	this->closeDiskCacheImmediately = closeDiskCacheImmediately;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));
	this->dateStamp = dateStamp;

	pResult = new SHLVWBACKGROUNDTHUMBNAILINFO;
	if(pResult) {
		ZeroMemory(pResult, sizeof(SHLVWBACKGROUNDTHUMBNAILINFO));
		pResult->itemID = itemID;
		pResult->targetThumbnailSize = imageSize;
		pResult->executableIconIndex = -1;
		pResult->pOverlayImageResource = NULL;
		pResult->mask = SIIF_IMAGE;
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
	ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDTHUMBNAILINFO);

	ATLASSUME(pExtractImage);
	pExtractImage->QueryInterface(IID_PPV_ARGS(&pExtractionTask));

	HRESULT hr = pExtractImage->Extract(&pResult->hThumbnailOrIcon);
	if(pExtractionTask) {
		pExtractionTask->Release();
		pExtractionTask = NULL;
	}

	// TODO: Maybe allow an interruption here?

	if(SUCCEEDED(hr)) {
		if(useThumbnailDiskCache && pThumbnailDiskCache) {
			// if the item is a file, add the image to the thumbnail disk cache
			if(itemIsFile && pResult->hThumbnailOrIcon) {
				DWORD lock;
				BOOL lockedDiskCache = FALSE;
				hr = pThumbnailDiskCache->Open(STGM_WRITE, &lock);
				if(hr == STG_E_FILENOTFOUND) {
					hr = pThumbnailDiskCache->Create(STGM_WRITE, &lock);
				}
				lockedDiskCache = SUCCEEDED(hr);
				if(SUCCEEDED(hr)) {
					hr = pThumbnailDiskCache->AddEntry(pItemPath, &dateStamp, STGM_WRITE, pResult->hThumbnailOrIcon);
					if(!closeDiskCacheImmediately) {
						if(IsWindow(hWndToNotify)) {
							closeDiskCacheImmediately = !SendMessage(hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, pResult->itemID, GetTickCount());
						}
					}
					if(lockedDiskCache) {
						hr = pThumbnailDiskCache->ReleaseLock(&lock);
					}
					if(closeDiskCacheImmediately) {
						pThumbnailDiskCache->Close(NULL);
					}
				}
			}
		}

		BITMAP bitmap = {0};
		if(GetObject(pResult->hThumbnailOrIcon, sizeof(bitmap), reinterpret_cast<LPVOID*>(&bitmap))) {
			if(bitmap.bmWidth != pResult->targetThumbnailSize.cx || bitmap.bmHeight != pResult->targetThumbnailSize.cy) {
				CDC memoryDC;
				memoryDC.CreateCompatibleDC();
				if(!memoryDC.IsNull()) {
					LPBITMAPINFO pBitmapInfo = static_cast<LPBITMAPINFO>(HeapAlloc(GetProcessHeap(), 0, sizeof(BITMAPINFO) + sizeof(RGBQUAD) * 256));
					if(pBitmapInfo) {
						ZeroMemory(pBitmapInfo, sizeof(BITMAPINFO) + sizeof(RGBQUAD) * 256);
						pBitmapInfo->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
						if(GetDIBits(memoryDC, pResult->hThumbnailOrIcon, 0, 0, NULL, pBitmapInfo, DIB_RGB_COLORS)) {
							// we have the header, now get the data
							LPRGBQUAD pBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pBitmapInfo->bmiHeader.biSizeImage));
							if(pBits) {
								if(GetDIBits(memoryDC, pResult->hThumbnailOrIcon, 0, pBitmapInfo->bmiHeader.biHeight, pBits, pBitmapInfo, DIB_RGB_COLORS)) {
									RECT boundingRectangle = {0, 0, bitmap.bmWidth, bitmap.bmHeight};
									CalculateAspectRatio(&pResult->targetThumbnailSize, &boundingRectangle);
									HBITMAP h = NULL;
									if(DrawOntoWhiteBackground(pBitmapInfo, pBits, &pResult->targetThumbnailSize, &boundingRectangle, &h)) {
										DeleteObject(pResult->hThumbnailOrIcon);
										pResult->hThumbnailOrIcon = h;
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

		EnterCriticalSection(pCriticalSection);
		#ifdef USE_STL
			pBackgroundThumbnailsQueue->push(pResult);
		#else
			pBackgroundThumbnailsQueue->AddTail(pResult);
		#endif
		pResult = NULL;
		LeaveCriticalSection(pCriticalSection);

		if(IsWindow(hWndToNotify)) {
			PostMessage(hWndToNotify, WM_TRIGGER_UPDATETHUMBNAIL, 0, 0);
		}
	}
	return NOERROR;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoKill(BOOL /*unused*/)
{
	if(pExtractionTask) {
		return pExtractionTask->Kill(FALSE);
	}
	return E_NOTIMPL;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoSuspend(void)
{
	if(pExtractionTask) {
		return pExtractionTask->Suspend();
	}
	return E_NOTIMPL;
}

STDMETHODIMP ShLvwLegacyExtractThumbnailTask::DoResume(void)
{
	if(pExtractionTask) {
		return pExtractionTask->Resume();
	}
	return DoInternalResume();
}