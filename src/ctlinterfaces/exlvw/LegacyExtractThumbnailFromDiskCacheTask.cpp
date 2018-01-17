// LegacyExtractThumbnailFromDiskCacheTask.cpp: A specialization of RunnableTask for extracting a thumbnail from the disk cache

#include "stdafx.h"
#include "LegacyExtractThumbnailFromDiskCacheTask.h"


ShLvwLegacyExtractThumbnailFromDiskCacheTask::ShLvwLegacyExtractThumbnailFromDiskCacheTask(void)
    : RunnableTask(FALSE)
{
	pThumbnailDiskCache = NULL;
	pItemPath[0] = L'\0';
	pBackgroundThumbnailsQueue = NULL;
	pResult = NULL;
	pCriticalSection = NULL;
}

void ShLvwLegacyExtractThumbnailFromDiskCacheTask::FinalRelease()
{
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->Release();
		pThumbnailDiskCache = NULL;
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
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize)
#else
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize)
#endif
{
	this->hWndToNotify = hWndToNotify;
	this->pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->pCriticalSection = pCriticalSection;
	pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->pThumbnailDiskCache));
	this->closeDiskCacheImmediately = closeDiskCacheImmediately;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));

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
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::CreateInstance(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize, IRunnableTask** ppTask)
#else
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::CreateInstance(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pBackgroundThumbnailsQueue || !pCriticalSection || !pItemPath || !pThumbnailDiskCache) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwLegacyExtractThumbnailFromDiskCacheTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwLegacyExtractThumbnailFromDiskCacheTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pBackgroundThumbnailsQueue, pCriticalSection, itemID, pItemPath, pThumbnailDiskCache, closeDiskCacheImmediately, imageSize);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwLegacyExtractThumbnailFromDiskCacheTask::DoRun(void)
{
	ATLASSERT_POINTER(pResult, SHLVWBACKGROUNDTHUMBNAILINFO);
	ATLASSUME(pThumbnailDiskCache);

	DWORD lock = 0;
	HRESULT hr = pThumbnailDiskCache->Open(STGM_READ, &lock);
	if(SUCCEEDED(hr)) {
		hr = pThumbnailDiskCache->GetEntry(pItemPath, STGM_READ, &pResult->hThumbnailOrIcon);
		if(!closeDiskCacheImmediately) {
			if(IsWindow(hWndToNotify)) {
				closeDiskCacheImmediately = !SendMessage(hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, pResult->itemID, GetTickCount());
			}
		}
		pThumbnailDiskCache->ReleaseLock(&lock);
		if(closeDiskCacheImmediately) {
			pThumbnailDiskCache->Close(NULL);
		}
	}

	// TODO: Maybe allow an interruption here?

	if(SUCCEEDED(hr)) {
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