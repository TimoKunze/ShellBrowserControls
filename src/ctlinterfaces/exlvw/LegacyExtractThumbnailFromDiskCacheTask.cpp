// LegacyExtractThumbnailFromDiskCacheTask.cpp: A specialization of RunnableTask for extracting a thumbnail from the disk cache

#include "stdafx.h"
#include "LegacyExtractThumbnailFromDiskCacheTask.h"


ShLvwLegacyExtractThumbnailFromDiskCacheTask::ShLvwLegacyExtractThumbnailFromDiskCacheTask(void)
    : RunnableTask(FALSE)
{
	properties.pThumbnailDiskCache = NULL;
	properties.pItemPath[0] = L'\0';
	properties.pBackgroundThumbnailsQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwLegacyExtractThumbnailFromDiskCacheTask::FinalRelease()
{
	if(properties.pThumbnailDiskCache) {
		properties.pThumbnailDiskCache->Release();
		properties.pThumbnailDiskCache = NULL;
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
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::Attach(HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize)
#else
	HRESULT ShLvwLegacyExtractThumbnailFromDiskCacheTask::Attach(HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, LONG itemID, LPWSTR pItemPath, LPSHELLIMAGESTORE pThumbnailDiskCache, BOOL closeDiskCacheImmediately, SIZE imageSize)
#endif
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	pThumbnailDiskCache->QueryInterface(IID_IShellImageStore, reinterpret_cast<LPVOID*>(&this->properties.pThumbnailDiskCache));
	this->properties.closeDiskCacheImmediately = closeDiskCacheImmediately;
	ATLVERIFY(SUCCEEDED(StringCchCopyNW(this->properties.pItemPath, 1024, pItemPath, lstrlenW(pItemPath))));

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
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTHUMBNAILINFO);
	ATLASSUME(properties.pThumbnailDiskCache);

	DWORD lock = 0;
	HRESULT hr = properties.pThumbnailDiskCache->Open(STGM_READ, &lock);
	if(SUCCEEDED(hr)) {
		hr = properties.pThumbnailDiskCache->GetEntry(properties.pItemPath, STGM_READ, &properties.pResult->hThumbnailOrIcon);
		if(!properties.closeDiskCacheImmediately) {
			if(IsWindow(properties.hWndToNotify)) {
				properties.closeDiskCacheImmediately = !SendMessage(properties.hWndToNotify, WM_REPORT_THUMBNAILDISKCACHEACCESS, properties.pResult->itemID, GetTickCount());
			}
		}
		properties.pThumbnailDiskCache->ReleaseLock(&lock);
		if(properties.closeDiskCacheImmediately) {
			properties.pThumbnailDiskCache->Close(NULL);
		}
	}

	// TODO: Maybe allow an interruption here?

	if(SUCCEEDED(hr)) {
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