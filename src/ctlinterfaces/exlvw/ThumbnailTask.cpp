// ThumbnailTask.cpp: A specialization of RunnableTask for retrieving listview item thumbnail details

#include "stdafx.h"
#include "ThumbnailTask.h"
#include "../../ClassFactory.h"


ShLvwThumbnailTask::ShLvwThumbnailTask(void)
    : RunnableTask(TRUE)
{
	properties.pShellItem = NULL;
	properties.pImageFactory = NULL;
	properties.pExecutableString = NULL;
	properties.pBackgroundThumbnailsQueue = NULL;
	properties.pResult = NULL;
	properties.pCriticalSection = NULL;
}

void ShLvwThumbnailTask::FinalRelease()
{
	if(properties.pShellItem) {
		properties.pShellItem->Release();
		properties.pShellItem = NULL;
	}
	if(properties.pImageFactory) {
		properties.pImageFactory->Release();
		properties.pImageFactory = NULL;
	}
	if(properties.pExecutableString) {
		HeapFree(GetProcessHeap(), 0, properties.pExecutableString);
		properties.pExecutableString = NULL;
	}
	if(properties.pResult) {
		if(properties.pResult->hThumbnailOrIcon) {
			DeleteObject(properties.pResult->hThumbnailOrIcon);
		}
		delete properties.pResult;
		properties.pResult = NULL;
	}
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
}


#ifdef USE_STL
	HRESULT ShLvwThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer)
#else
	HRESULT ShLvwThumbnailTask::Attach(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer)
#endif
{
	this->properties.hWndShellUIParentWindow = hWndShellUIParentWindow;
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pBackgroundThumbnailsQueue = pBackgroundThumbnailsQueue;
	this->properties.pCriticalSection = pCriticalSection;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.displayThumbnailAdornments = displayThumbnailAdornments;
	this->properties.displayFileTypeOverlays = displayFileTypeOverlays;
	this->properties.isComctl32600OrNewer = isComctl32600OrNewer;

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

	properties.status = TTS_NOTHINGDONE;
	return S_OK;
}

#ifdef USE_STL
	HRESULT ShLvwThumbnailTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer, IRunnableTask** ppTask)
#else
	HRESULT ShLvwThumbnailTask::CreateInstance(HWND hWndShellUIParentWindow, HWND hWndToNotify, CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO>* pBackgroundThumbnailsQueue, LPCRITICAL_SECTION pCriticalSection, PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE imageSize, ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments, ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays, BOOL isComctl32600OrNewer, IRunnableTask** ppTask)
#endif
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pBackgroundThumbnailsQueue || !pCriticalSection || !pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwThumbnailTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwThumbnailTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndShellUIParentWindow, hWndToNotify, pBackgroundThumbnailsQueue, pCriticalSection, pIDL, itemID, imageSize, displayThumbnailAdornments, displayFileTypeOverlays, isComctl32600OrNewer);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwThumbnailTask::DoInternalResume(void)
{
	ATLASSERT_POINTER(properties.pResult, SHLVWBACKGROUNDTHUMBNAILINFO);

	if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
		return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
	}

	HRESULT hr;

	if(properties.status == TTS_RETRIEVINGFLAGS1 || properties.status == TTS_RETRIEVINGIMAGE/* || properties.status == TTS_RETRIEVINGEXECUTABLEICON*/) {
		if(ShouldTaskBeAborted()) {
			properties.pResult->aborted = TRUE;
			properties.status = TTS_DONE;
		}
	}

	if(properties.status == TTS_RETRIEVINGFLAGS1) {
		if(properties.pResult->targetThumbnailSize.cx >= 32) {
			ATLASSUME(properties.pShellItem);
			properties.pShellItem->QueryInterface(IID_PPV_ARGS(&properties.pImageFactory));
			ATLASSUME(properties.pImageFactory);
			if(properties.pImageFactory) {
				HBITMAP hThumbnail = NULL;
				properties.pImageFactory->GetImage(properties.pResult->targetThumbnailSize, SIIGBF_THUMBNAILONLY, &hThumbnail);
				if(hThumbnail) {
					DeleteObject(hThumbnail);
					properties.pResult->flags |= AII_HASTHUMBNAIL;
				}
			}
			if(properties.pResult->flags & AII_HASTHUMBNAIL) {
				SFGAOF attributes = 0;
				hr = properties.pShellItem->GetAttributes(SFGAO_FOLDER, &attributes);
				if(SUCCEEDED(hr) && (attributes & SFGAO_FOLDER) == SFGAO_FOLDER) {
					properties.pResult->flags |= AII_NOFILETYPEOVERLAY;
				} else {
					LPWSTR pOverlayImageResource = NULL;
					DWORD thumbnailAdornment = 1;			// drop shadow is the default
					if(properties.displayThumbnailAdornments != sldtaNone && SUCCEEDED(GetThumbnailAdornment(properties.hWndShellUIParentWindow, properties.pIDL, TRUE, &thumbnailAdornment, &pOverlayImageResource))) {
						if(properties.displayThumbnailAdornments != sldtaAny) {
							switch(thumbnailAdornment) {
								case 1:			// drop shadow
									if(!(properties.displayThumbnailAdornments & sldtaDropShadow)) {
										thumbnailAdornment = 0;
									}
									break;
								case 2:			// photo border
									if(!(properties.displayThumbnailAdornments & sldtaPhotoBorder)) {
										thumbnailAdornment = 0;
									}
									break;
								case 3:			// video sprockets
									if(!(properties.displayThumbnailAdornments & sldtaVideoSprocket)) {
										thumbnailAdornment = 0;
									}
									break;
							}
						}
					} else {
						thumbnailAdornment = 0;
					}
					switch(thumbnailAdornment) {
						case 0:			// no adornment
							properties.pResult->flags |= AII_NOADORNMENT;
							break;
						case 1:			// drop shadow
							properties.pResult->flags |= AII_DROPSHADOWADORNMENT;
							break;
						case 2:			// photo border
							properties.pResult->flags |= AII_PHOTOBORDERADORNMENT;
							break;
						case 3:			// video sprockets
							properties.pResult->flags |= AII_VIDEOSPROCKETADORNMENT;
							break;
						default:
							ATLASSERT(FALSE && "Unknown adornment setting");
							properties.pResult->flags |= AII_DROPSHADOWADORNMENT;
							break;
					}
					// retrieve TypeOverlay setting
					properties.pResult->pOverlayImageResource = NULL;
					if(properties.displayFileTypeOverlays == sldftoFollowSystemSettings) {
						SHELLSTATE shellState = {0};
						SHGetSetSettings(&shellState, SSF_SHOWTYPEOVERLAY, FALSE);
						if(shellState.fShowTypeOverlay) {
							if(pOverlayImageResource) {
								properties.pResult->pOverlayImageResource = pOverlayImageResource;
								hr = S_OK;
							} else {
								hr = GetThumbnailTypeOverlay(properties.hWndShellUIParentWindow, properties.pIDL, TRUE, &properties.pResult->pOverlayImageResource);
							}

							if(SUCCEEDED(hr)) {
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

					if(pOverlayImageResource && pOverlayImageResource != properties.pResult->pOverlayImageResource) {
						CoTaskMemFree(pOverlayImageResource);
						pOverlayImageResource = NULL;
					}
				}
			}
		}
		properties.status = TTS_RETRIEVINGFLAGS2;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(properties.status == TTS_RETRIEVINGFLAGS2) {
		if(IsElevationRequired(properties.hWndShellUIParentWindow, properties.pIDL)) {
			properties.pResult->flags |= AII_DISPLAYELEVATIONOVERLAY;
			if(properties.pResult->flags & AII_HASTHUMBNAIL) {
				if((properties.pResult->flags & AII_FILETYPEOVERLAYMASK) == AII_NOFILETYPEOVERLAY) {
					/* The thumbnail may have a size that leads to complications when we draw the elevation shield
					 * directly onto it. Therefore we force display of executable icon. The elevation shield will be
					 * stamped onto the executable icon then.
					 */
					properties.pResult->flags &= ~AII_FILETYPEOVERLAYMASK;
					properties.pResult->flags |= AII_EXECUTABLEICONOVERLAY;
				}
			}
		}
		properties.status = (properties.pResult->targetThumbnailSize.cx >= 32 ? TTS_RETRIEVINGIMAGE : TTS_DONE);
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(properties.status == TTS_RETRIEVINGIMAGE) {
		HBITMAP hThumbnail = NULL;
		CComPtr<IThumbnailAdorner> pThumbnailAdorner = NULL;
		SIZE contentSize = properties.pResult->targetThumbnailSize;
		double aspectRatio = 1.0;
		if(properties.pResult->flags & AII_HASTHUMBNAIL) {
			switch(properties.pResult->flags & AII_ADORNMENTMASK) {
				case AII_DROPSHADOWADORNMENT:
					pThumbnailAdorner = ClassFactory::InitDropShadowAdorner();
					break;
				case AII_PHOTOBORDERADORNMENT:
					pThumbnailAdorner = ClassFactory::InitPhotoBorderAdorner();
					break;
				case AII_VIDEOSPROCKETADORNMENT:
					pThumbnailAdorner = ClassFactory::InitSprocketAdorner();
					break;
			}
			if(pThumbnailAdorner) {
				properties.pImageFactory->GetImage(properties.pResult->targetThumbnailSize, SIIGBF_BIGGERSIZEOK, &hThumbnail);
				if(hThumbnail) {
					BITMAP bmp = {0};
					GetObject(hThumbnail, sizeof(BITMAP), &bmp);
					DeleteObject(hThumbnail);
					hThumbnail = NULL;

					aspectRatio = static_cast<double>(bmp.bmHeight) / static_cast<double>(bmp.bmWidth);
				}

				// this will load the adorner icon
				ATLVERIFY(SUCCEEDED(pThumbnailAdorner->SetIconSize(properties.pResult->targetThumbnailSize, aspectRatio)));
				// now calculate the size of the content area
				ATLVERIFY(SUCCEEDED(pThumbnailAdorner->GetMaxContentSize(aspectRatio, &contentSize)));
			}
		}

		ATLASSUME(properties.pImageFactory);
		properties.pImageFactory->GetImage(contentSize, SIIGBF_BIGGERSIZEOK, &hThumbnail);
		if(hThumbnail) {
			BITMAP bmp = {0};
			GetObject(hThumbnail, sizeof(BITMAP), &bmp);
			if(bmp.bmWidth > contentSize.cx || bmp.bmHeight > contentSize.cy) {
				// we use SIIGBF_BIGGERSIZEOK only so that small thumbnails do not get up-scaled
				DeleteObject(hThumbnail);
				hThumbnail = NULL;

				properties.pImageFactory->GetImage(contentSize, 0, &hThumbnail);
			}
		}
		properties.pImageFactory->Release();
		properties.pImageFactory = NULL;
		if(hThumbnail) {
			BITMAP bmp = {0};
			GetObject(hThumbnail, sizeof(BITMAP), &bmp);
			contentSize.cx = bmp.bmWidth;
			contentSize.cy = bmp.bmHeight;
			SIZE adornedSize = contentSize;
			if(pThumbnailAdorner) {
				ATLVERIFY(SUCCEEDED(pThumbnailAdorner->GetAdornedSize(contentSize, &adornedSize)));
				adornedSize.cx = min(properties.pResult->targetThumbnailSize.cx, adornedSize.cx);
				adornedSize.cy = min(properties.pResult->targetThumbnailSize.cy, adornedSize.cy);
				ATLVERIFY(SUCCEEDED(pThumbnailAdorner->SetAdornedSize(adornedSize)));
			}
			CRect adornedRectangle(0, 0, adornedSize.cx, adornedSize.cy);
			CRect contentRectangle(0, 0, contentSize.cx, contentSize.cy);
			CRect contentAreaRectangle(0, 0, contentSize.cx, contentSize.cy);
			CRect contentToDrawRectangle(0, 0, contentSize.cx, contentSize.cy);
			if(pThumbnailAdorner) {
				ATLVERIFY(SUCCEEDED(pThumbnailAdorner->GetContentRectangle(adornedRectangle, contentSize, &contentRectangle, &contentAreaRectangle, &contentToDrawRectangle)));
			}

			CComPtr<IWICImagingFactory> pWICImagingFactory = NULL;
			CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
			if(pWICImagingFactory) {
				CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
				hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					ATLASSUME(pWICBitmapScaler);
					CComPtr<IWICBitmap> pWICBitmap = NULL;
					hr = pWICImagingFactory->CreateBitmapFromHBITMAP(hThumbnail, NULL, WICBitmapUsePremultipliedAlpha, &pWICBitmap);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						ATLASSUME(pWICBitmap);
						hr = pWICBitmapScaler->Initialize(pWICBitmap, contentRectangle.Width(), contentRectangle.Height(), WICBitmapInterpolationModeFant);
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							// create target bitmap
							BITMAPINFO bitmapInfo = {0};
							bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
							bitmapInfo.bmiHeader.biWidth = adornedSize.cx;
							bitmapInfo.bmiHeader.biHeight = -adornedSize.cy;
							bitmapInfo.bmiHeader.biPlanes = 1;
							bitmapInfo.bmiHeader.biBitCount = 32;
							bitmapInfo.bmiHeader.biCompression = BI_RGB;
							LPRGBQUAD pBits = NULL;
							HBITMAP hBitmap = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBits), NULL, 0);
							if(hBitmap) {
								ATLASSERT_POINTER(pBits, RGBQUAD);
								if(pBits) {
									if(pThumbnailAdorner) {
										BOOL drawAdornerFirst = TRUE;
										ATLVERIFY(SUCCEEDED(pThumbnailAdorner->GetDrawOrder(&drawAdornerFirst)));

										if(drawAdornerFirst) {
											ATLVERIFY(SUCCEEDED(pThumbnailAdorner->DrawIntoDIB(&adornedRectangle, pBits)));
										}
										LPRGBQUAD pThumbnailBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, contentToDrawRectangle.Width() * contentToDrawRectangle.Height() * sizeof(RGBQUAD)));
										if(pThumbnailBits) {
											WICRect rectangleToCopy = {contentToDrawRectangle.left - contentRectangle.left, contentToDrawRectangle.top - contentRectangle.top, contentToDrawRectangle.Width(), contentToDrawRectangle.Height()};
											ATLVERIFY(SUCCEEDED(pWICBitmapScaler->CopyPixels(&rectangleToCopy, rectangleToCopy.Width * sizeof(RGBQUAD), rectangleToCopy.Height * rectangleToCopy.Width * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pThumbnailBits))));

											if(IsPremultiplied(pThumbnailBits, rectangleToCopy.Height * rectangleToCopy.Width)) {
												// demultiply
												#pragma omp parallel for
												for(int i = 0; i < rectangleToCopy.Height * rectangleToCopy.Width; i++) {
													if(pThumbnailBits[i].rgbReserved) {
														pThumbnailBits[i].rgbBlue = static_cast<BYTE>((static_cast<int>(pThumbnailBits[i].rgbBlue) * 255 / static_cast<int>(pThumbnailBits[i].rgbReserved)) & 0xFF);
														pThumbnailBits[i].rgbGreen = static_cast<BYTE>((static_cast<int>(pThumbnailBits[i].rgbGreen) * 255 / static_cast<int>(pThumbnailBits[i].rgbReserved)) & 0xFF);
														pThumbnailBits[i].rgbRed = static_cast<BYTE>((static_cast<int>(pThumbnailBits[i].rgbRed) * 255 / static_cast<int>(pThumbnailBits[i].rgbReserved)) & 0xFF);
													}
												}
											}

											// copy pThumbnailBits into pBits
											POINT contentStartPoint = contentAreaRectangle.TopLeft();
											SIZE contentAreaSize = {contentAreaRectangle.Width(), contentAreaRectangle.Height()};
											#pragma omp parallel for
											for(int y = 0; y < contentAreaSize.cy; y++) {
												UINT offset1 = (y + contentStartPoint.y) * adornedSize.cx + contentStartPoint.x;
												UINT offset2 = y * rectangleToCopy.Width;
												CopyMemory(&pBits[offset1], &pThumbnailBits[offset2], contentAreaSize.cx * sizeof(RGBQUAD));
											}

											HeapFree(GetProcessHeap(), 0, pThumbnailBits);
										}

										if(!drawAdornerFirst) {
											ATLVERIFY(SUCCEEDED(pThumbnailAdorner->DrawIntoDIB(&adornedRectangle, pBits)));
										}

										HBITMAP h = NULL;
										if(SUCCEEDED(pThumbnailAdorner->PostProcess(hBitmap, pBits, &h))) {
											hBitmap = h;
										}
									} else {
										ATLVERIFY(SUCCEEDED(pWICBitmap->CopyPixels(NULL, contentSize.cx * sizeof(RGBQUAD), contentSize.cy * contentSize.cx * sizeof(RGBQUAD), reinterpret_cast<LPBYTE>(pBits))));

										if(IsPremultiplied(pBits, contentSize.cy * contentSize.cx)) {
											// demultiply
											#pragma omp parallel for
											for(int i = 0; i < contentSize.cy * contentSize.cx; i++) {
												if(pBits[i].rgbReserved) {
													pBits[i].rgbBlue = static_cast<BYTE>((static_cast<int>(pBits[i].rgbBlue) * 255 / static_cast<int>(pBits[i].rgbReserved)) & 0xFF);
													pBits[i].rgbGreen = static_cast<BYTE>((static_cast<int>(pBits[i].rgbGreen) * 255 / static_cast<int>(pBits[i].rgbReserved)) & 0xFF);
													pBits[i].rgbRed = static_cast<BYTE>((static_cast<int>(pBits[i].rgbRed) * 255 / static_cast<int>(pBits[i].rgbReserved)) & 0xFF);
												}
											}
										}
									}
								}
								properties.pResult->hThumbnailOrIcon = hBitmap;
							}
						}
					}
				}
			}
			DeleteObject(hThumbnail);
		}
		properties.status = TTS_RETRIEVINGEXECUTABLEICON;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}

	if(properties.status == TTS_RETRIEVINGEXECUTABLEICON) {
		if(!properties.pExecutableString) {
			ATLASSUME(properties.pShellItem);
			CComPtr<IQueryAssociations> pIQueryAssociations = NULL;
			SFGAOF attributes = 0;
			hr = properties.pShellItem->GetAttributes(SFGAO_LINK, &attributes);
			if(SUCCEEDED(hr) && (attributes & SFGAO_LINK) == SFGAO_LINK) {
				CComPtr<IShellItem> pTargetItem = NULL;
				hr = properties.pShellItem->BindToHandler(NULL, BHID_LinkTargetItem, IID_PPV_ARGS(&pTargetItem));
				if(SUCCEEDED(hr)) {
					ATLASSUME(pTargetItem);
					hr = pTargetItem->BindToHandler(NULL, BHID_AssociationArray, IID_PPV_ARGS(&pIQueryAssociations));
				}
			} else {
				hr = properties.pShellItem->BindToHandler(NULL, BHID_AssociationArray, IID_PPV_ARGS(&pIQueryAssociations));
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

			properties.status = TTS_DONE;
			HeapFree(GetProcessHeap(), 0, properties.pExecutableString);
			properties.pExecutableString = NULL;
		} else {
			properties.status = TTS_DONE;
		}
		properties.pShellItem->Release();
		properties.pShellItem = NULL;
		if(WaitForSingleObject(hDoneEvent, 0) == WAIT_OBJECT_0) {
			return (state == IRTIR_TASK_SUSPENDED ? E_PENDING : E_FAIL);
		}
	}
	ATLASSERT(properties.status == TTS_DONE);

	properties.pResult->mask = SIIF_OVERLAYIMAGE | SIIF_FLAGS | SIIF_IMAGE;
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
	return NOERROR;
}

STDMETHODIMP ShLvwThumbnailTask::DoRun(void)
{
	if(APIWrapper::IsSupported_SHCreateItemFromIDList()) {
		ATLVERIFY(SUCCEEDED(APIWrapper::SHCreateItemFromIDList(properties.pIDL, IID_PPV_ARGS(&properties.pShellItem), NULL)));
	}
	if(!properties.pShellItem) {
		return E_FAIL;
	}
	properties.status = TTS_RETRIEVINGFLAGS1;
	return E_PENDING;
}


BOOL ShLvwThumbnailTask::ShouldTaskBeAborted(void)
{
	/* TODO: Somehow detect that the item currently is not visible anymore. This should be done without
	 *       using SendMessage.
	 */

	BOOL abort = FALSE;
	if(properties.isComctl32600OrNewer && properties.pResult) {
		BOOL visible = TRUE;
		if(SUCCEEDED(SendMessage(properties.hWndToNotify, WM_ISLVWITEMVISIBLE, static_cast<WPARAM>(properties.pResult->itemID), reinterpret_cast<LPARAM>(&visible)))) {
			abort = !visible;
		}
	}
	return abort;
}