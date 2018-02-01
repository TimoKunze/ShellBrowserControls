// IconTask.cpp: A specialization of RunnableTask for retrieving treeview item icon indexes

#include "stdafx.h"
#include "IconTask.h"


ShTvwIconTask::ShTvwIconTask(void)
    : RunnableTask(FALSE)
{
}

void ShTvwIconTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
}


HRESULT ShTvwIconTask::Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage, UseGenericIconsConstants useGenericIcons)
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.itemHandle = itemHandle;
	this->properties.retrieveNormalImage = retrieveNormalImage;
	this->properties.retrieveSelectedImage = retrieveSelectedImage;
	this->properties.useGenericIcons = useGenericIcons;
	return S_OK;
}


HRESULT ShTvwIconTask::CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage, UseGenericIconsConstants useGenericIcons, IRunnableTask** ppTask)
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShTvwIconTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShTvwIconTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pIDL, itemHandle, retrieveNormalImage, retrieveSelectedImage, useGenericIcons);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShTvwIconTask::DoRun(void)
{
	BOOL isFolder = FALSE;
	BOOL getGenericIcons = FALSE;
	LPWSTR pPath = NULL;
	CComPtr<IShellFolder> pParentISF = NULL;
	PUITEMID_CHILD pRelativePIDL = NULL;

	if(properties.useGenericIcons != ugiNever) {
		HRESULT hr = SHBindToParent(properties.pIDL, IID_PPV_ARGS(&pParentISF), const_cast<PCUITEMID_CHILD*>(&pRelativePIDL));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

			switch(properties.useGenericIcons) {
				case ugiOnlyForSlowItems:
					getGenericIcons = IsSlowItem(pParentISF, pRelativePIDL, properties.pIDL, FALSE, FALSE);
					break;
				case ugiAlways:
					getGenericIcons = TRUE;
					break;
			}
			if(getGenericIcons) {
				SFGAOF shellAttributes = SFGAO_FOLDER | SFGAO_FILESYSTEM;
				ATLVERIFY(SUCCEEDED(pParentISF->GetAttributesOf(1, &pRelativePIDL, &shellAttributes)));
				getGenericIcons = (shellAttributes & SFGAO_FILESYSTEM);
				if(getGenericIcons) {
					isFolder = (shellAttributes & SFGAO_FOLDER);
					if(ILIsEmpty(pRelativePIDL)) {
						// desktop
						getGenericIcons = FALSE;
					}
					if(getGenericIcons) {
						BOOL isDrive = FALSE;
						ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
						if(pPath) {
							if(PathIsRootW(pPath)) {
								int len = lstrlenW(pPath);
								if(len >= 2 && pPath[1] == L':') {
									isDrive = (len == 2 || (len == 3 && pPath[2] == L'\\'));
								}
							}
						}
						getGenericIcons = !isDrive;
					}
				}
			}
		}
	}

	// If it's a folder, we don't need to retrieve generic icons, because the control already has set the correct icons.
	if(!(getGenericIcons && isFolder)) {
		int iconIndex = I_IMAGECALLBACK;
		int selectedIconIndex = I_IMAGECALLBACK;
		if(getGenericIcons) {
			SHFILEINFOW fileInfoData = {0};
			if(!pPath) {
				ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
			}
			if(pPath) {
				LPWSTR pExtension = PathFindExtensionW(pPath);
				if(properties.retrieveNormalImage) {
					SHGetFileInfoW(pExtension, FILE_ATTRIBUTE_NORMAL, &fileInfoData, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
					ATLASSERT(fileInfoData.iIcon >= 0);
					iconIndex = fileInfoData.iIcon;
				}
				if(properties.retrieveSelectedImage) {
					SHGetFileInfoW(pExtension, FILE_ATTRIBUTE_NORMAL, &fileInfoData, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SMALLICON | SHGFI_SYSICONINDEX | SHGFI_OPENICON);
					ATLASSERT(fileInfoData.iIcon >= 0);
					selectedIconIndex = fileInfoData.iIcon;
				}
			}
		} else {
			SHFILEINFO fileInfoData = {0};
			if(properties.retrieveNormalImage) {
				SHGetFileInfo(reinterpret_cast<LPCTSTR>(properties.pIDL), 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
				ATLASSERT(fileInfoData.iIcon >= 0);
				iconIndex = fileInfoData.iIcon;
			}
			if(properties.retrieveSelectedImage) {
				SHGetFileInfo(reinterpret_cast<LPCTSTR>(properties.pIDL), 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_SMALLICON | SHGFI_SYSICONINDEX | SHGFI_OPENICON);
				ATLASSERT(fileInfoData.iIcon >= 0);
				selectedIconIndex = fileInfoData.iIcon;
			}
		}

		if(iconIndex >= 0) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATEICON, reinterpret_cast<WPARAM>(properties.itemHandle), iconIndex);
		}
		if(selectedIconIndex >= 0) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATESELECTEDICON, reinterpret_cast<WPARAM>(properties.itemHandle), selectedIconIndex);
		}
	}

	if(pPath) {
		CoTaskMemFree(pPath);
	}
	return NOERROR;
}