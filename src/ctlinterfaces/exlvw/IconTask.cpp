// IconTask.cpp: A specialization of RunnableTask for retrieving listview item icon index

#include "stdafx.h"
#include "IconTask.h"


ShLvwIconTask::ShLvwIconTask(void)
    : RunnableTask(FALSE)
{
}

void ShLvwIconTask::FinalRelease()
{
	if(properties.pIDL) {
		ILFree(properties.pIDL);
		properties.pIDL = NULL;
	}
	if(properties.pParentISI) {
		properties.pParentISI->Release();
		properties.pParentISI = NULL;
	}
}


HRESULT ShLvwIconTask::Attach(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI, UseGenericIconsConstants useGenericIcons)
{
	this->properties.hWndToNotify = hWndToNotify;
	this->properties.itemID = itemID;
	this->properties.pIDL = ILCloneFull(pIDL);
	this->properties.pParentISI = pParentISI;
	if(pParentISI) {
		this->properties.pParentISI->AddRef();
	}
	this->properties.useGenericIcons = useGenericIcons;
	return S_OK;
}

HRESULT ShLvwIconTask::CreateInstance(HWND hWndToNotify, PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI, UseGenericIconsConstants useGenericIcons, IRunnableTask** ppTask)
{
	ATLASSERT_POINTER(ppTask, IRunnableTask*);
	if(!ppTask) {
		return E_POINTER;
	}
	if(!pIDL) {
		return E_INVALIDARG;
	}

	*ppTask = NULL;

	CComObject<ShLvwIconTask>* pNewTask = NULL;
	HRESULT hr = CComObject<ShLvwIconTask>::CreateInstance(&pNewTask);
	if(SUCCEEDED(hr)) {
		pNewTask->AddRef();
		hr = pNewTask->Attach(hWndToNotify, pIDL, itemID, pParentISI, useGenericIcons);
		if(SUCCEEDED(hr)) {
			pNewTask->QueryInterface(IID_PPV_ARGS(ppTask));
		}
		pNewTask->Release();
	}
	return hr;
}


STDMETHODIMP ShLvwIconTask::DoRun(void)
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
		if(getGenericIcons) {
			if(!pPath) {
				ATLVERIFY(SUCCEEDED(GetFileSystemPath(NULL, pParentISF, pRelativePIDL, FALSE, &pPath)));
			}
			if(pPath) {
				LPWSTR pExtension = PathFindExtensionW(pPath);
				SHFILEINFOW fileInfoData = {0};
				SHGetFileInfoW(pExtension, FILE_ATTRIBUTE_NORMAL, &fileInfoData, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
				ATLASSERT(fileInfoData.iIcon >= 0);
				iconIndex = fileInfoData.iIcon;
			}
		} else {
			if(properties.pParentISI) {
				// use IShellIcon
				if(FAILED(properties.pParentISI->GetIconOf(ILFindLastID(properties.pIDL), GIL_FORSHELL, &iconIndex))) {
					iconIndex = I_IMAGECALLBACK;
				}
			}

			if(iconIndex == I_IMAGECALLBACK) {
				// either IShellIcon failed or we couldn't use it
				SHFILEINFO fileInfoData = {0};
				SHGetFileInfo(reinterpret_cast<LPCTSTR>(properties.pIDL), 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_SMALLICON | SHGFI_SYSICONINDEX);
				ATLASSERT(fileInfoData.iIcon >= 0);
				iconIndex = fileInfoData.iIcon;
			}
		}

		if(iconIndex >= 0) {
			PostMessage(properties.hWndToNotify, WM_TRIGGER_UPDATEICON, properties.itemID, iconIndex);
		}
	}

	if(pPath) {
		CoTaskMemFree(pPath);
	}
	return NOERROR;
}