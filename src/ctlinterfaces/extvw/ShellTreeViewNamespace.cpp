// ShellTreeViewNamespace.cpp: A wrapper for existing shell namespaces.

#include "stdafx.h"
#include "ShellTreeViewNamespace.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellTreeViewNamespace::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShTreeViewNamespace, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellTreeViewNamespace::Properties::~Properties()
{
	if(pNamespaceEnumSettings) {
		pNamespaceEnumSettings->Release();
		pNamespaceEnumSettings = NULL;
	}
	if(pOwnerShTvw) {
		pOwnerShTvw->Release();
		pOwnerShTvw = NULL;
	}
}


void ShellTreeViewNamespace::Attach(PCIDLIST_ABSOLUTE pIDL, HTREEITEM hRootItem, INamespaceEnumSettings* pEnumSettings)
{
	properties.pIDL = pIDL;
	properties.hRootItem = hRootItem;
	if(pEnumSettings) {
		pEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&properties.pNamespaceEnumSettings));
	} else {
		properties.pNamespaceEnumSettings = NULL;
	}
}

void ShellTreeViewNamespace::Detach(void)
{
	properties.pIDL = NULL;
	properties.hRootItem = NULL;
	if(properties.pNamespaceEnumSettings) {
		properties.pNamespaceEnumSettings->Release();
		properties.pNamespaceEnumSettings = NULL;
	}
}

LRESULT ShellTreeViewNamespace::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir)
{
	ATLASSERT_POINTER(simplePIDLAddedDir, ITEMIDLIST_ABSOLUTE);
	if(ILIsParent(properties.pIDL, simplePIDLAddedDir, TRUE)) {
		ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to contain pIDL 0x%X (count: %i)\n"), properties.pIDL, simplePIDLAddedDir, ILCount(simplePIDLAddedDir));
		CComPtr<INamespaceEnumSettings> pEnumSettings;
		get_NamespaceEnumSettings(&pEnumSettings);
		ATLASSUME(pEnumSettings);
		return properties.pOwnerShTvw->OnSHChangeNotify_CREATE(simplePIDLAddedDir, properties.pIDL, properties.hRootItem, pEnumSettings);
	}
	return 0;
}

LRESULT ShellTreeViewNamespace::OnSHChangeNotify_DELETE(PCIDLIST_ABSOLUTE pIDLRemovedDir)
{
	ATLASSERT_POINTER(pIDLRemovedDir, ITEMIDLIST_ABSOLUTE);
	if(ILIsEqual(properties.pIDL, pIDLRemovedDir) || ILIsParent(pIDLRemovedDir, properties.pIDL, FALSE)) {
		ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be a child of pIDL 0x%X (count: %i)\n"), properties.pIDL, pIDLRemovedDir, ILCount(pIDLRemovedDir));
		properties.pOwnerShTvw->RemoveNamespace(properties.pIDL, TRUE, TRUE);
	}
	return 0;
}

HRESULT ShellTreeViewNamespace::UpdateEnumeration(void)
{
	CComPtr<INamespaceEnumSettings> pEnumSettingsToUse = NULL;
	if(properties.pNamespaceEnumSettings) {
		ATLVERIFY(SUCCEEDED(properties.pNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&pEnumSettingsToUse))));
	} else {
		ATLVERIFY(SUCCEEDED(properties.pOwnerShTvw->get_DefaultNamespaceEnumSettings(&pEnumSettingsToUse)));
	}
	ATLASSUME(pEnumSettingsToUse);

	CComPtr<IShellFolder> pNamespaceISF;
	HRESULT hr = BindToPIDL(properties.pIDL, IID_PPV_ARGS(&pNamespaceISF));
	if(pNamespaceISF) {
		CComPtr<IRunnableTask> pTask;
		#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
			if(RunTimeHelperEx::IsWin8()) {
				hr = ShTvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), properties.hRootItem, NULL, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, &pTask);
			} else {
				hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), properties.hRootItem, NULL, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, &pTask);
			}
		#else
			hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShTvw->attachedControl, &properties.pOwnerShTvw->properties.backgroundEnumResults, &properties.pOwnerShTvw->properties.backgroundEnumResultsCritSection, properties.pOwnerShTvw->GethWndShellUIParentWindow(), properties.hRootItem, NULL, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, &pTask);
		#endif
		if(SUCCEEDED(hr)) {
			INamespaceEnumTask* pEnumTask = NULL;
			pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
			ATLASSUME(pEnumTask);
			if(pEnumTask) {
				CComQIPtr<IShTreeViewNamespace> pNamespace = this;
				ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pNamespace)));
				pEnumTask->Release();
			}
			hr = properties.pOwnerShTvw->EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM);
			ATLASSERT(SUCCEEDED(hr));
		}
	}
	return hr;
}

void ShellTreeViewNamespace::SetOwner(ShellTreeView* pOwner)
{
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->Release();
	}
	properties.pOwnerShTvw = pOwner;
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->AddRef();
	}
}


STDMETHODIMP ShellTreeViewNamespace::get_AutoSortItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoSortItems);
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::put_AutoSortItems(VARIANT_BOOL newValue)
{
	properties.autoSortItems = VARIANTBOOL2BOOL(newValue);
	if(properties.autoSortItems) {
		SortItems(VARIANT_FALSE);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_Customizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(CanBeCustomized(properties.pIDL));
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_DefaultDisplayColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	BindToPIDL(properties.pIDL, IID_PPV_ARGS(&pNamespaceISF));
	if(pNamespaceISF) {
		CComQIPtr<IShellFolder2> pNamespaceISF2 = pNamespaceISF;
		if(pNamespaceISF2) {
			ULONG dummy;
			ATLVERIFY(SUCCEEDED(pNamespaceISF2->GetDefaultColumn(0, &dummy, reinterpret_cast<PULONG>(pValue))));
		}
	}
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_DefaultSortColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	BindToPIDL(properties.pIDL, IID_PPV_ARGS(&pNamespaceISF));
	if(pNamespaceISF) {
		CComQIPtr<IShellFolder2> pNamespaceISF2 = pNamespaceISF;
		if(pNamespaceISF2) {
			ULONG dummy;
			ATLVERIFY(SUCCEEDED(pNamespaceISF2->GetDefaultColumn(0, reinterpret_cast<PULONG>(pValue), &dummy)));
		}
	}
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_FullyQualifiedPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = *reinterpret_cast<LONG*>(&properties.pIDL);
	return (*pValue ? S_OK : E_FAIL);
}

STDMETHODIMP ShellTreeViewNamespace::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pOwnerShTvw->NamespacePIDLToIndex(properties.pIDL, TRUE);
	return (*pValue >= 0 ? S_OK : E_FAIL);
}

STDMETHODIMP ShellTreeViewNamespace::get_Items(IShTreeViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IShTreeViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	CComObject<ShellTreeViewItems>* pShTvwItemsObj = NULL;
	CComObject<ShellTreeViewItems>::CreateInstance(&pShTvwItemsObj);
	pShTvwItemsObj->AddRef();

	pShTvwItemsObj->SetOwner(properties.pOwnerShTvw);
	pShTvwItemsObj->put_FilterType(stfpNamespace, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(*reinterpret_cast<LONG*>(&properties.pIDL)));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;
	pShTvwItemsObj->put_Filter(stfpNamespace, filter);

	pShTvwItemsObj->QueryInterface(IID_IShTreeViewItems, reinterpret_cast<LPVOID*>(ppItems));
	pShTvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings)
{
	ATLASSERT_POINTER(ppEnumSettings, INamespaceEnumSettings*);
	if(!ppEnumSettings) {
		return E_POINTER;
	}

	if(properties.pNamespaceEnumSettings) {
		return properties.pNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(ppEnumSettings));
	} else {
		// return the default enum settings
		return properties.pOwnerShTvw->get_DefaultNamespaceEnumSettings(ppEnumSettings);
	}
}

STDMETHODIMP ShellTreeViewNamespace::get_ParentItemHandle(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hRootItem);
	return S_OK;
}

STDMETHODIMP ShellTreeViewNamespace::get_ParentTreeViewItemObject(IDispatch** ppItem)
{
	ATLASSERT_POINTER(ppItem, LPDISPATCH);
	if(!ppItem) {
		return E_POINTER;
	}

	if(properties.hRootItem) {
		return properties.pOwnerShTvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_GETITEMBYHANDLE, reinterpret_cast<WPARAM>(properties.hRootItem), reinterpret_cast<LPARAM>(ppItem));
	} else {
		*ppItem = NULL;
		return S_OK;
	}
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellTreeViewNamespace::get_SecurityZone(ISecurityZone** ppSecurityZone)
	{
		ATLASSERT_POINTER(ppSecurityZone, ISecurityZone*);
		if(!ppSecurityZone) {
			return E_POINTER;
		}

		URLZONE zone;
		HRESULT hr = GetZone(properties.pOwnerShTvw->attachedControl, properties.pIDL, &zone);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(properties.pOwnerShTvw);
			CComPtr<ISecurityZones> pZones = NULL;
			hr = properties.pOwnerShTvw->get_SecurityZones(&pZones);
			ATLASSERT(SUCCEEDED(hr));
			if(SUCCEEDED(hr)) {
				ATLASSUME(pZones);
				pZones->get_Item(zone, ppSecurityZone);
				hr = S_OK;
			}
		}
		return hr;
	}
#endif

STDMETHODIMP ShellTreeViewNamespace::get_SupportsNewFolders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(SupportsNewFolders(properties.pIDL));
	return S_OK;
}


STDMETHODIMP ShellTreeViewNamespace::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	ATLASSERT_POINTER(pMenu, OLE_HANDLE);
	if(!pMenu) {
		return E_POINTER;
	}

	HMENU hMenu = NULL;
	HRESULT hr = properties.pOwnerShTvw->CreateShellContextMenu(properties.pIDL, hMenu);
	*pMenu = HandleToLong(hMenu);
	return hr;
}

STDMETHODIMP ShellTreeViewNamespace::Customize(void)
{
	return CustomizeFolder(properties.pOwnerShTvw->GethWndShellUIParentWindow(), properties.pIDL);
}

STDMETHODIMP ShellTreeViewNamespace::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	POINT position = {x, y};
	properties.pOwnerShTvw->attachedControl.ClientToScreen(&position);
	return properties.pOwnerShTvw->DisplayShellContextMenu(properties.pIDL, position);
}

STDMETHODIMP ShellTreeViewNamespace::SortItems(VARIANT_BOOL recurse/* = VARIANT_FALSE*/)
{
	if(properties.hRootItem) {
		return properties.pOwnerShTvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, VARIANTBOOL2BOOL(recurse), reinterpret_cast<LPARAM>(properties.hRootItem));
	} else {
		return properties.pOwnerShTvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, VARIANTBOOL2BOOL(recurse), NULL);
	}
}