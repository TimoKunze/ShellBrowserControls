// ShellListViewNamespace.cpp: A wrapper for existing shell namespaces.

#include "stdafx.h"
#include "ShellListViewNamespace.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewNamespace::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewNamespace, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewNamespace::Properties::~Properties()
{
	if(pNamespaceEnumSettings) {
		pNamespaceEnumSettings->Release();
		pNamespaceEnumSettings = NULL;
	}
	if(pThumbnailDiskCache) {
		pThumbnailDiskCache->Close(NULL);
		pThumbnailDiskCache->Release();
		pThumbnailDiskCache = NULL;
	}
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}


void ShellListViewNamespace::Attach(PCIDLIST_ABSOLUTE pIDL, INamespaceEnumSettings* pEnumSettings)
{
	properties.pIDL = pIDL;
	if(pEnumSettings) {
		pEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&properties.pNamespaceEnumSettings));
	} else {
		properties.pNamespaceEnumSettings = NULL;
	}
}

void ShellListViewNamespace::Detach(void)
{
	properties.pIDL = NULL;
	if(properties.pNamespaceEnumSettings) {
		properties.pNamespaceEnumSettings->Release();
		properties.pNamespaceEnumSettings = NULL;
	}
}

LRESULT ShellListViewNamespace::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir)
{
	ATLASSERT_POINTER(simplePIDLAddedDir, ITEMIDLIST_ABSOLUTE);
	if(ILIsParent(properties.pIDL, simplePIDLAddedDir, TRUE)) {
		ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to contain pIDL 0x%X (count: %i)\n"), properties.pIDL, simplePIDLAddedDir, ILCount(simplePIDLAddedDir));
		CComPtr<INamespaceEnumSettings> pEnumSettings;
		get_NamespaceEnumSettings(&pEnumSettings);
		ATLASSUME(pEnumSettings);
		return properties.pOwnerShLvw->OnSHChangeNotify_CREATE(simplePIDLAddedDir, properties.pIDL, pEnumSettings);
	}
	return 0;
}

LRESULT ShellListViewNamespace::OnSHChangeNotify_DELETE(PCIDLIST_ABSOLUTE pIDLRemovedDir)
{
	ATLASSERT_POINTER(pIDLRemovedDir, ITEMIDLIST_ABSOLUTE);
	if(ILIsEqual(properties.pIDL, pIDLRemovedDir) || ILIsParent(pIDLRemovedDir, properties.pIDL, FALSE)) {
		ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be a child of pIDL 0x%X (count: %i)\n"), properties.pIDL, pIDLRemovedDir, ILCount(pIDLRemovedDir));
		properties.pOwnerShLvw->RemoveNamespace(properties.pIDL, TRUE, TRUE);
	}
	return 0;
}

HRESULT ShellListViewNamespace::UpdateEnumeration(void)
{
	CComPtr<INamespaceEnumSettings> pEnumSettingsToUse = NULL;
	if(properties.pNamespaceEnumSettings) {
		ATLVERIFY(SUCCEEDED(properties.pNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&pEnumSettingsToUse))));
	} else {
		ATLVERIFY(SUCCEEDED(properties.pOwnerShLvw->get_DefaultNamespaceEnumSettings(&pEnumSettingsToUse)));
	}
	ATLASSUME(pEnumSettingsToUse);

	CComPtr<IRunnableTask> pTask;
	HRESULT hr = S_OK;
	#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
		if(RunTimeHelperEx::IsWin8()) {
			hr = ShLvwBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), -1, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, FALSE, &pTask);
		} else {
			hr = ShLvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), -1, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, FALSE, &pTask);
		}
	#else
		hr = ShLvwLegacyBackgroundItemEnumTask::CreateInstance(properties.pOwnerShLvw->attachedControl, &properties.pOwnerShLvw->properties.backgroundItemEnumResults, &properties.pOwnerShLvw->properties.backgroundItemEnumResultsCritSection, properties.pOwnerShLvw->GethWndShellUIParentWindow(), -1, properties.pIDL, NULL, NULL, pEnumSettingsToUse, properties.pIDL, TRUE, FALSE, &pTask);
	#endif
	if(SUCCEEDED(hr)) {
		INamespaceEnumTask* pEnumTask = NULL;
		pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
		ATLASSUME(pEnumTask);
		if(pEnumTask) {
			CComQIPtr<IShListViewNamespace> pNamespace = this;
			ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pNamespace)));
			pEnumTask->Release();
		}
		hr = properties.pOwnerShLvw->EnqueueTask(pTask, TASKID_ShLvwBackgroundItemEnum, 0, TASK_PRIORITY_LV_BACKGROUNDENUM);
		ATLASSERT(SUCCEEDED(hr));
	}
	return hr;
}

void ShellListViewNamespace::SetOwner(ShellListView* pOwner)
{
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->Release();
	}
	properties.pOwnerShLvw = pOwner;
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->AddRef();
	}
}


STDMETHODIMP ShellListViewNamespace::get_AutoSortItems(AutoSortItemsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoSortItemsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoSortItems;
	return S_OK;
}

STDMETHODIMP ShellListViewNamespace::put_AutoSortItems(AutoSortItemsConstants newValue)
{
	properties.autoSortItems = newValue;
	if(properties.autoSortItems == asiAutoSort) {
		ATLASSUME(properties.pOwnerShLvw);
		/*LONG sortColumn;
		get_DefaultDisplayColumnIndex(&sortColumn);*/
		properties.pOwnerShLvw->SortItems(-1/*sortColumn*/);
	}
	return S_OK;
}

STDMETHODIMP ShellListViewNamespace::get_Customizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(CanBeCustomized(properties.pIDL));
	return S_OK;
}

STDMETHODIMP ShellListViewNamespace::get_DefaultDisplayColumnIndex(LONG* pValue)
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

STDMETHODIMP ShellListViewNamespace::get_DefaultSortColumnIndex(LONG* pValue)
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

STDMETHODIMP ShellListViewNamespace::get_FullyQualifiedPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = *reinterpret_cast<LONG*>(&properties.pIDL);
	return (*pValue ? S_OK : E_FAIL);
}

STDMETHODIMP ShellListViewNamespace::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pOwnerShLvw->NamespacePIDLToIndex(properties.pIDL, TRUE);
	return (*pValue >= 0 ? S_OK : E_FAIL);
}

STDMETHODIMP ShellListViewNamespace::get_Items(IShListViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IShListViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	CComObject<ShellListViewItems>* pShLvwItemsObj = NULL;
	CComObject<ShellListViewItems>::CreateInstance(&pShLvwItemsObj);
	pShLvwItemsObj->AddRef();

	pShLvwItemsObj->properties.pOwnerShLvw = properties.pOwnerShLvw;
	pShLvwItemsObj->put_FilterType(slfpNamespace, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(*reinterpret_cast<LONG*>(&properties.pIDL)));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;
	pShLvwItemsObj->put_Filter(slfpNamespace, filter);

	pShLvwItemsObj->QueryInterface(IID_IShListViewItems, reinterpret_cast<LPVOID*>(ppItems));
	pShLvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ShellListViewNamespace::get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings)
{
	ATLASSERT_POINTER(ppEnumSettings, INamespaceEnumSettings*);
	if(!ppEnumSettings) {
		return E_POINTER;
	}

	if(properties.pNamespaceEnumSettings) {
		return properties.pNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(ppEnumSettings));
	} else {
		// return the default enum settings
		return properties.pOwnerShLvw->get_DefaultNamespaceEnumSettings(ppEnumSettings);
	}
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellListViewNamespace::get_SecurityZone(ISecurityZone** ppSecurityZone)
	{
		ATLASSERT_POINTER(ppSecurityZone, ISecurityZone*);
		if(!ppSecurityZone) {
			return E_POINTER;
		}

		URLZONE zone;
		// REVIEW: Why attachedControl and not GethWndShellUIParentWindow()?
		HRESULT hr = GetZone(properties.pOwnerShLvw->attachedControl, properties.pIDL, &zone);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(properties.pOwnerShLvw);
			CComPtr<ISecurityZones> pZones = NULL;
			hr = properties.pOwnerShLvw->get_SecurityZones(&pZones);
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

STDMETHODIMP ShellListViewNamespace::get_SupportsNewFolders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(SupportsNewFolders(properties.pIDL));
	return S_OK;
}


STDMETHODIMP ShellListViewNamespace::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	ATLASSERT_POINTER(pMenu, OLE_HANDLE);
	if(!pMenu) {
		return E_POINTER;
	}

	HMENU hMenu = NULL;
	HRESULT hr = properties.pOwnerShLvw->CreateShellContextMenu(properties.pIDL, hMenu);
	*pMenu = HandleToLong(hMenu);
	return hr;
}

STDMETHODIMP ShellListViewNamespace::Customize(void)
{
	return CustomizeFolder(properties.pOwnerShLvw->GethWndShellUIParentWindow(), properties.pIDL);
}

STDMETHODIMP ShellListViewNamespace::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	POINT position = {x, y};
	properties.pOwnerShLvw->attachedControl.ClientToScreen(&position);
	return properties.pOwnerShLvw->DisplayShellContextMenu(properties.pIDL, position);
}