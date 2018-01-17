// ShellTreeViewItem.cpp: A wrapper for existing shell treeview items.

#include "stdafx.h"
#include "ShellTreeViewItem.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellTreeViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShTreeViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellTreeViewItem::Properties::~Properties()
{
	if(pOwnerShTvw) {
		pOwnerShTvw->Release();
		pOwnerShTvw = NULL;
	}
}


void ShellTreeViewItem::Attach(HTREEITEM itemHandle)
{
	properties.itemHandle = itemHandle;
}

void ShellTreeViewItem::Detach(void)
{
	properties.itemHandle = NULL;
}

void ShellTreeViewItem::SetOwner(ShellTreeView* pOwner)
{
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->Release();
	}
	properties.pOwnerShTvw = pOwner;
	if(properties.pOwnerShTvw) {
		properties.pOwnerShTvw->AddRef();
	}
}


STDMETHODIMP ShellTreeViewItem::get_Customizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	*pValue = BOOL2VARIANTBOOL(CanBeCustomized(pIDL));
	return S_OK;
}

STDMETHODIMP ShellTreeViewItem::get_DefaultDisplayColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	HRESULT hr = BindToPIDL(pIDL, IID_PPV_ARGS(&pNamespaceISF));
	if(pNamespaceISF) {
		CComQIPtr<IShellFolder2> pNamespaceISF2 = pNamespaceISF;
		if(pNamespaceISF2) {
			ULONG dummy;
			ATLVERIFY(SUCCEEDED(pNamespaceISF2->GetDefaultColumn(0, &dummy, reinterpret_cast<PULONG>(pValue))));
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_DefaultSortColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	HRESULT hr = BindToPIDL(pIDL, IID_PPV_ARGS(&pNamespaceISF));
	if(pNamespaceISF) {
		CComQIPtr<IShellFolder2> pNamespaceISF2 = pNamespaceISF;
		if(pNamespaceISF2) {
			ULONG dummy;
			ATLVERIFY(SUCCEEDED(pNamespaceISF2->GetDefaultColumn(0, reinterpret_cast<PULONG>(pValue), &dummy)));
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_DisplayName(DisplayNameTypeConstants displayNameType/* = dntDisplayName*/, VARIANT_BOOL relativeToDesktop/* = VARIANT_FALSE*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	PCIDLIST_ABSOLUTE pIDL = NULL;
	HRESULT hr = properties.pOwnerShTvw->SHBindToParent(properties.itemHandle, IID_PPV_ARGS(&pParentISF), &pRelativePIDL, &pIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	if(SUCCEEDED(hr)) {
		hr = GetDisplayName(pIDL, pParentISF, pRelativePIDL, displayNameType, relativeToDesktop, pValue);
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::put_DisplayName(DisplayNameTypeConstants displayNameType/* = dntDisplayName*/, VARIANT_BOOL relativeToDesktop/* = VARIANT_FALSE*/, BSTR newValue/* = NULL*/)
{
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShTvw->SHBindToParent(properties.itemHandle, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		if(!pRelativePIDL) {
			return E_FAIL;
		}
		SHGDNF flags = DisplayNameTypeToSHGDNFlags(displayNameType, relativeToDesktop);

		// TODO: do something useful with the returned relative pIDL
		PITEMID_CHILD pIDLNew = NULL;
		hr = pParentISF->SetNameOf(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pRelativePIDL, OLE2W_EX_DEF(newValue), flags, &pIDLNew);
		ATLASSERT(SUCCEEDED(hr));
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_FileAttributes(ItemFileAttributesConstants mask/* = static_cast<ItemFileAttributesConstants>(0x17FB7)*/, ItemFileAttributesConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShTvw->SHBindToParent(properties.itemHandle, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		hr = GetFileAttributes(pParentISF, pRelativePIDL, mask, pValue);
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::put_FileAttributes(ItemFileAttributesConstants mask/* = static_cast<ItemFileAttributesConstants>(0x17FB7)*/, ItemFileAttributesConstants newValue/* = static_cast<ItemFileAttributesConstants>(0)*/)
{
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShTvw->SHBindToParent(properties.itemHandle, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		hr = SetFileAttributes(pParentISF, pRelativePIDL, mask, newValue);
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_FullyQualifiedPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE typedPIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	*pValue = *reinterpret_cast<LONG*>(&typedPIDL);
	return (*pValue ? S_OK : E_FAIL);
}

STDMETHODIMP ShellTreeViewItem::get_Handle(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!properties.itemHandle) {
		return E_FAIL;
	}

	*pValue = HandleToLong(properties.itemHandle);
	return S_OK;
}

STDMETHODIMP ShellTreeViewItem::get_InfoTipText(InfoTipFlagsConstants flags/* = static_cast<InfoTipFlagsConstants>(itfDefault | itfAllowSlowInfoTipFollowSysSettings)*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	LPSHELLTREEVIEWITEMDATA pItemDetails = properties.pOwnerShTvw->GetItemDetails(properties.itemHandle);
	if(pItemDetails) {
		LPWSTR pInfoTip;
		GetInfoTip(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pItemDetails->pIDL, flags, &pInfoTip);
		*pValue = _bstr_t(pInfoTip).Detach();
		CoTaskMemFree(pInfoTip);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeViewItem::get_LinkTarget(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	TCHAR buffer[MAX_PATH];
	HRESULT hr = GetLinkTarget(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pIDL, buffer, MAX_PATH);
	*pValue = _bstr_t(buffer).Detach();
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_LinkTargetPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	PIDLIST_ABSOLUTE targetPIDL = NULL;
	HRESULT hr = GetLinkTarget(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pIDL, &targetPIDL);
	*pValue = *reinterpret_cast<LONG*>(&targetPIDL);
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_ManagedProperties(ShTvwManagedItemPropertiesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShTvwManagedItemPropertiesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	LPSHELLTREEVIEWITEMDATA pItemDetails = properties.pOwnerShTvw->GetItemDetails(properties.itemHandle);
	if(pItemDetails) {
		*pValue = pItemDetails->managedProperties;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeViewItem::put_ManagedProperties(ShTvwManagedItemPropertiesConstants newValue)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = properties.pOwnerShTvw->GetItemDetails(properties.itemHandle);
	if(pItemDetails) {
		ShTvwManagedItemPropertiesConstants buffer = pItemDetails->managedProperties;
		pItemDetails->managedProperties = newValue;
		HRESULT hr = properties.pOwnerShTvw->ApplyManagedProperties(properties.itemHandle);
		if(FAILED(hr)) {
			pItemDetails->managedProperties = buffer;
		}
		return hr;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeViewItem::get_Namespace(IShTreeViewNamespace** ppNamespace)
{
	ATLASSERT_POINTER(ppNamespace, IShTreeViewNamespace*);
	if(!ppNamespace) {
		return E_POINTER;
	}

	if(!properties.itemHandle) {
		return E_FAIL;
	}

	CComObject<ShellTreeViewNamespace>* pNamespaceObj = properties.pOwnerShTvw->ItemHandleToNamespace(properties.itemHandle);
	if(pNamespaceObj) {
		if(SUCCEEDED(pNamespaceObj->QueryInterface(IID_IShTreeViewNamespace, reinterpret_cast<LPVOID*>(ppNamespace)))) {
			return S_OK;
		}
	} else {
		*ppNamespace = NULL;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeViewItem::get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings)
{
	ATLASSERT_POINTER(ppEnumSettings, INamespaceEnumSettings*);
	if(!ppEnumSettings) {
		return E_POINTER;
	}

	if(!properties.itemHandle) {
		return E_FAIL;
	}

	*ppEnumSettings = properties.pOwnerShTvw->ItemHandleToNameSpaceEnumSettings(properties.itemHandle);
	return S_OK;
}

STDMETHODIMP ShellTreeViewItem::get_RequiresElevation(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	*pValue = BOOL2VARIANTBOOL(IsElevationRequired(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pIDL));
	return S_OK;
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellTreeViewItem::get_SecurityZone(ISecurityZone** ppSecurityZone)
	{
		ATLASSERT_POINTER(ppSecurityZone, ISecurityZone*);
		if(!ppSecurityZone) {
			return E_POINTER;
		}

		ATLASSUME(properties.pOwnerShTvw);
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);

		URLZONE zone;
		HRESULT hr = GetZone(properties.pOwnerShTvw->attachedControl, pIDL, &zone);
		if(SUCCEEDED(hr)) {
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

STDMETHODIMP ShellTreeViewItem::get_ShellAttributes(ItemShellAttributesConstants mask/* = static_cast<ItemShellAttributesConstants>(0xFEFFE17F)*/, VARIANT_BOOL validate/* = VARIANT_FALSE*/, ItemShellAttributesConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShTvw->SHBindToParent(properties.itemHandle, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		SFGAOF attributes = mask;
		if(validate != VARIANT_FALSE) {
			attributes |= SFGAO_VALIDATE;
		}
		hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
		ATLASSERT(SUCCEEDED(hr));

		if(SUCCEEDED(hr)) {
			*pValue = static_cast<ItemShellAttributesConstants>(attributes);
		}
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::get_SupportsNewFolders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	*pValue = BOOL2VARIANTBOOL(SupportsNewFolders(pIDL));
	return S_OK;
}

STDMETHODIMP ShellTreeViewItem::get_TreeViewItemObject(IDispatch** ppItem)
{
	ATLASSERT_POINTER(ppItem, LPDISPATCH);
	if(!ppItem) {
		return E_POINTER;
	}

	if(!properties.itemHandle) {
		return E_FAIL;
	}
	HRESULT hr = properties.pOwnerShTvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_GETITEMBYHANDLE, reinterpret_cast<WPARAM>(properties.itemHandle), reinterpret_cast<LPARAM>(ppItem));
	if(hr == E_INVALIDARG) {
		*ppItem = NULL;
		hr = S_OK;
	}
	return hr;
}


STDMETHODIMP ShellTreeViewItem::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = HandleToLong(properties.itemHandle);
	HRESULT hr = properties.pOwnerShTvw->CreateShellContextMenu(v, pMenu);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItem::Customize(void)
{
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShTvw->ItemHandleToPIDL(properties.itemHandle);
	return CustomizeFolder(properties.pOwnerShTvw->GethWndShellUIParentWindow(), pIDL);
}

STDMETHODIMP ShellTreeViewItem::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = HandleToLong(properties.itemHandle);
	HRESULT hr = properties.pOwnerShTvw->DisplayShellContextMenu(v, x, y);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItem::EnsureSubItemsAreLoaded(VARIANT_BOOL waitForLastItem/* = VARIANT_TRUE*/)
{
	ShTvwManagedItemPropertiesConstants managedProperties = static_cast<ShTvwManagedItemPropertiesConstants>(0);
	HRESULT hr = get_ManagedProperties(&managedProperties);
	if(SUCCEEDED(hr)) {
		if(managedProperties & stmipSubItems) {
			// we're managing this item's sub-items
			TVITEMEX item = {0};
			item.hItem = properties.itemHandle;
			item.mask = TVIF_HANDLE | TVIF_STATE;
			if(properties.pOwnerShTvw->attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				if((item.state & TVIS_EXPANDEDONCE) != TVIS_EXPANDEDONCE) {
					// it has not yet been expanded, so insert the sub-items
					BOOL isShellItem;
					properties.pOwnerShTvw->OnFirstTimeItemExpanding(properties.itemHandle, (waitForLastItem == VARIANT_FALSE ? tmTimeOutThreading : tmNoThreading), FALSE, isShellItem);
					if(isShellItem) {
						// ensure the items aren't inserted twice
						item.state = TVIS_EXPANDEDONCE;
						item.stateMask = TVIS_EXPANDEDONCE;
						properties.pOwnerShTvw->attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
					}
				}
			} else {
				hr = E_FAIL;
			}
		}
	}
	return hr;
}

STDMETHODIMP ShellTreeViewItem::InvokeDefaultShellContextMenuCommand(void)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = HandleToLong(properties.itemHandle);
	HRESULT hr = properties.pOwnerShTvw->InvokeDefaultShellContextMenuCommand(v);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellTreeViewItem::Validate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.pOwnerShTvw->ValidateItem(properties.itemHandle));
	return S_OK;
}