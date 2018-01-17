// ShellListViewItem.cpp: A wrapper for existing shell listview items.

#include "stdafx.h"
#include "ShellListViewItem.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellListViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ShellListViewItem::Properties::~Properties()
{
	if(pOwnerShLvw) {
		pOwnerShLvw->Release();
		pOwnerShLvw = NULL;
	}
}


void ShellListViewItem::Attach(LONG itemID)
{
	properties.itemID = itemID;
}

void ShellListViewItem::Detach(void)
{
	properties.itemID = -1;
}

void ShellListViewItem::SetOwner(ShellListView* pOwner)
{
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->Release();
	}
	properties.pOwnerShLvw = pOwner;
	if(properties.pOwnerShLvw) {
		properties.pOwnerShLvw->AddRef();
	}
}


STDMETHODIMP ShellListViewItem::get_Customizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	*pValue = BOOL2VARIANTBOOL(CanBeCustomized(pIDL));
	return S_OK;
}

STDMETHODIMP ShellListViewItem::get_DefaultDisplayColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
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

STDMETHODIMP ShellListViewItem::get_DefaultSortColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	CComPtr<IShellFolder> pNamespaceISF;
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
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

STDMETHODIMP ShellListViewItem::get_DisplayName(DisplayNameTypeConstants displayNameType/* = dntDisplayName*/, VARIANT_BOOL relativeToDesktop/* = VARIANT_FALSE*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	PCIDLIST_ABSOLUTE pIDL = NULL;
	HRESULT hr = properties.pOwnerShLvw->SHBindToParent(properties.itemID, IID_PPV_ARGS(&pParentISF), &pRelativePIDL, &pIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);

	if(SUCCEEDED(hr)) {
		hr = GetDisplayName(pIDL, pParentISF, pRelativePIDL, displayNameType, relativeToDesktop, pValue);
	}
	return hr;
}

STDMETHODIMP ShellListViewItem::put_DisplayName(DisplayNameTypeConstants displayNameType/* = dntDisplayName*/, VARIANT_BOOL relativeToDesktop/* = VARIANT_FALSE*/, BSTR newValue/* = NULL*/)
{
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShLvw->SHBindToParent(properties.itemID, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
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
		hr = pParentISF->SetNameOf(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pRelativePIDL, OLE2W_EX_DEF(newValue), flags, &pIDLNew);
		ATLASSERT(SUCCEEDED(hr));
	}
	return hr;
}

STDMETHODIMP ShellListViewItem::get_FileAttributes(ItemFileAttributesConstants mask/* = static_cast<ItemFileAttributesConstants>(0x17FB7)*/, ItemFileAttributesConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, ItemFileAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShLvw->SHBindToParent(properties.itemID, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		hr = GetFileAttributes(pParentISF, pRelativePIDL, mask, pValue);
	}
	return hr;
}

STDMETHODIMP ShellListViewItem::put_FileAttributes(ItemFileAttributesConstants mask/* = static_cast<ItemFileAttributesConstants>(0x17FB7)*/, ItemFileAttributesConstants newValue/* = static_cast<ItemFileAttributesConstants>(0)*/)
{
	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShLvw->SHBindToParent(properties.itemID, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
	ATLASSERT(SUCCEEDED(hr));
	ATLASSUME(pParentISF);
	ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

	if(SUCCEEDED(hr)) {
		hr = SetFileAttributes(pParentISF, pRelativePIDL, mask, newValue);
	}
	return hr;
}

STDMETHODIMP ShellListViewItem::get_FullyQualifiedPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE typedPIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	*pValue = *reinterpret_cast<LONG*>(&typedPIDL);
	return (*pValue ? S_OK : E_FAIL);
}

STDMETHODIMP ShellListViewItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemID == -1) {
		return E_FAIL;
	}

	*pValue = properties.itemID;
	return S_OK;
}

STDMETHODIMP ShellListViewItem::get_InfoTipText(InfoTipFlagsConstants flags/* = static_cast<InfoTipFlagsConstants>(itfDefault | itfNoInfoTipFollowSystemSettings | itfAllowSlowInfoTipFollowSysSettings)*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	LPSHELLLISTVIEWITEMDATA pItemDetails = properties.pOwnerShLvw->GetItemDetails(properties.itemID);
	if(pItemDetails) {
		LPWSTR pInfoTip;
		GetInfoTip(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pItemDetails->pIDL, flags, &pInfoTip);
		*pValue = _bstr_t(pInfoTip).Detach();
		CoTaskMemFree(pInfoTip);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewItem::get_LinkTarget(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	TCHAR buffer[MAX_PATH];
	HRESULT hr = GetLinkTarget(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pIDL, buffer, MAX_PATH);
	*pValue = _bstr_t(buffer).Detach();
	return hr;
}

STDMETHODIMP ShellListViewItem::get_LinkTargetPIDL(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	PIDLIST_ABSOLUTE targetPIDL = NULL;
	HRESULT hr = GetLinkTarget(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pIDL, &targetPIDL);
	*pValue = *reinterpret_cast<LONG*>(&targetPIDL);
	return hr;
}

STDMETHODIMP ShellListViewItem::get_ListViewItemObject(IDispatch** ppItem)
{
	ATLASSERT_POINTER(ppItem, LPDISPATCH);
	if(!ppItem) {
		return E_POINTER;
	}

	if(properties.itemID == -1) {
		return E_FAIL;
	}
	HRESULT hr = properties.pOwnerShLvw->properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETITEMBYID, properties.itemID, reinterpret_cast<LPARAM>(ppItem));
	if(hr == DISP_E_BADINDEX) {
		*ppItem = NULL;
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ShellListViewItem::get_ManagedProperties(ShLvwManagedItemPropertiesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwManagedItemPropertiesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	LPSHELLLISTVIEWITEMDATA pItemDetails = properties.pOwnerShLvw->GetItemDetails(properties.itemID);
	if(pItemDetails) {
		*pValue = pItemDetails->managedProperties;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewItem::put_ManagedProperties(ShLvwManagedItemPropertiesConstants newValue)
{
	LPSHELLLISTVIEWITEMDATA pItemDetails = properties.pOwnerShLvw->GetItemDetails(properties.itemID);
	if(pItemDetails) {
		ShLvwManagedItemPropertiesConstants buffer = pItemDetails->managedProperties;
		pItemDetails->managedProperties = newValue;
		HRESULT hr = properties.pOwnerShLvw->ApplyManagedProperties(properties.itemID);
		if(FAILED(hr)) {
			pItemDetails->managedProperties = buffer;
		}
		return hr;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewItem::get_Namespace(IShListViewNamespace** ppNamespace)
{
	ATLASSERT_POINTER(ppNamespace, IShListViewNamespace*);
	if(!ppNamespace) {
		return E_POINTER;
	}

	if(properties.itemID == -1) {
		return E_FAIL;
	}

	CComObject<ShellListViewNamespace>* pNamespaceObj = properties.pOwnerShLvw->ItemIDToNamespace(properties.itemID);
	if(pNamespaceObj) {
		if(SUCCEEDED(pNamespaceObj->QueryInterface(IID_IShListViewNamespace, reinterpret_cast<LPVOID*>(ppNamespace)))) {
			return S_OK;
		}
	} else {
		*ppNamespace = NULL;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListViewItem::get_RequiresElevation(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	*pValue = BOOL2VARIANTBOOL(IsElevationRequired(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pIDL));
	return S_OK;
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellListViewItem::get_SecurityZone(ISecurityZone** ppSecurityZone)
	{
		ATLASSERT_POINTER(ppSecurityZone, ISecurityZone*);
		if(!ppSecurityZone) {
			return E_POINTER;
		}

		ATLASSUME(properties.pOwnerShLvw);
		PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);

		URLZONE zone;
		HRESULT hr = GetZone(properties.pOwnerShLvw->attachedControl, pIDL, &zone);
		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
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

STDMETHODIMP ShellListViewItem::get_ShellAttributes(ItemShellAttributesConstants mask/* = static_cast<ItemShellAttributesConstants>(0xFEFFE17F)*/, VARIANT_BOOL validate/* = VARIANT_FALSE*/, ItemShellAttributesConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, ItemShellAttributesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IShellFolder> pParentISF = NULL;
	PCUITEMID_CHILD pRelativePIDL = NULL;
	HRESULT hr = properties.pOwnerShLvw->SHBindToParent(properties.itemID, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
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

STDMETHODIMP ShellListViewItem::get_SupportsNewFolders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	*pValue = BOOL2VARIANTBOOL(SupportsNewFolders(pIDL));
	return S_OK;
}


STDMETHODIMP ShellListViewItem::CreateShellContextMenu(OLE_HANDLE* pMenu)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = properties.itemID;
	HRESULT hr = properties.pOwnerShLvw->CreateShellContextMenu(v, pMenu);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItem::Customize(void)
{
	PCIDLIST_ABSOLUTE pIDL = properties.pOwnerShLvw->ItemIDToPIDL(properties.itemID);
	return CustomizeFolder(properties.pOwnerShLvw->GethWndShellUIParentWindow(), pIDL);
}

STDMETHODIMP ShellListViewItem::DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = properties.itemID;
	HRESULT hr = properties.pOwnerShLvw->DisplayShellContextMenu(v, x, y);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItem::InvokeDefaultShellContextMenuCommand(void)
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = properties.itemID;
	HRESULT hr = properties.pOwnerShLvw->InvokeDefaultShellContextMenuCommand(v);
	VariantClear(&v);
	return hr;
}

STDMETHODIMP ShellListViewItem::Validate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.pOwnerShLvw->ValidateItem(properties.itemID));
	return S_OK;
}