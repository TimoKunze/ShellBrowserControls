// ShellTreeView.cpp: Makes an ExplorerTreeView a shell treeview.

#include "stdafx.h"
#include "ShellTreeView.h"
#include "../../ClassFactory.h"


ShellTreeView::ShellTreeView()
{
	SIZEL size = {36, 36};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	HICON hIDEIcon = NULL;
	HRSRC hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_IDELOGOSHTVW), RT_GROUP_ICON);
	if(hResource) {
		HGLOBAL hMem = LoadResource(ModuleHelper::GetResourceInstance(), hResource);
		if(hMem) {
			#pragma pack(push)
			#pragma pack(2)

			typedef struct
			{
				BYTE width;
				BYTE height;
				BYTE colorCount;
				BYTE reserved;
				WORD planes;
				WORD bitCount;
				DWORD sizeInBytes;
				WORD ID;
			} GRPICONDIRENTRY, *LPGRPICONDIRENTRY;

			typedef struct
			{
				WORD reserved;
				WORD type;
				WORD count;
				GRPICONDIRENTRY entries[1];
			} GRPICONDIR, *LPGRPICONDIR;

			#pragma pack(pop)

			LPGRPICONDIR pIconGroupData = reinterpret_cast<LPGRPICONDIR>(LockResource(hMem));
			if(pIconGroupData) {
				for(int i = 0; i < pIconGroupData->count; ++i) {
					if(RunTimeHelper::IsCommCtrl6()) {
						// use the 32bpp icon
						if(pIconGroupData->entries[i].bitCount == 32) {
							hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(pIconGroupData->entries[i].ID), RT_ICON);
							break;
						}
					} else {
						// use the 8bpp icon
						if(pIconGroupData->entries[i].bitCount == 8) {
							hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(pIconGroupData->entries[i].ID), RT_ICON);
							break;
						}
					}
				}

				if(hResource) {
					hMem = LoadResource(ModuleHelper::GetResourceInstance(), hResource);
					if(hMem) {
						LPVOID pIconData = LockResource(hMem);
						if(pIconData) {
							hIDEIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(ModuleHelper::GetResourceInstance(), hResource), TRUE, 0x00030000, 32, 32, LR_DEFAULTCOLOR | LR_CREATEDIBSECTION);
						}
					}
				}
			}
		}
	}

	if(hIDEIcon) {
		if(RunTimeHelper::IsCommCtrl6()) {
			imglstIDEIcon.Create(32, 32, ILC_COLOR32 | ILC_MASK, 1, 0);
		} else {
			imglstIDEIcon.Create(32, 32, ILC_COLOR24 | ILC_MASK, 1, 0);
		}
		imglstIDEIcon.SetBkColor(CLR_NONE);
		imglstIDEIcon.AddIcon(hIDEIcon);
		DestroyIcon(hIDEIcon);
	}

	pThreadObject = NULL;
}

ShellTreeView::~ShellTreeView()
{
	ATLASSERT((attachedControl.m_hWnd == NULL) && "Did you call Detach()?");
	ATLASSERT((properties.pAttachedCOMControl == NULL) && "Did you call Detach()?");
	#ifdef USE_STL
		ATLASSERT((properties.items.size() == 0) && "Did you call Detach()?");
		ATLASSERT((properties.namespaces.size() == 0) && "Did you call Detach()?");
	#else
		ATLASSERT((properties.items.GetCount() == 0) && "Did you call Detach()?");
		ATLASSERT((properties.namespaces.GetCount() == 0) && "Did you call Detach()?");
	#endif
	if(!imglstIDEIcon.IsNull()) {
		imglstIDEIcon.Destroy();
	}
}

HRESULT ShellTreeView::FinalConstruct()
{
	if(properties.pDefaultNamespaceEnumSettings) {
		properties.pDefaultNamespaceEnumSettings->Release();
		properties.pDefaultNamespaceEnumSettings = NULL;
	}
	CComObject<NamespaceEnumSettings>* pDefaultNamespaceEnumSettingsObj = NULL;
	CComObject<NamespaceEnumSettings>::CreateInstance(&pDefaultNamespaceEnumSettingsObj);
	pDefaultNamespaceEnumSettingsObj->AddRef();
	DWORD flags = snseIncludeFolders | snseMayIncludeHiddenItems;
	if(RunTimeHelper::IsWin7()) {
		// TODO: Maybe this is available on Vista, too?
		flags |= snseEnumForNavigationPane;
	}
	pDefaultNamespaceEnumSettingsObj->put_EnumerationFlags(static_cast<ShNamespaceEnumerationConstants>(flags));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_IncludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	/* We initialize this to isaIsFolder because some spooky namespaces (e. g. Network Connections in the
	   Control Panel) enumerate file items although the enumerator wasn't created with the
	   snseIncludeNonFolders flag. */
	pDefaultNamespaceEnumSettingsObj->put_IncludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(isaIsFolder));
	pDefaultNamespaceEnumSettingsObj->put_IncludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_IncludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&properties.pDefaultNamespaceEnumSettings));
	pDefaultNamespaceEnumSettingsObj->Release();

	if(IsInDesignMode()) {
		IPersistentChildObject* pPersistentChildObject = NULL;
		properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistentChildObject, reinterpret_cast<LPVOID*>(&pPersistentChildObject));
		ATLASSUME(pPersistentChildObject);
		if(pPersistentChildObject) {
			pPersistentChildObject->SetOwnerControl(this);
			pPersistentChildObject->Release();
		}
	}

	return S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ShellTreeView::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShTreeView, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ShellTreeView::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	ATLASSUME(pPropertyBag);
	if(pPropertyBag == NULL) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	if(SUCCEEDED(pPropertyBag->Read(OLESTR("Version"), &propertyValue, pErrorLog))) {
		if(propertyValue.lVal > 0x0100) {
			return E_FAIL;
		}
	}

	propertyValue.vt = VT_BOOL;
	HRESULT hr = pPropertyBag->Read(OLESTR("AutoEditNewItems"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL autoEditNewItems = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("ColorCompressedItems"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL colorCompressedItems = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("ColorEncryptedItems"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL colorEncryptedItems = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("DefaultManagedItemProperties"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ShTvwManagedItemPropertiesConstants defaultManagedItemProperties = (SUCCEEDED(hr) ? static_cast<ShTvwManagedItemPropertiesConstants>(propertyValue.lVal) : stmipAll);

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	hr = pPropertyBag->Read(OLESTR("DefaultNamespaceEnumSettings"), &propertyValue, pErrorLog);
	CComPtr<INamespaceEnumSettings> pDefaultNamespaceEnumSettings = NULL;
	if(propertyValue.pdispVal) {
		propertyValue.pdispVal->QueryInterface(IID_PPV_ARGS(&pDefaultNamespaceEnumSettings));
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("DisabledEvents"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = (SUCCEEDED(hr) ? static_cast<DisabledEventsConstants>(propertyValue.lVal) : static_cast<DisabledEventsConstants>(deNamespacePIDLChangeEvents | deNamespaceInsertionEvents | deNamespaceDeletionEvents | deItemPIDLChangeEvents | deItemInsertionEvents | deItemDeletionEvents));
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("DisplayElevationShieldOverlays"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL displayElevationShieldOverlays = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("HandleOLEDragDrop"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	HandleOLEDragDropConstants handleOLEDragDrop = (SUCCEEDED(hr) ? static_cast<HandleOLEDragDropConstants>(propertyValue.lVal) : static_cast<HandleOLEDragDropConstants>(hoddSourcePart | hoddTargetPart | hoddTargetPartWithDropHilite));
	hr = pPropertyBag->Read(OLESTR("HiddenItemsStyle"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	HiddenItemsStyleConstants hiddenItemsStyle = (SUCCEEDED(hr) ? static_cast<HiddenItemsStyleConstants>(propertyValue.lVal) : hisGhostedOnDemand);
	hr = pPropertyBag->Read(OLESTR("InfoTipFlags"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	InfoTipFlagsConstants infoTipFlags = (SUCCEEDED(hr) ? static_cast<InfoTipFlagsConstants>(propertyValue.lVal) : itfNoInfoTip);
	hr = pPropertyBag->Read(OLESTR("ItemEnumerationTimeout"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	LONG itemEnumerationTimeout = (SUCCEEDED(hr) ? propertyValue.lVal : 3000);
	hr = pPropertyBag->Read(OLESTR("ItemTypeSortOrder"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ItemTypeSortOrderConstants itemTypeSortOrder = (SUCCEEDED(hr) ? static_cast<ItemTypeSortOrderConstants>(propertyValue.lVal) : itsoShellItemsFirst);
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("LimitLabelEditInput"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL limitLabelEditInput = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("LoadOverlaysOnDemand"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL loadOverlaysOnDemand = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("PreselectBasenameOnLabelEdit"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL preselectBasenameOnLabelEdit = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("ProcessShellNotifications"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL processShellNotifications = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("UseGenericIcons"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	UseGenericIconsConstants useGenericIcons = (SUCCEEDED(hr) ? static_cast<UseGenericIconsConstants>(propertyValue.lVal) : ugiOnlyForSlowItems);
	hr = pPropertyBag->Read(OLESTR("UseSystemImageList"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	UseSystemImageListConstants useSystemImageList = (SUCCEEDED(hr) ? static_cast<UseSystemImageListConstants>(propertyValue.lVal) : usilSmallImageList);

	hr = put_AutoEditNewItems(autoEditNewItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColorCompressedItems(colorCompressedItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColorEncryptedItems(colorEncryptedItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DefaultManagedItemProperties(defaultManagedItemProperties);
	if(FAILED(hr)) {
		return hr;
	}
	if(pDefaultNamespaceEnumSettings) {
		hr = putref_DefaultNamespaceEnumSettings(pDefaultNamespaceEnumSettings);
		if(FAILED(hr)) {
			return hr;
		}
	} else {
		properties.pDefaultNamespaceEnumSettings->put_EnumerationFlags(static_cast<ShNamespaceEnumerationConstants>(snseIncludeFolders | snseMayIncludeHiddenItems));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_IncludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		/* We initialize this to isaIsFolder because some spooky namespaces (e. g. Network Connections in the
			 Control Panel) enumerate file items although the enumerator wasn't created with the
			 snseIncludeNonFolders flag. */
		properties.pDefaultNamespaceEnumSettings->put_IncludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(isaIsFolder));
		properties.pDefaultNamespaceEnumSettings->put_IncludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_IncludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayElevationShieldOverlays(displayElevationShieldOverlays);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HandleOLEDragDrop(handleOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HiddenItemsStyle(hiddenItemsStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InfoTipFlags(infoTipFlags);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemEnumerationTimeout(itemEnumerationTimeout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemTypeSortOrder(itemTypeSortOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LimitLabelEditInput(limitLabelEditInput);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LoadOverlaysOnDemand(loadOverlaysOnDemand);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_PreselectBasenameOnLabelEdit(preselectBasenameOnLabelEdit);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessShellNotifications(processShellNotifications);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseGenericIcons(useGenericIcons);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemImageList(useSystemImageList);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ShellTreeView::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/)
{
	ATLASSUME(pPropertyBag);
	if(pPropertyBag == NULL) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	HRESULT hr = S_OK;
	propertyValue.vt = VT_I4;
	propertyValue.lVal = 0x0100;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Version"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoEditNewItems);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("AutoEditNewItems"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.colorCompressedItems);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ColorCompressedItems"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.colorEncryptedItems);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ColorEncryptedItems"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.defaultManagedItemProperties;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DefaultManagedItemProperties"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DefaultNamespaceEnumSettings"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisabledEvents"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayElevationShieldOverlays);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisplayElevationShieldOverlays"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.handleOLEDragDrop;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("HandleOLEDragDrop"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hiddenItemsStyle;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("HiddenItemsStyle"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.infoTipFlags;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("InfoTipFlags"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemEnumerationTimeout;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ItemEnumerationTimeout"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemTypeSortOrder;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ItemTypeSortOrder"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.limitLabelEditInput);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("LimitLabelEditInput"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.loadOverlaysOnDemand);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("LoadOverlaysOnDemand"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.preselectBasenameOnLabelEdit);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("PreselectBasenameOnLabelEdit"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processShellNotifications);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ProcessShellNotifications"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.useGenericIcons;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseGenericIcons"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.useSystemImageList;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseSystemImageList"), &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}

	return S_OK;
}

STDMETHODIMP ShellTreeView::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(pSize == NULL) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 9 VT_I4 properties...
	pSize->QuadPart += 9 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 8 VT_BOOL properties...
	pSize->QuadPart += 8 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

	// ...and 1 VT_DISPATCH property
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(FAILED(properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP ShellTreeView::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(pStream == NULL) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x06060606/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0101) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x02020202/*4x 0x02 (-> ShellTreeView)*/) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoEditNewItems = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL colorCompressedItems = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL colorEncryptedItems = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ShTvwManagedItemPropertiesConstants defaultManagedItemProperties = static_cast<ShTvwManagedItemPropertiesConstants>(propertyValue.lVal);

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CLSID clsid = CLSID_NULL;
	if(FAILED(hr = ReadClassStm(pStream, &clsid))) {
		return hr;
	}
	CComPtr<INamespaceEnumSettings> pDefaultNamespaceEnumSettings = NULL;
	CComObject<NamespaceEnumSettings>* pDefaultNamespaceEnumSettingsObj = NULL;
	CComObject<NamespaceEnumSettings>::CreateInstance(&pDefaultNamespaceEnumSettingsObj);
	pDefaultNamespaceEnumSettingsObj->AddRef();
	pDefaultNamespaceEnumSettingsObj->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&pDefaultNamespaceEnumSettings));
	pDefaultNamespaceEnumSettingsObj->Release();
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(FAILED(hr = pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		if(FAILED(hr = pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			return hr;
		}
	}
	if(pStreamInit) {
		/*CLSID storedCLSID;
		if(FAILED(hr = pStreamInit->GetClassID(&storedCLSID))) {
			return hr;
		}*/
		// {CC889E2B-5A0D-42f0-AA08-D5FD5863410C}
		CLSID CLSID_NamespaceEnumSettingsW = {0xCC889E2B, 0x5A0D, 0x42f0, {0xAA, 0x08, 0xD5, 0xFD, 0x58, 0x63, 0x41, 0x0C}};
		// {E21D9123-B1FE-4815-BBF5-1CBE1ED47AC4}
		CLSID CLSID_NamespaceEnumSettingsA = {0xE21D9123, 0xB1FE, 0x4815, {0xBB, 0xF5, 0x1C, 0xBE, 0x1E, 0xD4, 0x7A, 0xC4}};
		if((clsid == CLSID_NamespaceEnumSettingsW) || (clsid == CLSID_NamespaceEnumSettingsA)) {
			if(FAILED(hr = pStreamInit->Load(pStream))) {
				return hr;
			}
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayElevationShieldOverlays = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HandleOLEDragDropConstants handleOLEDragDrop = static_cast<HandleOLEDragDropConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HiddenItemsStyleConstants hiddenItemsStyle = static_cast<HiddenItemsStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	InfoTipFlagsConstants infoTipFlags = static_cast<InfoTipFlagsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG itemEnumerationTimeout = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemTypeSortOrderConstants itemTypeSortOrder = static_cast<ItemTypeSortOrderConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL limitLabelEditInput = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL loadOverlaysOnDemand = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL preselectBasenameOnLabelEdit = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processShellNotifications = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	UseSystemImageListConstants useSystemImageList = static_cast<UseSystemImageListConstants>(propertyValue.lVal);
	UseGenericIconsConstants useGenericIcons = ugiOnlyForSlowItems;
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		useGenericIcons = static_cast<UseGenericIconsConstants>(propertyValue.lVal);
	}

	hr = put_AutoEditNewItems(autoEditNewItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColorCompressedItems(colorCompressedItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColorEncryptedItems(colorEncryptedItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DefaultManagedItemProperties(defaultManagedItemProperties);
	if(FAILED(hr)) {
		return hr;
	}
	hr = putref_DefaultNamespaceEnumSettings(pDefaultNamespaceEnumSettings);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayElevationShieldOverlays(displayElevationShieldOverlays);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HandleOLEDragDrop(handleOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HiddenItemsStyle(hiddenItemsStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InfoTipFlags(infoTipFlags);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemEnumerationTimeout(itemEnumerationTimeout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemTypeSortOrder(itemTypeSortOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LimitLabelEditInput(limitLabelEditInput);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LoadOverlaysOnDemand(loadOverlaysOnDemand);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_PreselectBasenameOnLabelEdit(preselectBasenameOnLabelEdit);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessShellNotifications(processShellNotifications);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseGenericIcons(useGenericIcons);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemImageList(useSystemImageList);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ShellTreeView::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(pStream == NULL) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x06060606/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0101;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x02020202/*4x 0x02 (-> ShellTreeView)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoEditNewItems);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.colorCompressedItems);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.colorEncryptedItems);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.defaultManagedItemProperties;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStreamInit> pPersistInit = NULL;
	if(FAILED(hr = properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
		if(FAILED(hr = properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistInit) {
		CLSID clsid = CLSID_NULL;
		if(FAILED(hr = pPersistInit->GetClassID(&clsid))) {
			return hr;
		}
		if(FAILED(hr = WriteClassStm(pStream, clsid))) {
			return hr;
		}
		if(FAILED(hr = pPersistInit->Save(pStream, clearDirtyFlag))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayElevationShieldOverlays);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.handleOLEDragDrop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hiddenItemsStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.infoTipFlags;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemEnumerationTimeout;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemTypeSortOrder;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.limitLabelEditInput);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.loadOverlaysOnDemand);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.preselectBasenameOnLabelEdit);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processShellNotifications);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.useSystemImageList;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0101 starts here
	propertyValue.lVal = properties.useGenericIcons;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


ShellTreeView::Properties::Properties()
{
	VariantInit(&emptyVariant);
	VariantClear(&emptyVariant);

	ZeroMemory(&itemDataForFastInsertion, sizeof(itemDataForFastInsertion));
	itemDataForFastInsertion.structSize = sizeof(EXTVMADDITEMDATA);
	itemDataForFastInsertion.iconIndex = -1;
	itemDataForFastInsertion.selectedIconIndex = -1;
	itemDataForFastInsertion.expandedIconIndex = -2;
	itemDataForFastInsertion.heightIncrement = 1;

	hWndShellUIParentWindow = NULL;
	InitializeCriticalSectionAndSpinCount(&backgroundEnumResultsCritSection, 400);
	pDefaultNamespaceEnumSettings = NULL;

	pNamespaces = NULL;
	CComObject<ShellTreeViewNamespaces>::CreateInstance(&pNamespaces);
	ATLASSUME(pNamespaces);
	pNamespaces->AddRef();

	pTreeItems = NULL;
	CComObject<ShellTreeViewItems>::CreateInstance(&pTreeItems);
	ATLASSUME(pTreeItems);
	pTreeItems->AddRef();

	ResetToDefaults();

	pAttachedCOMControl = NULL;
	pAttachedInternalMessageListener = NULL;

	SHGetDesktopFolder(&pDesktopISF);
	ATLASSUME(pDesktopISF);

	SHFILEINFO fileInfo = {0};
	SHGetFileInfo(TEXT(".zyxwv12"), FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
	DEFAULTICON_FILE = fileInfo.iIcon;
	SHGetFileInfo(TEXT("folder"), FILE_ATTRIBUTE_DIRECTORY, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
	DEFAULTICON_FOLDER = fileInfo.iIcon;
}

ShellTreeView::Properties::~Properties()
{
	if(pEnumTaskScheduler) {
		ATLASSERT(pEnumTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(pNormalTaskScheduler) {
		ATLASSERT(pNormalTaskScheduler->CountTasks(TOID_NULL) == 0);
	}

	if(pDesktopISF) {
		pDesktopISF->Release();
		pDesktopISF = NULL;
	}

	if(pDefaultNamespaceEnumSettings) {
		ATLVERIFY(pDefaultNamespaceEnumSettings->Release() == 0);
		pDefaultNamespaceEnumSettings = NULL;
	}

	if(pNamespaces) {
		pNamespaces->Release();
		pNamespaces = NULL;
	}
	if(pTreeItems) {
		pTreeItems->Release();
		pTreeItems = NULL;
	}

	#ifdef USE_STL
		while(!shellItemDataOfInsertedItems.empty()) {
			delete shellItemDataOfInsertedItems.top();
			shellItemDataOfInsertedItems.pop();
		}
		EnterCriticalSection(&backgroundEnumResultsCritSection);
		while(!backgroundEnumResults.empty()) {
			LPSHTVWBACKGROUNDITEMENUMINFO pEnumResult = backgroundEnumResults.front();
			backgroundEnumResults.pop();
			if(pEnumResult) {
				if(pEnumResult->hPIDLBuffer) {
					DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAEnumeratedItemElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundEnumResultsCritSection);
	#else
		while(!shellItemDataOfInsertedItems.IsEmpty()) {
			delete shellItemDataOfInsertedItems.RemoveHead();
		}
		EnterCriticalSection(&backgroundEnumResultsCritSection);
		while(!backgroundEnumResults.IsEmpty()) {
			LPSHTVWBACKGROUNDITEMENUMINFO pEnumResult = backgroundEnumResults.RemoveHead();
			if(pEnumResult) {
				if(pEnumResult->hPIDLBuffer) {
					DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAEnumeratedItemElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundEnumResultsCritSection);
	#endif

	DeleteCriticalSection(&backgroundEnumResultsCritSection);
}


HRESULT ShellTreeView::OnDraw(ATL_DRAWINFO& drawInfo)
{
	SIZEL iconPosition = {drawInfo.prcBounds->left, drawInfo.prcBounds->top};
	if(drawInfo.bRectInHimetric) {
		SIZEL iconPositionHiMetric = iconPosition;
		AtlHiMetricToPixel(&iconPositionHiMetric, &iconPosition);
	}
	RECT rc = {drawInfo.prcBounds->left, drawInfo.prcBounds->top, drawInfo.prcBounds->right, drawInfo.prcBounds->bottom};
	FillRect(drawInfo.hdcDraw, &rc, GetSysColorBrush(COLOR_BTNFACE));
	imglstIDEIcon.Draw(drawInfo.hdcDraw, 0, iconPosition.cx + 2, iconPosition.cy + 2, ILD_NORMAL);

	SetTextColor(drawInfo.hdcDraw, RGB(0, 0, 255));
	SetBkMode(drawInfo.hdcDraw, TRANSPARENT);
	RECT textRectangle = {iconPosition.cx, iconPosition.cy, iconPosition.cx + 36, iconPosition.cy + 36};
	#ifdef UNICODE
		LPCTSTR pText = TEXT("U");
	#else
		LPCTSTR pText = TEXT("A");
	#endif
	DrawText(drawInfo.hdcDraw, pText, 1, &textRectangle, DT_LEFT | DT_TOP);
	return S_OK;
}

STDMETHODIMP ShellTreeView::IOleObject_SetExtent(DWORD /*drawAspect*/, SIZEL* /*pSize*/)
{
	return E_FAIL;
}

HWND ShellTreeView::GethWndShellUIParentWindow(void)
{
	return properties.hWndShellUIParentWindow;
}

//////////////////////////////////////////////////////////////////////
// implementation of IContextMenuCB
STDMETHODIMP ShellTreeView::CallBack(LPSHELLFOLDER pShellFolder, HWND hWnd, LPDATAOBJECT pDataObject, UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: Check whether we've to handle e. g. DFM_WM_DRAWITEM to make owner-drawn menus work.
	//ATLASSERT(message != DFM_WM_MEASUREITEM);
	//ATLASSERT(message != DFM_WM_DRAWITEM);
	//ATLASSERT(message != DFM_WM_INITMENUPOPUP);
	//ATLASSERT(message != DFM_INVOKECOMMANDEX);
	ATLTRACE2(shtvwTraceContextMenus, 3, TEXT("Callback was called with pShellFolder=0x%X, hWnd=0x%X, pDataObject=0x%X, message=0x%X, wParam=0x%X, lParam=0x%X\n"), pShellFolder, hWnd, pDataObject, message, wParam, lParam);
	return ContextMenuCallback(pShellFolder, hWnd, pDataObject, message, wParam, lParam);
}
// implementation of IContextMenuCB
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ShellTreeView::GetCategoryName(PROPCAT /*category*/, LCID /*languageID*/, BSTR* /*pName*/)
{
	return E_FAIL;
}

STDMETHODIMP ShellTreeView::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_SHTVW_COLORCOMPRESSEDITEMS:
		case DISPID_SHTVW_COLORENCRYPTEDITEMS:
		case DISPID_SHTVW_HWNDSHELLUIPARENTWINDOW:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_SHTVW_AUTOEDITNEWITEMS:
		case DISPID_SHTVW_DEFAULTMANAGEDITEMPROPERTIES:
		case DISPID_SHTVW_DEFAULTNAMESPACEENUMSETTINGS:
		case DISPID_SHTVW_DISABLEDEVENTS:
		case DISPID_SHTVW_DISPLAYELEVATIONSHIELDOVERLAYS:
		case DISPID_SHTVW_HANDLEOLEDRAGDROP:
		case DISPID_SHTVW_HIDDENITEMSSTYLE:
		case DISPID_SHTVW_INFOTIPFLAGS:
		case DISPID_SHTVW_ITEMENUMERATIONTIMEOUT:
		case DISPID_SHTVW_ITEMTYPESORTORDER:
		case DISPID_SHTVW_LIMITLABELEDITINPUT:
		case DISPID_SHTVW_LOADOVERLAYSONDEMAND:
		case DISPID_SHTVW_PRESELECTBASENAMEONLABELEDIT:
		case DISPID_SHTVW_PROCESSSHELLNOTIFICATIONS:
		case DISPID_SHTVW_USEGENERICICONS:
		case DISPID_SHTVW_USESYSTEMIMAGELIST:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_SHTVW_APPID:
		case DISPID_SHTVW_APPNAME:
		case DISPID_SHTVW_APPSHORTNAME:
		case DISPID_SHTVW_BUILD:
		case DISPID_SHTVW_CHARSET:
		case DISPID_SHTVW_HIMAGELIST:
		case DISPID_SHTVW_HWND:
		case DISPID_SHTVW_ISRELEASE:
		case DISPID_SHTVW_PROGRAMMER:
		case DISPID_SHTVW_TESTER:
		case DISPID_SHTVW_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_SHTVW_NAMESPACES:
		case DISPID_SHTVW_TREEITEMS:
		#ifdef ACTIVATE_SECZONE_FEATURES
			case DISPID_SHTVW_SECURITYZONES:
		#endif
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ShellTreeView::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ShellTreeView::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ShellTreeView::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ShellBrowserControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ShellTreeView::GetSpecialThanks(void)
{
	return TEXT("Jim Barry, Geoff Chappell, Wine Headquarters");
}

CAtlString ShellTreeView::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ShellTreeView::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ShellTreeView::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_SHTVW_HANDLEOLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, properties.handleOLEDragDrop, description);
			break;
		case DISPID_SHTVW_HIDDENITEMSSTYLE:
			hr = GetDisplayStringForSetting(property, properties.hiddenItemsStyle, description);
			break;
		case DISPID_SHTVW_ITEMTYPESORTORDER:
			hr = GetDisplayStringForSetting(property, properties.itemTypeSortOrder, description);
			break;
		case DISPID_SHTVW_USEGENERICICONS:
			hr = GetDisplayStringForSetting(property, properties.useGenericIcons, description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellTreeView>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ShellTreeView::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_SHTVW_ITEMTYPESORTORDER:
			c = 2;
			break;
		case DISPID_SHTVW_HIDDENITEMSSTYLE:
		case DISPID_SHTVW_USEGENERICICONS:
			c = 3;
			break;
		case DISPID_SHTVW_HANDLEOLEDRAGDROP:
			c = 6;
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellTreeView>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(property == DISPID_SHTVW_HANDLEOLEDRAGDROP && iDescription >= 4) {
			propertyValue += 2;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_SHTVW_HANDLEOLEDRAGDROP:
		case DISPID_SHTVW_HIDDENITEMSSTYLE:
		case DISPID_SHTVW_ITEMTYPESORTORDER:
		case DISPID_SHTVW_USEGENERICICONS:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellTreeView>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	if(!pPropertyPage) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_SHTVW_DEFAULTMANAGEDITEMPROPERTIES:
			*pPropertyPage = CLSID_ShTvwDefaultManagedItemPropertiesProperties;
			return S_OK;
			break;
		case DISPID_SHTVW_DEFAULTNAMESPACEENUMSETTINGS:
			*pPropertyPage = CLSID_NamespaceEnumSettingsProperties;
			return S_OK;
			break;
		case DISPID_SHTVW_DISABLEDEVENTS:
		case DISPID_SHTVW_INFOTIPFLAGS:
		case DISPID_SHTVW_USESYSTEMIMAGELIST:
			*pPropertyPage = CLSID_CommonProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<ShellTreeView>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ShellTreeView::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_SHTVW_HANDLEOLEDRAGDROP:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPNONE, 0);
					break;
				case hoddSourcePart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPSOURCEPART, hoddSourcePart);
					break;
				case hoddTargetPart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPTARGETPART, hoddTargetPart);
					break;
				case hoddTargetPartWithDropHilite:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPTARGETPARTEX, hoddTargetPartWithDropHilite);
					break;
				case hoddSourcePart | hoddTargetPart:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPBOTHPARTS, hoddSourcePart | hoddTargetPart);
					break;
				case hoddSourcePart | hoddTargetPartWithDropHilite:
					description = GetResStringWithNumber(IDP_HANDLEOLEDRAGDROPBOTHPARTSEX, hoddSourcePart | hoddTargetPartWithDropHilite);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHTVW_HIDDENITEMSSTYLE:
			switch(cookie) {
				case hisNormal:
					description = GetResStringWithNumber(IDP_HIDDENITEMSSTYLENORMAL, hisNormal);
					break;
				case hisGhosted:
					description = GetResStringWithNumber(IDP_HIDDENITEMSSTYLEGHOSTED, hisGhosted);
					break;
				case hisGhostedOnDemand:
					description = GetResStringWithNumber(IDP_HIDDENITEMSSTYLEGHOSTEDONDEMAND, hisGhostedOnDemand);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHTVW_ITEMTYPESORTORDER:
			switch(cookie) {
				case itsoShellItemsFirst:
					description = GetResStringWithNumber(IDP_ITEMTYPESORTORDERSHELLITEMSFIRST, itsoShellItemsFirst);
					break;
				case itsoShellItemsLast:
					description = GetResStringWithNumber(IDP_ITEMTYPESORTORDERSHELLITEMSLAST, itsoShellItemsLast);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHTVW_USEGENERICICONS:
			switch(cookie) {
				case ugiNever:
					description = GetResStringWithNumber(IDP_USEGENERICICONSNEVER, ugiNever);
					break;
				case ugiOnlyForSlowItems:
					description = GetResStringWithNumber(IDP_USEGENERICICONSONLYFORSLOWITEMS, ugiOnlyForSlowItems);
					break;
				case ugiAlways:
					description = GetResStringWithNumber(IDP_USEGENERICICONSALWAYS, ugiAlways);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP ShellTreeView::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 3;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_ShTvwDefaultManagedItemPropertiesProperties;
		pPropertyPages->pElems[2] = CLSID_NamespaceEnumSettingsProperties;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ShellTreeView::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ShellTreeView>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ShellTreeView::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[2] = {
	    {OLEIVERB_PROPERTIES, L"&Properties...", MF_STRING, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", MF_STRING, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 2, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

HRESULT ShellTreeView::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}


STDMETHODIMP ShellTreeView::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 6;
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ShellBrowserControls");
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ShBrowserCtls");
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_AutoEditNewItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoEditNewItems);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_AutoEditNewItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_AUTOEDITNEWITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoEditNewItems != b) {
		properties.autoEditNewItems = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_AUTOEDITNEWITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_ColorCompressedItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.colorCompressedItems);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_ColorCompressedItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_COLORCOMPRESSEDITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.colorCompressedItems != b) {
		properties.colorCompressedItems = b;
		SetDirty(TRUE);

		if(attachedControl.IsWindow()) {
			attachedControl.Invalidate();
		}
		FireOnChanged(DISPID_SHTVW_COLORCOMPRESSEDITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_ColorEncryptedItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.colorEncryptedItems);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_ColorEncryptedItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_COLORENCRYPTEDITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.colorEncryptedItems != b) {
		properties.colorEncryptedItems = b;
		SetDirty(TRUE);

		if(attachedControl.IsWindow()) {
			attachedControl.Invalidate();
		}
		FireOnChanged(DISPID_SHTVW_COLORENCRYPTEDITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_DefaultManagedItemProperties(ShTvwManagedItemPropertiesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShTvwManagedItemPropertiesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.defaultManagedItemProperties;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_DefaultManagedItemProperties(ShTvwManagedItemPropertiesConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_DEFAULTMANAGEDITEMPROPERTIES);
	if(properties.defaultManagedItemProperties != newValue) {
		properties.defaultManagedItemProperties = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_DEFAULTMANAGEDITEMPROPERTIES);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_DefaultNamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings)
{
	ATLASSERT_POINTER(ppEnumSettings, INamespaceEnumSettings*);
	if(!ppEnumSettings) {
		return E_POINTER;
	}

	*ppEnumSettings = NULL;
	__try {
		if(properties.pDefaultNamespaceEnumSettings) {
			return properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(ppEnumSettings));
		}
	} __except(EXCEPTION_ACCESS_VIOLATION) {
		return E_NOINTERFACE;
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::putref_DefaultNamespaceEnumSettings(INamespaceEnumSettings* pEnumSettings)
{
	PUTPROPPROLOG(DISPID_SHTVW_DEFAULTNAMESPACEENUMSETTINGS);
	if(!pEnumSettings) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.pDefaultNamespaceEnumSettings != pEnumSettings) {
		if(properties.pDefaultNamespaceEnumSettings) {
			IPersistentChildObject* pPersistentChildObject = NULL;
			properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistentChildObject, reinterpret_cast<LPVOID*>(&pPersistentChildObject));
			ATLASSUME(pPersistentChildObject);
			if(pPersistentChildObject) {
				pPersistentChildObject->SetOwnerControl(NULL);
				pPersistentChildObject->Release();
			}
			properties.pDefaultNamespaceEnumSettings->Release();
			properties.pDefaultNamespaceEnumSettings = NULL;
		}
		pEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(&properties.pDefaultNamespaceEnumSettings));
		if(properties.pDefaultNamespaceEnumSettings && IsInDesignMode()) {
			IPersistentChildObject* pPersistentChildObject = NULL;
			properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_IPersistentChildObject, reinterpret_cast<LPVOID*>(&pPersistentChildObject));
			ATLASSUME(pPersistentChildObject);
			if(pPersistentChildObject) {
				pPersistentChildObject->SetOwnerControl(this);
				pPersistentChildObject->Release();
			}
		}
	} else {
		pEnumSettings->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SHTVW_DEFAULTNAMESPACEENUMSETTINGS);
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_DisplayElevationShieldOverlays(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayElevationShieldOverlays);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_DisplayElevationShieldOverlays(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_DISPLAYELEVATIONSHIELDOVERLAYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayElevationShieldOverlays != b) {
		properties.displayElevationShieldOverlays = b;
		SetDirty(TRUE);

		BOOL invalidate = FALSE;
		if(properties.pUnifiedImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			pImgLstPriv->SetOptions(ILOF_DISPLAYELEVATIONSHIELDS, properties.displayElevationShieldOverlays ? ILOF_DISPLAYELEVATIONSHIELDS : 0);
			invalidate = TRUE;
		}
		if(invalidate && attachedControl.IsWindow()) {
			attachedControl.Invalidate();
		}

		FireOnChanged(DISPID_SHTVW_DISPLAYELEVATIONSHIELDOVERLAYS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_HandleOLEDragDrop(HandleOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HandleOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.handleOLEDragDrop;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_HandleOLEDragDrop(HandleOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_HANDLEOLEDRAGDROP);
	if(properties.handleOLEDragDrop != newValue) {
		properties.handleOLEDragDrop = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_HANDLEOLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_HiddenItemsStyle(HiddenItemsStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HiddenItemsStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hiddenItemsStyle;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_HiddenItemsStyle(HiddenItemsStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_HIDDENITEMSSTYLE);
	if(properties.hiddenItemsStyle != newValue) {
		properties.hiddenItemsStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_HIDDENITEMSSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(attachedControl));
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
		case ilNonShellItems:
			*pValue = HandleToLong(properties.hImageList[0]);
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_HIMAGELIST);

	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilNonShellItems:
			properties.hImageList[0] = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			if(properties.pUnifiedImageList) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
			}
			fireViewChange = TRUE;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_SHTVW_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_hWndShellUIParentWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hWndShellUIParentWindow);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_hWndShellUIParentWindow(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_HWNDSHELLUIPARENTWINDOW);
	if(properties.hWndShellUIParentWindow != static_cast<HWND>(LongToHandle(newValue))) {
		properties.hWndShellUIParentWindow = static_cast<HWND>(LongToHandle(newValue));
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_HWNDSHELLUIPARENTWINDOW);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_InfoTipFlags(InfoTipFlagsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, InfoTipFlagsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.infoTipFlags;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_InfoTipFlags(InfoTipFlagsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_INFOTIPFLAGS);
	if(properties.infoTipFlags != newValue) {
		properties.infoTipFlags = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_INFOTIPFLAGS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_ItemEnumerationTimeout(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemEnumerationTimeout;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_ItemEnumerationTimeout(LONG newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_ITEMENUMERATIONTIMEOUT);
	if((newValue < 1000) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.itemEnumerationTimeout != newValue) {
		properties.itemEnumerationTimeout = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_ITEMENUMERATIONTIMEOUT);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_ItemTypeSortOrder(ItemTypeSortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemTypeSortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemTypeSortOrder;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_ItemTypeSortOrder(ItemTypeSortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_ITEMTYPESORTORDER);
	if(properties.itemTypeSortOrder != newValue) {
		properties.itemTypeSortOrder = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_ITEMTYPESORTORDER);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_LimitLabelEditInput(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.limitLabelEditInput);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_LimitLabelEditInput(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_LIMITLABELEDITINPUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.limitLabelEditInput != b) {
		properties.limitLabelEditInput = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_LIMITLABELEDITINPUT);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_LoadOverlaysOnDemand(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.loadOverlaysOnDemand);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_LoadOverlaysOnDemand(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_LOADOVERLAYSONDEMAND);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.loadOverlaysOnDemand != b) {
		properties.loadOverlaysOnDemand = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_LOADOVERLAYSONDEMAND);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_Namespaces(IShTreeViewNamespaces** ppNamespaces)
{
	ATLASSERT_POINTER(ppNamespaces, IShTreeViewNamespaces*);
	if(!ppNamespaces) {
		return E_POINTER;
	}

	if(attachedControl.IsWindow()) {
		return properties.pNamespaces->QueryInterface(IID_IShTreeViewNamespaces, reinterpret_cast<LPVOID*>(ppNamespaces));
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeView::get_PreselectBasenameOnLabelEdit(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.preselectBasenameOnLabelEdit);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_PreselectBasenameOnLabelEdit(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_PRESELECTBASENAMEONLABELEDIT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.preselectBasenameOnLabelEdit != b) {
		properties.preselectBasenameOnLabelEdit = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_PRESELECTBASENAMEONLABELEDIT);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_ProcessShellNotifications(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processShellNotifications);
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_ProcessShellNotifications(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_PROCESSSHELLNOTIFICATIONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processShellNotifications != b) {
		properties.processShellNotifications = b;
		SetDirty(TRUE);

		if(attachedControl.IsWindow()) {
			if(properties.processShellNotifications) {
				RegisterForShellNotifications();
			} else {
				DeregisterForShellNotifications();
			}
		}
		FireOnChanged(DISPID_SHTVW_PROCESSSHELLNOTIFICATIONS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellTreeView::get_SecurityZones(ISecurityZones** ppZones)
	{
		ATLASSERT_POINTER(ppZones, ISecurityZones*);
		if(!ppZones) {
			return E_POINTER;
		}

		CComObject<SecurityZones>* pSecurityZonesObj = NULL;
		CComObject<SecurityZones>::CreateInstance(&pSecurityZonesObj);
		pSecurityZonesObj->AddRef();
		HRESULT hr = pSecurityZonesObj->QueryInterface(IID_ISecurityZones, reinterpret_cast<LPVOID*>(ppZones));
		pSecurityZonesObj->Release();
		return hr;
	}
#endif

STDMETHODIMP ShellTreeView::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_TreeItems(IShTreeViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IShTreeViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	if(attachedControl.IsWindow()) {
		return properties.pTreeItems->QueryInterface(IID_IShTreeViewItems, reinterpret_cast<LPVOID*>(ppItems));
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeView::get_UseGenericIcons(UseGenericIconsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UseGenericIconsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.useGenericIcons;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_UseGenericIcons(UseGenericIconsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_USEGENERICICONS);
	if(properties.useGenericIcons != newValue) {
		properties.useGenericIcons = newValue;

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener) {
			FlushIcons(-1, FALSE);
		}
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_USEGENERICICONS);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_UseSystemImageList(UseSystemImageListConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UseSystemImageListConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.useSystemImageList;
	return S_OK;
}

STDMETHODIMP ShellTreeView::put_UseSystemImageList(UseSystemImageListConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHTVW_USESYSTEMIMAGELIST);
	if(properties.useSystemImageList != newValue) {
		properties.useSystemImageList = newValue;
		SetSystemImageLists();
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHTVW_USESYSTEMIMAGELIST);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}


STDMETHODIMP ShellTreeView::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ShellTreeView::Attach(OLE_HANDLE hWnd)
{
	ATLASSERT((attachedControl.m_hWnd == NULL) && "Call Detach() before attaching to a new control.");
	if(attachedControl.m_hWnd) {
		return E_FAIL;
	}
	ATLASSERT((properties.pAttachedCOMControl == NULL) && "Call Detach() before attaching to a new control.");
	if(properties.pAttachedCOMControl) {
		return E_FAIL;
	}
	#ifdef USE_STL
		ATLASSERT((properties.items.size() == 0) && "Call Detach() before attaching to a new control.");
		ATLASSERT((properties.namespaces.size() == 0) && "Call Detach() before attaching to a new control.");
	#else
		ATLASSERT((properties.items.GetCount() == 0) && "Call Detach() before attaching to a new control.");
		ATLASSERT((properties.namespaces.GetCount() == 0) && "Call Detach() before attaching to a new control.");
	#endif

	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(this);
	ATLASSUME(properties.pTreeItems);
	properties.pTreeItems->SetOwner(this);

	SHBHANDSHAKE handshakeDetails = {0};
	handshakeDetails.structSize = sizeof(SHBHANDSHAKE);
	handshakeDetails.processed = FALSE;
	handshakeDetails.errorCode = E_FAIL;
	handshakeDetails.shellBrowserBuildNumber = VERSION_BUILD;
	#ifdef UNICODE
		handshakeDetails.shellBrowserSupportsUnicode = TRUE;
	#else
		handshakeDetails.shellBrowserSupportsUnicode = FALSE;
	#endif
	handshakeDetails.pMessageHook = this;
	handshakeDetails.pShBInternalMessageHook = this;
	handshakeDetails.pCtlInterface = reinterpret_cast<LPVOID*>(&properties.pAttachedCOMControl);
	handshakeDetails.ppCtlInternalMessageHook = &properties.pAttachedInternalMessageListener;
	handshakeDetails.pCtlBuildNumber = &properties.attachedControlBuildNumber;

	HWND hWndTvw = static_cast<HWND>(LongToHandle(hWnd));
	if(SendMessage(hWndTvw, SHBM_HANDSHAKE, 0, reinterpret_cast<LPARAM>(&handshakeDetails))) {
		ATLTRACE2(shtvwTraceCtlCommunication, 3, TEXT("Handshake with 0x%X succeeded\n"), hWndTvw);
		attachedControl.Attach(hWndTvw);
		AddRef();
		ConfigureControl();
		return S_OK;
	} else {
		if(handshakeDetails.processed) {
			ATLTRACE2(shtvwTraceCtlCommunication, 0, TEXT("Handshake with 0x%X failed (0x%X)\n"), hWndTvw, handshakeDetails.errorCode);
			return handshakeDetails.errorCode;
		} else {
			ATLTRACE2(shtvwTraceCtlCommunication, 0, TEXT("Handshake with 0x%X failed (no response)\n"), hWndTvw);
		}
	}
	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(NULL);
	ATLASSUME(properties.pTreeItems);
	properties.pTreeItems->SetOwner(NULL);
	return E_FAIL;
}

STDMETHODIMP ShellTreeView::CompareItems(IShTreeViewItem* pFirstItem, IShTreeViewItem* pSecondItem, LONG sortColumnIndex/* = 0*/, LONG* pResult/* = NULL*/)
{
	ATLASSERT_POINTER(pFirstItem, IShTreeViewItem);
	ATLASSERT_POINTER(pSecondItem, IShTreeViewItem);
	ATLASSERT_POINTER(pResult, LONG);
	if(!pFirstItem || !pSecondItem || !pResult) {
		return E_POINTER;
	}

	HTREEITEM hFirstItem = NULL;
	HTREEITEM hSecondItem = NULL;

	OLE_HANDLE h = NULL;
	pFirstItem->get_Handle(&h);
	hFirstItem = static_cast<HTREEITEM>(LongToHandle(h));
	h = NULL;
	pSecondItem->get_Handle(&h);
	hSecondItem = static_cast<HTREEITEM>(LongToHandle(h));
	if(!hFirstItem || !hSecondItem) {
		return E_INVALIDARG;
	}

	PCIDLIST_ABSOLUTE pIDL1 = ItemHandleToPIDL(hFirstItem);
	PCIDLIST_ABSOLUTE pIDL2 = ItemHandleToPIDL(hSecondItem);
	if(!pIDL1 || !pIDL2) {
		return E_INVALIDARG;
	}

	// get the parent IShellFolder
	PIDLIST_ABSOLUTE pIDLParent1 = ILCloneFull(pIDL1);
	ATLASSERT_POINTER(pIDLParent1, ITEMIDLIST_ABSOLUTE);
	ILRemoveLastID(pIDLParent1);
	if(ILIsParent(pIDLParent1, pIDL2, TRUE)) {
		// same parent
		CComPtr<IShellFolder> pParentISF = NULL;
		ATLVERIFY(SUCCEEDED(BindToPIDL(pIDLParent1, IID_PPV_ARGS(&pParentISF))));
		ILFree(pIDLParent1);
		*pResult = static_cast<short>(HRESULT_CODE(pParentISF->CompareIDs(MAKELPARAM(max(sortColumnIndex, 0), 0), ILFindLastID(pIDL1), ILFindLastID(pIDL2))));
		return S_OK;
	}
	ILFree(pIDLParent1);

	/* TODO: On Vista, try to use SHCompareIDsFull (but check with "My Computer" whether sorting is still
	         right then!) */
	// fallback to the Desktop's IShellFolder interface
	*pResult = static_cast<short>(HRESULT_CODE(properties.pDesktopISF->CompareIDs(MAKELPARAM(max(sortColumnIndex, 0), 0), pIDL1, pIDL2)));
	return S_OK;
}

STDMETHODIMP ShellTreeView::CreateShellContextMenu(VARIANT items, OLE_HANDLE* pMenu)
{
	ATLASSERT_POINTER(pMenu, OLE_HANDLE);
	if(!pMenu) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	if(items.vt == VT_DISPATCH && items.pdispVal) {
		CComQIPtr<IShTreeViewNamespace> pShTvwNamespace = items.pdispVal;
		if(pShTvwNamespace) {
			// a single ShellTreeViewNamespace object
			LONG pIDLNamespace = 0;
			hr = pShTvwNamespace->get_FullyQualifiedPIDL(&pIDLNamespace);
			if(SUCCEEDED(hr)) {
				HMENU hMenu = NULL;
				hr = CreateShellContextMenu(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDLNamespace), hMenu);
				*pMenu = HandleToLong(hMenu);
			}
			return hr;
		}
	}

	UINT itemCount = 0;
	HTREEITEM* pItems = NULL;
	hr = VariantToItemHandles(items, pItems, itemCount);
	if(SUCCEEDED(hr)) {
		if(itemCount > 0) {
			ATLASSERT_ARRAYPOINTER(pItems, LONG, itemCount);
		}
		HMENU hMenu = NULL;
		hr = CreateShellContextMenu(pItems, itemCount, CMF_NORMAL, hMenu);
		*pMenu = HandleToLong(hMenu);
	}
	if(pItems) {
		delete[] pItems;
	}
	return hr;
}

STDMETHODIMP ShellTreeView::DestroyShellContextMenu(void)
{
	HRESULT hr = E_FAIL;
	if(contextMenuData.pContextMenuItems) {
		if(contextMenuData.currentShellContextMenu.IsMenu()) {
			Raise_DestroyingShellContextMenu(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)));
			hr = S_OK;
			contextMenuData.currentShellContextMenu.DestroyMenu();
		}
		contextMenuData.pContextMenuItems->Release();
		contextMenuData.pContextMenuItems = NULL;
	}
	if(contextMenuData.pIContextMenu) {
		contextMenuData.pIContextMenu->Release();
		contextMenuData.pIContextMenu = NULL;
	}
	contextMenuData.pMultiNamespaceParentISF = NULL;
	contextMenuData.pMultiNamespaceDataObject = NULL;
	return hr;
}

STDMETHODIMP ShellTreeView::Detach(void)
{
	ATLASSERT(attachedControl.IsWindow());

	flags.detaching = TRUE;
	flags.acceptNewTasks = FALSE;
	if(attachedControl.m_hWnd) {
		DeregisterForShellNotifications();
	}
	if(properties.pEnumTaskScheduler) {
		properties.pEnumTaskScheduler->RemoveTasks(TASKID_ShTvwBackgroundItemEnum, NULL, TRUE);
		ATLASSERT(properties.pEnumTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(properties.pNormalTaskScheduler) {
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShTvwAutoUpdate, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ElevationShield, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShTvwIcon, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShTvwOverlay, NULL, TRUE);
		ATLASSERT(properties.pNormalTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(properties.useSystemImageList & usilSmallImageList) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETIMAGELIST, 1/*ilItems*/, NULL)));
	}

	if(attachedControl.IsWindow()) {
		SHBHANDSHAKE handshakeDetails = {0};
		handshakeDetails.structSize = sizeof(SHBHANDSHAKE);
		handshakeDetails.processed = FALSE;
		handshakeDetails.errorCode = E_FAIL;
		handshakeDetails.shellBrowserBuildNumber = VERSION_BUILD;
		#ifdef UNICODE
			handshakeDetails.shellBrowserSupportsUnicode = TRUE;
		#else
			handshakeDetails.shellBrowserSupportsUnicode = FALSE;
		#endif
		handshakeDetails.pMessageHook = NULL;
		handshakeDetails.pShBInternalMessageHook = NULL;

		ATLVERIFY(attachedControl.SendMessage(SHBM_HANDSHAKE, 0, reinterpret_cast<LPARAM>(&handshakeDetails)));
	}
	if(properties.pAttachedCOMControl) {
		properties.pAttachedCOMControl->Release();
		properties.pAttachedCOMControl = NULL;
	}
	properties.pAttachedInternalMessageListener = NULL;
	properties.attachedControlBuildNumber = 0;
	if(attachedControl.m_hWnd) {
		//DeregisterForShellNotifications();
		attachedControl.KillTimer(timers.ID_BACKGROUNDENUMTIMEOUT);
		attachedControl.Detach();
	}

	// free all pIDLs and namespaces
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		#ifdef USE_STL
			if(properties.items.size() > 0) {
		#else
			if(properties.items.GetCount() > 0) {
		#endif
			Raise_RemovingItem(NULL);
		}
	}
	if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
		#ifdef USE_STL
			if(properties.namespaces.size() > 0) {
		#else
			if(properties.namespaces.GetCount() > 0) {
		#endif
			Raise_RemovingNamespace(NULL);
		}
	}
	#ifdef USE_STL
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			delete iter->second;
		}
		properties.items.clear();

		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			ILFree(const_cast<PIDLIST_ABSOLUTE>(iter->first));
			iter->second->Release();
		}
		properties.namespaces.clear();

		for(std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter = properties.backgroundEnums.cbegin(); iter != properties.backgroundEnums.cend(); ++iter) {
			if(iter->second->pTargetObject) {
				iter->second->pTargetObject->Release();
			}
			delete iter->second;
		}
		properties.backgroundEnums.clear();

		while(properties.shellItemDataOfInsertedItems.size() > 0) {
			delete properties.shellItemDataOfInsertedItems.top();
			properties.shellItemDataOfInsertedItems.pop();
		}

		properties.namespacesOfInsertedItems.clear();
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			delete properties.items.GetValueAt(p);
			properties.items.GetNextValue(p);
		}
		properties.items.RemoveAll();

		p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair1 = properties.namespaces.GetAt(p);
			ILFree(const_cast<PIDLIST_ABSOLUTE>(pPair1->m_key));
			pPair1->m_value->Release();
			properties.namespaces.GetNext(p);
		}
		properties.namespaces.RemoveAll();

		p = properties.backgroundEnums.GetStartPosition();
		while(p) {
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair2 = properties.backgroundEnums.GetAt(p);
			if(pPair2->m_value->pTargetObject) {
				pPair2->m_value->pTargetObject->Release();
			}
			delete pPair2->m_value;
			properties.backgroundEnums.GetNext(p);
		}
		properties.backgroundEnums.RemoveAll();

		for(size_t i = 0; i < properties.shellItemDataOfInsertedItems.GetCount(); ++i) {
			delete properties.shellItemDataOfInsertedItems.GetAt(p);
		}
		properties.shellItemDataOfInsertedItems.RemoveAll();

		properties.namespacesOfInsertedItems.RemoveAll();
	#endif

	DestroyShellContextMenu();
	contextMenuData.ResetToDefaults();
	dragDropStatus.ResetToDefaults();
	properties.pUnifiedImageList = NULL;
	if(pThreadObject) {
		pThreadObject->Release();
		if(threadReferenceCounter != -1) {
			// we were the ones that have set the object
			APIWrapper::SHSetThreadRef(NULL, NULL);
		}
		pThreadObject = NULL;
	}
	flags.detaching = FALSE;
	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(NULL);
	ATLASSUME(properties.pTreeItems);
	properties.pTreeItems->SetOwner(NULL);
	Release();
	return S_OK;
}

STDMETHODIMP ShellTreeView::DisplayShellContextMenu(VARIANT items, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	POINT position = {x, y};
	attachedControl.ClientToScreen(&position);

	HRESULT hr = S_OK;
	if(items.vt == VT_DISPATCH && items.pdispVal) {
		CComQIPtr<IShTreeViewNamespace> pShTvwNamespace = items.pdispVal;
		if(pShTvwNamespace) {
			// a single ShellTreeViewNamespace object
			LONG pIDLNamespace = 0;
			hr = pShTvwNamespace->get_FullyQualifiedPIDL(&pIDLNamespace);
			if(SUCCEEDED(hr)) {
				hr = DisplayShellContextMenu(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDLNamespace), position);
			}
			return hr;
		}
	}

	UINT itemCount = 0;
	HTREEITEM* pItems = NULL;
	hr = VariantToItemHandles(items, pItems, itemCount);
	if(SUCCEEDED(hr)) {
		if(itemCount > 0) {
			ATLASSERT_ARRAYPOINTER(pItems, LONG, itemCount);
		}
		hr = DisplayShellContextMenu(pItems, itemCount, position);
	}
	if(pItems) {
		delete[] pItems;
	}
	return hr;
}

STDMETHODIMP ShellTreeView::EnsureItemIsLoaded(VARIANT pIDLOrParsingName, IShTreeViewItem** ppLastValidItem)
{
	ATLASSERT_POINTER(ppLastValidItem, IShTreeViewItem*);
	if(!ppLastValidItem) {
		return E_POINTER;
	}

	BOOL isPIDL = TRUE;
	PCUIDLIST_RELATIVE typedPIDL = NULL;
	ATL::CString parsingNameToFind;

	VARIANT v;
	VariantInit(&v);
	if(pIDLOrParsingName.vt == VT_BSTR) {
		if(FAILED(VariantCopy(&v, &pIDLOrParsingName))) {
			return E_INVALIDARG;
		}
		isPIDL = FALSE;
		parsingNameToFind = v.bstrVal;
		VariantClear(&v);
	} else {
		if(FAILED(VariantChangeType(&v, &pIDLOrParsingName, 0, VT_I4))) {
			return E_INVALIDARG;
		}
		typedPIDL = *reinterpret_cast<PCUIDLIST_RELATIVE*>(&v.lVal);
		ATLASSERT_POINTER(typedPIDL, ITEMIDLIST_RELATIVE);
		VariantClear(&v);
		if(!typedPIDL) {
			return E_INVALIDARG;
		}
	}

	HTREEITEM hLastExistingItem = NULL;
	ATL::CString foundParsingName;
	if(isPIDL) {
		PCUIDLIST_RELATIVE pIDLBackup = typedPIDL;

		PIDLIST_RELATIVE pIDLPartial = GetDesktopPIDL();
		ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_RELATIVE);
		if(pIDLPartial) {
			HTREEITEM hItem = PIDLToItemHandle(reinterpret_cast<PCUIDLIST_ABSOLUTE>(pIDLPartial), FALSE);
			if(hItem) {
				hLastExistingItem = hItem;
			}

			UINT itemCount = ILCount(typedPIDL);
			for(UINT i = 0; i < itemCount; ++i) {
				pIDLPartial = ILAppendID(pIDLPartial, &typedPIDL->mkid, TRUE);
				ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_RELATIVE);
				typedPIDL = ILNext(typedPIDL);
				ATLASSERT_POINTER(typedPIDL, ITEMIDLIST_RELATIVE);

				// pIDLPartial is now a pIDL consisting of the first i items of typedPIDL
				hItem = PIDLToItemHandle(reinterpret_cast<PCUIDLIST_ABSOLUTE>(pIDLPartial), FALSE);
				if(hItem) {
					hLastExistingItem = hItem;
				}
			}
			ILFree(pIDLPartial);
		}
		// we've changed typedPIDL
		typedPIDL = pIDLBackup;
	} else {
		// look for a direct hit
		ATL::CString parsingName = parsingNameToFind;
		CT2CW converter(parsingName);
		LPCWSTR pParsingName = converter;
		HTREEITEM hItem = ParsingNameToItemHandle(pParsingName);
		if(hItem) {
			hLastExistingItem = hItem;
			foundParsingName = parsingName;
		}
		if(!hLastExistingItem) {
			// okay, so find the last existing parent item
			BOOL exitLoop = FALSE;
			do {
				// TODO: Maybe we should also look for "/"?
				int pos = parsingName.ReverseFind(TEXT('\\'));
				if(pos == -1) {
					pos = parsingName.GetLength();
					exitLoop = TRUE;
				}
				parsingName = parsingName.Left(pos);
				CT2CW converter(parsingName);
				pParsingName = converter;
				hItem = ParsingNameToItemHandle(pParsingName);
				if(hItem) {
					hLastExistingItem = hItem;
					foundParsingName = parsingName;
				}
			} while(!hLastExistingItem && !exitLoop);
		}
	}

	HTREEITEM hParentItem = NULL;
	UINT firstItemIDToInsert = 0;
	if(hLastExistingItem) {
		if(isPIDL) {
			LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(hLastExistingItem);
			ATLASSERT_POINTER(pItemDetails, SHELLTREEVIEWITEMDATA);
			if(ILIsEqual(pItemDetails->pIDL, reinterpret_cast<PCIDLIST_ABSOLUTE>(typedPIDL)) || !(pItemDetails->managedProperties & stmipSubItems)) {
				// nothing else to do
				ClassFactory::InitShellTreeItem(hLastExistingItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastValidItem));
				return S_OK;
			}

			// hLastExistingItem isn't the last item and we're managing this item's sub-items
			hParentItem = hLastExistingItem;
			firstItemIDToInsert = ILCount(pItemDetails->pIDL);
		} else {
			if(foundParsingName == parsingNameToFind) {
				// nothing else to do
				ClassFactory::InitShellTreeItem(hLastExistingItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastValidItem));
				return S_OK;
			}

			// hLastExistingItem isn't the last item and we're managing this item's sub-items
			hParentItem = hLastExistingItem;
		}
	} else {
		// TODO: search the namespaces
		ATLASSERT(FALSE);
		return E_FAIL;
	}

	if(isPIDL) {
		// insert the remaining items
		PIDLIST_RELATIVE pIDLPartial = GetDesktopPIDL();
		ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_ABSOLUTE);
		if(pIDLPartial) {
			CComPtr<IShTreeViewItems> pTreeItems = NULL;
			get_TreeItems(&pTreeItems);
			if(pTreeItems) {
				ATLASSUME(pTreeItems);
				BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();
				UINT itemCount = ILCount(typedPIDL);
				for(UINT i = 0; i < itemCount; ++i) {
					pIDLPartial = ILAppendID(pIDLPartial, &typedPIDL->mkid, TRUE);
					ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_RELATIVE);
					typedPIDL = ILNext(typedPIDL);
					ATLASSERT_POINTER(typedPIDL, ITEMIDLIST_RELATIVE);
					if(i < firstItemIDToInsert) {
						continue;
					}

					// pIDLPartial is now a pIDL consisting of the first i items of typedPIDL
					CComPtr<IShTreeViewItem> pTreeItem = NULL;
					PIDLIST_RELATIVE p = ILClone(pIDLPartial);
					v.vt = VT_I4;
					v.lVal = *reinterpret_cast<LONG*>(&p);

					LONG hasExpando = 0/*ExTVw::heNo*/;
					if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
						hasExpando = -2/*ExTVw::heAuto*/;
					}
					INamespaceEnumSettings* pEnumSettingsToUse = ItemHandleToNameSpaceEnumSettings(hParentItem);
					ATLASSUME(pEnumSettingsToUse);
					CachedEnumSettings cachedEnumSettings = CacheEnumSettings(pEnumSettingsToUse);
					if(pEnumSettingsToUse) {
						pEnumSettingsToUse->Release();
						pEnumSettingsToUse = NULL;
					}
					HasSubItems hasSubItems = HasAtLeastOneSubItem(reinterpret_cast<PIDLIST_ABSOLUTE>(p), &cachedEnumSettings, TRUE);
					if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
						hasExpando = 1/*heYes*/;
					}

					pTreeItems->Add(v, HandleToLong(hParentItem), NULL, static_cast<ShTvwManagedItemPropertiesConstants>(-1), NULL, hasExpando, -2, -2, -2, 0, VARIANT_FALSE, 1, &pTreeItem);
					VariantClear(&v);
					if(pTreeItem) {
						// start adding sibling items and also ensure the item has an expando
						TVITEMEX item = {0};
						item.hItem = hParentItem;
						item.mask = TVIF_HANDLE | TVIF_STATE | TVIF_CHILDREN;
						if(attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
							item.mask &= ~TVIF_STATE;
							if((item.state & TVIS_EXPANDEDONCE) != TVIS_EXPANDEDONCE) {
								// it has not yet been expanded, so insert the sub-items
								BOOL isShellItem;
								OnFirstTimeItemExpanding(hParentItem, tmImmediateThreading, TRUE, isShellItem);
								if(isShellItem) {
									// ensure the items aren't inserted twice
									item.mask |= TVIF_STATE;
									item.state = TVIS_EXPANDEDONCE;
									item.stateMask = TVIS_EXPANDEDONCE;
								}
							}
							if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
								item.cChildren = -2/*ExTVw::heAuto*/;
							} else {
								item.cChildren = 1/*ExTVw::heYes*/;
							}
							attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
						}

						OLE_HANDLE h = NULL;
						pTreeItem->get_Handle(&h);
						hParentItem = static_cast<HTREEITEM>(LongToHandle(h));
					} else {
						ILFree(p);
					}
				}
			}
			ILFree(pIDLPartial);
		}
	} else {
		CComPtr<IShellFolder> pParentISF;
		PCIDLIST_ABSOLUTE pIDL = ItemHandleToPIDL(hParentItem);
		HRESULT hr = BindToPIDL(pIDL, IID_PPV_ARGS(&pParentISF));
		PIDLIST_RELATIVE pIDLPartial = ILClone(pIDL);
		ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_ABSOLUTE);
		if(pIDLPartial) {
			CComPtr<IShTreeViewItems> pTreeItems = NULL;
			get_TreeItems(&pTreeItems);
			if(pTreeItems) {
				ATLASSUME(pTreeItems);
				BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();
				BOOL exitLoop = FALSE;
				while(SUCCEEDED(hr) && !exitLoop) {
					ATL::CString token = parsingNameToFind.Mid(foundParsingName.GetLength() + 1);
					// TODO: Maybe we should also look for "/"?
					int pos = token.ReverseFind(TEXT('\\'));
					if(pos == -1) {
						pos = token.GetLength();
						exitLoop = TRUE;
					}
					token = token.Left(pos);
					CT2W converter(token);
					LPWSTR pToken = converter;
					PIDLIST_RELATIVE pIDLRelative = NULL;
					hr = pParentISF->ParseDisplayName(GethWndShellUIParentWindow(), NULL, pToken, NULL, &pIDLRelative, NULL);
					if(SUCCEEDED(hr)) {
						CComPtr<IShellFolder> pItem;
						hr = pParentISF->BindToObject(pIDLRelative, NULL, IID_PPV_ARGS(&pItem));
						pParentISF = pItem;
						if(SUCCEEDED(hr)) {
							foundParsingName = parsingNameToFind.Left(foundParsingName.GetLength() + 1 + token.GetLength());
							pIDLPartial = ILAppendID(pIDLPartial, &pIDLRelative->mkid, TRUE);
							ATLASSERT_POINTER(pIDLPartial, ITEMIDLIST_RELATIVE);
							CComPtr<IShTreeViewItem> pTreeItem = NULL;
							PIDLIST_RELATIVE p = ILClone(pIDLPartial);
							v.vt = VT_I4;
							v.lVal = *reinterpret_cast<LONG*>(&p);

							LONG hasExpando = 0/*ExTVw::heNo*/;
							if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
								hasExpando = -2/*ExTVw::heAuto*/;
							}
							INamespaceEnumSettings* pEnumSettingsToUse = ItemHandleToNameSpaceEnumSettings(hParentItem);
							ATLASSUME(pEnumSettingsToUse);
							CachedEnumSettings cachedEnumSettings = CacheEnumSettings(pEnumSettingsToUse);
							if(pEnumSettingsToUse) {
								pEnumSettingsToUse->Release();
								pEnumSettingsToUse = NULL;
							}
							HasSubItems hasSubItems = HasAtLeastOneSubItem(reinterpret_cast<PIDLIST_ABSOLUTE>(p), &cachedEnumSettings, TRUE);
							if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
								hasExpando = 1/*heYes*/;
							}

							pTreeItems->Add(v, HandleToLong(hParentItem), NULL, static_cast<ShTvwManagedItemPropertiesConstants>(-1), NULL, hasExpando, -2, -2, -2, 0, VARIANT_FALSE, 1, &pTreeItem);
							VariantClear(&v);
							if(pTreeItem) {
								// start adding sibling items and also ensure the item has an expando
								TVITEMEX item = {0};
								item.hItem = hParentItem;
								item.mask = TVIF_HANDLE | TVIF_STATE | TVIF_CHILDREN;
								if(attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
									item.mask &= ~TVIF_STATE;
									if((item.state & TVIS_EXPANDEDONCE) != TVIS_EXPANDEDONCE) {
										// it has not yet been expanded, so insert the sub-items
										BOOL isShellItem;
										OnFirstTimeItemExpanding(hParentItem, tmImmediateThreading, TRUE, isShellItem);
										if(isShellItem) {
											// ensure the items aren't inserted twice
											item.mask |= TVIF_STATE;
											item.state = TVIS_EXPANDEDONCE;
											item.stateMask = TVIS_EXPANDEDONCE;
										}
									}
									if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
										item.cChildren = -2/*ExTVw::heAuto*/;
									} else {
										item.cChildren = 1/*ExTVw::heYes*/;
									}
									attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
								}

								OLE_HANDLE h = NULL;
								pTreeItem->get_Handle(&h);
								hParentItem = static_cast<HTREEITEM>(LongToHandle(h));
							} else {
								ILFree(p);
							}
						}
					}
				}
			}
			ILFree(pIDLPartial);
		}
	}

	if(hParentItem) {
		ClassFactory::InitShellTreeItem(hParentItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(ppLastValidItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ShellTreeView::FlushManagedIcons(VARIANT_BOOL includeOverlays/* = VARIANT_TRUE*/)
{
	FlushIcons(-1, VARIANTBOOL2BOOL(includeOverlays));
	return S_OK;
}

STDMETHODIMP ShellTreeView::GetShellContextMenuItemString(LONG commandID, VARIANT* pItemDescription/* = NULL*/, VARIANT* pItemVerb/* = NULL*/, VARIANT_BOOL* pCommandIDWasValid/* = NULL*/)
{
	ATLASSERT_POINTER(pCommandIDWasValid, VARIANT_BOOL);
	if(!pCommandIDWasValid) {
		return E_POINTER;
	}
	ATLASSUME(contextMenuData.pIContextMenu);
	if(!contextMenuData.pIContextMenu) {
		return E_FAIL;
	}

	if(pItemDescription) {
		VariantInit(pItemDescription);
	}
	if(pItemVerb) {
		VariantInit(pItemVerb);
	}

	*pCommandIDWasValid = VARIANT_FALSE;
	if(commandID < MIN_CONTEXTMENUEXTENSION_CMDID || commandID > MAX_CONTEXTMENUEXTENSION_CMDID) {
		return S_OK;
	}
	commandID -= MIN_CONTEXTMENUEXTENSION_CMDID;

	// implementation of GCS_VALIDATE seems to be very buggy, so don't count on it
	/*TCHAR pBuffer[2048];
	HRESULT hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_VALIDATE, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
	if(hr == S_OK) {
		*pCommandIDWasValid = VARIANT_TRUE;
		if(pItemDescription) {
			ZeroMemory(pBuffer, 2048 * sizeof(TCHAR));
			ATLVERIFY(SUCCEEDED(contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_HELPTEXT, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048)));
			pItemDescription->vt = VT_BSTR;
			CT2OLE converter(pBuffer);
			LPOLESTR convertedString = converter;
			pItemDescription->bstrVal = SysAllocString(convertedString);
		}
		if(pItemVerb) {
			ZeroMemory(pBuffer, 2048 * sizeof(TCHAR));
			ATLVERIFY(SUCCEEDED(contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048)));
			pItemVerb->vt = VT_BSTR;
			CT2OLE converter(pBuffer);
			LPOLESTR convertedString = converter;
			pItemVerb->bstrVal = SysAllocString(convertedString);
		}
	}
	return S_OK;*/
	TCHAR pBuffer[2048];
	if(pItemDescription) {
		ZeroMemory(pBuffer, 2048 * sizeof(TCHAR));
		HRESULT hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_HELPTEXT, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
		*pCommandIDWasValid = BOOL2VARIANTBOOL(hr == S_OK);
		//if(hr == S_OK) {
			pItemDescription->vt = VT_BSTR;
			CT2OLE converter(pBuffer);
			LPOLESTR convertedString = converter;
			pItemDescription->bstrVal = SysAllocString(convertedString);
		//}
	}
	if(pItemVerb) {
		ZeroMemory(pBuffer, 2048 * sizeof(TCHAR));
		HRESULT hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
		*pCommandIDWasValid = BOOL2VARIANTBOOL(hr == S_OK);
		//if(hr == S_OK) {
			pItemVerb->vt = VT_BSTR;
			CT2OLE converter(pBuffer);
			LPOLESTR convertedString = converter;
			pItemVerb->bstrVal = SysAllocString(convertedString);
		//}
	} else if(!pItemDescription) {
		HRESULT hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_VALIDATE, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
		if(hr != S_OK) {
			hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_HELPTEXT, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
			if(hr != S_OK) {
				hr = contextMenuData.pIContextMenu->GetCommandString(commandID, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pBuffer), 2048);
			}
		}
		*pCommandIDWasValid = BOOL2VARIANTBOOL(hr == S_OK);
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::InvokeDefaultShellContextMenuCommand(VARIANT items)
{
	HRESULT hr = S_OK;

	UINT itemCount = 0;
	HTREEITEM* pItems = NULL;
	hr = VariantToItemHandles(items, pItems, itemCount);
	if(SUCCEEDED(hr)) {
		if(itemCount > 0) {
			ATLASSERT_ARRAYPOINTER(pItems, LONG, itemCount);
		}
		hr = InvokeDefaultShellContextMenuCommand(pItems, itemCount);
	}
	if(pItems) {
		HeapFree(GetProcessHeap(), 0, pItems);
	}
	return hr;
}

STDMETHODIMP ShellTreeView::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ShellTreeView::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of IMessageListener
HRESULT ShellTreeView::PostMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, DWORD cookie, BOOL eatenMessage)
{
	if(!pResult) {
		return E_POINTER;
	}

	BOOL wasHandled = TRUE;
	if(hWnd == attachedControl.m_hWnd) {
		LRESULT lr = *pResult;
		switch(message) {
			case WM_DRAWITEM:
			case WM_INITMENUPOPUP:
			case WM_MEASUREITEM:
			case WM_MENUCHAR:
			case WM_MENUSELECT:
			case WM_NEXTMENU:
				lr = OnMenuMessage(message, wParam, lParam, wasHandled, cookie, eatenMessage, FALSE);
				break;
			case TVM_GETITEM:
				lr = OnGetItem(message, wParam, lParam, wasHandled, cookie, eatenMessage, FALSE);
				break;
			case OCM_NOTIFY:
				lr = OnReflectedNotify(message, wParam, lParam, wasHandled, cookie, eatenMessage, FALSE);
				break;
			default:
				wasHandled = FALSE;
				break;
		}
		if(wasHandled) {
			*pResult = lr;
		}
	}

	return (wasHandled ? S_OK : S_FALSE);
}

HRESULT ShellTreeView::PreMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, LPDWORD pCookie)
{
	if(!pResult || !pCookie) {
		return E_POINTER;
	}

	BOOL eatMessage = FALSE;
	if(hWnd == attachedControl.m_hWnd) {
		BOOL wasHandled = TRUE;
		LRESULT lr = *pResult;
		__try {
			switch(message) {
				case TVM_SETITEM:
					lr = OnSetItem(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case OCM_NOTIFY:
					lr = OnReflectedNotify(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_SHCHANGENOTIFY:
					eatMessage = TRUE;
					lr = OnSHChangeNotify(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TIMER:
					lr = OnTimer(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_ITEMENUMCOMPLETE:
					eatMessage = TRUE;
					lr = OnTriggerItemEnumComplete(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_SETELEVATIONSHIELD:
					eatMessage = TRUE;
					lr = OnTriggerSetElevationShield(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATEICON:
					eatMessage = TRUE;
					lr = OnTriggerUpdateIcon(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATESELECTEDICON:
					eatMessage = TRUE;
					lr = OnTriggerUpdateSelectedIcon(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATEOVERLAY:
					eatMessage = TRUE;
					lr = OnTriggerUpdateOverlay(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_ENQUEUETASK:
					eatMessage = TRUE;
					lr = OnTriggerEnqueueTask(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				default:
					wasHandled = FALSE;
					break;
			}
		} __except(EXCEPTION_EXECUTE_HANDLER) {
			//
		}
		if(wasHandled) {
			*pResult = lr;
		}
	}

	return (eatMessage ? S_FALSE : S_OK);
}
// implementation of IMessageListener
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IInternalMessageListener
HRESULT ShellTreeView::ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch(message) {
		case SHTVM_ALLOCATEMEMORY:
			return OnInternalAllocateMemory(message, wParam, lParam);
			break;
		case SHTVM_FREEMEMORY:
			return OnInternalFreeMemory(message, wParam, lParam);
			break;
		case SHTVM_INSERTINGITEM:
			return OnInternalInsertingItem(message, wParam, lParam);
			break;
		case SHTVM_INSERTEDITEM:
			return OnInternalInsertedItem(message, wParam, lParam);
			break;
		case SHTVM_FREEITEM:
			return OnInternalFreeItem(message, wParam, lParam);
			break;
		case SHTVM_UPDATEDITEMHANDLE:
			return OnInternalUpdatedItemHandle(message, wParam, lParam);
			break;
		case SHTVM_GETSHTREEVIEWITEM:
			return OnInternalGetShTreeViewItem(message, wParam, lParam);
			break;
		case SHTVM_RENAMEDITEM:
			return OnInternalRenamedItem(message, wParam, lParam);
			break;
		case SHTVM_COMPAREITEMS:
			return OnInternalCompareItems(message, wParam, lParam);
			break;
		case SHTVM_BEGINLABELEDITA:
		case SHTVM_BEGINLABELEDITW:
			return OnInternalBeginLabelEdit(message, wParam, lParam);
			break;
		case SHTVM_BEGINDRAGA:
		case SHTVM_BEGINDRAGW:
			return OnInternalBeginDrag(message, wParam, lParam);
			break;
		case SHTVM_HANDLEDRAGEVENT:
			return OnInternalHandleDragEvent(message, wParam, lParam);
			break;
		case SHTVM_CONTEXTMENU:
			return OnInternalContextMenu(message, wParam, lParam);
			break;
		case SHTVM_GETINFOTIP:
			return OnInternalGetInfoTip(message, wParam, lParam);
			break;
		case SHTVM_CUSTOMDRAW:
			return OnInternalCustomDraw(message, wParam, lParam);
			break;
	}
	return E_NOTIMPL;
}
// implementation of IInternalMessageListener
//////////////////////////////////////////////////////////////////////


LRESULT ShellTreeView::OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(message == WM_MENUSELECT) {
		UINT commandID = 0;
		if((HIWORD(wParam) == 0xFFFF) && (lParam == NULL)) {
			// the menu has been closed
		} else if(HIWORD(wParam) & MF_POPUP) {
			CMenuHandle menu = reinterpret_cast<HMENU>(lParam);
			CMenuItemInfo menuItemInfo;
			menuItemInfo.fMask = MIIM_ID;
			menu.GetMenuItemInfo(LOWORD(wParam), TRUE, &menuItemInfo);
			commandID = menuItemInfo.wID;
		} else {
			commandID = LOWORD(wParam);
		}
		BOOL isSystemMenuItem = FALSE;
		if(contextMenuData.currentShellContextMenu.IsMenu()) {
			isSystemMenuItem = ((commandID >= MIN_CONTEXTMENUEXTENSION_CMDID) && (commandID <= MAX_CONTEXTMENUEXTENSION_CMDID));
		}
		if(commandID != 0) {
			if(isSystemMenuItem) {
				contextMenuData.clickedMenu = reinterpret_cast<HMENU>(lParam);
			} else {
				contextMenuData.clickedMenu.Detach();
			}
		}
		//if((commandID == 0) || isSystemMenuItem) {
			/* TODO: Think about what's better: pass the passed menu handle which might be a sub-menu or pass the
			         main menu handle (which we would have to buffer in a class-wide variable). */
			Raise_ChangedContextMenuSelection(HandleToLong(reinterpret_cast<HMENU>(lParam)), commandID, BOOL2VARIANTBOOL(!isSystemMenuItem));
		//}
	}

	if(contextMenuData.pIContextMenu) {
		CComQIPtr<IContextMenu3> pIContextMenu3 = contextMenuData.pIContextMenu;
		if(pIContextMenu3) {
			LRESULT lr = 0;
			if(SUCCEEDED(pIContextMenu3->HandleMenuMsg2(message, wParam, lParam, &lr))) {
				return lr;
			}
		} else {
			CComQIPtr<IContextMenu2> pIContextMenu2 = contextMenuData.pIContextMenu;
			if(pIContextMenu2) {
				if(SUCCEEDED(pIContextMenu2->HandleMenuMsg(message, wParam, lParam))) {
					switch(message) {
						case WM_DRAWITEM:
						case WM_MEASUREITEM:
							return TRUE;
						default:
							return 0;
					}
				}
			}
		}
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ShellTreeView::OnGetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	LPTVITEMEX pItemData = reinterpret_cast<LPTVITEMEX>(lParam);
	if(!pItemData) {
		wasHandled = FALSE;
		return TRUE;
	}

	LONG itemID = -1;
	if(pItemData->mask & TVIF_IMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			itemID = ItemHandleToID(pItemData->hItem);
			if(itemID != -1) {
				// get real icon index
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_NONSHELLITEMICON, &pItemData->iImage, NULL, NULL, NULL)));
				}
			}
		}
	}
	if(pItemData->mask & TVIF_SELECTEDIMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			if(itemID == -1) {
				itemID = ItemHandleToID(pItemData->hItem);
			}
			if(itemID != -1) {
				// get real icon index
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_SELECTEDNONSHELLITEMICON, &pItemData->iSelectedImage, NULL, NULL, NULL)));
				}
			}
		}
	}
	if(pItemData->mask & TVIF_EXPANDEDIMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			if(itemID == -1) {
				itemID = ItemHandleToID(pItemData->hItem);
			}
			if(itemID != -1) {
				// get real icon index
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_EXPANDEDNONSHELLITEMICON, &pItemData->iExpandedImage, NULL, NULL, NULL)));
				}
			}
		}
	}
	return TRUE;
}

LRESULT ShellTreeView::OnSetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	LPTVITEMEX pItemData = reinterpret_cast<LPTVITEMEX>(lParam);
	if(!pItemData) {
		wasHandled = FALSE;
		return TRUE;
	}

	BOOL redraw = FALSE;
	LONG itemID = -1;
	if(pItemData->mask & TVIF_IMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			itemID = ItemHandleToID(pItemData->hItem);
			if(itemID != -1) {
				// update icon index
				BOOL setID = FALSE;
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_NONSHELLITEMICON, NULL, pItemData->iImage, 0, NULL, 0)));
					setID = TRUE;
				}
				if(setID) {
					// set the item's icon index to its ID
					pItemData->iImage = itemID;
					// since in most cases we don't actually change the icon index, we need to force a redraw manually
					redraw = TRUE;
				}
			}
		}
	}
	if(pItemData->mask & TVIF_SELECTEDIMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			if(itemID == -1) {
				itemID = ItemHandleToID(pItemData->hItem);
			}
			if(itemID != -1) {
				// update icon index
				BOOL setID = FALSE;
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SELECTEDNONSHELLITEMICON, NULL, pItemData->iSelectedImage, 0, NULL, 0)));
					setID = TRUE;
				}
				if(setID) {
					// set the item's icon index to its ID
					pItemData->iSelectedImage = itemID;
					// since in most cases we don't actually change the icon index, we need to force a redraw manually
					redraw = TRUE;
				}
			}
		}
	}
	if(pItemData->mask & TVIF_EXPANDEDIMAGE) {
		if(!IsShellItem(pItemData->hItem)) {
			if(itemID == -1) {
				itemID = ItemHandleToID(pItemData->hItem);
			}
			if(itemID != -1) {
				// update icon index
				BOOL setID = FALSE;
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_EXPANDEDNONSHELLITEMICON, NULL, pItemData->iExpandedImage, 0, NULL, 0)));
					setID = TRUE;
				}
				if(setID) {
					// set the item's icon index to its ID
					pItemData->iExpandedImage = itemID;
					// since in most cases we don't actually change the icon index, we need to force a redraw manually
					redraw = TRUE;
				}
			}
		}
	}
	if(redraw) {
		RECT rc = {0};
		*reinterpret_cast<HTREEITEM*>(&rc) = pItemData->hItem;
		if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
			attachedControl.InvalidateRect(&rc);
		}
	}
	return TRUE;
}

LRESULT ShellTreeView::OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& cookie, BOOL& eatenMessage, BOOL preProcessing)
{
	LPNMHDR pDetails = reinterpret_cast<LPNMHDR>(lParam);
	switch(pDetails->code) {
		case TVN_GETDISPINFOA:
		case TVN_GETDISPINFOW:
			if(preProcessing) {
				return OnGetDispInfoNotification(pDetails->code, pDetails, wasHandled, cookie, preProcessing);
			}
			break;
		case TVN_ITEMEXPANDINGA:
		case TVN_ITEMEXPANDINGW:
			if(!preProcessing && !eatenMessage) {
				return OnItemExpandingNotification(pDetails->code, pDetails, wasHandled, cookie, preProcessing);
			}
			break;
		case TVN_ITEMEXPANDEDA:
		case TVN_ITEMEXPANDEDW:
			if(!preProcessing && !eatenMessage) {
				return OnItemExpandedNotification(pDetails->code, pDetails, wasHandled, cookie, preProcessing);
			}
			break;
		case TVN_SELCHANGEDA:
		case TVN_SELCHANGEDW:
			if(preProcessing) {
				return OnCaretChangedNotification(pDetails->code, pDetails, wasHandled, cookie, preProcessing);
			}
			break;
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	LRESULT ret = 0;
	if(flags.detaching) {
		return ret;
	}

	LONG eventID = 0;
	PCIDLIST_ABSOLUTE_ARRAY ppIDLs = NULL;
	HANDLE hNotifyLock = SHChangeNotification_Lock(reinterpret_cast<HANDLE>(wParam), static_cast<DWORD>(lParam), const_cast<PIDLIST_ABSOLUTE**>(&ppIDLs), &eventID);
	if(hNotifyLock) {
		/*if(eventID & SHCNE_INTERRUPT) {
			eventID &= 0xFFFFFFF;
		}*/
		ATLTRACE2(shtvwTraceAutoUpdate, 2, TEXT("Received notification 0x%X\n"), eventID);

		if(eventID == SHCNE_MEDIAREMOVED) {
			if(ppIDLs && ppIDLs[0]) {
				SHFILEINFO fileInfoData = {0};
				SHGetFileInfo(reinterpret_cast<LPCTSTR>(ppIDLs[0]), 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_DISPLAYNAME);
				if(lstrlen(fileInfoData.szDisplayName) == 0) {
					// invalid pIDL - probably a drive has been removed instead only a media
					eventID = SHCNE_DRIVEREMOVED;
				}
			}
		}
		if(eventID == SHCNE_NETSHARE || eventID == SHCNE_NETUNSHARE) {
			if(ppIDLs && ppIDLs[0]) {
				// flush cache
				APIWrapper::IsPathShared(NULL, TRUE, NULL);
			}
		}

		switch(eventID) {
			case SHCNE_ASSOCCHANGED:
				ret = OnSHChangeNotify_ASSOCCHANGED(ppIDLs);
				break;
			case SHCNE_CREATE:
			case SHCNE_MKDIR:
				ret = OnSHChangeNotify_CREATE(ppIDLs);
				break;
			case SHCNE_DELETE:
			case SHCNE_RMDIR:
				ret = OnSHChangeNotify_DELETE(ppIDLs);
				break;
			case SHCNE_DRIVEADD:
			case SHCNE_DRIVEADDGUI:
				ret = OnSHChangeNotify_DRIVEADD(ppIDLs);
				break;
			case SHCNE_DRIVEREMOVED:
				ret = OnSHChangeNotify_DRIVEREMOVED(ppIDLs);
				break;
			case SHCNE_MEDIAINSERTED:
				ret = OnSHChangeNotify_MEDIAINSERTED(ppIDLs);
				break;
			case SHCNE_MEDIAREMOVED:
				/* Windows Explorer removes any sub-items of the affected item on SHCNE_SERVERDISCONNECT. This is
				   exactly what we do in OnSHChangeNotify_MEDIAREMOVED, so no need to implement it again. */
			case SHCNE_SERVERDISCONNECT:
				ret = OnSHChangeNotify_MEDIAREMOVED(ppIDLs);
				break;
			case SHCNE_NETSHARE:
			case SHCNE_NETUNSHARE:
				ret = OnSHChangeNotify_CHANGEDSHARE(ppIDLs);
				break;
			case SHCNE_RENAMEITEM:
			case SHCNE_RENAMEFOLDER:
				ret = OnSHChangeNotify_RENAME(ppIDLs);
				break;
			case SHCNE_UPDATEDIR:
				ret = OnSHChangeNotify_UPDATEDIR(ppIDLs);
				break;
			case SHCNE_UPDATEIMAGE:
				ret = OnSHChangeNotify_UPDATEIMAGE(ppIDLs);
				break;
			case SHCNE_ATTRIBUTES:
			case SHCNE_UPDATEITEM:
				ret = OnSHChangeNotify_UPDATEITEM(ppIDLs);
				break;
		}
		SHChangeNotification_Unlock(hNotifyLock);
	}
	return ret;
}

LRESULT ShellTreeView::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& eatenMessage, BOOL /*preProcessing*/)
{
	if(wParam == timers.ID_BACKGROUNDENUMTIMEOUT) {
		eatenMessage = TRUE;
		#ifdef USE_STL
			for(std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter = properties.backgroundEnums.cbegin(); iter != properties.backgroundEnums.cend(); ++iter) {
				LPBACKGROUNDENUMERATION pBackgroundEnum = iter->second;
		#else
			POSITION p = properties.backgroundEnums.GetStartPosition();
			while(p) {
				CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair = properties.backgroundEnums.GetAt(p);
				LPBACKGROUNDENUMERATION pBackgroundEnum = pPair->m_value;
		#endif

			DWORD now = GetTickCount();
			DWORD diff = now - pBackgroundEnum->startTime;
			if(now < pBackgroundEnum->startTime) {
				diff = ULONG_MAX - pBackgroundEnum->startTime + now;
			}

			#ifndef USE_STL
				properties.backgroundEnums.GetNext(p);
			#endif

			if(diff >= static_cast<DWORD>(properties.itemEnumerationTimeout) && !pBackgroundEnum->timedOut) {
				if(properties.itemEnumerationTimeout != -1) {
					pBackgroundEnum->timedOut = TRUE;
					if(pBackgroundEnum->pTargetObject) {
						CComQIPtr<IShTreeViewItem> pTvwItem = pBackgroundEnum->pTargetObject;
						if(pTvwItem) {
							OLE_HANDLE h = NULL;
							pTvwItem->get_Handle(&h);
							if(h) {
								LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(static_cast<HTREEITEM>(LongToHandle(h)));
								if(pItemDetails) {
									pItemDetails->isLoadingSubItems = FALSE;
								}
							}
						}
					}
					if(pBackgroundEnum->raiseEvents) {
						Raise_ItemEnumerationTimedOut(pBackgroundEnum->pTargetObject);
					}
				}
			}
		}
	} else {
		wasHandled = FALSE;
	}
	return 0;
}

LRESULT ShellTreeView::OnTriggerEnqueueTask(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	ATLASSERT_POINTER(lParam, SHCTLSBACKGROUNDTASKINFO);
	ReplyMessage(TRUE);

	LPSHCTLSBACKGROUNDTASKINFO pLParam = reinterpret_cast<LPSHCTLSBACKGROUNDTASKINFO>(lParam);
	if(pLParam) {
		ATLASSUME(pLParam->pTask);
		if(pLParam->pTask) {
			ATLVERIFY(SUCCEEDED(EnqueueTask(pLParam->pTask, pLParam->taskID, 0, pLParam->taskPriority)));
			pLParam->pTask->Release();
		}
		delete pLParam;
		pLParam = NULL;
	}
	return 0;
}

LRESULT ShellTreeView::OnTriggerItemEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHTVWBACKGROUNDITEMENUMINFO pEnumResult = NULL;
		EnterCriticalSection(&properties.backgroundEnumResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.backgroundEnumResults.empty()) {
				pEnumResult = properties.backgroundEnumResults.front();
				properties.backgroundEnumResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.backgroundEnumResults.IsEmpty()) {
				pEnumResult = properties.backgroundEnumResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.backgroundEnumResultsCritSection);

		if(empty) {
			break;
		} else if(!pEnumResult) {
			continue;
		}
		HTREEITEM hInsertAfter = pEnumResult->hInsertAfter;

		LPBACKGROUNDENUMERATION pBackgroundEnum = NULL;
		BOOL namespaceHasBeenRemoved = FALSE;
		ATLASSERT(pEnumResult->taskID);
		#ifdef USE_STL
			std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter = properties.backgroundEnums.find(pEnumResult->taskID);
			namespaceHasBeenRemoved = (iter == properties.backgroundEnums.cend());
			if(!namespaceHasBeenRemoved) {
				pBackgroundEnum = iter->second;
			}
			properties.backgroundEnums.erase(pEnumResult->taskID);
		#else
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair = properties.backgroundEnums.Lookup(pEnumResult->taskID);
			namespaceHasBeenRemoved = !pPair;
			if(!namespaceHasBeenRemoved) {
				pBackgroundEnum = pPair->m_value;
			}
			properties.backgroundEnums.RemoveKey(pEnumResult->taskID);
		#endif

		if(pEnumResult->hPIDLBuffer) {
			// insert the items
			if(namespaceHasBeenRemoved) {
				ATLTRACE2(shtvwTraceThreading, 3, TEXT("Ignoring insertion of %i sub-items of item 0x%X due to removed namespace\n"), DPA_GetPtrCount(pEnumResult->hPIDLBuffer), pEnumResult->hParentItem);
				DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAEnumeratedItemElement, NULL);
			} else {
				#ifdef USE_STL
					std::vector<std::pair<HTREEITEM, PCIDLIST_ABSOLUTE>> immediateSubItems;
				#else
					CAtlMap<HTREEITEM, PCIDLIST_ABSOLUTE> immediateSubItems;
				#endif
				OLE_HANDLE parentItem = HandleToLong(pEnumResult->hParentItem);
				HTREEITEM hItemToEdit = NULL;
				BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();
				int itemCount = 0;
				for(int i = 0; TRUE; ++i) {
					LPENUMERATEDITEM pItem = reinterpret_cast<LPENUMERATEDITEM>(DPA_GetPtr(pEnumResult->hPIDLBuffer, i));
					if(!pItem) {
						break;
					}

					if(pEnumResult->checkForDuplicates) {
						if(i == 0) {
							// cache all immediate sub-items
							HTREEITEM hChildItem = NULL;
							if(pEnumResult->hParentItem) {
								hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(pEnumResult->hParentItem)));
							} else {
								hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
							}
							while(hChildItem) {
								#ifdef USE_STL
									immediateSubItems.push_back(std::pair<HTREEITEM, PCIDLIST_ABSOLUTE>(hChildItem, GetItemDetails(hChildItem)->pIDL));
								#else
									immediateSubItems[hChildItem] = GetItemDetails(hChildItem)->pIDL;
								#endif
								hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hChildItem)));
							}
						}

						HTREEITEM hExistingItem = NULL;  //ItemFindImmediateSubItem(pEnumResult->hParentItem, pItem->pIDL, FALSE);
						#ifdef USE_STL
							std::vector<std::pair<HTREEITEM, PCIDLIST_ABSOLUTE>>::const_iterator foundEntry;
							BOOL done = FALSE;
							#pragma omp parallel
							{
								int this_thread = omp_get_thread_num();
								int num_threads = omp_get_num_threads();
								int size = immediateSubItems.size();
								int begin = (this_thread + 0) * size / num_threads;
								int end = (this_thread + 1) * size / num_threads;

								for(int i = begin; i < end; ++i) {
									#pragma omp flush(done)
									if(done) {
										break;
									}
									if(ILIsEqual(immediateSubItems[i].second, pItem->pIDL)) {
										done = TRUE;
										#pragma omp flush(done)
										hExistingItem = immediateSubItems[i].first;
										foundEntry = immediateSubItems.cbegin();
										std::advance(foundEntry, i);
									}
								}
							}
							if(hExistingItem) {
								immediateSubItems.erase(foundEntry);
							}
						#else
							POSITION p = immediateSubItems.GetStartPosition();
							while(p) {
								CAtlMap<HTREEITEM, PCIDLIST_ABSOLUTE>::CPair* pPair = immediateSubItems.GetAt(p);
								if(ILIsEqual(pPair->m_value, pItem->pIDL)) {
									hExistingItem = pPair->m_key;
									immediateSubItems.RemoveAtPos(p);
									break;
								}
								immediateSubItems.GetNext(p);
							}
						#endif
						if(hExistingItem) {
							ATLTRACE2(shtvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X because it's a duplicate of item 0x%X\n"), pItem->pIDL, hExistingItem);
							RemoveBufferedShellItemNamespace(pItem->pIDL);
							ILFree(pItem->pIDL);
							delete pItem;
							//DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
							continue;
						}
					}
					if(!(properties.disabledEvents & deItemInsertionEvents)) {
						VARIANT_BOOL cancel = VARIANT_FALSE;
						Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&pItem->pIDL), &cancel);
						if(cancel != VARIANT_FALSE) {
							ATLTRACE2(shtvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X as requested by client app\n"), pItem->pIDL);
							RemoveBufferedShellItemNamespace(pItem->pIDL);
							ILFree(pItem->pIDL);
							delete pItem;
							//DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
							continue;
						}
					}

					LONG hasExpando = 0/*ExTVw::heNo*/;
					if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
						hasExpando = -2/*ExTVw::heAuto*/;
					}
					if((pItem->hasSubItems == hsiYes) || (pItem->hasSubItems == hsiMaybe)) {
						hasExpando = 1/*heYes*/;
					}

					if(pEnumResult->namespacePIDLToSet) {
						BufferShellItemNamespace(pItem->pIDL, pEnumResult->namespacePIDLToSet);
					}
					hInsertAfter = FastInsertShellItem(parentItem, pItem->pIDL, hInsertAfter, hasExpando);
					if(!hInsertAfter) {
						RemoveBufferedShellItemNamespace(pItem->pIDL);
						ILFree(pItem->pIDL);
					} else {
						#ifdef DEBUG
							ATLASSERT(hInsertAfter);
							ATLASSERT(PIDLToItemHandle(pItem->pIDL, TRUE) == hInsertAfter);
							++itemCount;
						#endif
						if(pItem->pIDL == pEnumResult->pIDLToLabelEdit) {
							hItemToEdit = hInsertAfter;
						}
					}

					delete pItem;
					DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
				}
				itemCount;
				ATLTRACE2(shtvwTraceThreading, 3, TEXT("Added %i out of %i sub-items to item 0x%X\n"), itemCount, DPA_GetPtrCount(pEnumResult->hPIDLBuffer), pEnumResult->hParentItem);
				#ifdef USE_STL
					immediateSubItems.clear();
				#else
					immediateSubItems.RemoveAll();
				#endif

				LPSHELLTREEVIEWITEMDATA pItemDetails = NULL;
				if(!attachedControl.SendMessage(TVM_GETEDITCONTROL, 0, 0)) {
					BOOL sendMessage = TRUE;
					if(pBackgroundEnum->pTargetObject) {
						CComQIPtr<ShellTreeViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
						if(pNamespace) {
							sendMessage = FALSE;
							VARIANT_BOOL autoSort = VARIANT_FALSE;
							pNamespace->get_AutoSortItems(&autoSort);
							if(autoSort != VARIANT_FALSE) {
								pNamespace->SortItems();
							}
						}
					}
					if(sendMessage) {
						pItemDetails = GetItemDetails(pEnumResult->hParentItem);
						if(pItemDetails) {
							if(pItemDetails->managedProperties & stmipSubItemsSorting) {
								ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, FALSE, reinterpret_cast<LPARAM>(LongToHandle(parentItem)))));
							}
						}
					}
				}

				attachedControl.InvalidateRect(NULL);

				DPA_DeleteAllPtrs(pEnumResult->hPIDLBuffer);
				DPA_Destroy(pEnumResult->hPIDLBuffer);

				if(pBackgroundEnum->pTargetObject) {
					if(!pItemDetails) {
						pItemDetails = GetItemDetails(pEnumResult->hParentItem);
					}
					if(pItemDetails) {
						pItemDetails->isLoadingSubItems = FALSE;
					}
					if(pBackgroundEnum->raiseEvents) {
						Raise_ItemEnumerationCompleted(pBackgroundEnum->pTargetObject, VARIANT_FALSE);
					}
					pBackgroundEnum->pTargetObject->Release();
				}
				delete pBackgroundEnum;
				pBackgroundEnum = NULL;

				if(hItemToEdit) {
					attachedControl.SendMessage(TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(hItemToEdit));
				}
			}
		}

		if(!attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(pEnumResult->hParentItem))) {
			// remove the expando
			TVITEMEX item = {0};
			item.mask = TVIF_HANDLE | TVIF_CHILDREN;
			item.hItem = pEnumResult->hParentItem;
			attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
		}

		if(pEnumResult) {
			delete pEnumResult;
		}
		if(pBackgroundEnum) {
			if(pBackgroundEnum->pTargetObject) {
				pBackgroundEnum->pTargetObject->Release();
			}
			delete pBackgroundEnum;
		}
	}
	return 0;
}

LRESULT ShellTreeView::OnTriggerSetElevationShield(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	LONG itemID = static_cast<LONG>(wParam);
	BOOL redraw = FALSE;
	if(properties.pUnifiedImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);

		UINT flags = 0;
		ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &flags)));
		if(lParam) {
			flags |= AII_DISPLAYELEVATIONOVERLAY;
		} else {
			flags &= ~AII_DISPLAYELEVATIONOVERLAY;
		}
		ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, flags)));
		redraw = TRUE;
	}
	if(redraw) {
		HTREEITEM hItem = ItemIDToHandle(itemID);
		RECT rc = {0};
		*reinterpret_cast<HTREEITEM*>(&rc) = hItem;
		if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
			attachedControl.InvalidateRect(&rc);
		}
	}
	return 0;
}

LRESULT ShellTreeView::OnTriggerUpdateIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	if(properties.pUnifiedImageList) {
		LONG itemID = ItemHandleToID(reinterpret_cast<HTREEITEM>(wParam));
		if(itemID != -1) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, static_cast<int>(lParam), 0, NULL, 0)));

			if(!RunTimeHelper::IsCommCtrl6()) {
				UINT flags = 0;
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &flags)));
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, flags | AII_USELEGACYDISPLAYCODE)));
			}
			if(properties.displayElevationShieldOverlays) {
				PCIDLIST_ABSOLUTE pIDL = ItemHandleToPIDL(reinterpret_cast<HTREEITEM>(wParam));
				if(pIDL) {
					ATLVERIFY(SUCCEEDED(GetElevationShieldAsync(itemID, pIDL)));
				}
			}

			RECT rc = {0};
			*reinterpret_cast<HTREEITEM*>(&rc) = reinterpret_cast<HTREEITEM>(wParam);
			if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
				attachedControl.InvalidateRect(&rc);
			}
		}
	} else {
		TVITEMEX item = {0};
		item.hItem = reinterpret_cast<HTREEITEM>(wParam);
		item.iImage = static_cast<int>(lParam);
		item.mask = TVIF_HANDLE | TVIF_IMAGE;
		attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
	}
	ATLTRACE2(shtvwTraceThreading, 3, TEXT("Updated icon of item 0x%X to %i\n"), wParam, lParam);
	return 0;
}

LRESULT ShellTreeView::OnTriggerUpdateSelectedIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	if(properties.pUnifiedImageList) {
		LONG itemID = ItemHandleToID(reinterpret_cast<HTREEITEM>(wParam));
		if(itemID != -1) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SELECTEDSYSTEMICON, NULL, static_cast<int>(lParam), 0, NULL, 0)));

			RECT rc = {0};
			*reinterpret_cast<HTREEITEM*>(&rc) = reinterpret_cast<HTREEITEM>(wParam);
			if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
				attachedControl.InvalidateRect(&rc);
			}
		}
	} else {
		TVITEMEX item = {0};
		item.hItem = reinterpret_cast<HTREEITEM>(wParam);
		item.iSelectedImage = static_cast<int>(lParam);
		item.mask = TVIF_HANDLE | TVIF_SELECTEDIMAGE;
		attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
	}
	ATLTRACE2(shtvwTraceThreading, 3, TEXT("Updated selected icon of item 0x%X to %i\n"), wParam, lParam);
	return 0;
}

LRESULT ShellTreeView::OnTriggerUpdateOverlay(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	TVITEMEX item = {0};
	item.hItem = reinterpret_cast<HTREEITEM>(wParam);
	item.state = INDEXTOOVERLAYMASK(static_cast<int>(lParam));
	item.stateMask = TVIS_OVERLAYMASK;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
	ATLTRACE2(shtvwTraceThreading, 3, TEXT("Updated overlay of item 0x%X to %i\n"), wParam, lParam);
	return 0;
}


HRESULT ShellTreeView::OnInternalAllocateMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	ATLASSERT(wParam > 0);
	return reinterpret_cast<HRESULT>(HeapAlloc(GetProcessHeap(), 0, wParam));
}

HRESULT ShellTreeView::OnInternalFreeMemory(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	ATLASSERT(lParam);
	HeapFree(GetProcessHeap(), 0, reinterpret_cast<LPVOID>(lParam));
	return S_OK;
}

HRESULT ShellTreeView::OnInternalInsertingItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	properties.IDOfItemBeingInserted = static_cast<LONG>(wParam);
	return 0;
}

HRESULT ShellTreeView::OnInternalInsertedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	HTREEITEM hItem = reinterpret_cast<HTREEITEM>(wParam);
	LPTVINSERTSTRUCT pItemData = reinterpret_cast<LPTVINSERTSTRUCT>(lParam);
	if(pItemData->itemex.mask & TVIF_ISASHELLITEM) {
		#ifdef USE_STL
			LPSHELLTREEVIEWITEMDATA pShellItemData = properties.shellItemDataOfInsertedItems.top();
			properties.shellItemDataOfInsertedItems.pop();
			AddItem(hItem, pShellItemData);
			std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter = properties.namespacesOfInsertedItems.find(pShellItemData->pIDL);
			if(iter != properties.namespacesOfInsertedItems.cend()) {
				pShellItemData->pIDLNamespace = iter->second;
				properties.namespacesOfInsertedItems.erase(iter);
			}
		#else
			LPSHELLTREEVIEWITEMDATA pShellItemData = properties.shellItemDataOfInsertedItems.RemoveHead();
			AddItem(hItem, pShellItemData);
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::CPair* pPair = properties.namespacesOfInsertedItems.Lookup(pShellItemData->pIDL);
			if(pPair) {
				pShellItemData->pIDLNamespace = pPair->m_value;
				properties.namespacesOfInsertedItems.RemoveKey(pPair->m_key);
			}
		#endif
	} else if(pItemData->itemex.mask & TVIF_IMAGE) {
		// store icon index
		BOOL setIconIndex = FALSE;
		if(properties.pUnifiedImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(properties.IDOfItemBeingInserted, SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, pItemData->itemex.iImage, 0, NULL, AII_NONSHELLITEM | AII_NOADORNMENT | AII_NOFILETYPEOVERLAY)));
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(properties.IDOfItemBeingInserted, SIIF_SELECTEDNONSHELLITEMICON, NULL, pItemData->itemex.iSelectedImage, 0, NULL, AII_NONSHELLITEM | AII_NOADORNMENT | AII_NOFILETYPEOVERLAY)));
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(properties.IDOfItemBeingInserted, SIIF_EXPANDEDNONSHELLITEMICON, NULL, pItemData->itemex.iExpandedImage, 0, NULL, AII_NONSHELLITEM | AII_NOADORNMENT | AII_NOFILETYPEOVERLAY)));
			setIconIndex = TRUE;
		}
		if(setIconIndex) {
			// set the item's icon index to its ID
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETITEMICONINDEX, reinterpret_cast<WPARAM>(hItem), properties.IDOfItemBeingInserted)));
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETITEMSELECTEDICONINDEX, reinterpret_cast<WPARAM>(hItem), properties.IDOfItemBeingInserted)));
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETITEMEXPANDEDICONINDEX, reinterpret_cast<WPARAM>(hItem), properties.IDOfItemBeingInserted)));
		}
	}
	properties.IDOfItemBeingInserted = 0;
	return 0;
}

HRESULT ShellTreeView::OnInternalFreeItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	HTREEITEM hItem = reinterpret_cast<HTREEITEM>(wParam);
	#ifdef USE_STL
		std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.find(hItem);
		if(iter != properties.items.cend()) {
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(hItem);
		if(pPair) {
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
	#endif
		// handle outstanding background enums
		#ifdef USE_STL
			std::vector<ULONGLONG> tasksToRemove;
			for(std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter2 = properties.backgroundEnums.cbegin(); iter2 != properties.backgroundEnums.cend(); ++iter2) {
				ULONGLONG taskID = iter2->first;
				LPBACKGROUNDENUMERATION pBackgroundEnum = iter2->second;
		#else
			CAtlArray<ULONGLONG> tasksToRemove;
			POSITION p = properties.backgroundEnums.GetStartPosition();
			while(p) {
				CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair2 = properties.backgroundEnums.GetAt(p);
				ULONGLONG taskID = pPair2->m_key;
				LPBACKGROUNDENUMERATION pBackgroundEnum = pPair2->m_value;
				properties.backgroundEnums.GetNext(p);
		#endif
			if(pBackgroundEnum->pTargetObject) {
				CComQIPtr<IShTreeViewItem> pItem = pBackgroundEnum->pTargetObject;
				if(pItem) {
					LONG pIDL = 0;
					pItem->get_FullyQualifiedPIDL(&pIDL);
					if(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDL) == pItemDetails->pIDL) {
						#ifdef USE_STL
							tasksToRemove.push_back(taskID);
						#else
							tasksToRemove.Add(taskID);
						#endif
						pItemDetails->isLoadingSubItems = FALSE;
						if(pBackgroundEnum->raiseEvents) {
							Raise_ItemEnumerationCompleted(pItem, VARIANT_TRUE);
						}
						pBackgroundEnum->pTargetObject->Release();
						pBackgroundEnum->pTargetObject = NULL;
					}
				}
			}
		}     // for/while
		#ifdef USE_STL
			for(size_t i = 0; i < tasksToRemove.size(); ++i) {
		#else
			for(size_t i = 0; i < tasksToRemove.GetCount(); ++i) {
		#endif
			ULONGLONG taskID = tasksToRemove[i];
			LPBACKGROUNDENUMERATION pData = properties.backgroundEnums[taskID];
			#ifdef USE_STL
				properties.backgroundEnums.erase(taskID);
			#else
				properties.backgroundEnums.RemoveKey(taskID);
			#endif
			delete pData;
		}

		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			CComPtr<IShTreeViewItem> pItem;
			ClassFactory::InitShellTreeItem(hItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
			Raise_RemovingItem(pItem);
		}
	#ifdef USE_STL
			LPSHELLTREEVIEWITEMDATA pData = iter->second;
			properties.items.erase(iter);
	#else
			LPSHELLTREEVIEWITEMDATA pData = pPair->m_value;
			properties.items.RemoveKey(hItem);
	#endif
		delete pData;
	}
	if(properties.pUnifiedImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);
		ATLVERIFY(SUCCEEDED(pImgLstPriv->RemoveIconInfo(static_cast<UINT>(lParam))));
	}
	properties.pTreeItems->Reset();

	return 0;
}

HRESULT ShellTreeView::OnInternalUpdatedItemHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	#ifdef USE_STL
		std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.find(reinterpret_cast<HTREEITEM>(wParam));
		if(iter != properties.items.cend()) {
			// re-insert the entry using the new key
			LPSHELLTREEVIEWITEMDATA tmp = iter->second;
			properties.items.erase(iter);
			properties.items[reinterpret_cast<HTREEITEM>(lParam)] = tmp;
		}
	#else
		CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pEntry = properties.items.Lookup(reinterpret_cast<HTREEITEM>(wParam));
		if(pEntry) {
			// re-insert the entry using the new key
			LPSHELLTREEVIEWITEMDATA tmp = pEntry->m_value;
			properties.items.RemoveKey(reinterpret_cast<HTREEITEM>(wParam));
			properties.items[reinterpret_cast<HTREEITEM>(lParam)] = tmp;
		}
	#endif
	return 0;
}

HRESULT ShellTreeView::OnInternalGetShTreeViewItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPDISPATCH* ppItem = reinterpret_cast<LPDISPATCH*>(lParam);
	ATLASSERT_POINTER(ppItem, LPDISPATCH);
	if(!ppItem) {
		return S_FALSE;
	}

	if(IsShellItem(reinterpret_cast<HTREEITEM>(wParam))) {
		ClassFactory::InitShellTreeItem(reinterpret_cast<HTREEITEM>(wParam), this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(ppItem));
	} else {
		*ppItem = NULL;
	}
	return S_OK;
}

HRESULT ShellTreeView::OnInternalRenamedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	BOOL success = TRUE;

	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(reinterpret_cast<HTREEITEM>(wParam));
	if(pItemDetails) {
		if(pItemDetails->managedProperties & stmipRenaming) {
			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
			if(SUCCEEDED(hr)) {
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

				// TODO: handle the pIDL change
				PITEMID_CHILD pIDLNew = NULL;
				success = SUCCEEDED(pParentISF->SetNameOf(GethWndShellUIParentWindow(), pRelativePIDL, reinterpret_cast<LPCWSTR>(lParam), 0, &pIDLNew));
			}
		}
	}
	return (success ? S_OK : S_FALSE);
}

HRESULT ShellTreeView::OnInternalCompareItems(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	#define ResultFromShort(i) ResultFromScode(MAKE_SCODE(SEVERITY_SUCCESS, 0, static_cast<USHORT>(i)))

	ATLASSERT_ARRAYPOINTER(wParam, HTREEITEM, 2);
	ATLASSERT_POINTER(lParam, BOOL);
	if(!wParam || !lParam) {
		return E_INVALIDARG;
	}

	HTREEITEM* pItems = reinterpret_cast<HTREEITEM*>(wParam);
	LPBOOL pWasHandled = reinterpret_cast<LPBOOL>(lParam);

	PCIDLIST_ABSOLUTE pIDL1 = ItemHandleToPIDL(pItems[0]);
	PCIDLIST_ABSOLUTE pIDL2 = ItemHandleToPIDL(pItems[1]);
	if(pIDL1) {
		if(pIDL2) {
			// both items are shell items - let the shell compare them
			*pWasHandled = TRUE;
			// get the parent IShellFolder
			PIDLIST_ABSOLUTE pIDLParent1 = ILCloneFull(pIDL1);
			ATLASSERT_POINTER(pIDLParent1, ITEMIDLIST_ABSOLUTE);
			ILRemoveLastID(pIDLParent1);
			if(ILIsParent(pIDLParent1, pIDL2, TRUE)) {
				// same parent
				CComPtr<IShellFolder> pParentISF = NULL;
				ATLVERIFY(SUCCEEDED(BindToPIDL(pIDLParent1, IID_PPV_ARGS(&pParentISF))));
				ILFree(pIDLParent1);
				return pParentISF->CompareIDs(0, ILFindLastID(pIDL1), ILFindLastID(pIDL2));
			}
			ILFree(pIDLParent1);

			/* TODO: On Vista, try to use SHCompareIDsFull (but check with "My Computer" whether sorting is still
			         right then!) */
			// fallback to the Desktop's IShellFolder interface
			return properties.pDesktopISF->CompareIDs(0, pIDL1, pIDL2);
		} else {
			// only the first item is a shell item
			*pWasHandled = TRUE;
			switch(properties.itemTypeSortOrder) {
				case itsoShellItemsFirst:
					return ResultFromShort(-1);
					break;
				case itsoShellItemsLast:
					return ResultFromShort(1);
					break;
			}
		}
	} else if(pIDL2) {
		// only the second item is a shell item
		*pWasHandled = TRUE;
		switch(properties.itemTypeSortOrder) {
			case itsoShellItemsFirst:
				return ResultFromShort(1);
				break;
			case itsoShellItemsLast:
				return ResultFromShort(-1);
				break;
		}
	} else {
		// no item is a shell item - that's not our party
	}
	return ResultFromShort(0);
}

HRESULT ShellTreeView::OnInternalBeginLabelEdit(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPNMTVDISPINFO pNotificationDetails = reinterpret_cast<LPNMTVDISPINFO>(lParam);
	//ATLASSERT_POINTER(pNotificationDetails, NMTVDISPINFOA);
	ATLASSERT(pNotificationDetails);
	if(!pNotificationDetails) {
		return E_INVALIDARG;
	}
	BOOL cancel = static_cast<BOOL>(wParam);

	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(pNotificationDetails->item.hItem);
	if(pItemDetails) {
		if(pItemDetails->managedProperties & stmipRenaming) {
			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
			if(SUCCEEDED(hr)) {
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

				SFGAOF attributes = SFGAO_CANRENAME | (properties.preselectBasenameOnLabelEdit ? SFGAO_FILESYSTEM | SFGAO_FOLDER : 0);
				hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				if(SUCCEEDED(hr)) {
					cancel = cancel || ((attributes & SFGAO_CANRENAME) == 0);

					HWND hWndEdit = reinterpret_cast<HWND>(SendMessage(pNotificationDetails->hdr.hwndFrom, TVM_GETEDITCONTROL, 0, 0));
					if(hWndEdit) {
						if(properties.limitLabelEditInput) {
							APIWrapper::SHLimitInputEdit(hWndEdit, pParentISF, NULL);
						}

						if(pItemDetails->managedProperties & stmipText) {
							LPWSTR pDisplayName = NULL;
							hr = GetDisplayName(pItemDetails->pIDL, pParentISF, pRelativePIDL, dntEditingName, FALSE, &pDisplayName);
							ATLASSERT(SUCCEEDED(hr));
							if(pDisplayName) {
								::SetWindowTextW(hWndEdit, pDisplayName);
								if(properties.preselectBasenameOnLabelEdit && (attributes & (SFGAO_FILESYSTEM | SFGAO_FOLDER)) == SFGAO_FILESYSTEM) {
									PathRemoveExtensionW(pDisplayName);
									SendMessage(hWndEdit, EM_SETSEL, 0, lstrlenW(pDisplayName));
								}
								CoTaskMemFree(pDisplayName);
							}
						}
					}
				}
			}
		}
	}
	return (cancel ? S_FALSE : S_OK);
}

HRESULT ShellTreeView::OnInternalBeginDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPNMTREEVIEW pNotificationDetails = reinterpret_cast<LPNMTREEVIEW>(lParam);
	//ATLASSERT_POINTER(pNotificationDetails, NMTREEVIEW);
	ATLASSERT(pNotificationDetails);
	if(!pNotificationDetails) {
		return 0;
	}

	LPSHELLTREEVIEWITEMDATA pItemDetails = NULL;
	if(pNotificationDetails->itemNew.mask & TVIF_HANDLE) {
		pItemDetails = GetItemDetails(pNotificationDetails->itemNew.hItem);
	}
	if(pItemDetails) {
		if(properties.handleOLEDragDrop & hoddSourcePart) {
			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
			if(SUCCEEDED(hr)) {
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

				// retrieve the item's data object
				LPDATAOBJECT pDataObject = NULL;
				hr = pParentISF->GetUIObjectOf(GethWndShellUIParentWindow(), 1, &pRelativePIDL, IID_IDataObject, NULL, reinterpret_cast<LPVOID*>(&pDataObject));
				if(SUCCEEDED(hr)) {
					ATLASSUME(pDataObject);
					// retrieve the supported verbs
					SFGAOF attributes = SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE;
					hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
					if(SUCCEEDED(hr)) {
						int supportedEffects = attributes & (SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE);
						if(supportedEffects != 0/*odeNone*/) {
							// create the item container
							ATLASSUME(properties.pAttachedInternalMessageListener);
							EXTVMCREATEITEMCONTAINERDATA containerData = {0};
							containerData.structSize = sizeof(EXTVMCREATEITEMCONTAINERDATA);
							containerData.numberOfItems = 1;
							HTREEITEM* pItems = new HTREEITEM[1];
							pItems[0] = pNotificationDetails->itemNew.hItem;
							containerData.pItems = pItems;
							hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_CREATEITEMCONTAINER, 0, reinterpret_cast<LPARAM>(&containerData));
							delete[] pItems;
							if(SUCCEEDED(hr)) {
								ATLASSUME(containerData.pContainer);

								// finally call OLEDrag
								// TODO: What shall we do with <performedEffects>?
								int performedEffects = 0/*odeNone*/;
								EXTVMOLEDRAGDATA oleDragData = {0};
								oleDragData.structSize = sizeof(EXTVMOLEDRAGDATA);
								oleDragData.useSHDoDragDrop = TRUE;
								oleDragData.useShellDropSource = TRUE;  // TODO: ((properties.handleOLEDragDrop & hoddSourcePartDefaultDropSource) == hoddSourcePartDefaultDropSource)
								oleDragData.pDataObject = pDataObject;
								oleDragData.supportedEffects = supportedEffects;
								oleDragData.pDraggedItems = containerData.pContainer;
								oleDragData.performedEffects = performedEffects;
								hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_OLEDRAG, 0, reinterpret_cast<LPARAM>(&oleDragData));
								if(SUCCEEDED(hr)) {
									performedEffects = oleDragData.performedEffects;
								} else if(oleDragData.useSHDoDragDrop) {
									oleDragData.useSHDoDragDrop = FALSE;
									oleDragData.useShellDropSource = FALSE;
									hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_OLEDRAG, 0, reinterpret_cast<LPARAM>(&oleDragData));
									if(SUCCEEDED(hr)) {
										performedEffects = oleDragData.performedEffects;
									}
								}
							}
							if(containerData.pContainer) {
								containerData.pContainer->Release();
							}
						}
					}
					pDataObject->Release();
				}
			}
		}
	}
	return 0;
}

HRESULT ShellTreeView::OnInternalHandleDragEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	if(!(properties.handleOLEDragDrop & hoddTargetPart)) {
		return S_FALSE;
	}

	LPSHTVMDRAGDROPEVENTDATA pEventDetails = reinterpret_cast<LPSHTVMDRAGDROPEVENTDATA>(lParam);
	ATLASSERT(pEventDetails);
	if(pEventDetails->structSize < sizeof(SHTVMDRAGDROPEVENTDATA)) {
		return S_FALSE;
	}
	ATLASSUME(pEventDetails->pDataObject);
	if(!pEventDetails->pDataObject) {
		return S_FALSE;
	}

	if(wParam != OLEDRAGEVENT_DROP) {
		dragDropStatus.lastKeyState = pEventDetails->keyState;
	}
	CComQIPtr<IDataObject> pDataObject = pEventDetails->pDataObject;
	HTREEITEM hDropTargetBeforeEvent = pEventDetails->hDropTarget;

	#ifdef DEBUG
		CTreeViewCtrl treeview = attachedControl;
		CAtlString itemText;
		treeview.GetItemText(hDropTargetBeforeEvent, itemText);
		ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Handling drag event %i (drop target item: 0x%X - %s)\n"), wParam, hDropTargetBeforeEvent, itemText);

		if(wParam == OLEDRAGEVENT_DRAGENTER) {
			ATLASSERT(dragDropStatus.pCurrentDropTarget == NULL);
			ATLASSERT(dragDropStatus.hCurrentDropTarget == NULL);
		}
	#endif

	CComPtr<IDropTarget> pNewDropTarget = NULL;
	BOOL isSameDropTarget = (hDropTargetBeforeEvent == dragDropStatus.hCurrentDropTarget);
	LPSHELLTREEVIEWITEMDATA pItemDetails = NULL;
	if(hDropTargetBeforeEvent) {
		pItemDetails = GetItemDetails(hDropTargetBeforeEvent);
	}
	if(pItemDetails) {
		if(isSameDropTarget) {
			pNewDropTarget = dragDropStatus.pCurrentDropTarget;
		} else {
			::GetDropTarget(GethWndShellUIParentWindow(), pItemDetails->pIDL, &pNewDropTarget, FALSE);
		}
	}
	if(!pNewDropTarget) {
		pEventDetails->effect = DROPEFFECT_NONE;
	}

	LPDROPTARGET pEnteredDropTarget = NULL;
	if(wParam != OLEDRAGEVENT_DRAGLEAVE) {
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), pNewDropTarget);
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver((wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &pEventDetails->effect)));
			} else {
				ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Entering drop target 0x%X\n"), pNewDropTarget);
				pNewDropTarget->DragEnter(pDataObject, (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &pEventDetails->effect);
				pEnteredDropTarget = pNewDropTarget;
			}
		}
	}

	// fire the event
	properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_FIREDRAGDROPEVENT, wParam, lParam);

	// handle the easy OLEDRAGEVENT_DRAGLEAVE case
	if(wParam == OLEDRAGEVENT_DRAGLEAVE) {
		if(dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragLeave()));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragLeave()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragLeave()));
		}
		dragDropStatus.hCurrentDropTarget = NULL;
		if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETDROPHILITEDITEM, 0, NULL)));
		}
		return S_OK;
	}

	// the other cases are a bit more complicated

	// Did the event handler change the drop target?
	HTREEITEM hDropTargetAfterEvent = pEventDetails->hDropTarget;
	BOOL eventHandlerHasChangedTarget = (hDropTargetAfterEvent != hDropTargetBeforeEvent);
	isSameDropTarget = (hDropTargetAfterEvent == dragDropStatus.hCurrentDropTarget);
	if(eventHandlerHasChangedTarget) {
		// update pNewDropTarget
		if(pNewDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), pNewDropTarget);
			pNewDropTarget->DragLeave();
			pNewDropTarget = NULL;
		}
		if(hDropTargetAfterEvent) {
			pItemDetails = (isSameDropTarget ? pItemDetails : GetItemDetails(hDropTargetAfterEvent));
		} else {
			pItemDetails = NULL;
		}
		if(isSameDropTarget) {
			pNewDropTarget = dragDropStatus.pCurrentDropTarget;
		} else if(pItemDetails) {
			::GetDropTarget(GethWndShellUIParentWindow(), pItemDetails->pIDL, &pNewDropTarget, FALSE);
		}
	}

	// now we don't have to worry about hDropTargetBeforeEvent anymore
	if(isSameDropTarget) {
		if(wParam == OLEDRAGEVENT_DRAGOVER) {
			// just call DragOver
			if(eventHandlerHasChangedTarget) {
				// we entered this target on a previous call of this method, but did not yet call DragOver this time
				if(dragDropStatus.pCurrentDropTarget) {
					ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
					DWORD effect = pEventDetails->effect;
					ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(pEventDetails->keyState, pEventDetails->cursorPosition, &effect)));
				}
			} else {
				// we already called DragOver on this target before firing the event
			}
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragOver()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragOver(reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
		} else {
			if(wParam == OLEDRAGEVENT_DRAGENTER) {
				if(pEventDetails->pDropTargetHelper) {
					ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragEnter()\n"));
					ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragEnter(attachedControl, pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
				}
			}
			// nothing else to do here (will be done below)
		}
	} else {
		// leave the current drop target
		if(dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			dragDropStatus.pCurrentDropTarget->DragLeave();
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(wParam != OLEDRAGEVENT_DRAGENTER && pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragLeave()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragLeave()));
		}

		// enter the new drop target
		dragDropStatus.pCurrentDropTarget = pNewDropTarget;
		if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Entering drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			DWORD effect = pEventDetails->effect;
			dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &effect);
		}
		if(pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragEnter()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragEnter(attachedControl, pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
		}
		if(wParam == OLEDRAGEVENT_DROP && dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			DWORD effect = pEventDetails->effect;
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver((wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &effect)));
		}
	}
	dragDropStatus.hCurrentDropTarget = hDropTargetAfterEvent;

	if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
		if(LOWORD(pEventDetails->effect) != DROPEFFECT_NONE) {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETDROPHILITEDITEM, 0, reinterpret_cast<LPARAM>(pEventDetails->pDropTarget))));
		} else {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETDROPHILITEDITEM, 0, NULL)));
		}
	}

	/* Now we have called DragEnter/DragOver for the correct drop target and left the old drop target if
	 * necessary. We've also made the necessary calls to IDropTargetHelper.
	 * Now we can handle OLEDRAGEVENT_DROP.
	 */

	if(wParam == OLEDRAGEVENT_DROP) {
		if(!(dragDropStatus.lastKeyState & MK_LBUTTON)) {
			// a popup menu is likely to be displayed, so make sure the drag image is already hidden
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::Drop()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
			dragDropStatus.hCurrentDropTarget = NULL;
			if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETDROPHILITEDITEM, 0, NULL)));
			}
		}
		if(dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Dropping on drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->Drop(pDataObject, pEventDetails->keyState, pEventDetails->cursorPosition, &pEventDetails->effect)));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(dragDropStatus.lastKeyState & MK_LBUTTON) {
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shtvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::Drop()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
			dragDropStatus.hCurrentDropTarget = NULL;
			if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETDROPHILITEDITEM, 0, NULL)));
			}
		}
	}
	return S_OK;
}

HRESULT ShellTreeView::OnInternalContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPPOINT pMenuPosition = reinterpret_cast<LPPOINT>(lParam);
	ATLASSERT_POINTER(pMenuPosition, POINT);
	if(!pMenuPosition) {
		return E_INVALIDARG;
	}

	HTREEITEM hContextMenuItem = NULL;
	if((pMenuPosition->x == -1) && (pMenuPosition->y == -1)) {
		hContextMenuItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, TVGN_CARET, 0));
		if(hContextMenuItem) {
			WTL::CRect itemRectangle;
			*reinterpret_cast<HTREEITEM*>(&itemRectangle) = hContextMenuItem;
			if(attachedControl.SendMessage(TVM_GETITEMRECT, TRUE, reinterpret_cast<LPARAM>(&itemRectangle))) {
				*pMenuPosition = itemRectangle.CenterPoint();
			}
		}
	} else {
		EXTVMHITTESTDATA hitTestData = {0};
		hitTestData.structSize = sizeof(EXTVMHITTESTDATA);
		hitTestData.pointToTest = *pMenuPosition;
		HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestData));
		ATLASSERT(SUCCEEDED(hr));
		if(FAILED(hr)) {
			return 0;
		}
		hContextMenuItem = hitTestData.hHitItem;
	}

	if(hContextMenuItem) {
		HTREEITEM pItems[1] = {hContextMenuItem};
		attachedControl.ClientToScreen(pMenuPosition);
		DisplayShellContextMenu(pItems, 1, *pMenuPosition);
	}
	return 0;
}

HRESULT ShellTreeView::OnInternalGetInfoTip(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPWSTR* ppInfoTip = reinterpret_cast<LPWSTR*>(lParam);
	ATLASSERT_POINTER(ppInfoTip, LPWSTR);

	if(ppInfoTip) {
		LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(reinterpret_cast<HTREEITEM>(wParam));
		if(pItemDetails) {
			if(pItemDetails->managedProperties & stmipInfoTip) {
				GetInfoTip(GethWndShellUIParentWindow(), pItemDetails->pIDL, properties.infoTipFlags, ppInfoTip);
			}
		}
	}
	return 0;
}

HRESULT ShellTreeView::OnInternalCustomDraw(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPNMTVCUSTOMDRAW pNotificationDetails = reinterpret_cast<LPNMTVCUSTOMDRAW>(lParam);
	ATLASSERT_POINTER(pNotificationDetails, NMTVCUSTOMDRAW);
	if(!pNotificationDetails) {
		return CDRF_DODEFAULT;
	}

	BOOL colorCompressedItems = FALSE;
	BOOL colorEncryptedItems = FALSE;
	if(properties.colorCompressedItems) {
		// TODO: Buffer this and the colors (see below)?
		SHELLSTATE shellState = {0};
		SHGetSetSettings(&shellState, SSF_SHOWCOMPCOLOR, FALSE);
		colorCompressedItems = shellState.fShowCompColor;
	}
	colorEncryptedItems = properties.colorCompressedItems;

	if(colorCompressedItems || colorEncryptedItems) {
		switch(pNotificationDetails->nmcd.dwDrawStage) {
			case CDDS_PREPAINT:
				return CDRF_NOTIFYITEMDRAW;
			case CDDS_ITEMPREPAINT: {
				if(pNotificationDetails->clrText == static_cast<COLORREF>(SendMessage(pNotificationDetails->nmcd.hdr.hwndFrom, TVM_GETTEXTCOLOR, 0, 0))) {
					LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(reinterpret_cast<HTREEITEM>(pNotificationDetails->nmcd.dwItemSpec));
					if(pItemDetails && !IsSlowItem(const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), TRUE, TRUE)) {
						CComPtr<IShellFolder> pParentISF = NULL;
						PCUITEMID_CHILD pRelativePIDL = NULL;
						HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
						if(SUCCEEDED(hr)) {
							ATLASSUME(pParentISF);
							ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

							SFGAOF attributes = SFGAO_COMPRESSED | SFGAO_ENCRYPTED;
							hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
							if(SUCCEEDED(hr)) {
								if(colorEncryptedItems) {
									if(attributes & SFGAO_ENCRYPTED) {
										COLORREF encryptedColor = RGB(19, 146, 13);
										DWORD dataSize = sizeof(encryptedColor);
										SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer"), TEXT("AltEncryptionColor"), NULL, &encryptedColor, &dataSize);
										pNotificationDetails->clrText = encryptedColor;
										return CDRF_DODEFAULT | CDRF_NEWFONT;
									}
								}
								if(colorCompressedItems) {
									if(attributes & SFGAO_COMPRESSED) {
										COLORREF compressedColor = RGB(0, 0, 255);
										DWORD dataSize = sizeof(compressedColor);
										SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer"), TEXT("AltColor"), NULL, &compressedColor, &dataSize);
										pNotificationDetails->clrText = compressedColor;
										return CDRF_DODEFAULT | CDRF_NEWFONT;
									}
								}
							}
						}
					}
				}
				break;
			}
		}
	}

	return CDRF_DODEFAULT;
}


LRESULT ShellTreeView::OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/)
{
	LPNMTVDISPINFO pNotificationData = reinterpret_cast<LPNMTVDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_GETDISPINFOA) {
			ATLASSERT_POINTER(pNotificationData, NMTVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pNotificationData, NMTVDISPINFOW);
		}
	#endif

	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(pNotificationData->item.hItem);
	if(!pItemDetails) {
		if((properties.IDOfItemBeingInserted != 0) && (pNotificationData->item.lParam == properties.IDOfItemBeingInserted)) {
			#ifdef USE_STL
				pItemDetails = properties.shellItemDataOfInsertedItems.top();
			#else
				pItemDetails = properties.shellItemDataOfInsertedItems.GetHead();
			#endif
		}
	}
	if(!pItemDetails) {
		return 0;
	}

	BOOL needsText = (((pNotificationData->item.mask & TVIF_TEXT) == TVIF_TEXT) && ((pItemDetails->managedProperties & stmipText) == stmipText));
	BOOL needsImage = (((pNotificationData->item.mask & TVIF_IMAGE) == TVIF_IMAGE) && ((pItemDetails->managedProperties & stmipIconIndex) == stmipIconIndex));
	BOOL needsSelectedImage = (((pNotificationData->item.mask & TVIF_SELECTEDIMAGE) == TVIF_SELECTEDIMAGE) && ((pItemDetails->managedProperties & stmipSelectedIconIndex) == stmipSelectedIconIndex));
	BOOL needsAnyImage = (needsImage || needsSelectedImage);
	BOOL needsOverlay = needsAnyImage && properties.loadOverlaysOnDemand && ((pItemDetails->managedProperties & stmipOverlayIndex) == stmipOverlayIndex);
	BOOL needsState = needsAnyImage && (properties.hiddenItemsStyle == hisGhostedOnDemand) && ((pItemDetails->managedProperties & stmipGhosted) == stmipGhosted);
	BOOL needsAnything = (needsText || needsAnyImage || needsOverlay || needsState);

	if(needsAnything) {
		pNotificationData->item.stateMask = 0;
		pNotificationData->item.state = 0;

		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
		if(SUCCEEDED(hr)) {
			CComQIPtr<IShellIconOverlay> pParentISIO;
			SFGAOF attributes = 0;
			if(needsAnyImage) {
				attributes |= SFGAO_FOLDER;
			}
			if(needsState) {
				attributes |= SFGAO_GHOSTED;
				pNotificationData->item.mask |= TVIF_STATE;
				pNotificationData->item.stateMask |= TVIS_CUT;
			}
			if(needsOverlay) {
				pParentISIO = pParentISF;
				if(!pParentISIO) {
					attributes |= SFGAO_LINK | SFGAO_SHARE;
				}
				pNotificationData->item.mask |= TVIF_STATE;
				pNotificationData->item.stateMask |= TVIS_OVERLAYMASK;
			}
			if(attributes != 0) {
				hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				ATLASSERT(SUCCEEDED(hr));
			}

			// finally retrieve the data
			if(needsText) {
				LPWSTR pDisplayName = NULL;
				hr = GetDisplayName(pItemDetails->pIDL, pParentISF, pRelativePIDL, dntDisplayName, FALSE, &pDisplayName);
				ATLASSERT(SUCCEEDED(hr));
				if(pDisplayName) {
					// NOTE: We'll have a small problem here if pDisplayName is shorter and not null-terminated
					if(pNotificationDetails->code == TVN_GETDISPINFOA) {
						CW2A converter(pDisplayName);
						LPSTR pBuffer = converter;
						lstrcpynA(reinterpret_cast<LPNMTVDISPINFOA>(pNotificationData)->item.pszText, pBuffer, pNotificationData->item.cchTextMax);
					} else {
						lstrcpynW(reinterpret_cast<LPNMTVDISPINFOW>(pNotificationData)->item.pszText, pDisplayName, pNotificationData->item.cchTextMax);
					}
					CoTaskMemFree(pDisplayName);
				}
			}

			UINT itemID = 0;
			if(needsAnyImage && properties.pUnifiedImageList) {
				itemID = ItemHandleToID(pNotificationData->item.hItem);
			}
			if(needsImage) {
				if(properties.pUnifiedImageList && itemID != -1) {
					pNotificationData->item.iImage = itemID;
				} else {
					if(attributes & SFGAO_FOLDER) {
						pNotificationData->item.iImage = properties.DEFAULTICON_FOLDER;
					} else {
						pNotificationData->item.iImage = properties.DEFAULTICON_FILE;
					}
				}
			}
			if(needsSelectedImage) {
				if(properties.pUnifiedImageList && itemID != -1) {
					pNotificationData->item.iSelectedImage = itemID;
				} else {
					if(attributes & SFGAO_FOLDER) {
						pNotificationData->item.iSelectedImage = properties.DEFAULTICON_FOLDER;
					} else {
						pNotificationData->item.iSelectedImage = properties.DEFAULTICON_FILE;
					}
				}
			}
			if(needsAnyImage) {
				if(properties.pUnifiedImageList && itemID != -1) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, (attributes & SFGAO_FOLDER ? properties.DEFAULTICON_FOLDER : properties.DEFAULTICON_FILE), 0, NULL, 0)));
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SELECTEDSYSTEMICON, NULL, (attributes & SFGAO_FOLDER ? properties.DEFAULTICON_FOLDER : properties.DEFAULTICON_FILE), 0, NULL, 0)));
				}
				ATLVERIFY(SUCCEEDED(GetIconsAsync(pItemDetails->pIDL, pNotificationData->item.hItem, needsImage, needsSelectedImage)));
			}

			if(needsState) {
				if(attributes & SFGAO_GHOSTED) {
					pNotificationData->item.state |= TVIS_CUT;
				}
			}

			if(needsOverlay) {
				int overlayIndex = OI_ASYNC;
				if(pParentISIO) {
					hr = pParentISIO->GetOverlayIndex(pRelativePIDL, &overlayIndex);
					if(hr == E_PENDING) {
						ATLVERIFY(SUCCEEDED(GetOverlayAsync(pItemDetails->pIDL, pNotificationData->item.hItem)));
					} else if(hr != S_OK) {
						overlayIndex = 0;
					}
				} else {
					if(attributes & SFGAO_LINK) {
						overlayIndex = 2;
					} else if(attributes & SFGAO_SHARE) {
						overlayIndex = 1;
					}
				}
				if(overlayIndex > 0) {
					pNotificationData->item.state |= INDEXTOOVERLAYMASK(overlayIndex);
				}
			}
		}

		pNotificationData->item.mask |= TVIF_DI_SETITEM;
	}

	return 0;
}

LRESULT ShellTreeView::OnItemExpandingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled, DWORD& /*cookie*/, BOOL /*preProcessing*/)
{
	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_ITEMEXPANDINGA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	if((pDetails->action == TVE_EXPAND) && ((pDetails->itemNew.state & TVIS_EXPANDEDONCE) == 0)) {
		return OnFirstTimeItemExpanding(pDetails->itemNew.hItem, tmTimeOutThreading, FALSE, wasHandled);
	}
	return FALSE;
}

LRESULT ShellTreeView::OnItemExpandedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/)
{
	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_ITEMEXPANDEDA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	if(properties.pUnifiedImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);

		if((pDetails->itemNew.mask & TVIF_HANDLE) && pDetails->itemNew.hItem && !IsShellItem(pDetails->itemNew.hItem)) {
			UINT flags = 0;
			LONG itemID = ItemHandleToID(pDetails->itemNew.hItem);
			if(itemID != -1) {
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &flags)));
				if(pDetails->action & TVE_COLLAPSE) {
					flags &= ~AII_USEEXPANDEDICON;
				} else if(pDetails->action & TVE_EXPAND) {
					flags |= AII_USEEXPANDEDICON;
				}
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, flags)));

				if((pDetails->itemOld.mask & TVIF_HANDLE) && pDetails->itemOld.hItem) {
					RECT rc = {0};
					*reinterpret_cast<HTREEITEM*>(&rc) = pDetails->itemOld.hItem;
					if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
						attachedControl.InvalidateRect(&rc);
					}
				}
			}
		}
	}
	return 0;
}

LRESULT ShellTreeView::OnCaretChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/)
{
	LPNMTREEVIEW pDetails = reinterpret_cast<LPNMTREEVIEW>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TVN_SELCHANGEDA) {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTREEVIEWW);
		}
	#endif

	if(properties.pUnifiedImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);

		if(pDetails->itemOld.hItem && !IsShellItem(pDetails->itemOld.hItem)) {
			UINT flags = 0;
			LONG itemID = ItemHandleToID(pDetails->itemOld.hItem);
			if(itemID != -1) {
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &flags)));
				flags &= ~AII_USESELECTEDICON;
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, flags)));

				RECT rc = {0};
				*reinterpret_cast<HTREEITEM*>(&rc) = pDetails->itemOld.hItem;
				if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
					attachedControl.InvalidateRect(&rc);
				}
			}
		}
		if(pDetails->itemNew.hItem && !IsShellItem(pDetails->itemNew.hItem)) {
			UINT flags = 0;
			LONG itemID = ItemHandleToID(pDetails->itemNew.hItem);
			if(itemID != -1) {
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &flags)));
				flags |= AII_USESELECTEDICON;
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, flags)));

				RECT rc = {0};
				*reinterpret_cast<HTREEITEM*>(&rc) = pDetails->itemNew.hItem;
				if(attachedControl.SendMessage(TVM_GETITEMRECT, FALSE, reinterpret_cast<LPARAM>(&rc))) {
					attachedControl.InvalidateRect(&rc);
				}
			}
		}
	}
	return 0;
}


inline HRESULT ShellTreeView::Raise_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedContextMenuSelection(hContextMenu, commandID, isCustomMenuItem);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_ChangedItemPIDL(HTREEITEM itemHandle, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<IShTreeViewItem> pItem;
		ClassFactory::InitShellTreeItem(itemHandle, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		return Fire_ChangedItemPIDL(pItem, *reinterpret_cast<LONG*>(&previousPIDL), *reinterpret_cast<LONG*>(&newPIDL));
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_ChangedNamespacePIDL(IShTreeViewNamespace* pNamespace, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedNamespacePIDL(pNamespace, *reinterpret_cast<LONG*>(&previousPIDL), *reinterpret_cast<LONG*>(&newPIDL));
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_CreatedShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedShellContextMenu(pItems, hContextMenu, minimumCustomCommandID, pCancelMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_CreatingShellContextMenu(LPDISPATCH pItems, ShellContextMenuStyleConstants* pContextMenuStyle, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CreatingShellContextMenu(pItems, pContextMenuStyle, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_DestroyingShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyingShellContextMenu(pItems, hContextMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_InsertedItem(IShTreeViewItem* pTreeItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pTreeItem);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_InsertedNamespace(IShTreeViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedNamespace(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_InsertingItem(OLE_HANDLE parentItemHandle, LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(parentItemHandle, fullyQualifiedPIDL, pCancelInsertion);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_InvokedShellContextMenuCommand(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants usedWindowMode, CommandInvocationFlagsConstants usedInvocationFlags)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InvokedShellContextMenuCommand(pItems, hContextMenu, commandID, usedWindowMode, usedInvocationFlags);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_ItemEnumerationCompleted(LPDISPATCH pNamespace, VARIANT_BOOL aborted)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationCompleted(pNamespace, aborted);
	//}
	//return S_OK;
}

HRESULT ShellTreeView::Raise_ItemEnumerationStarted(LPDISPATCH pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationStarted(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_ItemEnumerationTimedOut(LPDISPATCH pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationTimedOut(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_RemovingItem(IShTreeViewItem* pTreeItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pTreeItem);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_RemovingNamespace(IShTreeViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingNamespace(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellTreeView::Raise_SelectedShellContextMenuItem(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants* pWindowModeToUse, CommandInvocationFlagsConstants* pInvocationFlagsToUse, VARIANT_BOOL* pExecuteCommand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedShellContextMenuItem(pItems, hContextMenu, commandID, pWindowModeToUse, pInvocationFlagsToUse, pExecuteCommand);
	//}
	//return S_OK;
}


void ShellTreeView::ConfigureControl(void)
{
	threadReferenceCounter = -1;
	pThreadObject = NULL;
	HRESULT hr = E_NOINTERFACE;
	APIWrapper::SHGetThreadRef(&pThreadObject, &hr);
	if(FAILED(hr)) {
		threadReferenceCounter = 0;
		APIWrapper::SHCreateThreadRef(&threadReferenceCounter, &pThreadObject, NULL);
		if(pThreadObject) {
			APIWrapper::SHSetThreadRef(pThreadObject, NULL);
		}
	}

	flags.acceptNewTasks = TRUE;
	SetSystemImageLists();
	RegisterForShellNotifications();
	attachedControl.SetTimer(timers.ID_BACKGROUNDENUMTIMEOUT, timers.INT_BACKGROUNDENUMTIMEOUT);
}

BOOL ShellTreeView::OnFirstTimeItemExpanding(HTREEITEM itemHandle, ThreadingMode threadingMode, BOOL checkForDuplicates, BOOL& isShellItem)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(itemHandle);

	if(pItemDetails && !pItemDetails->isLoadingSubItems && !InterlockedCompareExchange(&pItemDetails->waitingForLoadingSubItems, TRUE, FALSE)) {
		/* TODO: We can run into a race condition here and therefore should use a sync mechanism.
		 *       Imagine the item is expanded for the first time. We'll enter this branch here,
		 *       because isLoadingSubItems is FALSE. But the way up to isLoadingSubItems = TRUE
		 *       costs some time, because we are within a slow network. In this time, the client
		 *       app calls EnsureSubItemsAreLoaded, so we'll end up here again, because
		 *       isLoadingSubItems is still FALSE.
		 *       Okay, but won't the call to EnsureSubItemsAreLoaded hang because it is in the
		 *       same thread?
		 * NOTE: Should be fixed now by using waitingForLoadingSubItems.
		 */
		isShellItem = TRUE;

		/* Too often we crash in ItemHandleToNameSpaceEnumSettings with an access violation.
		 * Hopefully using a copy does help.
		 */
		SHELLTREEVIEWITEMDATA itemDetails = *pItemDetails;
		// will be freed by the destructor
		itemDetails.pIDL = ILCloneFull(itemDetails.pIDL);

		// insert sub-items
		CComPtr<IShTreeViewItem> pExpandedItem = NULL;
		BOOL completedSynchronously = TRUE;
		BOOL insertedSubItem = FALSE;
		INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(itemHandle, &itemDetails);
		if(pEnumSettings) {
			HTREEITEM hParentItem = itemHandle;
			OLE_HANDLE parentItem = HandleToLong(hParentItem);
			HTREEITEM hInsertAfter = NULL;

			BOOL doEnum = TRUE;
			BOOL startThread = (threadingMode == tmImmediateThreading);
			if(!startThread) {
				startThread = (threadingMode != tmNoThreading) && IsSlowItem(const_cast<PIDLIST_ABSOLUTE>(itemDetails.pIDL), TRUE, TRUE);
			}

			ClassFactory::InitShellTreeItem(hParentItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pExpandedItem));
			Raise_ItemEnumerationStarted(pExpandedItem);
			pItemDetails->isLoadingSubItems = TRUE;

			if(startThread) {
				// switch to background enumeration immediately
				// TODO: a binding context with STR_GET_ASYNC_HANDLER might help here
				CComPtr<IRunnableTask> pTask;
				HRESULT hr = E_FAIL;
				#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
					if(RunTimeHelperEx::IsWin8()) {
						hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hParentItem, hInsertAfter, itemDetails.pIDL, NULL, NULL, pEnumSettings, itemDetails.pIDLNamespace, checkForDuplicates, &pTask);
					} else {
						hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hParentItem, hInsertAfter, itemDetails.pIDL, NULL, NULL, pEnumSettings, itemDetails.pIDLNamespace, checkForDuplicates, &pTask);
					}
				#else
					hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hParentItem, hInsertAfter, itemDetails.pIDL, NULL, NULL, pEnumSettings, itemDetails.pIDLNamespace, checkForDuplicates, &pTask);
				#endif
				if(SUCCEEDED(hr)) {
					INamespaceEnumTask* pEnumTask = NULL;
					pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
					ATLASSUME(pEnumTask);
					if(pEnumTask) {
						ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pExpandedItem)));
						pEnumTask->Release();
					}
					ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM, GetTickCount())));
					doEnum = FALSE;
					insertedSubItem = TRUE;
					completedSynchronously = FALSE;

					TVITEMEX item = {0};
					item.hItem = hParentItem;
					/* Until revision 541 we effectively did *remove* TVIS_EXPANDEDONCE here. This caused duplicate items
					 * when calling EnsureSubItemsAreLoaded. The code had been introduced in revision 93 and never changed,
					 * so it probably has been a bug. */
					item.state = TVIS_EXPANDED | TVIS_EXPANDEDONCE;
					item.stateMask = TVIS_EXPANDED | TVIS_EXPANDEDONCE | TVIS_EXPANDPARTIAL;
					item.mask = TVIF_HANDLE | TVIF_STATE;
					attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
				}
			}

			BOOL didEnum = FALSE;
			if(doEnum) {
				// get the IShellFolder implementation of <itemHandle>
				CComPtr<IShellFolder> pNamespaceISF = NULL;
				CComPtr<IShellItem> pNamespaceISI = NULL;
				#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
					if(RunTimeHelperEx::IsWin8() && APIWrapper::IsSupported_SHCreateItemFromIDList()) {
						APIWrapper::SHCreateItemFromIDList(itemDetails.pIDL, IID_PPV_ARGS(&pNamespaceISI), NULL);
					} else {
						BindToPIDL(itemDetails.pIDL, IID_PPV_ARGS(&pNamespaceISF));
					}
				#else
					BindToPIDL(itemDetails.pIDL, IID_PPV_ARGS(&pNamespaceISF));
				#endif

				if(pNamespaceISI || pNamespaceISF) {
					DWORD enumerationBegin = GetTickCount();
					CachedEnumSettings enumSettings = CacheEnumSettings(pEnumSettings);
					CComPtr<IRunnableTask> pTask = NULL;
					#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
						if(pNamespaceISI) {
					#else
						if(FALSE) {
					#endif
						// get the enumerator
						CComPtr<IEnumShellItems> pEnum = NULL;
						CComPtr<IBindCtx> pBindContext;
						// Note that STR_ENUM_ITEMS_FLAGS is ignored on Windows Vista and 7! That's why we use this code path only on Windows 8 and newer.
						if(SUCCEEDED(CreateDwordBindCtx(STR_ENUM_ITEMS_FLAGS, enumSettings.enumerationFlags, &pBindContext))) {
							pNamespaceISI->BindToHandler(pBindContext, BHID_EnumItems, IID_PPV_ARGS(&pEnum));
						}
						// will be NULL if access was denied
						if(pEnum) {
							didEnum = TRUE;
							ULONG fetched = 0;
							BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();

							IShellItem* pChildItem = NULL;
							while((pEnum->Next(1, &pChildItem, &fetched) == S_OK) && (fetched > 0)) {
								if(ShouldShowItem(pChildItem, &enumSettings)) {
									BOOL insert = TRUE;
									PIDLIST_ABSOLUTE pIDLSubItem = NULL;
									if(FAILED(CComQIPtr<IPersistIDList>(pChildItem)->GetIDList(&pIDLSubItem)) || !pIDLSubItem) {
										pIDLSubItem = NULL;
										insert = FALSE;
									} else {
										ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);
									}

									if(insert && checkForDuplicates) {
										HTREEITEM hExistingItem = ItemFindImmediateSubItem(static_cast<HTREEITEM>(LongToHandle(parentItem)), pIDLSubItem, FALSE);
										if(hExistingItem) {
											ATLTRACE2(shtvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X because it's a duplicate of item 0x%X\n"), pIDLSubItem, hExistingItem);
											ILFree(pIDLSubItem);
											pIDLSubItem = NULL;
											insert = FALSE;
										}
									}

									if(insert) {
										if(!(properties.disabledEvents & deItemInsertionEvents)) {
											VARIANT_BOOL cancel = VARIANT_FALSE;
											Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
											insert = !VARIANTBOOL2BOOL(cancel);
										}
									}

									if(insert) {
										LONG hasExpando = 0/*ExTVw::heNo*/;
										if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
											hasExpando = -2/*ExTVw::heAuto*/;
										}
										HasSubItems hasSubItems = HasAtLeastOneSubItem(pChildItem, &enumSettings, TRUE);
										if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
											hasExpando = 1/*heYes*/;
										}

										if(itemDetails.pIDLNamespace) {
											BufferShellItemNamespace(pIDLSubItem, itemDetails.pIDLNamespace);
										}
										hInsertAfter = FastInsertShellItem(parentItem, pIDLSubItem, hInsertAfter, hasExpando);
										if(!hInsertAfter) {
											RemoveBufferedShellItemNamespace(pIDLSubItem);
											ILFree(pIDLSubItem);
											pIDLSubItem = NULL;
										}
										#ifdef DEBUG
											else {
												ATLASSERT(hInsertAfter);
												ATLASSERT(PIDLToItemHandle(pIDLSubItem, TRUE) == hInsertAfter);
											}
										#endif
										insertedSubItem = (hInsertAfter != NULL);
									} else if(pIDLSubItem) {
										ILFree(pIDLSubItem);
									}
								}
								pChildItem->Release();
								pChildItem = NULL;

								if(threadingMode == tmTimeOutThreading && GetTickCount() - enumerationBegin >= 200) {
									// switch to background enumeration
									// TODO: If thread B uses an IShellFolder object that has been created in thread A, thread B sometimes blocks thread A.
									LPSTREAM pMarshaledNamespaceSHI = NULL;
									CoMarshalInterThreadInterfaceInStream(IID_IShellItem, pNamespaceISI, &pMarshaledNamespaceSHI);
									HRESULT hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, NULL, hParentItem, hInsertAfter, itemDetails.pIDL, pMarshaledNamespaceSHI, pEnum, pEnumSettings, itemDetails.pIDLNamespace, checkForDuplicates, &pTask);
									if(SUCCEEDED(hr)) {
										break;
									} else {
										pTask = NULL;
									}
								}
							}
						}
					} else {
						// get the enumerator
						CComPtr<IEnumIDList> pEnum = NULL;
						pNamespaceISF->EnumObjects(GethWndShellUIParentWindow(), enumSettings.enumerationFlags, &pEnum);
						// will be NULL if access was denied
						if(pEnum) {
							didEnum = TRUE;
							ULONG fetched = 0;
							BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();

							PITEMID_CHILD pRelativePIDL = NULL;
							while((pEnum->Next(1, &pRelativePIDL, &fetched) == S_OK) && (fetched > 0)) {
								if(ShouldShowItem(pNamespaceISF, pRelativePIDL, &enumSettings)) {
									// build a fully qualified pIDL
									PCIDLIST_ABSOLUTE pIDLClone = ILCloneFull(itemDetails.pIDL);
									ATLASSERT_POINTER(pIDLClone, ITEMIDLIST_ABSOLUTE);
									PIDLIST_ABSOLUTE pIDLSubItem = reinterpret_cast<PIDLIST_ABSOLUTE>(ILAppendID(const_cast<PIDLIST_ABSOLUTE>(pIDLClone), reinterpret_cast<LPCSHITEMID>(pRelativePIDL), TRUE));
									ATLASSERT_POINTER(pIDLSubItem, ITEMIDLIST_ABSOLUTE);

									if(checkForDuplicates) {
										HTREEITEM hExistingItem = ItemFindImmediateSubItem(static_cast<HTREEITEM>(LongToHandle(parentItem)), pIDLSubItem, FALSE);
										if(hExistingItem) {
											ATLTRACE2(shtvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X because it's a duplicate of item 0x%X\n"), pIDLSubItem, hExistingItem);
											ILFree(pIDLSubItem);
											continue;
										}
									}

									BOOL insert = TRUE;
									if(!(properties.disabledEvents & deItemInsertionEvents)) {
										VARIANT_BOOL cancel = VARIANT_FALSE;
										Raise_InsertingItem(parentItem, *reinterpret_cast<LONG*>(&pIDLSubItem), &cancel);
										insert = !VARIANTBOOL2BOOL(cancel);
									}

									if(insert) {
										LONG hasExpando = 0/*ExTVw::heNo*/;
										if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
											hasExpando = -2/*ExTVw::heAuto*/;
										}
										HasSubItems hasSubItems = HasAtLeastOneSubItem(pNamespaceISF, pRelativePIDL, pIDLSubItem, &enumSettings, TRUE);
										if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
											hasExpando = 1/*heYes*/;
										}

										if(itemDetails.pIDLNamespace) {
											BufferShellItemNamespace(pIDLSubItem, itemDetails.pIDLNamespace);
										}
										hInsertAfter = FastInsertShellItem(parentItem, pIDLSubItem, hInsertAfter, hasExpando);
										if(!hInsertAfter) {
											RemoveBufferedShellItemNamespace(pIDLSubItem);
											ILFree(pIDLSubItem);
										}
										#ifdef DEBUG
											else {
												ATLASSERT(hInsertAfter);
												ATLASSERT(PIDLToItemHandle(pIDLSubItem, TRUE) == hInsertAfter);
											}
										#endif
										insertedSubItem = (hInsertAfter != NULL);
									} else {
										ILFree(pIDLSubItem);
									}
								}
								ILFree(pRelativePIDL);

								if(threadingMode == tmTimeOutThreading && GetTickCount() - enumerationBegin >= 200) {
									// switch to background enumeration
									// TODO: If thread B uses an IShellFolder object that has been created in thread A, thread B sometimes blocks thread A.
									LPSTREAM pMarshaledNamespaceSHF = NULL;
									CoMarshalInterThreadInterfaceInStream(IID_IShellFolder, pNamespaceISF, &pMarshaledNamespaceSHF);
									HRESULT hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, NULL, hParentItem, hInsertAfter, itemDetails.pIDL, pMarshaledNamespaceSHF, pEnum, pEnumSettings, itemDetails.pIDLNamespace, checkForDuplicates, &pTask);
									if(SUCCEEDED(hr)) {
										break;
									} else {
										pTask = NULL;
									}
								}
							}
						}
					}
					if(pTask) {
						INamespaceEnumTask* pEnumTask = NULL;
						pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
						ATLASSUME(pEnumTask);
						if(pEnumTask) {
							ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pExpandedItem)));
							pEnumTask->Release();
						}
						ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM, enumerationBegin)));
						insertedSubItem = TRUE;
						completedSynchronously = FALSE;
					}
				}
			}

			if(didEnum) {
				if(itemDetails.managedProperties & stmipSubItemsSorting) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, FALSE, reinterpret_cast<LPARAM>(LongToHandle(parentItem)))));
				}
			}
			pEnumSettings->Release();
			pEnumSettings = NULL;
		}

		if(!insertedSubItem && !attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(itemHandle))) {
			// remove the expando
			TVITEMEX item = {0};
			item.mask = TVIF_HANDLE | TVIF_CHILDREN;
			item.hItem = itemHandle;
			attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
		}

		if(completedSynchronously && pExpandedItem) {
			pItemDetails->isLoadingSubItems = FALSE;
			Raise_ItemEnumerationCompleted(pExpandedItem, VARIANT_FALSE);
		}

		InterlockedCompareExchange(&pItemDetails->waitingForLoadingSubItems, FALSE, TRUE);
	} else {
		isShellItem = (pItemDetails != NULL);
	}

	/* This makes the control recalculate the scrollbars. It "fixes" a bug in DateiCommander where the
	 * vertical scrollbar gets the wrong range. The undocumented message TVMP_CALCSCROLLBARS does not
	 * solve the problem. */
	attachedControl.SendMessage(WM_SETREDRAW, TRUE, 0);
	return FALSE;
}

void ShellTreeView::SetSystemImageLists(void)
{
	if(properties.pAttachedInternalMessageListener) {
		if(properties.useSystemImageList & usilSmallImageList) {
			HIMAGELIST hSmallImageList = NULL;
			if(APIWrapper::IsSupported_SHGetImageList()) {
				APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hSmallImageList), NULL);
			} else {
				SHFILEINFO details = {0};
				hSmallImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
			}
			if(!properties.pUnifiedImageList) {
				UnifiedImageList_CreateInstance(&properties.pUnifiedImageList);
			}
			ATLASSUME(properties.pUnifiedImageList);
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);

			UINT flags = (properties.displayElevationShieldOverlays ? ILOF_DISPLAYELEVATIONSHIELDS : 0);
			if(!RunTimeHelper::IsCommCtrl6()) {
				flags |= ILOF_IGNOREEXTRAALPHA;
			}
			pImgLstPriv->SetOptions(0, flags);
			pImgLstPriv->SetImageList(AIL_SHELLITEMS, hSmallImageList, NULL);
			pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
			HIMAGELIST hImageList = IImageListToHIMAGELIST(properties.pUnifiedImageList);

			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETIMAGELIST, 1/*ilItems*/, reinterpret_cast<LPARAM>(hImageList))));
		}
	}
}


void ShellTreeView::AddItem(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemData)
{
	ATLASSERT(itemHandle);
	ATLASSERT_POINTER(pItemData, SHELLTREEVIEWITEMDATA);

	properties.items[itemHandle] = pItemData;

	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		CComPtr<IShTreeViewItem> pItem;
		ClassFactory::InitShellTreeItem(itemHandle, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		Raise_InsertedItem(pItem);
	}
}

HRESULT ShellTreeView::ApplyManagedProperties(HTREEITEM itemHandle, BOOL dontTouchSubItems/* = FALSE*/)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(itemHandle);
	if(!pItemDetails) {
		return E_INVALIDARG;
	}

	CComPtr<INamespaceEnumSettings> pEnumSettings;
	if(pItemDetails->pIDLNamespace) {
		// return the namespace's enum settings
		CComObject<ShellTreeViewNamespace>* pNamespaceObj = NamespacePIDLToNamespace(pItemDetails->pIDLNamespace, TRUE);
		ATLASSUME(pNamespaceObj);
		ATLVERIFY(SUCCEEDED(pNamespaceObj->get_NamespaceEnumSettings(&pEnumSettings)));
	} else {
		// return the default enum settings
		ATLVERIFY(SUCCEEDED(get_DefaultNamespaceEnumSettings(&pEnumSettings)));
	}
	CachedEnumSettings enumSettings = CacheEnumSettings(pEnumSettings);
	if(!ShouldShowItem(pItemDetails->pIDL, &enumSettings)) {
		// remove the item
		attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(itemHandle));
		return S_OK;
	}

	if(pItemDetails->managedProperties != 0) {
		// update the managed properties
		HRESULT hr = E_FAIL;

		if(attachedControl.IsWindow()) {
			// retrieve the current settings
			TVITEMEX item = {0};
			item.hItem = itemHandle;

			BOOL retrieveStateNow = ((properties.hiddenItemsStyle == hisGhosted) && ((pItemDetails->managedProperties & stmipGhosted) == stmipGhosted));
			retrieveStateNow = retrieveStateNow || (!properties.loadOverlaysOnDemand && ((pItemDetails->managedProperties & stmipOverlayIndex) == stmipOverlayIndex));
			// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
			retrieveStateNow = retrieveStateNow || (!(pItemDetails->managedProperties & stmipIconIndex) && !(pItemDetails->managedProperties & stmipSelectedIconIndex));
			retrieveStateNow = retrieveStateNow || (attachedControl.SendMessage(TVM_GETIMAGELIST, TVSIL_NORMAL, 0) == NULL);

			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			if(retrieveStateNow) {
				ATLVERIFY(SUCCEEDED(::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
			}

			if((pItemDetails->managedProperties & stmipGhosted) == stmipGhosted) {
				if(properties.hiddenItemsStyle == hisGhosted || (properties.hiddenItemsStyle == hisGhostedOnDemand && retrieveStateNow)) {
					// set the 'cut' state here and now
					item.mask |= TVIF_STATE;
					item.stateMask |= TVIS_CUT;

					if(pParentISF && pRelativePIDL) {
						SFGAOF attributes = SFGAO_GHOSTED;
						hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							if(attributes & SFGAO_GHOSTED) {
								item.state |= TVIS_CUT;
							}
						}
					}
				}
			}
			if(pItemDetails->managedProperties & stmipIconIndex) {
				item.iImage = I_IMAGECALLBACK;
				item.mask |= TVIF_IMAGE;
			}
			if(pItemDetails->managedProperties & stmipSelectedIconIndex) {
				item.iSelectedImage = I_IMAGECALLBACK;
				item.mask |= TVIF_SELECTEDIMAGE;
			}
			if(pItemDetails->managedProperties & stmipText) {
				item.pszText = LPSTR_TEXTCALLBACK;
				item.mask |= TVIF_TEXT;
			}
			if((!properties.loadOverlaysOnDemand || retrieveStateNow) && ((pItemDetails->managedProperties & stmipOverlayIndex) == stmipOverlayIndex)) {
				// load the overlay here and now
				item.stateMask |= TVIS_OVERLAYMASK;
				item.mask |= TVIF_STATE;

				if(pParentISF && pRelativePIDL) {
					int overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
					if(overlayIndex > 0) {
						item.state |= INDEXTOOVERLAYMASK(overlayIndex);
					}
				}
			}

			// apply the new settings
			if(attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				hr = S_OK;
			}

			if(!dontTouchSubItems) {
				if(pItemDetails->managedProperties & stmipSubItems) {
					if(!pParentISF) {
						ATLVERIFY(SUCCEEDED(::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
						ATLASSUME(pParentISF);
						ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
					}
					INamespaceEnumSettings* pEnumSettingsToUse = ItemHandleToNameSpaceEnumSettings(itemHandle, pItemDetails);
					ATLASSUME(pEnumSettingsToUse);
					CachedEnumSettings cachedEnumSettings = CacheEnumSettings(pEnumSettingsToUse);
					LONG hasExpando = 0/*heNo*/;
					HasSubItems hasSubItems = HasAtLeastOneSubItem(pParentISF, const_cast<PITEMID_CHILD>(pRelativePIDL), const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), &cachedEnumSettings, TRUE);
					if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
						hasExpando = 1/*heYes*/;
					}
					item.mask = TVIF_HANDLE | TVIF_CHILDREN;
					item.cChildren = hasExpando;

					BOOL hasBeenExpandedOnce = ItemHasBeenExpandedOnce(itemHandle);
					if(hasBeenExpandedOnce) {
						// validate all shell sub-items
						HTREEITEM hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(itemHandle)));
						while(hChildItem) {
							HTREEITEM hChildItemTemp = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hChildItem)));

							LPSHELLTREEVIEWITEMDATA pChildItemDetails = GetItemDetails(hChildItem);
							if(pChildItemDetails) {
								if(!ValidatePIDL(pChildItemDetails->pIDL)) {
									// remove the item
									attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(hChildItem));
								}
							}
							hChildItem = hChildItemTemp;
						}
					}
					BOOL hasChildren = (attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(itemHandle)) != NULL);
					BOOL removedAllSubItems = (hasBeenExpandedOnce && !hasChildren);
					BOOL searchForNewItems = FALSE;
					if((hasExpando == 1/*heYes*/) || ((hasExpando == 0/*heNo*/) && !hasChildren)) {
						if(hasExpando == 0/*heNo*/ || removedAllSubItems) {
							attachedControl.SendMessage(TVM_EXPAND, TVE_COLLAPSE | TVE_COLLAPSERESET, reinterpret_cast<LPARAM>(itemHandle));
							item.mask |= TVIF_STATE;
							item.stateMask = TVIS_EXPANDED | TVIS_EXPANDEDONCE;
							item.state = 0;
						} else {
							searchForNewItems = hasBeenExpandedOnce;
						}
						attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
					}

					if(searchForNewItems) {
						// search for new sub-items
						CComPtr<IRunnableTask> pTask;
						#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
							if(RunTimeHelperEx::IsWin8()) {
								hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), itemHandle, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettingsToUse, pItemDetails->pIDLNamespace, TRUE, &pTask);
							} else {
								hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), itemHandle, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettingsToUse, pItemDetails->pIDLNamespace, TRUE, &pTask);
							}
						#else
							hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), itemHandle, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettingsToUse, pItemDetails->pIDLNamespace, TRUE, &pTask);
						#endif
						if(SUCCEEDED(hr)) {
							INamespaceEnumTask* pEnumTask = NULL;
							pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
							ATLASSUME(pEnumTask);
							if(pEnumTask) {
								CComPtr<IShTreeViewItem> pItem;
								VARIANT v;
								VariantInit(&v);
								v.vt = VT_I4;
								v.lVal = HandleToLong(itemHandle);
								properties.pTreeItems->get_Item(v, stiitHandle, &pItem);
								VariantClear(&v);
								ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pItem)));
								pEnumTask->Release();
							}
							hr = EnqueueTask(pTask, TASKID_ShTvwBackgroundItemEnum, 0, TASK_PRIORITY_TV_BACKGROUNDENUM);
							ATLASSERT(SUCCEEDED(hr));
						}
					}
					pEnumSettingsToUse->Release();
					pEnumSettingsToUse = NULL;
				}

				// has been handled implicitly above with stmipSubItems
				/*if(((pItemDetails->managedProperties & stmipSubItemsSorting) == stmipSubItemsSorting) && SUCCEEDED(innerResult)) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, FALSE, reinterpret_cast<LPARAM>(itemHandle))));
				}*/
			}
		}
		return hr;
	}

	return S_OK;
}

void ShellTreeView::BufferShellItemData(LPSHELLTREEVIEWITEMDATA pItemData)
{
	ATLASSERT_POINTER(pItemData, SHELLTREEVIEWITEMDATA);

	#ifdef USE_STL
		properties.shellItemDataOfInsertedItems.push(pItemData);
	#else
		properties.shellItemDataOfInsertedItems.AddTail(pItemData);
	#endif
}

HTREEITEM ShellTreeView::FastInsertShellItem(OLE_HANDLE hParentItem, PIDLIST_ABSOLUTE pIDL, HTREEITEM hInsertAfter, LONG hasExpando)
{
	properties.itemDataForFastInsertion.parentItem = hParentItem;
	properties.itemDataForFastInsertion.insertAfter = HandleToLong((hInsertAfter ? hInsertAfter : TVI_LAST));
	properties.itemDataForFastInsertion.hasExpando = hasExpando;
	properties.itemDataForFastInsertion.isGhosted = VARIANT_FALSE;
	properties.itemDataForFastInsertion.overlayIndex = 0;

	BOOL noImages = (attachedControl.SendMessage(TVM_GETIMAGELIST, TVSIL_NORMAL, 0) == NULL);
	BOOL retrieveOverlayNow = !properties.loadOverlaysOnDemand || noImages;
	BOOL retrieveStateNow = (properties.hiddenItemsStyle == hisGhosted) || (noImages && properties.hiddenItemsStyle == hisGhostedOnDemand);
	if(retrieveOverlayNow || retrieveStateNow) {
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		ATLVERIFY(SUCCEEDED(::SHBindToParent(pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
		ATLASSUME(pParentISF);
		ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

		if(pParentISF && pRelativePIDL) {
			if(retrieveStateNow) {
				// set the 'cut' state here and now
				SFGAOF attributes = SFGAO_GHOSTED;
				HRESULT hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				ATLASSERT(SUCCEEDED(hr));
				if(SUCCEEDED(hr)) {
					properties.itemDataForFastInsertion.isGhosted = BOOL2VARIANTBOOL(attributes & SFGAO_GHOSTED);
				}
			}
			if(retrieveOverlayNow) {
				// load the overlay here and now
				properties.itemDataForFastInsertion.overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
			}
		}
	}

	LPSHELLTREEVIEWITEMDATA pItemData = new SHELLTREEVIEWITEMDATA();
	pItemData->pIDL = pIDL;
	pItemData->managedProperties = properties.defaultManagedItemProperties;
	BufferShellItemData(pItemData);

	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_ADDITEM, TRUE, reinterpret_cast<LPARAM>(&properties.itemDataForFastInsertion));
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		return properties.itemDataForFastInsertion.hInsertedItem;
	} else {
		// clean-up buffer
		#ifdef USE_STL
			ATLASSERT(pItemData == properties.shellItemDataOfInsertedItems.top());
			properties.shellItemDataOfInsertedItems.pop();
		#else
			#ifdef DEBUG
				ATLASSERT(pItemData == properties.shellItemDataOfInsertedItems.RemoveHead());
			#else
				properties.shellItemDataOfInsertedItems.RemoveHeadNoReturn();
			#endif
		#endif
		delete pItemData;
	}
	return NULL;
}

LPSHELLTREEVIEWITEMDATA ShellTreeView::GetItemDetails(HTREEITEM itemHandle)
{
	LPSHELLTREEVIEWITEMDATA pItemData = NULL;
	//__try {
		#ifdef USE_STL
			std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.find(itemHandle);
			if(iter != properties.items.cend()) {
				pItemData = iter->second;
			}
		#else
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemHandle);
			if(pPair) {
				pItemData = pPair->m_value;
			}
		#endif
	//} __except(GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION ? EXCEPTION_EXECUTE_HANDLER : EXCEPTION_CONTINUE_SEARCH) {
	//	pItemData = NULL;
	//}
	return pItemData;
}

BOOL ShellTreeView::IsShellItem(HTREEITEM itemHandle)
{
	#ifdef USE_STL
		std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.find(itemHandle);
		if(iter != properties.items.cend()) {
			return TRUE;
		}
	#else
		CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemHandle);
		if(pPair) {
			return TRUE;
		}
	#endif
	return FALSE;
}

LONG ShellTreeView::ItemHandleToID(HTREEITEM itemHandle)
{
	LONG itemID = -1;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_ITEMHANDLETOID, reinterpret_cast<WPARAM>(itemHandle), reinterpret_cast<LPARAM>(&itemID));
	if(SUCCEEDED(hr)) {
		return itemID;
	}
	return -1;
}

HTREEITEM ShellTreeView::ItemIDToHandle(LONG itemID)
{
	HTREEITEM hItem = NULL;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_IDTOITEMHANDLE, itemID, reinterpret_cast<LPARAM>(&hItem));
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		return hItem;
	}
	return NULL;
}

PCIDLIST_ABSOLUTE ShellTreeView::ItemHandleToPIDL(HTREEITEM itemHandle)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(itemHandle);
	return (pItemDetails ? pItemDetails->pIDL : NULL);
}

UINT ShellTreeView::ItemHandlesToPIDLs(HTREEITEM* pItemHandles, UINT itemCount, PIDLIST_ABSOLUTE*& ppIDLs, BOOL keepNonShellItems)
{
	ATLASSERT_ARRAYPOINTER(pItemHandles, HTREEITEM, itemCount);
	if(!pItemHandles) {
		ppIDLs = NULL;
		return 0;
	}

	ppIDLs = new PIDLIST_ABSOLUTE[itemCount];
	UINT pIDLCount = 0;
	for(UINT item = 0; item < itemCount; ++item) {
		LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(pItemHandles[item]);
		if(pItemDetails) {
			ppIDLs[pIDLCount++] = const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL);
		} else if(keepNonShellItems) {
			ppIDLs[pIDLCount++] = NULL;
		}
	}
	if(pIDLCount < itemCount) {
		// truncate
		PIDLIST_ABSOLUTE* p = ppIDLs;
		ppIDLs = new PIDLIST_ABSOLUTE[pIDLCount];
		CopyMemory(ppIDLs, p, pIDLCount * sizeof(PIDLIST_ABSOLUTE));
		delete[] p;
	}
	return pIDLCount;
}

CComObject<ShellTreeViewNamespace>* ShellTreeView::ItemHandleToNamespace(HTREEITEM itemHandle)
{
	return NamespacePIDLToNamespace(ItemHandleToNamespacePIDL(itemHandle), TRUE);
}

PCIDLIST_ABSOLUTE ShellTreeView::ItemHandleToNamespacePIDL(HTREEITEM itemHandle)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(itemHandle);
	return (pItemDetails ? pItemDetails->pIDLNamespace : NULL);
}

INamespaceEnumSettings* ShellTreeView::ItemHandleToNameSpaceEnumSettings(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemDetails/* = NULL*/)
{
	INamespaceEnumSettings* pEnumSettings = NULL;

	if(!pItemDetails) {
		pItemDetails = GetItemDetails(itemHandle);
	}
	if(pItemDetails) {
		if(pItemDetails->managedProperties & stmipSubItems) {
			if(pItemDetails->pIDLNamespace) {
				// return the namespace's enum settings
				CComObject<ShellTreeViewNamespace>* pNamespaceObj = NamespacePIDLToNamespace(pItemDetails->pIDLNamespace, TRUE);
				ATLASSUME(pNamespaceObj);
				ATLVERIFY(SUCCEEDED(pNamespaceObj->get_NamespaceEnumSettings(&pEnumSettings)));
			} else {
				// return the default enum settings
				ATLVERIFY(SUCCEEDED(get_DefaultNamespaceEnumSettings(&pEnumSettings)));
			}
		} else {
			// sub-items are not managed by the control
		}
	}
	return pEnumSettings;
}

HTREEITEM ShellTreeView::PIDLToItemHandle(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	#ifdef USE_STL
		if(exactMatch) {
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(iter->second->pIDL == pIDL) {
					return iter->first;
				}
			}
		} else {
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(ILIsEqual(iter->second->pIDL, pIDL)) {
					return iter->first;
				}
			}
		}
	#else
		if(exactMatch) {
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(pPair->m_value->pIDL == pIDL) {
					return pPair->m_key;
				}
				properties.items.GetNext(p);
			}
		} else {
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(ILIsEqual(pPair->m_value->pIDL, pIDL)) {
					return pPair->m_key;
				}
				properties.items.GetNext(p);
			}
		}
	#endif
	return NULL;
}

HTREEITEM ShellTreeView::ParsingNameToItemHandle(LPCWSTR pParsingName)
{
	#ifdef USE_STL
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LPWSTR pCompareTo = NULL;
			if(SUCCEEDED(GetDisplayName(iter->second->pIDL, properties.pDesktopISF, reinterpret_cast<PCUITEMID_CHILD>(iter->second->pIDL), dntParsingName, TRUE, &pCompareTo))) {
				if(lstrcmpiW(pParsingName, pCompareTo) == 0) {
					CoTaskMemFree(pCompareTo);
					return iter->first;
				}
				CoTaskMemFree(pCompareTo);
			}
		}
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LPWSTR pCompareTo = NULL;
			if(SUCCEEDED(GetDisplayName(pPair->m_value->pIDL, properties.pDesktopISF, reinterpret_cast<PCUITEMID_CHILD>(pPair->m_value->pIDL), dntParsingName, TRUE, &pCompareTo))) {
				if(lstrcmpiW(pParsingName, pCompareTo) == 0) {
					CoTaskMemFree(pCompareTo);
					return pPair->m_key;
				}
				CoTaskMemFree(pCompareTo);
			}
			properties.items.GetNext(p);
		}
	#endif
	return NULL;
}

HRESULT ShellTreeView::VariantToItemHandles(VARIANT& items, HTREEITEM*& pItemHandles, UINT& itemCount)
{
	HRESULT hr = S_OK;
	itemCount = 0;
	pItemHandles = NULL;
	switch(items.vt) {
		case VT_DISPATCH:
			// invalid arg - raise VB runtime error 380
			hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			if(items.pdispVal) {
				CComQIPtr<IShTreeViewItem> pShTvwItem = items.pdispVal;
				if(pShTvwItem) {
					// a single ShellTreeViewItem object
					pItemHandles = static_cast<HTREEITEM*>(HeapAlloc(GetProcessHeap(), 0, sizeof(HTREEITEM)));
					OLE_HANDLE h = NULL;
					hr = pShTvwItem->get_Handle(&h);
					pItemHandles[0] = static_cast<HTREEITEM>(LongToHandle(h));
					itemCount = 1;
				} else {
					CComQIPtr<IShTreeViewItems> pShTvwItems = items.pdispVal;
					if(pShTvwItems) {
						// a ShellTreeViewItems collection
						CComQIPtr<IEnumVARIANT> pEnumerator = pShTvwItems;
						LONG c = 0;
						pShTvwItems->Count(&c);
						itemCount = c;
						if(itemCount > 0 && pEnumerator) {
							pItemHandles = static_cast<HTREEITEM*>(HeapAlloc(GetProcessHeap(), 0, itemCount * sizeof(HTREEITEM)));
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							for(UINT i = 0; i < itemCount && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
								pItemHandles[i] = NULL;
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pShTvwItem = item.pdispVal;
									if(pShTvwItem) {
										OLE_HANDLE h = NULL;
										pShTvwItem->get_Handle(&h);
										pItemHandles[i] = static_cast<HTREEITEM>(LongToHandle(h));
									}
								}
								VariantClear(&item);
							}
							hr = S_OK;
						}
					} else {
						// let ExplorerTreeView try its luck
						EXTVMGETITEMHANDLESDATA data = {0};
						data.structSize = sizeof(EXTVMGETITEMHANDLESDATA);
						data.pVariant = &items;
						hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_GETITEMHANDLESFROMVARIANT, 0, reinterpret_cast<LPARAM>(&data));
						itemCount = data.numberOfItems;
						pItemHandles = data.pItemHandles;
					}
				}
			}
			break;
		case VT_ARRAY | VT_VARIANT:
			hr = E_FAIL;
			if(items.parray) {
				LONG lBound = 0;
				LONG uBound = 0;
				if((SafeArrayGetLBound(items.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(items.parray, 1, &uBound) == S_OK)) {
					// now that we have the bounds, iterate the array
					VARIANT element;
					BOOL isValid = TRUE;
					itemCount = (uBound >= lBound ? (uBound - lBound + 1) : 0);
					if(itemCount > 0) {
						pItemHandles = static_cast<HTREEITEM*>(HeapAlloc(GetProcessHeap(), 0, itemCount * sizeof(HTREEITEM)));
						itemCount = 0;
						for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
							pItemHandles[itemCount] = NULL;
							if(SafeArrayGetElement(items.parray, &i, &element) == S_OK) {
								if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_I4))) {
									pItemHandles[itemCount++] = static_cast<HTREEITEM>(LongToHandle(element.lVal));
								} else {
									// the element is invalid - abort
									isValid = FALSE;
								}
								VariantClear(&element);
							} else {
								isValid = FALSE;
							}
						}
					}
					hr = S_OK;
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			hr = VariantChangeType(&v, &items, 0, VT_I4);
			if(SUCCEEDED(hr)) {
				pItemHandles = static_cast<HTREEITEM*>(HeapAlloc(GetProcessHeap(), 0, sizeof(HTREEITEM)));
				pItemHandles[0] = static_cast<HTREEITEM>(LongToHandle(items.lVal));
				itemCount = 1;
			} else {
				// invalid arg - raise VB runtime error 380
				hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
	}

	return hr;
}

HRESULT ShellTreeView::SHBindToParent(HTREEITEM itemHandle, REFIID requiredInterface, LPVOID* ppInterfaceImpl, PCUITEMID_CHILD* pPIDLLast, PCIDLIST_ABSOLUTE* pPIDL/* = NULL*/)
{
	LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(itemHandle);
	if(pItemDetails) {
		if(pPIDL) {
			*pPIDL = pItemDetails->pIDL;
		}
		return ::SHBindToParent(pItemDetails->pIDL, requiredInterface, ppInterfaceImpl, pPIDLLast);
	}
	return E_INVALIDARG;
}

BOOL ShellTreeView::ValidateItem(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemDetails/* = NULL*/)
{
	if(!pItemDetails) {
		pItemDetails = GetItemDetails(itemHandle);
	}
	if(pItemDetails) {
		return ValidatePIDL(pItemDetails->pIDL);
	}
	return FALSE;
}

void ShellTreeView::UpdateItemPIDL(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemDetails, PIDLIST_ABSOLUTE pIDLNew)
{
	UINT itemIDsUpdatedItem = ILCount(pIDLNew);
	if(itemIDsUpdatedItem == ILCount(pItemDetails->pIDL)) {
		PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
		pItemDetails->pIDL = ILCloneFull(pIDLNew);
		if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
			Raise_ChangedItemPIDL(itemHandle, tmp, pItemDetails->pIDL);
		}
		ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
	} else {
		ATLASSERT(itemIDsUpdatedItem < ILCount(pItemDetails->pIDL));
		// replace the first <itemIDsUpdatedItem> item IDs of the item's pIDL with the new pIDL
		// find the start of the "tail"
		PCUIDLIST_RELATIVE pIDLTail = pItemDetails->pIDL;
		for(UINT i = 0; i < itemIDsUpdatedItem; ++i) {
			pIDLTail = ILNext(pIDLTail);
		}
		ATLASSERT_POINTER(pIDLTail, ITEMIDLIST_RELATIVE);
		// NOTE: ILCombine doesn't change pIDLNew
		PIDLIST_ABSOLUTE pIDLItemNew = ILCombine(pIDLNew, pIDLTail);
		ATLASSERT(ILCount(pIDLItemNew) == ILCount(pItemDetails->pIDL));
		PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
		pItemDetails->pIDL = pIDLItemNew;
		if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
			Raise_ChangedItemPIDL(itemHandle, tmp, pItemDetails->pIDL);
		}
		ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
	}
}

void ShellTreeView::RemoveCollapsedSubTree(HTREEITEM hParentItem)
{
	if(hParentItem) {
		TVITEMEX item = {0};
		item.hItem = hParentItem;
		item.mask = TVIF_HANDLE | TVIF_STATE | TVIF_CHILDREN;
		if(attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			if((item.cChildren == 1) && !(item.state & (TVIS_EXPANDED | TVIS_EXPANDPARTIAL))) {
				// item is collapsed, so maybe remove its sub-items
				LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(hParentItem);
				if(pItemDetails && (pItemDetails->managedProperties & stmipSubItems)) {
					attachedControl.SendMessage(TVM_EXPAND, TVE_COLLAPSE | TVE_COLLAPSERESET, reinterpret_cast<LPARAM>(hParentItem));
				}
			} else {
				HTREEITEM hSubItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
				while(hSubItem) {
					RemoveCollapsedSubTree(hSubItem);
					hSubItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hSubItem)));
				}
			}
		}
	}
}


void ShellTreeView::AddNamespace(PCIDLIST_ABSOLUTE pIDL, CComObject<ShellTreeViewNamespace>* pNamespace)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pNamespace, CComObject<ShellTreeViewNamespace>);

	properties.namespaces[pIDL] = pNamespace;

	if(!(properties.disabledEvents & deNamespaceInsertionEvents)) {
		CComQIPtr<IShTreeViewNamespace> pNamespaceItf = pNamespace;
		Raise_InsertedNamespace(pNamespaceItf);
	}
}

void ShellTreeView::BufferShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem, PCIDLIST_ABSOLUTE pIDLNamespace)
{
	ATLASSERT_POINTER(pIDLItem, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);

	properties.namespacesOfInsertedItems[pIDLItem] = pIDLNamespace;
}

void ShellTreeView::RemoveBufferedShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem)
{
	#ifdef USE_STL
		std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter = properties.namespacesOfInsertedItems.find(pIDLItem);
		if(iter != properties.namespacesOfInsertedItems.cend()) {
			properties.namespacesOfInsertedItems.erase(iter);
		}
	#else
		properties.namespacesOfInsertedItems.RemoveKey(pIDLItem);
	#endif
}

CComObject<ShellTreeViewNamespace>* ShellTreeView::IndexToNamespace(UINT index)
{
	#ifdef USE_STL
		if((index >= 0) && (index < properties.namespaces.size())) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin();
			std::advance(iter, index);
			if(iter != properties.namespaces.cend()) {
				return iter->second;
			}
		}
	#else
		if((index >= 0) && (index < properties.namespaces.GetCount())) {
			POSITION p = properties.namespaces.GetStartPosition();
			for(UINT i = 0; i < index; ++i) {
				properties.namespaces.GetNext(p);
			}
			return properties.namespaces.GetValueAt(p);
		}
	#endif
	return NULL;
}

PCIDLIST_ABSOLUTE ShellTreeView::IndexToNamespacePIDL(UINT index)
{
	#ifdef USE_STL
		if((index >= 0) && (index < properties.namespaces.size())) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin();
			std::advance(iter, index);
			if(iter != properties.namespaces.cend()) {
				return iter->first;
			}
		}
	#else
		if((index >= 0) && (index < properties.namespaces.GetCount())) {
			POSITION p = properties.namespaces.GetStartPosition();
			for(UINT i = 0; i < index; ++i) {
				properties.namespaces.GetNext(p);
			}
			return properties.namespaces.GetKeyAt(p);
		}
	#endif
	return NULL;
}

LONG ShellTreeView::NamespacePIDLToIndex(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	LONG index = 0;
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDL);
			if(iter != properties.namespaces.cend()) {
				index = static_cast<LONG>(std::distance(properties.namespaces.cbegin(), iter));
				return index;
			}
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter, ++index) {
				if(ILIsEqual(iter->first, pIDL)) {
					return index;
				}
			}
		}
	#else
		if(exactMatch) {
			for(POSITION p = properties.namespaces.GetStartPosition(); p != NULL; properties.namespaces.GetNext(p), ++index) {
				if(properties.namespaces.GetKeyAt(p) == pIDL) {
					return index;
				}
			}
		} else {
			for(POSITION p = properties.namespaces.GetStartPosition(); p != NULL; properties.namespaces.GetNext(p), ++index) {
				if(ILIsEqual(properties.namespaces.GetKeyAt(p), pIDL)) {
					return index;
				}
			}
		}
	#endif
	return -1;
}

CComObject<ShellTreeViewNamespace>* ShellTreeView::NamespacePIDLToNamespace(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDL);
			if(iter != properties.namespaces.cend()) {
				return iter->second;
			}
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				if(ILIsEqual(iter->first, pIDL)) {
					return iter->second;
				}
			}
		}
	#else
		if(exactMatch) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.Lookup(pIDL);
			if(pPair) {
				return pPair->m_value;
			}
		} else {
			POSITION p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				if(ILIsEqual(pPair->m_key, pIDL)) {
					return pPair->m_value;
				}
				properties.namespaces.GetNext(p);
			}
		}
	#endif
	return NULL;
}

CComObject<ShellTreeViewNamespace>* ShellTreeView::NamespaceParsingNameToNamespace(LPWSTR pParsingName)
{
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			LPWSTR pCompareTo = NULL;
			if(SUCCEEDED(GetDisplayName(iter->first, properties.pDesktopISF, reinterpret_cast<PCUITEMID_CHILD>(iter->first), dntParsingName, TRUE, &pCompareTo))) {
				if(lstrcmpiW(pParsingName, pCompareTo) == 0) {
					CoTaskMemFree(pCompareTo);
					return iter->second;
				}
				CoTaskMemFree(pCompareTo);
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			LPWSTR pCompareTo = NULL;
			if(SUCCEEDED(GetDisplayName(pPair->m_key, properties.pDesktopISF, reinterpret_cast<PCUITEMID_CHILD>(pPair->m_key), dntParsingName, TRUE, &pCompareTo))) {
				if(lstrcmpiW(pParsingName, pCompareTo) == 0) {
					CoTaskMemFree(pCompareTo);
					return pPair->m_value;
				}
				CoTaskMemFree(pCompareTo);
			}
			properties.namespaces.GetNext(p);
		}
	#endif
	return NULL;
}

HRESULT ShellTreeView::RemoveNamespace(PCIDLIST_ABSOLUTE pIDLNamespace, BOOL exactMatch, BOOL removeItemsFromTreeView)
{
	HRESULT hr = S_OK;

	// handle outstanding background enums
	#ifdef USE_STL
		std::vector<ULONGLONG> tasksToRemove;
		for(std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter = properties.backgroundEnums.cbegin(); iter != properties.backgroundEnums.cend(); ++iter) {
			ULONGLONG taskID = iter->first;
			LPBACKGROUNDENUMERATION pBackgroundEnum = iter->second;
	#else
		CAtlArray<ULONGLONG> tasksToRemove;
		POSITION p = properties.backgroundEnums.GetStartPosition();
		while(p) {
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair = properties.backgroundEnums.GetAt(p);
			ULONGLONG taskID = pPair->m_key;
			LPBACKGROUNDENUMERATION pBackgroundEnum = pPair->m_value;
			properties.backgroundEnums.GetNext(p);
	#endif
		if(pBackgroundEnum->pTargetObject) {
			CComQIPtr<IShTreeViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
			if(pNamespace) {
				LONG pIDL = 0;
				pNamespace->get_FullyQualifiedPIDL(&pIDL);
				if(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDL) == pIDLNamespace) {
					#ifdef USE_STL
						tasksToRemove.push_back(taskID);
					#else
						tasksToRemove.Add(taskID);
					#endif
					if(pBackgroundEnum->raiseEvents) {
						Raise_ItemEnumerationCompleted(pNamespace, VARIANT_TRUE);
					}
					pBackgroundEnum->pTargetObject->Release();
					pBackgroundEnum->pTargetObject = NULL;
				}
			}
		}
	}
	#ifdef USE_STL
		for(size_t i = 0; i < tasksToRemove.size(); ++i) {
	#else
		for(size_t i = 0; i < tasksToRemove.GetCount(); ++i) {
	#endif
		ULONGLONG taskID = tasksToRemove[i];
		LPBACKGROUNDENUMERATION pData = properties.backgroundEnums[taskID];
		#ifdef USE_STL
			properties.backgroundEnums.erase(taskID);
		#else
			properties.backgroundEnums.RemoveKey(taskID);
		#endif
		delete pData;
	}

	// collect the affected items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
		if(exactMatch) {
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(iter->second->pIDLNamespace == pIDLNamespace) {
					itemsToRemove.push_back(iter->first);
				}
			}
		} else {
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(ILIsEqual(iter->second->pIDLNamespace, pIDLNamespace)) {
					itemsToRemove.push_back(iter->first);
				}
			}
		}
	#else
		CAtlList<HTREEITEM> itemsToRemove;
		p = properties.items.GetStartPosition();
		if(exactMatch) {
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(pPair->m_value->pIDLNamespace == pIDLNamespace) {
					itemsToRemove.AddTail(pPair->m_key);
				}
				properties.items.GetNext(p);
			}
		} else {
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(ILIsEqual(pPair->m_value->pIDLNamespace, pIDLNamespace)) {
					itemsToRemove.AddTail(pPair->m_key);
				}
				properties.items.GetNext(p);
			}
		}
	#endif

	if(removeItemsFromTreeView) {
		// remove the items from the treeview
		SelectNewCaret(itemsToRemove, NULL);
		RemoveItems(itemsToRemove);
		hr = S_OK;
	} else {
		// transfer the items to normal items - free the data
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToRemove.size(); ++i) {
				std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.find(itemsToRemove[i]);
				if(iter != properties.items.cend()) {
					HTREEITEM hItem = iter->first;
		#else
			p = itemsToRemove.GetHeadPosition();
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemsToRemove.GetAt(p));
				if(pPair) {
					HTREEITEM hItem = pPair->m_key;
		#endif
				if(!(properties.disabledEvents & deItemDeletionEvents)) {
					CComPtr<IShTreeViewItem> pItem;
					ClassFactory::InitShellTreeItem(hItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
					Raise_RemovingItem(pItem);
				}
		#ifdef USE_STL
					LPSHELLTREEVIEWITEMDATA pData = iter->second;
					properties.items.erase(iter);
					delete pData;
				}
			}
		#else
					LPSHELLTREEVIEWITEMDATA pData = pPair->m_value;
					properties.items.RemoveKey(hItem);
					delete pData;
				}
				itemsToRemove.GetNext(p);
			}
		#endif
		hr = S_OK;
	}

	// remove the namespace
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDLNamespace);
			if(iter != properties.namespaces.cend()) {
				if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
					CComQIPtr<IShTreeViewNamespace> pNamespaceItf = iter->second;
					Raise_RemovingNamespace(pNamespaceItf);
				}
				CComObject<ShellTreeViewNamespace>* pElement = iter->second;
				properties.namespaces.erase(iter);
				pElement->Release();
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDLNamespace));
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				if(ILIsEqual(iter->first, pIDLNamespace)) {
					if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
						CComQIPtr<IShTreeViewNamespace> pNamespaceItf = iter->second;
						Raise_RemovingNamespace(pNamespaceItf);
					}
					CComObject<ShellTreeViewNamespace>* pElement = iter->second;
					PIDLIST_ABSOLUTE pIDL = const_cast<PIDLIST_ABSOLUTE>(iter->first);
					properties.namespaces.erase(iter);
					pElement->Release();
					ILFree(pIDL);
					break;
				}
			}
		}
	#else
		if(exactMatch) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.Lookup(pIDLNamespace);
			CComObject<ShellTreeViewNamespace>* pElement = NULL;
			if(pPair) {
				if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
					CComQIPtr<IShTreeViewNamespace> pNamespaceItf = pPair->m_value;
					Raise_RemovingNamespace(pNamespaceItf);
				}
				pElement = pPair->m_value;
			}
			properties.namespaces.RemoveKey(pIDLNamespace);
			if(pElement) {
				pElement->Release();
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDLNamespace));
		} else {
			p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				if(ILIsEqual(pPair->m_key, pIDLNamespace)) {
					if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
						CComQIPtr<IShTreeViewNamespace> pNamespaceItf = pPair->m_value;
						Raise_RemovingNamespace(pNamespaceItf);
					}
					CComObject<ShellTreeViewNamespace>* pElement = pPair->m_value;
					PIDLIST_ABSOLUTE pIDL = const_cast<PIDLIST_ABSOLUTE>(pPair->m_key);
					properties.namespaces.RemoveAtPos(p);
					pElement->Release();
					ILFree(pIDL);
					break;
				}
				properties.namespaces.GetNext(p);
			}
		}
	#endif
	return hr;
}

HRESULT ShellTreeView::RemoveAllNamespaces(BOOL removeItemsFromTreeView)
{
	HRESULT hr = S_OK;
	#ifdef USE_STL
		while(properties.namespaces.size() > 0) {
			hr = RemoveNamespace(properties.namespaces.cbegin()->first, TRUE, removeItemsFromTreeView);
	#else
		while(properties.namespaces.GetCount() > 0) {
			hr = RemoveNamespace(properties.namespaces.GetKeyAt(properties.namespaces.GetStartPosition()), TRUE, removeItemsFromTreeView);
	#endif
		if(FAILED(hr)) {
			break;
		}
	}
	return hr;
}


HRESULT ShellTreeView::CreateShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, HMENU& hMenu)
{
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);

	DestroyShellContextMenu();

	contextMenuData.bufferedCtrlDown = FALSE;
	contextMenuData.bufferedShiftDown = FALSE;
	UINT flags = CMF_NORMAL;
	if(GetKeyState(VK_SHIFT) & 0x8000) {
		flags |= CMF_EXTENDEDVERBS;
		contextMenuData.bufferedShiftDown = TRUE;
	}
	if(GetKeyState(VK_CONTROL) & 0x8000) {
		contextMenuData.bufferedCtrlDown = TRUE;
	}

	CComPtr<IShellFolder> pParentISF;
	HRESULT hr = BindToPIDL(pIDLNamespace, IID_PPV_ARGS(&pParentISF));
	//ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		ATLASSUME(pParentISF);
		if(FAILED(pParentISF->CreateViewObject(GethWndShellUIParentWindow(), IID_PPV_ARGS(&contextMenuData.pIContextMenu)))) {
			contextMenuData.pIContextMenu = NULL;
		}
	}

	//ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		//ATLASSUME(contextMenuData.pIContextMenu);

		CComObject<ShellTreeViewNamespace>* pNamespace = NamespacePIDLToNamespace(pIDLNamespace, TRUE);
		ATLASSUME(pNamespace);
		pNamespace->QueryInterface<IDispatch>(&contextMenuData.pContextMenuItems);

		// allow customization of 'flags'
		ShellContextMenuStyleConstants contextMenuStyle = static_cast<ShellContextMenuStyleConstants>(flags);
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_CreatingShellContextMenu(contextMenuData.pContextMenuItems, &contextMenuStyle, &cancel);
		flags = static_cast<UINT>(contextMenuStyle);

		if((cancel == VARIANT_FALSE) && contextMenuData.currentShellContextMenu.CreatePopupMenu()) {
			// fill the menu
			if(contextMenuData.pIContextMenu) {
				hr = contextMenuData.pIContextMenu->QueryContextMenu(contextMenuData.currentShellContextMenu, 0, MIN_CONTEXTMENUEXTENSION_CMDID, MAX_CONTEXTMENUEXTENSION_CMDID, flags);
				ATLASSERT(SUCCEEDED(hr));
			}
			if(SUCCEEDED(hr)) {
				// allow custom items
				cancel = VARIANT_FALSE;
				Raise_CreatedShellContextMenu(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), MAX_CONTEXTMENUEXTENSION_CMDID + 1, &cancel);
				if(cancel == VARIANT_FALSE) {
					PrettifyMenu(contextMenuData.currentShellContextMenu);

					hMenu = contextMenuData.currentShellContextMenu;
					hr = S_OK;
				}
			}
		}
	}

	return hr;
}

HRESULT ShellTreeView::CreateShellContextMenu(HTREEITEM* pItems, UINT itemCount, UINT predefinedFlags, HMENU& hMenu)
{
	ATLASSERT_ARRAYPOINTER(pItems, HTREEITEM, itemCount);
	if(!pItems) {
		return E_POINTER;
	}

	DestroyShellContextMenu();

	PIDLIST_ABSOLUTE* ppIDLs = NULL;
	UINT pIDLCount = ItemHandlesToPIDLs(pItems, itemCount, ppIDLs, FALSE);
	if(pIDLCount == 0) {
		return S_OK;
	}
	ATLASSERT_ARRAYPOINTER(ppIDLs, PIDLIST_ABSOLUTE, pIDLCount);
	if(!ppIDLs) {
		return E_OUTOFMEMORY;
	}
	BOOL singleNamespace = HaveSameParent(ppIDLs, pIDLCount);

	HRESULT hr;
	contextMenuData.bufferedCtrlDown = FALSE;
	contextMenuData.bufferedShiftDown = FALSE;
	UINT flags = predefinedFlags;
	if(GetKeyState(VK_SHIFT) & 0x8000) {
		flags |= CMF_EXTENDEDVERBS;
		contextMenuData.bufferedShiftDown = TRUE;
	}
	if(GetKeyState(VK_CONTROL) & 0x8000) {
		contextMenuData.bufferedCtrlDown = TRUE;
	}

	if(singleNamespace) {
		CComPtr<IShellFolder> pParentISF = NULL;
		hr = ::SHBindToParent(ppIDLs[0], IID_PPV_ARGS(&pParentISF), NULL);
		ATLTRACE2(shtvwTraceContextMenus, 3, TEXT("SHBindToParent returned 0x%X\n"), hr);
		//ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			PUITEMID_CHILD* pRelativePIDLs = new PUITEMID_CHILD[pIDLCount];
			ATLASSERT_ARRAYPOINTER(pRelativePIDLs, PUITEMID_CHILD, pIDLCount);
			if(pRelativePIDLs) {
				for(UINT i = 0; i < pIDLCount; ++i) {
					pRelativePIDLs[i] = ILFindLastID(ppIDLs[i]);
					ATLASSERT_POINTER(pRelativePIDLs[i], ITEMID_CHILD);
				}

				if(attachedControl.GetStyle() & TVS_EDITLABELS) {
					SFGAOF attributes = SFGAO_CANRENAME;
					hr = pParentISF->GetAttributesOf(pIDLCount, pRelativePIDLs, &attributes);
					if(SUCCEEDED(hr)) {
						if(attributes & SFGAO_CANRENAME) {
							flags |= CMF_CANRENAME;
						}
					}
				}
				if(FAILED(pParentISF->GetUIObjectOf(GethWndShellUIParentWindow(), pIDLCount, pRelativePIDLs, IID_IContextMenu, NULL, reinterpret_cast<LPVOID*>(&contextMenuData.pIContextMenu)))) {
					contextMenuData.pIContextMenu = NULL;
				}
				hr = S_OK;
				if(pRelativePIDLs) {
					delete[] pRelativePIDLs;
				}
			} else {
				hr = E_OUTOFMEMORY;
			}
		}
	} else {
		CComQIPtr<IContextMenuCB> pContextMenuCB = this;
		hr = GetCrossNamespaceContextMenu(GethWndShellUIParentWindow(), ppIDLs, pIDLCount, pContextMenuCB, &contextMenuData.pMultiNamespaceParentISF, &contextMenuData.pIContextMenu);
		if(FAILED(hr)) {
			contextMenuData.pIContextMenu = NULL;
		}
		ATLASSUME(contextMenuData.pMultiNamespaceParentISF);
		SFGAOF attributes = SFGAO_CANRENAME;
		if(SUCCEEDED(contextMenuData.pMultiNamespaceParentISF->GetAttributesOf(pIDLCount, reinterpret_cast<PCUITEMID_CHILD_ARRAY>(ppIDLs), &attributes))) {
			if(attributes & SFGAO_CANRENAME) {
				flags |= CMF_CANRENAME;
			}
		}
		ATLVERIFY(SUCCEEDED(contextMenuData.pMultiNamespaceParentISF->GetUIObjectOf(GethWndShellUIParentWindow(), pIDLCount, reinterpret_cast<PCUITEMID_CHILD_ARRAY>(ppIDLs), IID_IDataObject, 0, reinterpret_cast<LPVOID*>(&contextMenuData.pMultiNamespaceDataObject))));
		ATLASSUME(contextMenuData.pMultiNamespaceDataObject);
	}

	if(ppIDLs) {
		delete[] ppIDLs;
		ppIDLs = NULL;
	}

	//ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		//ATLASSUME(contextMenuData.pIContextMenu);

		if(RunTimeHelper::IsVista()) {
			flags |= CMF_ITEMMENU;
		}

		ATLASSUME(properties.pAttachedInternalMessageListener);
		EXTVMCREATEITEMCONTAINERDATA containerData = {0};
		containerData.structSize = sizeof(EXTVMCREATEITEMCONTAINERDATA);
		containerData.numberOfItems = itemCount;
		containerData.pItems = pItems;
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_CREATEITEMCONTAINER, 0, reinterpret_cast<LPARAM>(&containerData))));
		if(containerData.pContainer) {
			contextMenuData.pContextMenuItems = containerData.pContainer;
		}

		// allow customization of 'flags'
		ShellContextMenuStyleConstants contextMenuStyle = static_cast<ShellContextMenuStyleConstants>(flags);
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_CreatingShellContextMenu(contextMenuData.pContextMenuItems, &contextMenuStyle, &cancel);
		flags = static_cast<UINT>(contextMenuStyle);

		if((cancel == VARIANT_FALSE) && contextMenuData.currentShellContextMenu.CreatePopupMenu()) {
			// fill the menu
			if(contextMenuData.pIContextMenu) {
				hr = contextMenuData.pIContextMenu->QueryContextMenu(contextMenuData.currentShellContextMenu, 0, MIN_CONTEXTMENUEXTENSION_CMDID, MAX_CONTEXTMENUEXTENSION_CMDID, flags);
				ATLASSERT(SUCCEEDED(hr));
			}
			if(SUCCEEDED(hr)) {
				// allow custom items
				cancel = VARIANT_FALSE;
				Raise_CreatedShellContextMenu(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), MAX_CONTEXTMENUEXTENSION_CMDID + 1, &cancel);
				if(cancel == VARIANT_FALSE) {
					PrettifyMenu(contextMenuData.currentShellContextMenu);

					hMenu = contextMenuData.currentShellContextMenu;
					hr = S_OK;
				}
			}
		}
	}

	return hr;
}

HRESULT ShellTreeView::DisplayShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, POINT& position)
{
	HMENU hMenu;
	HRESULT hr = CreateShellContextMenu(pIDLNamespace, hMenu);
	if(SUCCEEDED(hr)) {
		// display the menu
		UINT chosenCommand = contextMenuData.currentShellContextMenu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, position.x, position.y, attachedControl);
		if(chosenCommand != 0) {
			VARIANT_BOOL executeCommand = VARIANT_TRUE;
			WindowModeConstants windowModeToUse = wmDefault;
			CommandInvocationFlagsConstants invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(cifAllowAsynchronousExecution | cifLogUsage);
			if(contextMenuData.bufferedShiftDown/*flags & CMF_EXTENDEDVERBS*/) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifSHIFTKeyPressed);
			}
			if(contextMenuData.bufferedCtrlDown) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifCTRLKeyPressed);
			}
			Raise_SelectedShellContextMenuItem(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, &windowModeToUse, &invocationFlagsToUse, &executeCommand);
			if(executeCommand != VARIANT_FALSE) {
				if((chosenCommand >= MIN_CONTEXTMENUEXTENSION_CMDID) && (chosenCommand <= MAX_CONTEXTMENUEXTENSION_CMDID)) {
					chosenCommand -= MIN_CONTEXTMENUEXTENSION_CMDID;
					ATLASSUME(contextMenuData.pIContextMenu);

					// TODO: move this code to its own method?
					BOOL invoke = TRUE;
					BOOL invoked = FALSE;
					TCHAR pVerb[MAX_CONTEXTMENUVERB];
					ZeroMemory(pVerb, MAX_CONTEXTMENUVERB * sizeof(TCHAR));
					hr = contextMenuData.pIContextMenu->GetCommandString(chosenCommand, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pVerb), MAX_CONTEXTMENUVERB);

					if(invoke) {
						BOOL isSubItemOfNewMenu = FALSE;
						if(properties.autoEditNewItems) {
							if(FALSE/*IsWindows2000()*/) {
								// On Windows 2000, the "new folder" command doesn't have a command string.
								isSubItemOfNewMenu = TRUE;
							} else {
								if(!contextMenuData.clickedMenu.IsNull()) {
									UINT mayBeNewFolderCommand = contextMenuData.clickedMenu.GetMenuItemID(0);
									if(mayBeNewFolderCommand != static_cast<UINT>(-1)) {
										mayBeNewFolderCommand -= MIN_CONTEXTMENUEXTENSION_CMDID;
										TCHAR pBuffer[MAX_CONTEXTMENUVERB];
										ZeroMemory(pBuffer, MAX_CONTEXTMENUVERB * sizeof(TCHAR));
										contextMenuData.pIContextMenu->GetCommandString(mayBeNewFolderCommand, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pBuffer), MAX_CONTEXTMENUVERB);
										if(lstrcmpi(pBuffer, CMDSTR_NEWFOLDER) == 0) {
											isSubItemOfNewMenu = TRUE;
										}
									}
								}
							}
						}

						properties.timeOfLastItemCreatorInvocation = 0;
						if(isSubItemOfNewMenu) {
							properties.timeOfLastItemCreatorInvocation = GetTickCount();
						}

						// execute the command
						CMINVOKECOMMANDINFOEX invokeData = {0};
						invokeData.cbSize = sizeof(CMINVOKECOMMANDINFOEX);
						invokeData.fMask = invocationFlagsToUse | CMIC_MASK_PTINVOKE;
						invokeData.hwnd = GethWndShellUIParentWindow();
						if(!invokeData.hwnd) {
							invokeData.hwnd = attachedControl;
						}
						invokeData.lpVerb = MAKEINTRESOURCEA(chosenCommand);
						invokeData.nShow = windowModeToUse;
						invokeData.ptInvoke = position;
						hr = contextMenuData.pIContextMenu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&invokeData));
						/* Release the context menu object as early as possible, while the pointer hopefully is still valid,
						   even if it was e.g. the delete command. */
						contextMenuData.pIContextMenu->Release();
						contextMenuData.pIContextMenu = NULL;
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							invoked = TRUE;
						}
					}

					if(invoked) {
						// don't forget that we've subtracted MIN_CONTEXTMENUEXTENSION_CMDID above
						chosenCommand += MIN_CONTEXTMENUEXTENSION_CMDID;
						Raise_InvokedShellContextMenuCommand(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, windowModeToUse, invocationFlagsToUse);
					}
				}
			}
		}
		ATLVERIFY(SUCCEEDED(DestroyShellContextMenu()));
	}

	return hr;
}

HRESULT ShellTreeView::DisplayShellContextMenu(HTREEITEM* pItems, UINT itemCount, POINT& position)
{
	HMENU hMenu;
	HRESULT hr = CreateShellContextMenu(pItems, itemCount, CMF_NORMAL, hMenu);
	if(SUCCEEDED(hr) && contextMenuData.currentShellContextMenu.IsMenu()) {
		// display the menu
		attachedControl.SendMessage(TVM_SELECTITEM, TVGN_DROPHILITE, reinterpret_cast<LPARAM>(pItems[0]));
		UINT chosenCommand = contextMenuData.currentShellContextMenu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, position.x, position.y, attachedControl);
		attachedControl.SendMessage(TVM_SELECTITEM, TVGN_DROPHILITE, 0);

		if(chosenCommand != 0) {
			VARIANT_BOOL executeCommand = VARIANT_TRUE;
			WindowModeConstants windowModeToUse = wmDefault;
			CommandInvocationFlagsConstants invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(cifAllowAsynchronousExecution | cifLogUsage);
			if(contextMenuData.bufferedShiftDown/*flags & CMF_EXTENDEDVERBS*/) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifSHIFTKeyPressed);
			}
			if(contextMenuData.bufferedCtrlDown) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifCTRLKeyPressed);
			}
			Raise_SelectedShellContextMenuItem(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, &windowModeToUse, &invocationFlagsToUse, &executeCommand);
			if(executeCommand != VARIANT_FALSE) {
				if((chosenCommand >= MIN_CONTEXTMENUEXTENSION_CMDID) && (chosenCommand <= MAX_CONTEXTMENUEXTENSION_CMDID)) {
					chosenCommand -= MIN_CONTEXTMENUEXTENSION_CMDID;
					ATLASSUME(contextMenuData.pIContextMenu);

					// TODO: move this code to its own method?
					BOOL invoke = TRUE;
					BOOL invoked = FALSE;
					TCHAR pVerb[MAX_CONTEXTMENUVERB];
					ZeroMemory(pVerb, MAX_CONTEXTMENUVERB * sizeof(TCHAR));
					hr = contextMenuData.pIContextMenu->GetCommandString(chosenCommand, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pVerb), MAX_CONTEXTMENUVERB);
					if(SUCCEEDED(hr)) {
						if(lstrcmpi(pVerb, TEXT("rename")) == 0) {
							invoked = (attachedControl.SendMessage(TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(pItems[0])) != NULL);
							invoke = FALSE;
						} else if(contextMenuData.pMultiNamespaceDataObject && lstrcmpi(pVerb, TEXT("properties")) == 0) {
							// WORKAROUND: we must execute the properties verb ourselves
							invoked = SUCCEEDED(SHMultiFileProperties(contextMenuData.pMultiNamespaceDataObject, 0));
							invoke = !invoked;   // on error give IContextMenu a try
						}
					}

					if(invoke) {
						// execute the command
						CMINVOKECOMMANDINFOEX invokeData = {0};
						invokeData.cbSize = sizeof(CMINVOKECOMMANDINFOEX);
						invokeData.fMask = invocationFlagsToUse | CMIC_MASK_PTINVOKE;
						invokeData.hwnd = attachedControl;
						invokeData.lpVerb = MAKEINTRESOURCEA(chosenCommand);
						invokeData.nShow = windowModeToUse;
						invokeData.ptInvoke = position;
						CHAR pBufferA[MAX_PATH];
						WCHAR pBufferW[MAX_PATH];
						if(itemCount == 1) {
							/* Some files don't like being called without a working dir. For an example see mails of
							   C. Lütgens from August 2010. */
							LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(pItems[0]);
							if(pItemDetails) {
								PIDLIST_RELATIVE pIDL = ILClone(pItemDetails->pIDL);
								ILRemoveLastID(pIDL);
								if(pIDL) {
									// TODO: Prefer SHGetPathFromIDListEx.
									SHGetPathFromIDListA(reinterpret_cast<PIDLIST_ABSOLUTE>(pIDL), pBufferA);
									SHGetPathFromIDListW(reinterpret_cast<PIDLIST_ABSOLUTE>(pIDL), pBufferW);
									if(lstrlenA(pBufferA) > 0) {
										invokeData.lpDirectory = pBufferA;
									}
									if(lstrlenW(pBufferW) > 0) {
										invokeData.lpDirectoryW = pBufferW;
									}
									ILFree(pIDL);
								}
							}
						}
						hr = contextMenuData.pIContextMenu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&invokeData));
						/* Release the context menu object as early as possible, while the pointer hopefully is still valid,
						   even if it was e.g. the delete command. */
						contextMenuData.pIContextMenu->Release();
						contextMenuData.pIContextMenu = NULL;
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							invoked = TRUE;
							if(lstrcmpi(pVerb, TEXT("cut")) == 0) {
								// TODO: maybe we should raise an event CutShellItem or similar?
							}
						}
					}

					if(invoked) {
						// don't forget that we've subtracted MIN_CONTEXTMENUEXTENSION_CMDID above
						chosenCommand += MIN_CONTEXTMENUEXTENSION_CMDID;
						Raise_InvokedShellContextMenuCommand(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, windowModeToUse, invocationFlagsToUse);
					}
				}
			}
		}
		ATLVERIFY(SUCCEEDED(DestroyShellContextMenu()));
	}

	return hr;
}

HRESULT ShellTreeView::InvokeDefaultShellContextMenuCommand(HTREEITEM* pItems, UINT itemCount)
{
	HMENU hMenu;
	HRESULT hr = CreateShellContextMenu(pItems, itemCount, CMF_DEFAULTONLY, hMenu);
	if(SUCCEEDED(hr)) {
		ATLASSERT(contextMenuData.currentShellContextMenu.IsMenu());
		// get the default item
		UINT chosenCommand = contextMenuData.currentShellContextMenu.GetMenuDefaultItem(0, GMDI_GOINTOPOPUPS);

		if(chosenCommand != 0) {
			VARIANT_BOOL executeCommand = VARIANT_TRUE;
			WindowModeConstants windowModeToUse = wmDefault;
			CommandInvocationFlagsConstants invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(cifAllowAsynchronousExecution | cifLogUsage);
			if(contextMenuData.bufferedShiftDown/*flags & CMF_EXTENDEDVERBS*/) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifSHIFTKeyPressed);
			}
			if(contextMenuData.bufferedCtrlDown) {
				invocationFlagsToUse = static_cast<CommandInvocationFlagsConstants>(invocationFlagsToUse | cifCTRLKeyPressed);
			}
			Raise_SelectedShellContextMenuItem(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, &windowModeToUse, &invocationFlagsToUse, &executeCommand);
			if(executeCommand != VARIANT_FALSE) {
				if((chosenCommand >= MIN_CONTEXTMENUEXTENSION_CMDID) && (chosenCommand <= MAX_CONTEXTMENUEXTENSION_CMDID)) {
					chosenCommand -= MIN_CONTEXTMENUEXTENSION_CMDID;
					ATLASSUME(contextMenuData.pIContextMenu);

					// TODO: move this code to its own method?
					BOOL invoke = TRUE;
					BOOL invoked = FALSE;
					TCHAR pVerb[MAX_CONTEXTMENUVERB];
					ZeroMemory(pVerb, MAX_CONTEXTMENUVERB * sizeof(TCHAR));
					hr = contextMenuData.pIContextMenu->GetCommandString(chosenCommand, GCS_VERB, NULL, reinterpret_cast<LPSTR>(pVerb), MAX_CONTEXTMENUVERB);
					if(SUCCEEDED(hr)) {
						if(lstrcmpi(pVerb, TEXT("rename")) == 0) {
							invoked = (attachedControl.SendMessage(TVM_EDITLABEL, 0, reinterpret_cast<LPARAM>(pItems[0])) != NULL);
							invoke = FALSE;
						} else if(contextMenuData.pMultiNamespaceDataObject && lstrcmpi(pVerb, TEXT("properties")) == 0) {
							// WORKAROUND: we must execute the properties verb ourselves
							invoked = SUCCEEDED(SHMultiFileProperties(contextMenuData.pMultiNamespaceDataObject, 0));
							invoke = !invoked;   // on error give IContextMenu a try
						}
					}

					if(invoke) {
						// execute the command
						CMINVOKECOMMANDINFOEX invokeData = {0};
						invokeData.cbSize = sizeof(CMINVOKECOMMANDINFOEX);
						invokeData.fMask = invocationFlagsToUse;
						invokeData.hwnd = attachedControl;
						invokeData.lpVerb = MAKEINTRESOURCEA(chosenCommand);
						invokeData.nShow = windowModeToUse;
						CHAR pBufferA[MAX_PATH];
						WCHAR pBufferW[MAX_PATH];
						if(itemCount == 1) {
							/* Some files don't like being called without a working dir. For an example see mails of
							   C. Lütgens from August 2010. */
							LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(pItems[0]);
							if(pItemDetails) {
								PIDLIST_RELATIVE pIDL = ILClone(pItemDetails->pIDL);
								ILRemoveLastID(pIDL);
								if(pIDL) {
									// TODO: Prefer SHGetPathFromIDListEx.
									SHGetPathFromIDListA(reinterpret_cast<PIDLIST_ABSOLUTE>(pIDL), pBufferA);
									SHGetPathFromIDListW(reinterpret_cast<PIDLIST_ABSOLUTE>(pIDL), pBufferW);
									if(lstrlenA(pBufferA) > 0) {
										invokeData.lpDirectory = pBufferA;
									}
									if(lstrlenW(pBufferW) > 0) {
										invokeData.lpDirectoryW = pBufferW;
									}
									ILFree(pIDL);
								}
							}
						}
						hr = contextMenuData.pIContextMenu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&invokeData));
						/* Release the context menu object as early as possible, while the pointer hopefully is still valid,
						   even if it was e.g. the delete command. */
						contextMenuData.pIContextMenu->Release();
						contextMenuData.pIContextMenu = NULL;
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							invoked = TRUE;
							if(lstrcmpi(pVerb, TEXT("cut")) == 0) {
								// TODO: maybe we should raise an event CutShellItem or similar?
							}
						}
					}

					if(invoked) {
						// don't forget that we've subtracted MIN_CONTEXTMENUEXTENSION_CMDID above
						chosenCommand += MIN_CONTEXTMENUEXTENSION_CMDID;
						Raise_InvokedShellContextMenuCommand(contextMenuData.pContextMenuItems, HandleToLong(static_cast<HMENU>(contextMenuData.currentShellContextMenu)), chosenCommand, windowModeToUse, invocationFlagsToUse);
					}
				}
			}
		}
		ATLVERIFY(SUCCEEDED(DestroyShellContextMenu()));
	}

	return hr;
}


HRESULT ShellTreeView::DeregisterForShellNotifications(void)
{
	if(!attachedControl.IsWindow() || !shellNotificationData.IsRegistered()) {
		return S_OK;
	}

	if(SHChangeNotifyDeregister(shellNotificationData.registrationID)) {
		ATLTRACE2(shtvwTraceAutoUpdate, 3, TEXT("Deregistered 0x%X\n"), shellNotificationData.registrationID);
		return S_OK;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Failed to deregister 0x%X\n"), shellNotificationData.registrationID);
	return E_FAIL;
}

HRESULT ShellTreeView::RegisterForShellNotifications(void)
{
	if(!attachedControl.IsWindow()) {
		return S_OK;
	}

	ATLVERIFY(SUCCEEDED(DeregisterForShellNotifications()));
	if(!properties.processShellNotifications) {
		return S_OK;
	}

	SHChangeNotifyEntry registrationDetails = {0};
	registrationDetails.fRecursive = TRUE;
	PIDLIST_ABSOLUTE pIDLDesktop = GetDesktopPIDL();
	registrationDetails.pidl = pIDLDesktop;
	shellNotificationData.registrationID = static_cast<ULONG>(SHChangeNotifyRegister(attachedControl, SHCNRF_ShellLevel | SHCNRF_InterruptLevel | SHCNRF_NewDelivery, SHCNE_ALLEVENTS, WM_SHCHANGENOTIFY, 1, &registrationDetails));
	ILFree(pIDLDesktop);
	ATLTRACE2(shtvwTraceAutoUpdate, 3, TEXT("SHChangeNotifyRegister returned 0x%X\n"), shellNotificationData.registrationID);
	return (shellNotificationData.registrationID > 0 ? S_OK : E_FAIL);
}


void ShellTreeView::FlushIcons(int iconIndex, BOOL flushOverlays)
{
	BOOL fullFlush = (iconIndex == -1);
	ShTvwManagedItemPropertiesConstants flagsToCheckFor;
	if(fullFlush) {
		flagsToCheckFor = static_cast<ShTvwManagedItemPropertiesConstants>(stmipIconIndex | stmipSelectedIconIndex | stmipOverlayIndex);
		if(properties.useSystemImageList & usilSmallImageList) {
			if(!RunTimeHelperEx::IsWin8()) {
				// flush the system imagelist
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SETIMAGELIST, 1/*ilItems*/, NULL)));
				FlushSystemImageList();
				SetSystemImageLists();
			}
		}

		SHFILEINFO fileInfo = {0};
		SHGetFileInfo(TEXT(".zyxwv12"), FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
		properties.DEFAULTICON_FILE = fileInfo.iIcon;
		SHGetFileInfo(TEXT("folder"), FILE_ATTRIBUTE_DIRECTORY, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
		properties.DEFAULTICON_FOLDER = fileInfo.iIcon;
	} else {
		flagsToCheckFor = static_cast<ShTvwManagedItemPropertiesConstants>(stmipIconIndex | stmipSelectedIconIndex);

		SHFILEINFO fileInfo = {0};
		if(properties.DEFAULTICON_FILE == iconIndex) {
			SHGetFileInfo(TEXT(".zyxwv12"), FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
			properties.DEFAULTICON_FILE = fileInfo.iIcon;
		}
		if(properties.DEFAULTICON_FOLDER == iconIndex) {
			SHGetFileInfo(TEXT("folder"), FILE_ATTRIBUTE_DIRECTORY, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
			properties.DEFAULTICON_FOLDER = fileInfo.iIcon;
		}
	}
	HIMAGELIST hImageList = NULL;
	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_GETIMAGELIST, 1/*ilItems*/, reinterpret_cast<LPARAM>(&hImageList))));

	TVITEMEX item = {0};
	#ifdef USE_STL
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			item.hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			item.hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(pItemDetails->managedProperties & flagsToCheckFor) {
			item.mask = TVIF_HANDLE;
			item.stateMask = 0;
			item.state = 0;
			if(pItemDetails->managedProperties & stmipIconIndex) {
				item.mask |= TVIF_IMAGE;
			}
			if(pItemDetails->managedProperties & stmipSelectedIconIndex) {
				item.mask |= TVIF_SELECTEDIMAGE;
			}
			if(flushOverlays && (pItemDetails->managedProperties & stmipOverlayIndex)) {
				item.mask |= TVIF_STATE;
				item.stateMask = TVIS_OVERLAYMASK;
			}

			BOOL update = fullFlush;
			if(fullFlush) {
				item.iImage = I_IMAGECALLBACK;
				item.iSelectedImage = I_IMAGECALLBACK;
				if(flushOverlays && (pItemDetails->managedProperties & stmipOverlayIndex)) {
					// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
					BOOL retrieveOverlayNow = !properties.loadOverlaysOnDemand || (!(pItemDetails->managedProperties & stmipIconIndex) && !(pItemDetails->managedProperties & stmipSelectedIconIndex)) || !hImageList;
					if(retrieveOverlayNow) {
						// load the overlay here and now
						int overlayIndex = GetOverlayIndex(pItemDetails->pIDL);
						if(overlayIndex > 0) {
							item.state = INDEXTOOVERLAYMASK(overlayIndex);
						}
					}
				}
			} else {
				attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));

				if((pItemDetails->managedProperties & stmipIconIndex) && (item.iImage == iconIndex)) {
					item.iImage = I_IMAGECALLBACK;
					update = TRUE;
				}
				if((pItemDetails->managedProperties & stmipSelectedIconIndex) && (item.iSelectedImage == iconIndex)) {
					item.iSelectedImage = I_IMAGECALLBACK;
					update = TRUE;
				}
				if(flushOverlays && (pItemDetails->managedProperties & stmipOverlayIndex)) {
					// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
					BOOL retrieveOverlayNow = !properties.loadOverlaysOnDemand || (!(pItemDetails->managedProperties & stmipIconIndex) && !(pItemDetails->managedProperties & stmipSelectedIconIndex)) || !hImageList;
					if(retrieveOverlayNow) {
						// load the overlay here and now
						int overlayIndex = GetOverlayIndex(pItemDetails->pIDL);
						if(overlayIndex > 0) {
							item.state = INDEXTOOVERLAYMASK(overlayIndex);
						}
						update = TRUE;
					}
				}
			}

			if(update) {
				attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
			}
		}
	}
}

HTREEITEM ShellTreeView::ItemFindImmediateSubItem(HTREEITEM hParentItem, PIDLIST_ABSOLUTE pIDLToCheck, BOOL simplePIDL)
{
	ATLASSERT_POINTER(pIDLToCheck, ITEMIDLIST_ABSOLUTE);
	if(!pIDLToCheck) {
		return NULL;
	}

	HTREEITEM hChildItem = NULL;
	if(hParentItem) {
		hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CHILD), reinterpret_cast<LPARAM>(hParentItem)));
	} else {
		hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
	}
	BOOL firstShellItem = TRUE;
	BOOL freePIDL = FALSE;
	while(hChildItem) {
		LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(hChildItem);
		if(pItemDetails) {
			if(firstShellItem) {
				if(simplePIDL) {
					// build a full pIDL
					pIDLToCheck = SimpleToRealPIDL(pIDLToCheck);
					freePIDL = TRUE;
				}
				firstShellItem = FALSE;
			}
			// a shell item - check it
			// NOTE: This method requires pIDL to be an immediate child of the pIDL of hParentItem!
			if(ILIsEqual(pItemDetails->pIDL, pIDLToCheck)) {
				// found it
				break;
			}
		}
		hChildItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hChildItem)));
	}
	if(freePIDL) {
		ILFree(pIDLToCheck);
	}
	return hChildItem;
}

BOOL ShellTreeView::ItemHasBeenExpandedOnce(HTREEITEM itemHandle)
{
	TVITEMEX item = {0};
	item.hItem = itemHandle;
	item.mask = TVIF_HANDLE | TVIF_STATE;
	if(attachedControl.SendMessage(TVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return ((item.state & TVIS_EXPANDEDONCE) == TVIS_EXPANDEDONCE);
	}
	return FALSE;
}

#ifdef USE_STL
	void ShellTreeView::RemoveItems(std::vector<HTREEITEM>& itemsToRemove)
	{
		for(size_t i = 0; i < itemsToRemove.size(); ++i) {
			attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(itemsToRemove[i]));
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	}
#else
	void ShellTreeView::RemoveItems(CAtlList<HTREEITEM>& itemsToRemove)
	{
		POSITION p = itemsToRemove.GetHeadPosition();
		while(p) {
			attachedControl.SendMessage(TVM_DELETEITEM, 0, reinterpret_cast<LPARAM>(itemsToRemove.GetAt(p)));
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
			itemsToRemove.GetNext(p);
		}
	}
#endif

#ifdef USE_STL
	void ShellTreeView::SelectNewCaret(std::vector<HTREEITEM>& itemsBeingRemoved, PCIDLIST_ABSOLUTE pIDLToAvoid)
#else
	void ShellTreeView::SelectNewCaret(CAtlList<HTREEITEM>& itemsBeingRemoved, PCIDLIST_ABSOLUTE pIDLToAvoid)
#endif
{
	HTREEITEM hCaretItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_CARET), 0));
	if(hCaretItem) {
		BOOL willHaveCaretChange = FALSE;
		HTREEITEM hParentItem = hCaretItem;
		while(!willHaveCaretChange && hParentItem) {
			#ifdef USE_STL
				willHaveCaretChange = (std::find(itemsBeingRemoved.cbegin(), itemsBeingRemoved.cend(), hParentItem) != itemsBeingRemoved.cend());
			#else
				willHaveCaretChange = (itemsBeingRemoved.Find(hParentItem) != NULL);
			#endif
			hParentItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hParentItem)));
		}

		if(willHaveCaretChange) {
			/* We'll have a caret change. Change the caret to an item not affected by this notification before
			   removing the items. This reduces the risk that a silly "insert disc" message is displayed. */
			BOOL foundNewCaretItem = FALSE;
			HTREEITEM hNewCaretItem = hCaretItem;
			LPSHELLTREEVIEWITEMDATA pTempItemDetails;
			HTREEITEM hCurrentNewCaretItem = hNewCaretItem;
			BOOL searchDown = TRUE;
			while(!foundNewCaretItem && hNewCaretItem) {
				// we've found a new caret item if it won't be removed and isn't the pIDL to avoid
				#ifdef USE_STL
					foundNewCaretItem = (std::find(itemsBeingRemoved.cbegin(), itemsBeingRemoved.cend(), hNewCaretItem) == itemsBeingRemoved.cend());
				#else
					foundNewCaretItem = (itemsBeingRemoved.Find(hNewCaretItem) == NULL);
				#endif
				if(foundNewCaretItem && pIDLToAvoid) {
					// the item won't be removed
					pTempItemDetails = GetItemDetails(hNewCaretItem);
					if(pTempItemDetails && ILIsEqual(pTempItemDetails->pIDL, pIDLToAvoid)) {
						// but it's the pIDL to avoid
						foundNewCaretItem = FALSE;
					}
				}

				if(foundNewCaretItem) {
					// we've a candidate, check the parent items
					hParentItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hNewCaretItem)));
					HTREEITEM hTemp = NULL;
					while(foundNewCaretItem && hParentItem) {
						hTemp = hParentItem;
						#ifdef USE_STL
							foundNewCaretItem = (std::find(itemsBeingRemoved.cbegin(), itemsBeingRemoved.cend(), hParentItem) == itemsBeingRemoved.cend());
						#else
							foundNewCaretItem = (itemsBeingRemoved.Find(hParentItem) == NULL);
						#endif
						hParentItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hParentItem)));
					}
					if(!foundNewCaretItem) {
						// one of the parent items (hTemp) will be removed - jump to it and continue
						hNewCaretItem = hTemp;
						hCurrentNewCaretItem = hNewCaretItem;
						searchDown = TRUE;
					}
				} else {
					// try next
					if(searchDown) {
						hNewCaretItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hNewCaretItem)));
						if(!hNewCaretItem) {
							searchDown = FALSE;
							hNewCaretItem = hCurrentNewCaretItem;
						}
					}
					if(!searchDown) {
						hNewCaretItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PREVIOUS), reinterpret_cast<LPARAM>(hNewCaretItem)));
						if(!hNewCaretItem) {
							searchDown = TRUE;
							hNewCaretItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_PARENT), reinterpret_cast<LPARAM>(hCurrentNewCaretItem)));
							hCurrentNewCaretItem = hNewCaretItem;
						}
					}
				}
			}

			attachedControl.SendMessage(TVM_SELECTITEM, static_cast<WPARAM>(TVGN_CARET | TVSI_NOSINGLEEXPAND), reinterpret_cast<LPARAM>(hNewCaretItem));
		}
	}
}


LRESULT ShellTreeView::OnSHChangeNotify_ASSOCCHANGED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}

	if(ILIsEmpty(ppIDLs[0])) {
		// changed icon
		ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED - flushing all icons\n"));
		FlushIcons(-1, FALSE);
	} else {
		if(ppIDLs[0] && reinterpret_cast<LPSHChangeDWORDAsIDList>(const_cast<PIDLIST_ABSOLUTE>(ppIDLs[0]))->dwItem1 == 0) {
			// the desktop might have a new icon or lost one (happens on Windows 2000 if TweakUI is used)
			PIDLIST_ABSOLUTE pIDLDesktop = GetDesktopPIDL();
			// update namespaces
			#ifdef USE_STL
				std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
				for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
					if(ILIsEmpty(iter->first)) {
						iter->second->UpdateEnumeration();
					} else if(ILIsParent(pIDLDesktop, iter->first, FALSE)) {
						// validate the namespace
						// TODO: ValidatePIDL doesn't return FALSE if the shell item has been hidden
						if(!ValidatePIDL(iter->first)) {
							namespacesToRemove.push_back(iter->first);
						}
					}
				}
				for(std::vector<PCIDLIST_ABSOLUTE>::const_iterator iter = namespacesToRemove.cbegin(); iter != namespacesToRemove.cend(); ++iter) {
					RemoveNamespace(*iter, TRUE, TRUE);
				}
			#else
				CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
				POSITION p = properties.namespaces.GetStartPosition();
				while(p) {
					CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
					if(ILIsEmpty(pPair->m_key)) {
						pPair->m_value->UpdateEnumeration();
					} else if(ILIsParent(pIDLDesktop, pPair->m_key, FALSE)) {
						// validate the namespace
						// TODO: ValidatePIDL doesn't return FALSE if the shell item has been hidden
						if(!ValidatePIDL(pPair->m_key)) {
							namespacesToRemove.Add(pPair->m_key);
						}
					}
					properties.namespaces.GetNext(p);
				}
				for(size_t i = 0; i < namespacesToRemove.GetCount(); ++i) {
					RemoveNamespace(namespacesToRemove[i], TRUE, TRUE);
				}
			#endif

			// update all items
			#ifdef USE_STL
				std::vector<HTREEITEM> itemsToRemove;
				for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					HTREEITEM hItem = iter->first;
					LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
			#else
				CAtlList<HTREEITEM> itemsToRemove;
				p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					HTREEITEM hItem = pPair->m_key;
					LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(ILIsEmpty(pItemDetails->pIDL)) {
					if(pItemDetails->managedProperties & stmipSubItems) {
						// Has the item been expanded at least once?
						if(ItemHasBeenExpandedOnce(hItem)) {
							// yes, so update the sub-items
							INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
							ATLASSUME(pEnumSettings);
							HRESULT hr = E_FAIL;
							CComPtr<IRunnableTask> pTask;
							#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
								if(RunTimeHelperEx::IsWin8()) {
									hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
								} else {
									hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
								}
							#else
								hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
							#endif
							if(SUCCEEDED(hr)) {
								ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
							}
							pEnumSettings->Release();
							pEnumSettings = NULL;
						}
					}
				} else if(ILIsParent(pIDLDesktop, pItemDetails->pIDL, FALSE)) {
					// validate the item
					// TODO: ValidatePIDL doesn't return FALSE if the shell item has been hidden
					if(!ValidatePIDL(pItemDetails->pIDL)) {
						// remove the item
						#ifdef USE_STL
							itemsToRemove.push_back(hItem);
						#else
							itemsToRemove.AddTail(hItem);
						#endif
					}
				}
			}
			SelectNewCaret(itemsToRemove, NULL);
			RemoveItems(itemsToRemove);
			ILFree(pIDLDesktop);
		}

		SHFILEINFO fileInfoData = {0};
		SHGetFileInfo(reinterpret_cast<LPCTSTR>(ppIDLs[0]), 0, &fileInfoData, sizeof(SHFILEINFO), SHGFI_PIDL | SHGFI_DISPLAYNAME);
		if(lstrlen(fileInfoData.szDisplayName) == 0) {
			// changed sort order of "My Computer" and "My Documents"
			/* TODO: This works only if the order is changed from "My Computer" to "My Documents", not the other
			         way around. */

			// we must exchange the pIDL of the "My Documents" entry to get the new sort order
			PIDLIST_ABSOLUTE pIDLMyDocs = GetMyDocsPIDL();

			#ifdef USE_STL
				for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					HTREEITEM hItem = iter->first;
					LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
			#else
				POSITION p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					HTREEITEM hItem = pPair->m_key;
					LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(ILIsEqual(pItemDetails->pIDL, pIDLMyDocs)) {
					PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
					pItemDetails->pIDL = ILCloneFull(pIDLMyDocs);
					if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
						Raise_ChangedItemPIDL(hItem, tmp, pItemDetails->pIDL);
					}
					ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				}
			}
			ILFree(pIDLMyDocs);

			// re-sort sub-items of desktop item
			#ifdef USE_STL
				for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					HTREEITEM hItem = iter->first;
					LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
			#else
				p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					HTREEITEM hItem = pPair->m_key;
					LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(ILIsEmpty(pItemDetails->pIDL)) {
					if(pItemDetails->managedProperties & stmipSubItemsSorting) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXTVM_SORTITEMS, FALSE, reinterpret_cast<LPARAM>(hItem))));
					}
				}
			}

			// re-sort desktop namespace
			#ifdef USE_STL
				for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
					if(ILIsEmpty(iter->first)) {
						VARIANT_BOOL autoSort = VARIANT_FALSE;
						iter->second->get_AutoSortItems(&autoSort);
						if(autoSort != VARIANT_FALSE) {
							iter->second->SortItems();
						}
					}
				}
			#else
				p = properties.namespaces.GetStartPosition();
				while(p) {
					CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
					if(ILIsEmpty(pPair->m_key)) {
						VARIANT_BOOL autoSort = VARIANT_FALSE;
						pPair->m_value->get_AutoSortItems(&autoSort);
						if(autoSort != VARIANT_FALSE) {
							pPair->m_value->SortItems();
						}
					}
					properties.namespaces.GetNext(p);
				}
			#endif
		} else {
			// changed file type link
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED - flushing all icons\n"));
			FlushIcons(-1, FALSE);
		}
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_CHANGEDSHARE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling changed share ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling changed share pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsParent(ppIDLs[0], iter->first, FALSE)) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsParent(ppIDLs[0], pPair->m_key, FALSE)) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToUpdate;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<HTREEITEM> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(hItem);
			#else
				itemsToUpdate.Add(hItem);
			#endif
			// update the pIDL
			UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
		} else if(ILIsParent(pItemDetails->pIDL, ppIDLs[0], TRUE)) {
			if(pItemDetails->managedProperties & stmipSubItems) {
				// Has the item been expanded at least once?
				if(ItemHasBeenExpandedOnce(hItem)) {
					// yes, so update the sub-items
					INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
					ATLASSUME(pEnumSettings);
					HRESULT hr = E_FAIL;
					CComPtr<IRunnableTask> pTask;
					#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
						if(RunTimeHelperEx::IsWin8()) {
							hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						} else {
							hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						}
					#else
						hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
					#endif
					if(SUCCEEDED(hr)) {
						ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
					}
					pEnumSettings->Release();
					pEnumSettings = NULL;
				}
			}
		}
	}
	ILFree(pIDLNew);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i]);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item creation ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item creation pIDL=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	// sanity check
	ATLASSERT(!IsSubItemOfGlobalRecycleBin(ppIDLs[0]));
	BOOL updateGlobalRecycler = IsSubItemOfDriveRecycleBin(ppIDLs[0]);
	#ifdef DEBUG
		if(updateGlobalRecycler) {
			ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	BOOL autoLabelEdit = FALSE;
	if(GetTickCount() - properties.timeOfLastItemCreatorInvocation <= 250) {
		autoLabelEdit = TRUE;
	}
	BOOL isComctl600 = RunTimeHelper::IsCommCtrl6();

	// update the sub-items of any item of which we manage the sub-items
	#ifdef USE_STL
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(pItemDetails->managedProperties & stmipSubItems) {
			// Is the item an immediate parent item of the added item?
			if(ILIsParent(pItemDetails->pIDL, ppIDLs[0], TRUE)) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found item 0x%X to be immediate parent of pIDL 0x%X (count: %i)\n"), hItem, ppIDLs[0], ILCount(ppIDLs[0]));
				// Has the item been expanded at least once?
				INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
				ATLASSUME(pEnumSettings);
				if(ItemHasBeenExpandedOnce(hItem)) {
					// yes, so we must add the item
					CComPtr<IRunnableTask> pTask;
					HRESULT hr = ShTvwInsertSingleItemTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, hItem, NULL, ppIDLs[0], TRUE, pItemDetails->pIDL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, autoLabelEdit, &pTask);
					if(SUCCEEDED(hr)) {
						CComPtr<IShTreeViewItem> pItem = NULL;
						ClassFactory::InitShellTreeItem(hItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
						INamespaceEnumTask* pEnumTask = NULL;
						pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
						ATLASSUME(pEnumTask);
						if(pEnumTask) {
							ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pItem)));
							pEnumTask->Release();
						}
						ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
					}
				} else {
					// maybe we must update the item's expando
					CComPtr<IShellFolder> pParentISF = NULL;
					PCUITEMID_CHILD pRelativePIDL = NULL;
					HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
					if(SUCCEEDED(hr)) {
						ATLASSUME(pParentISF);
						ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

						CachedEnumSettings enumSettings = CacheEnumSettings(pEnumSettings);
						LONG hasExpando = 0/*ExTVw::heNo*/;
						if(isComctl600 && properties.attachedControlBuildNumber >= 408) {
							hasExpando = -2/*ExTVw::heAuto*/;
						}
						HasSubItems hasSubItems = HasAtLeastOneSubItem(pParentISF, const_cast<PITEMID_CHILD>(pRelativePIDL), const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), &enumSettings, TRUE);
						if((hasSubItems == hsiYes) || (hasSubItems == hsiMaybe)) {
							hasExpando = 1/*heYes*/;
						}

						TVITEMEX item = {0};
						item.mask = TVIF_HANDLE | TVIF_CHILDREN;
						item.hItem = hItem;
						item.cChildren = hasExpando;
						attachedControl.SendMessage(TVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
					}
				}
				pEnumSettings->Release();
				pEnumSettings = NULL;
			// NOTE: We're assuming that the sub-items of the global recycler can't have sub-items themselves.
			} else if(updateGlobalRecycler && ILIsEqual(pIDLRecycler, pItemDetails->pIDL)) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found item 0x%X to be the global recycler\n"), hItem);
				// Has the item been expanded at least once?
				if(ItemHasBeenExpandedOnce(hItem)) {
					// yes, so we must update it
					INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
					ATLASSUME(pEnumSettings);
					HRESULT hr = E_FAIL;
					CComPtr<IRunnableTask> pTask;
					#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
						if(RunTimeHelperEx::IsWin8()) {
							hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						} else {
							hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						}
					#else
						hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
					#endif
					if(SUCCEEDED(hr)) {
						ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
					}
					pEnumSettings->Release();
					pEnumSettings = NULL;
				}
			}
		}
	}

	// now update all namespaces, but non-recursively
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && (ILIsEqual(pIDLRecycler, iter->first) || ILIsParent(pIDLRecycler, iter->first, FALSE))) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be equal or child of global recycler\n"), iter->first);
				iter->second->UpdateEnumeration();
			} else {
				iter->second->OnSHChangeNotify_CREATE(ppIDLs[0]);
			}
		}
	#else
		p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && (ILIsEqual(pIDLRecycler, pPair->m_key) || ILIsParent(pIDLRecycler, pPair->m_key, FALSE))) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be equal or child of global recycler\n"), pPair->m_key);
				pPair->m_value->UpdateEnumeration();
			} else {
				pPair->m_value->OnSHChangeNotify_CREATE(ppIDLs[0]);
			}
			properties.namespaces.GetNext(p);
		}
	#endif
	ILFree(pIDLRecycler);

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir, PCIDLIST_ABSOLUTE pIDLNamespace, HTREEITEM hParentItem, INamespaceEnumSettings* pEnumSettings)
{
	ATLASSERT_POINTER(simplePIDLAddedDir, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);
	ATLASSUME(pEnumSettings);

	BOOL autoLabelEdit = FALSE;
	if(GetTickCount() - properties.timeOfLastItemCreatorInvocation <= 250) {
		autoLabelEdit = TRUE;
	}

	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShTvwInsertSingleItemTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, hParentItem, NULL, simplePIDLAddedDir, TRUE, pIDLNamespace, pEnumSettings, pIDLNamespace, TRUE, autoLabelEdit, &pTask);
	if(SUCCEEDED(hr)) {
		CComPtr<IShTreeViewItem> pParentItem = NULL;
		ClassFactory::InitShellTreeItem(hParentItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pParentItem));
		INamespaceEnumTask* pEnumTask = NULL;
		pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
		ATLASSUME(pEnumTask);
		if(pEnumTask) {
			ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pParentItem)));
			pEnumTask->Release();
		}
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
	}
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_DELETE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item deletion ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, (ppIDLs[1] ? 2 : 1));
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item deletion pIDL1=0x%X (count: %i), pIDL2=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	// sanity check
	ATLASSERT(!IsSubItemOfGlobalRecycleBin(ppIDLs[0]));
	BOOL updateGlobalRecycler = IsSubItemOfDriveRecycleBin(ppIDLs[0]);
	#ifdef DEBUG
		if(updateGlobalRecycler) {
			ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	// update all namespaces
	#ifdef USE_STL
		std::vector<CComObject<ShellTreeViewNamespace>*> namespacesToCheck;
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && ILIsParent(pIDLRecycler, iter->first, FALSE)) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), iter->first);
				// validate the namespace
				if(!ValidatePIDL(iter->first)) {
					namespacesToRemove.push_back(iter->first);
				} else {
					namespacesToCheck.push_back(iter->second);
				}
			} else {
				namespacesToCheck.push_back(iter->second);
			}
		}
		for(std::vector<PCIDLIST_ABSOLUTE>::const_iterator iter = namespacesToRemove.cbegin(); iter != namespacesToRemove.cend(); ++iter) {
			RemoveNamespace(*iter, TRUE, TRUE);
		}
		for(std::vector<CComObject<ShellTreeViewNamespace>*>::const_iterator iter = namespacesToCheck.cbegin(); iter != namespacesToCheck.cend(); ++iter) {
			(*iter)->OnSHChangeNotify_DELETE(ppIDLs[0]);
		}
	#else
		CAtlArray<CComObject<ShellTreeViewNamespace>*> namespacesToCheck;
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pPair->m_key, FALSE)) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), pPair->m_key);
				// validate the namespace
				if(!ValidatePIDL(pPair->m_key)) {
					namespacesToRemove.Add(pPair->m_key);
				} else {
					namespacesToCheck.Add(pPair->m_value);
				}
			} else {
				namespacesToCheck.Add(pPair->m_value);
			}
			properties.namespaces.GetNext(p);
		}
		for(size_t i = 0; i < namespacesToRemove.GetCount(); ++i) {
			RemoveNamespace(namespacesToRemove[i], TRUE, TRUE);
		}
		for(size_t i = 0; i < namespacesToCheck.GetCount(); ++i) {
			namespacesToCheck[i]->OnSHChangeNotify_DELETE(ppIDLs[0]);
		}
	#endif

	// update all items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlList<HTREEITEM> itemsToRemove;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// remove the item
			#ifdef USE_STL
				itemsToRemove.push_back(hItem);
			#else
				itemsToRemove.AddTail(hItem);
			#endif
		} else if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pItemDetails->pIDL, FALSE)) {
			ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found item 0x%X to be child of global recycler\n"), hItem);
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(hItem);
				#else
					itemsToRemove.AddTail(hItem);
				#endif
			}
		}
	}
	SelectNewCaret(itemsToRemove, NULL);
	RemoveItems(itemsToRemove);
	ILFree(pIDLRecycler);

	if(ppIDLs[1]) {
		ATLASSERT_POINTER(ppIDLs[1], ITEMIDLIST_ABSOLUTE);
		PCIDLIST_ABSOLUTE* pTemp = new PCIDLIST_ABSOLUTE[2];
		pTemp[0] = ppIDLs[1];
		pTemp[1] = NULL;
		OnSHChangeNotify_CREATE(pTemp);
		delete[] pTemp;
	}

	// NOTE: ppIDL[0] and ppIDL[1] seem to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_DRIVEADD(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling new drive ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling new drive pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsEqual(iter->first, pIDLMyComputer)) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsEqual(pPair->m_key, pIDLMyComputer)) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToUpdate;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<HTREEITEM> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(hItem);
			#else
				itemsToUpdate.Add(hItem);
			#endif
			// update the pIDL
			UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
		} else if(ILIsEqual(pItemDetails->pIDL, pIDLMyComputer)) {
			if(pItemDetails->managedProperties & stmipSubItems) {
				// Has the item been expanded at least once?
				if(ItemHasBeenExpandedOnce(hItem)) {
					// yes, so update the sub-items
					INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
					ATLASSUME(pEnumSettings);
					HRESULT hr = E_FAIL;
					CComPtr<IRunnableTask> pTask;
					#ifdef USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
						if(RunTimeHelperEx::IsWin8()) {
							hr = ShTvwBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						} else {
							hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
						}
					#else
						hr = ShTvwLegacyBackgroundItemEnumTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, GethWndShellUIParentWindow(), hItem, NULL, pItemDetails->pIDL, NULL, NULL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, &pTask);
					#endif
					if(SUCCEEDED(hr)) {
						ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
					}
					pEnumSettings->Release();
					pEnumSettings = NULL;
				}
			}
		}
	}
	ILFree(pIDLMyComputer);
	ILFree(pIDLNew);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i]);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_DRIVEREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling drive removal ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling drive removal pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}

	// update namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsParent(ppIDLs[0], iter->first, FALSE)) {
				namespacesToRemove.push_back(iter->first);
			}
		}
		for(std::vector<PCIDLIST_ABSOLUTE>::const_iterator iter = namespacesToRemove.cbegin(); iter != namespacesToRemove.cend(); ++iter) {
			RemoveNamespace(*iter, TRUE, TRUE);
		}
	#else
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsParent(ppIDLs[0], pPair->m_key, FALSE)) {
				namespacesToRemove.Add(pPair->m_key);
			}
			properties.namespaces.GetNext(p);
		}
		for(size_t i = 0; i < namespacesToRemove.GetCount(); ++i) {
			RemoveNamespace(namespacesToRemove[i], TRUE, TRUE);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlList<HTREEITEM> itemsToRemove;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(hItem);
				#else
					itemsToRemove.AddTail(hItem);
				#endif
			}
		}
	}
	SelectNewCaret(itemsToRemove, NULL);
	RemoveItems(itemsToRemove);

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_MEDIAINSERTED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling media insertion ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling media insertion pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0])) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0])) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToUpdate;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<HTREEITEM> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(hItem);
			#else
				itemsToUpdate.Add(hItem);
			#endif
			// update the pIDL
			UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
		}
	}
	ILFree(pIDLNew);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i], FALSE);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_MEDIAREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling media removal ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling media removal pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsParent(ppIDLs[0], iter->first, FALSE)) {
				if(!ILIsEqual(ppIDLs[0], iter->first)) {
					namespacesToRemove.push_back(iter->first);
				}
			}
		}
		for(std::vector<PCIDLIST_ABSOLUTE>::const_iterator iter = namespacesToRemove.cbegin(); iter != namespacesToRemove.cend(); ++iter) {
			RemoveNamespace(*iter, TRUE, TRUE);
		}
	#else
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsParent(ppIDLs[0], pPair->m_key, FALSE)) {
				if(!ILIsEqual(ppIDLs[0], pPair->m_key)) {
					namespacesToRemove.Add(pPair->m_key);
				}
			}
			properties.namespaces.GetNext(p);
		}
		for(size_t i = 0; i < namespacesToRemove.GetCount(); ++i) {
			RemoveNamespace(namespacesToRemove[i], TRUE, TRUE);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
		std::vector<HTREEITEM> itemsToUpdate;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlList<HTREEITEM> itemsToRemove;
		CAtlArray<HTREEITEM> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(hItem);
			#else
				itemsToUpdate.Add(hItem);
			#endif
			// update the pIDL
			UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
		} else if(ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			#ifdef USE_STL
				itemsToRemove.push_back(hItem);
			#else
				itemsToRemove.AddTail(hItem);
			#endif
		}
	}
	ILFree(pIDLNew);

	SelectNewCaret(itemsToRemove, ppIDLs[0]);
	RemoveItems(itemsToRemove);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i], FALSE);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_RENAME(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item rename ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 2);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item rename pIDL1=0x%X (items: %i), pIDL2=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(ppIDLs[1], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0] || !ppIDLs[1]) {
		// shouldn't happen
		return 0;
	}

	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToCheckForNewSubItem;
		std::vector<HTREEITEM> itemsToUpdate;
		std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacePIDLsToReplace;
		std::vector<CComObject<ShellTreeViewNamespace>*> namespacesToUpdate;
	#else
		CAtlArray<HTREEITEM> itemsToCheckForNewSubItem;
		CAtlArray<HTREEITEM> itemsToUpdate;
		CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacePIDLsToReplace;
		CAtlArray<CComObject<ShellTreeViewNamespace>*> namespacesToUpdate;
		POSITION p;
	#endif

	PIDLIST_ABSOLUTE pIDLOld = ILCloneFull(ppIDLs[0]);
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[1]);
	BOOL isMove = !HaveSameParent(ppIDLs, 2);
	if(isMove) {
		/* TODO: Maybe someday we find a way to do a real move using ExplorerTreeView's item moving
		         capabilities. This is a bit difficult because the moved item may exist multiple times
		         within the treeview. Also it's possible only if we manage the sub-items of the new parent
		         item and it already has been expanded at least once. */
		PCIDLIST_ABSOLUTE* pTemp = new PCIDLIST_ABSOLUTE[2];
		pTemp[0] = ppIDLs[1];
		pTemp[1] = NULL;
		OnSHChangeNotify_CREATE(pTemp);
		pTemp[0] = pIDLOld;
		pTemp[1] = NULL;
		OnSHChangeNotify_DELETE(pTemp);
		delete[] pTemp;
	} else {
		// update namespace pIDLs
		#ifdef USE_STL
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				PCIDLIST_ABSOLUTE pIDLNamespaceOld = iter->first;
				CComObject<ShellTreeViewNamespace>* pNamespaceObj = iter->second;
		#else
			p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				PCIDLIST_ABSOLUTE pIDLNamespaceOld = pPair->m_key;
				CComObject<ShellTreeViewNamespace>* pNamespaceObj = pPair->m_value;
				properties.namespaces.GetNext(p);
		#endif
			if(ILIsParent(pIDLNamespaceOld, pIDLNew, TRUE)) {
				// maybe add the item
				/* NOTE: OnTriggerItemEnumComplete will sort the items. If this method isn't reached, it is because
				         ShTvwBackgroundItemEnumTask produced a serious error. */
				#ifdef USE_STL
					namespacesToUpdate.push_back(pNamespaceObj);
				#else
					namespacesToUpdate.Add(pNamespaceObj);
				#endif
			} else if(ILIsEqual(pIDLNamespaceOld, pIDLOld)) {
				namespacePIDLsToReplace[pIDLNamespaceOld] = ILCloneFull(pIDLNew);
			} else if(ILIsParent(pIDLOld, pIDLNamespaceOld, FALSE)) {
				UINT itemIDsRenamedItem = ILCount(pIDLOld);
				ATLASSERT(itemIDsRenamedItem == ILCount(pIDLNew));
				ATLASSERT(itemIDsRenamedItem < ILCount(pIDLNamespaceOld));
				// replace the first <itemIDsRenamedItem> item IDs of the namespace's pIDL with the new pIDL
				// find the start of the "tail"
				PCUIDLIST_RELATIVE pIDLTail = pIDLNamespaceOld;
				for(UINT i = 0; i < itemIDsRenamedItem; ++i) {
					pIDLTail = ILNext(pIDLTail);
				}
				ATLASSERT_POINTER(pIDLTail, ITEMIDLIST_RELATIVE);
				// NOTE: ILCombine doesn't change pIDLNew
				PIDLIST_ABSOLUTE pIDLNamespaceNew = ILCombine(pIDLNew, pIDLTail);
				ATLASSERT(ILCount(pIDLNamespaceNew) == ILCount(pIDLNamespaceOld));
				namespacePIDLsToReplace[pIDLNamespaceOld] = pIDLNamespaceNew;
			}
		}

		// update the items
		#ifdef USE_STL
			for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				HTREEITEM hItem = iter->first;
				LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;

				PCIDLIST_ABSOLUTE pIDLNamespaceNew = NULL;
				std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter2 = namespacePIDLsToReplace.find(pItemDetails->pIDLNamespace);
				if(iter2 != namespacePIDLsToReplace.cend()) {
					pIDLNamespaceNew = iter2->second;
				}
		#else
			p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				HTREEITEM hItem = pPair->m_key;
				LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
				properties.items.GetNext(p);

				PCIDLIST_ABSOLUTE pIDLNamespaceNew = NULL;
				namespacePIDLsToReplace.Lookup(pItemDetails->pIDLNamespace, pIDLNamespaceNew);
		#endif
			if(pIDLNamespaceNew) {
				// exchange the namespace pIDL
				pItemDetails->pIDLNamespace = pIDLNamespaceNew;
			}
			// update the item's pIDL
			if(ILIsParent(pItemDetails->pIDL, pIDLNew, TRUE)) {
				if(pItemDetails->managedProperties & stmipSubItems) {
					// maybe add the item
					if(ItemHasBeenExpandedOnce(hItem)) {
						/* NOTE: OnTriggerItemEnumComplete will sort the items. If this method isn't reached, it is
						         because ShTvwInsertSingleItemTask decides the new pIDL should not be shown. In this
						         case ApplyManagedProperties() will have decided the same for the maybe-duplicate and
						         removed it, so everything is fine. */
						#ifdef USE_STL
							itemsToCheckForNewSubItem.push_back(hItem);
						#else
							itemsToCheckForNewSubItem.Add(hItem);
						#endif
					}
				}
			} else if(ILIsEqual(pItemDetails->pIDL, pIDLOld)) {
				// exchange the pIDL and re-apply any managed properties
				PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
				pItemDetails->pIDL = ILCloneFull(pIDLNew);
				if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
					Raise_ChangedItemPIDL(hItem, tmp, pItemDetails->pIDL);
				}
				ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
				#ifdef USE_STL
					itemsToUpdate.push_back(hItem);
				#else
					itemsToUpdate.Add(hItem);
				#endif
			} else if(ILIsParent(pIDLOld, pItemDetails->pIDL, FALSE)) {
				// update the pIDL and re-apply any managed properties

				UINT itemIDsRenamedItem = ILCount(pIDLOld);
				ATLASSERT(itemIDsRenamedItem == ILCount(pIDLNew));
				ATLASSERT(itemIDsRenamedItem < ILCount(pItemDetails->pIDL));
				// replace the first <itemIDsRenamedItem> item IDs of the item's pIDL with the new pIDL
				// find the start of the "tail"
				PCUIDLIST_RELATIVE pIDLTail = pItemDetails->pIDL;
				for(UINT i = 0; i < itemIDsRenamedItem; ++i) {
					pIDLTail = ILNext(pIDLTail);
				}
				ATLASSERT_POINTER(pIDLTail, ITEMIDLIST_RELATIVE);
				// NOTE: ILCombine doesn't change pIDLNew
				PIDLIST_ABSOLUTE pIDLItemNew = ILCombine(pIDLNew, pIDLTail);
				ATLASSERT(ILCount(pIDLItemNew) == ILCount(pItemDetails->pIDL));

				// exchange the pIDL and re-apply any managed properties
				PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
				pItemDetails->pIDL = pIDLItemNew;
				if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
					Raise_ChangedItemPIDL(hItem, tmp, pItemDetails->pIDL);
				}
				ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
				#ifdef USE_STL
					itemsToUpdate.push_back(hItem);
				#else
					itemsToUpdate.Add(hItem);
				#endif
			}
		}
	}

	// store the new namespace pIDLs
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter = namespacePIDLsToReplace.cbegin(); iter != namespacePIDLsToReplace.cend(); ++iter) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter2 = properties.namespaces.find(iter->first);
			if(iter2 != properties.namespaces.cend()) {
				iter2->second->properties.pIDL = iter->second;
				properties.namespaces[iter->second] = iter2->second;
				properties.namespaces.erase(iter->first);
				if(!(properties.disabledEvents & deNamespacePIDLChangeEvents)) {
					CComQIPtr<IShTreeViewNamespace> pNamespaceItf = iter2->second;
					Raise_ChangedNamespacePIDL(pNamespaceItf, iter->first, iter->second);
				}
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(iter->first));
		}
	#else
		p = namespacePIDLsToReplace.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::CPair* pPair = namespacePIDLsToReplace.GetAt(p);
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair2 = properties.namespaces.Lookup(pPair->m_key);
			pPair2->m_value->properties.pIDL = pPair->m_value;
			properties.namespaces[pPair->m_value] = pPair2->m_value;
			properties.namespaces.RemoveKey(pPair->m_key);
			if(!(properties.disabledEvents & deNamespacePIDLChangeEvents)) {
				CComQIPtr<IShTreeViewNamespace> pNamespaceItf = pPair2->m_value;
				Raise_ChangedNamespacePIDL(pNamespaceItf, pPair->m_key, pPair->m_value);
			}

			ILFree(const_cast<PIDLIST_ABSOLUTE>(pPair->m_key));
			namespacePIDLsToReplace.GetNext(p);
		}
	#endif

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i], TRUE);
	}
	/* NOTE: It's very important that we start the tasks which search for new sub-items AFTER we have updated
	         all pIDLs. Otherwise the task may finish before the item, that's actually a duplicate, of the
	         new item gets its new pIDL thus being undetected by the duplicates check. */
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToCheckForNewSubItem.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToCheckForNewSubItem.GetCount(); ++i) {
	#endif
		HTREEITEM hItem = itemsToCheckForNewSubItem[i];
		LPSHELLTREEVIEWITEMDATA pItemDetails = GetItemDetails(hItem);
		ATLASSERT_POINTER(pItemDetails, SHELLTREEVIEWITEMDATA);
		if(pItemDetails) {
			INamespaceEnumSettings* pEnumSettings = ItemHandleToNameSpaceEnumSettings(hItem, pItemDetails);
			ATLASSUME(pEnumSettings);
			CComPtr<IRunnableTask> pTask;
			HRESULT hr = ShTvwInsertSingleItemTask::CreateInstance(attachedControl, &properties.backgroundEnumResults, &properties.backgroundEnumResultsCritSection, hItem, NULL, pIDLNew, FALSE, pItemDetails->pIDL, pEnumSettings, pItemDetails->pIDLNamespace, TRUE, FALSE, &pTask);
			if(SUCCEEDED(hr)) {
				CComPtr<IShTreeViewItem> pItem = NULL;
				ClassFactory::InitShellTreeItem(hItem, this, IID_IShTreeViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
				INamespaceEnumTask* pEnumTask = NULL;
				pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
				ATLASSUME(pEnumTask);
				if(pEnumTask) {
					ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pItem)));
					pEnumTask->Release();
				}
				ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwAutoUpdate, 0, TASK_PRIORITY_TV_BACKGROUNDENUM)));
			}
			pEnumSettings->Release();
			pEnumSettings = NULL;
		}
	}
	ILFree(pIDLOld);
	ILFree(pIDLNew);

	// search the namespaces for new items
	#ifdef USE_STL
		for(size_t i = 0; i < namespacesToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < namespacesToUpdate.GetCount(); ++i) {
	#endif
		namespacesToUpdate[i]->UpdateEnumeration();
	}

	// NOTE: ppIDL[0] and ppIDL[1] seem to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_UPDATEDIR(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling directory update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling directory update pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	BOOL removeCollapsedSubTrees = FALSE;
	#ifdef USE_STL
		if(properties.items.size() >= 2000) {
			removeCollapsedSubTrees = TRUE; //(properties.items.size() == static_cast<DWORD>(attachedControl.SendMessage(TVM_GETCOUNT, 0, 0)));
		}
	#else
		if(properties.items.GetCount() >= 2000) {
			removeCollapsedSubTrees = TRUE; //(properties.items.GetCount() == static_cast<DWORD>(attachedControl.SendMessage(TVM_GETCOUNT, 0, 0)));
		}
	#endif
	// sanity check
	// Happens on Vista: ATLASSERT(!IsSubItemOfGlobalRecycleBin(ppIDLs[0]));
	BOOL updateGlobalRecycler/* = IsSubItemOfDriveRecycleBin(ppIDLs[0])*/;   // TODO: "IsSubItemOfGlobalRecycleBin(ppIDLs[0]) || "?
	#ifdef DEBUG
		updateGlobalRecycler = IsSubItemOfDriveRecycleBin(ppIDLs[0]);
		if(updateGlobalRecycler) {
			ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	/* HACK: If many files are restored from the recycler at once, the shell often sends a SHCNE_UPDATEDIR
	         for the target directory only. */
	updateGlobalRecycler = TRUE;
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	if(removeCollapsedSubTrees) {
		HTREEITEM hRootItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_ROOT), 0));
		while(hRootItem) {
			RemoveCollapsedSubTree(hRootItem);
			hRootItem = reinterpret_cast<HTREEITEM>(attachedControl.SendMessage(TVM_GETNEXTITEM, static_cast<WPARAM>(TVGN_NEXT), reinterpret_cast<LPARAM>(hRootItem)));
		}
	}

	// update all namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		std::vector<CComObject<ShellTreeViewNamespace>*> namespacesToUpdate;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && (ILIsEqual(iter->first, pIDLRecycler) || ILIsParent(pIDLRecycler, iter->first, FALSE))) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), iter->first);
				// validate the namespace
				if(!ValidatePIDL(iter->first)) {
					namespacesToRemove.push_back(iter->first);
				} else {
					namespacesToUpdate.push_back(iter->second);
				}
			} else if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsParent(ppIDLs[0], iter->first, FALSE)) {
				// validate the namespace
				if(!ValidatePIDL(iter->first)) {
					namespacesToRemove.push_back(iter->first);
				} else {
					namespacesToUpdate.push_back(iter->second);
				}
			}
		}
		for(std::vector<PCIDLIST_ABSOLUTE>::const_iterator iter = namespacesToRemove.cbegin(); iter != namespacesToRemove.cend(); ++iter) {
			RemoveNamespace(*iter, TRUE, TRUE);
		}
		for(std::vector<CComObject<ShellTreeViewNamespace>*>::const_iterator iter = namespacesToUpdate.cbegin(); iter != namespacesToUpdate.cend(); ++iter) {
			(*iter)->UpdateEnumeration();
		}
	#else
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		CAtlArray<CComObject<ShellTreeViewNamespace>*> namespacesToUpdate;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && (ILIsEqual(pPair->m_key, pIDLRecycler) || ILIsParent(pIDLRecycler, pPair->m_key, FALSE))) {
				ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), pPair->m_key);
				// validate the namespace
				if(!ValidatePIDL(pPair->m_key)) {
					namespacesToRemove.Add(pPair->m_key);
				} else {
					namespacesToUpdate.Add(pPair->m_value);
				}
			} else if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsParent(ppIDLs[0], pPair->m_key, FALSE)) {
				// validate the namespace
				if(!ValidatePIDL(pPair->m_key)) {
					namespacesToRemove.Add(pPair->m_key);
				} else {
					namespacesToUpdate.Add(pPair->m_value);
				}
			}
			properties.namespaces.GetNext(p);
		}
		for(size_t i = 0; i < namespacesToRemove.GetCount(); ++i) {
			RemoveNamespace(namespacesToRemove[i], TRUE, TRUE);
		}
		for(size_t i = 0; i < namespacesToUpdate.GetCount(); ++i) {
			namespacesToUpdate[i]->UpdateEnumeration();
		}
	#endif

	// update all items
	#ifdef USE_STL
		std::vector<HTREEITEM> itemsToRemove;
		std::vector<HTREEITEM> itemsToUpdate;
		for(std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			HTREEITEM hItem = iter->first;
			LPSHELLTREEVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlList<HTREEITEM> itemsToRemove;
		CAtlArray<HTREEITEM> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			HTREEITEM hItem = pPair->m_key;
			LPSHELLTREEVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				#ifdef USE_STL
					itemsToRemove.push_back(hItem);
				#else
					itemsToRemove.AddTail(hItem);
				#endif
			} else {
				#ifdef USE_STL
					itemsToUpdate.push_back(hItem);
				#else
					itemsToUpdate.Add(hItem);
				#endif
				// update the pIDL
				UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
			}
		} else if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pItemDetails->pIDL, FALSE)) {
			ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Found item 0x%X to be child of global recycler\n"), hItem);
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(hItem);
				#else
					itemsToRemove.AddTail(hItem);
				#endif
			} else {
				#ifdef USE_STL
					itemsToUpdate.push_back(hItem);
				#else
					itemsToUpdate.Add(hItem);
				#endif
				// update the pIDL
				// TODO: Is calling UpdateItemPIDL() necessary? Which pIDL should be passed?
				//UpdateItemPIDL(hItem, pItemDetails, pIDLNew);
			}
		}
	}
	ILFree(pIDLRecycler);
	ILFree(pIDLNew);

	SelectNewCaret(itemsToRemove, NULL);
	RemoveItems(itemsToRemove);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i], FALSE);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_UPDATEIMAGE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling image update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, (ppIDLs[1] ? 2 : 1));
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling image update pIDL1=0x%X (count: %i), pIDL2=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
	ATLASSERT_POINTER(ppIDLs[0], SHChangeDWORDAsIDList);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	int oldIconIndex = reinterpret_cast<LPSHChangeDWORDAsIDList>(const_cast<PIDLIST_ABSOLUTE>(ppIDLs[0]))->dwItem1;
	if(ppIDLs[1]) {
		ATLASSERT_POINTER(ppIDLs[1], SHChangeUpdateImageIDList);
		oldIconIndex = SHHandleUpdateImage(ppIDLs[1]);
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling update of image %i\n"), oldIconIndex);

	// update all items
	FlushIcons(oldIconIndex, oldIconIndex == -1);

	// NOTE: ppIDL[0] and ppIDL[1] seem to be freed by the shell
	return 0;
}

LRESULT ShellTreeView::OnSHChangeNotify_UPDATEITEM(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shtvwTraceAutoUpdate, 1, TEXT("Handling item update pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}

	BOOL fullRefresh = ILIsEmpty(ppIDLs[0]);
	BOOL isPartOfFileSystem = FALSE;
	if(!fullRefresh) {
		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		HRESULT hr = ::SHBindToParent(ppIDLs[0], IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		if(SUCCEEDED(hr)) {
			ATLASSUME(pParentISF);
			ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);

			SFGAOF attributes = SFGAO_FILESYSTEM;
			hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
			if(SUCCEEDED(hr)) {
				isPartOfFileSystem = (attributes & SFGAO_FILESYSTEM);
			}
		}
	}

	if(fullRefresh) {
		OnSHChangeNotify_UPDATEDIR(ppIDLs);
	} else if(isPartOfFileSystem) {
		if(!IsSubItemOfDriveRecycleBin(ppIDLs[0])) {
			OnSHChangeNotify_UPDATEDIR(ppIDLs);
		}
	} else {
		// updated virtual item
		OnSHChangeNotify_UPDATEDIR(ppIDLs);
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}


HRESULT ShellTreeView::EnqueueTask(IRunnableTask* pTask, REFTASKOWNERID taskGroupID, DWORD_PTR lParam, DWORD priority, DWORD operationStart/* = 0*/)
{
	if(!flags.acceptNewTasks) {
		return E_FAIL;
	}

	HRESULT hr;

	ATLASSUME(pTask);
	ATLASSERT(priority > 0);

	IShellTaskScheduler* pTaskSchedulerToUse = NULL;

	if(taskGroupID == TASKID_ShTvwBackgroundItemEnum) {
		if(!properties.pEnumTaskScheduler) {
			// create the scheduler
			hr = CoCreateInstance(CLSID_ShellTaskScheduler, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&properties.pEnumTaskScheduler));
			ATLASSERT(SUCCEEDED(hr));
			ATLASSUME(properties.pEnumTaskScheduler);
			if(FAILED(hr)) {
				return hr;
			}

			// immediately stop execution of the current task when the scheduler is released
			properties.pEnumTaskScheduler->Status(ITSSFLAG_KILL_ON_DESTROY | ITSSFLAG_THREAD_TERMINATE_TIMEOUT, THREAD_TIMEOUT_DEFAULT);
		}
		pTaskSchedulerToUse = properties.pEnumTaskScheduler;
	} else {
		if(!properties.pNormalTaskScheduler) {
			// create the scheduler
			hr = CoCreateInstance(CLSID_ShellTaskScheduler, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&properties.pNormalTaskScheduler));
			ATLASSERT(SUCCEEDED(hr));
			ATLASSUME(properties.pNormalTaskScheduler);
			if(FAILED(hr)) {
				return hr;
			}

			// immediately stop execution of the current task when the scheduler is released
			properties.pNormalTaskScheduler->Status(ITSSFLAG_KILL_ON_DESTROY | ITSSFLAG_THREAD_TERMINATE_TIMEOUT, THREAD_TIMEOUT_DEFAULT);
		}
		pTaskSchedulerToUse = properties.pNormalTaskScheduler;
	}

	LPBACKGROUNDENUMERATION pBackgroundEnum = NULL;
	INamespaceEnumTask* pEnumTask = NULL;
	pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
	ULONGLONG taskID = 0;
	if(pEnumTask) {
		ATLVERIFY(SUCCEEDED(pEnumTask->GetTaskID(&taskID)));
		LPDISPATCH pTarget = NULL;
		pEnumTask->GetTarget(IID_PPV_ARGS(&pTarget));

		pBackgroundEnum = new BACKGROUNDENUMERATION(BET_ITEMS, pTarget, (operationStart != 0), operationStart);
		ATLASSERT_POINTER(pBackgroundEnum, BACKGROUNDENUMERATION);
		properties.backgroundEnums[taskID] = pBackgroundEnum;
	}

	hr = pTaskSchedulerToUse->AddTask(pTask, taskGroupID, lParam, priority);
	ATLASSERT(SUCCEEDED(hr));
	#ifdef DEBUG
		if(taskGroupID == TASKID_ShTvwBackgroundItemEnum) {
			ATLTRACE2(shtvwTraceThreading, 4, TEXT("Added background item enum task\n"));
		} else if(taskGroupID == TASKID_ShTvwIcon) {
			ATLTRACE2(shtvwTraceThreading, 4, TEXT("Added icon task\n"));
		} else if(taskGroupID == TASKID_ShTvwOverlay) {
			ATLTRACE2(shtvwTraceThreading, 4, TEXT("Added overlay task\n"));
		} else if(taskGroupID == TASKID_ElevationShield) {
			ATLTRACE2(shtvwTraceThreading, 4, TEXT("Added elevation shield task\n"));
		} else if(taskGroupID == TASKID_ShTvwAutoUpdate) {
			ATLTRACE2(shtvwTraceThreading, 4, TEXT("Added AutoUpdate related task\n"));
		}
	#endif

	if(pEnumTask) {
		pEnumTask->Release();
	}
	return hr;
}

HRESULT ShellTreeView::GetIconsAsync(PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShTvwIconTask::CreateInstance(attachedControl, pIDL, itemHandle, retrieveNormalImage, retrieveSelectedImage, properties.useGenericIcons, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwIcon, 0, TASK_PRIORITY_TV_GET_ICON)));
	}
	return hr;
}

HRESULT ShellTreeView::GetOverlayAsync(PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShTvwOverlayTask::CreateInstance(attachedControl, pIDL, itemHandle, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShTvwOverlay, 0, TASK_PRIORITY_TV_GET_OVERLAY)));
	}
	return hr;
}

HRESULT ShellTreeView::GetElevationShieldAsync(LONG itemID, PCIDLIST_ABSOLUTE pIDL)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ElevationShieldTask::CreateInstance(GethWndShellUIParentWindow(), attachedControl, pIDL, itemID, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ElevationShield, 0, TASK_PRIORITY_GET_ELEVATIONSHIELD)));
	}
	return hr;
}


BOOL ShellTreeView::IsInDesignMode(void)
{
	BOOL b = TRUE;
	this->GetAmbientUserMode(b);
	return !b;
}