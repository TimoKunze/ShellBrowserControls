// ShellListView.cpp: Makes an ExplorerListView a shell listview.

#include "stdafx.h"
#include "ShellListView.h"
#include "../../ClassFactory.h"


ShellListView::ShellListView()
{
	SIZEL size = {36, 36};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	HICON hIDEIcon = NULL;
	HRSRC hResource = FindResource(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_IDELOGOSHLVW), RT_GROUP_ICON);
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

ShellListView::~ShellListView()
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

HRESULT ShellListView::FinalConstruct()
{
	if(properties.pDefaultNamespaceEnumSettings) {
		properties.pDefaultNamespaceEnumSettings->Release();
		properties.pDefaultNamespaceEnumSettings = NULL;
	}
	CComObject<NamespaceEnumSettings>* pDefaultNamespaceEnumSettingsObj = NULL;
	CComObject<NamespaceEnumSettings>::CreateInstance(&pDefaultNamespaceEnumSettingsObj);
	pDefaultNamespaceEnumSettingsObj->AddRef();
	pDefaultNamespaceEnumSettingsObj->put_EnumerationFlags(static_cast<ShNamespaceEnumerationConstants>(snseIncludeFolders | snseMayIncludeHiddenItems | snseIncludeNonFolders));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_ExcludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_IncludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
	pDefaultNamespaceEnumSettingsObj->put_IncludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
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
STDMETHODIMP ShellListView::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IShListView, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ShellListView::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
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
	hr = pPropertyBag->Read(OLESTR("AutoInsertColumns"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL autoInsertColumns = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
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
	hr = pPropertyBag->Read(OLESTR("ColumnEnumerationTimeout"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	LONG columnEnumerationTimeout = (SUCCEEDED(hr) ? propertyValue.lVal : 3000);
	hr = pPropertyBag->Read(OLESTR("DefaultManagedItemProperties"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ShLvwManagedItemPropertiesConstants defaultManagedItemProperties = (SUCCEEDED(hr) ? static_cast<ShLvwManagedItemPropertiesConstants>(propertyValue.lVal) : slmipAll);

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
	DisabledEventsConstants disabledEvents = (SUCCEEDED(hr) ? static_cast<DisabledEventsConstants>(propertyValue.lVal) : static_cast<DisabledEventsConstants>(deNamespacePIDLChangeEvents | deNamespaceInsertionEvents | deNamespaceDeletionEvents | deItemPIDLChangeEvents | deItemInsertionEvents | deItemDeletionEvents | deColumnLoadingEvents | deColumnVisibilityEvents));
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("DisplayElevationShieldOverlays"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL displayElevationShieldOverlays = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("DisplayFileTypeOverlays"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays = (SUCCEEDED(hr) ? static_cast<ShLvwDisplayFileTypeOverlaysConstants>(propertyValue.lVal) : static_cast<ShLvwDisplayFileTypeOverlaysConstants>(sldftoFollowSystemSettings));
	hr = pPropertyBag->Read(OLESTR("DisplayThumbnailAdornments"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments = (SUCCEEDED(hr) ? static_cast<ShLvwDisplayThumbnailAdornmentsConstants>(propertyValue.lVal) : sldtaAny);
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("DisplayThumbnails"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL displayThumbnails = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_FALSE);
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
	InfoTipFlagsConstants infoTipFlags = (SUCCEEDED(hr) ? static_cast<InfoTipFlagsConstants>(propertyValue.lVal) : static_cast<InfoTipFlagsConstants>(itfDefault | itfNoInfoTipFollowSystemSettings | itfAllowSlowInfoTipFollowSysSettings));
	hr = pPropertyBag->Read(OLESTR("InitialSortOrder"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	SortOrderConstants initialSortOrder = (SUCCEEDED(hr) ? static_cast<SortOrderConstants>(propertyValue.lVal) : soAscending);
	hr = pPropertyBag->Read(OLESTR("ItemEnumerationTimeout"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	LONG itemEnumerationTimeout = (SUCCEEDED(hr) ? propertyValue.lVal : 3000);
	hr = pPropertyBag->Read(OLESTR("ItemTypeSortOrder"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ItemTypeSortOrderConstants itemTypeSortOrder = (SUCCEEDED(hr) ? static_cast<ItemTypeSortOrderConstants>(propertyValue.lVal) : itsoShellItemsLast);
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
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("PersistColumnSettingsAcrossNamespaces"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	ShLvwPersistColumnSettingsAcrossNamespacesConstants persistColumnSettingsAcrossNamespaces = (SUCCEEDED(hr) ? static_cast<ShLvwPersistColumnSettingsAcrossNamespacesConstants>(propertyValue.lVal) : slpcsanDontPersist);
	propertyValue.vt = VT_BOOL;
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
	hr = pPropertyBag->Read(OLESTR("SelectSortColumn"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL selectSortColumn = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("SetSortArrows"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL setSortArrows = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("SortOnHeaderClick"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL sortOnHeaderClick = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("ThumbnailSize"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	LONG thumbnailSize = (SUCCEEDED(hr) ? propertyValue.lVal : -1);
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("UseColumnResizability"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL useColumnResizability = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("UseFastButImpreciseItemHandling"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL useFastButImpreciseItemHandling = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
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
	UseSystemImageListConstants useSystemImageList = (SUCCEEDED(hr) ? static_cast<UseSystemImageListConstants>(propertyValue.lVal) : static_cast<UseSystemImageListConstants>(usilSmallImageList | usilLargeImageList | usilExtraLargeImageList));
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("UseThreadingForSlowColumns"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL useThreadingForSlowColumns = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);
	hr = pPropertyBag->Read(OLESTR("UseThumbnailDiskCache"), &propertyValue, pErrorLog);
	if(FAILED(hr) && hr != E_INVALIDARG) {
		return hr;
	}
	VARIANT_BOOL useThumbnailDiskCache = (SUCCEEDED(hr) ? propertyValue.boolVal : VARIANT_TRUE);

	hr = put_AutoEditNewItems(autoEditNewItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoInsertColumns(autoInsertColumns);
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
	hr = put_ColumnEnumerationTimeout(columnEnumerationTimeout);
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
		properties.pDefaultNamespaceEnumSettings->put_EnumerationFlags(static_cast<ShNamespaceEnumerationConstants>(snseIncludeFolders | snseMayIncludeHiddenItems | snseIncludeNonFolders));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_ExcludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_IncludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(0));
		properties.pDefaultNamespaceEnumSettings->put_IncludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(0));
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
	hr = put_DisplayFileTypeOverlays(displayFileTypeOverlays);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayThumbnailAdornments(displayThumbnailAdornments);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayThumbnails(displayThumbnails);
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
	hr = put_InitialSortOrder(initialSortOrder);
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
	hr = put_PersistColumnSettingsAcrossNamespaces(persistColumnSettingsAcrossNamespaces);
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
	hr = put_SelectSortColumn(selectSortColumn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SetSortArrows(setSortArrows);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SortOnHeaderClick(sortOnHeaderClick);
	if(FAILED(hr)) {
		return hr;
	}
	// TODO: Does this lead to a second cache flush after put_DisplayThumbnails?
	hr = put_ThumbnailSize(thumbnailSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseColumnResizability(useColumnResizability);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseFastButImpreciseItemHandling(useFastButImpreciseItemHandling);
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
	hr = put_UseThreadingForSlowColumns(useThreadingForSlowColumns);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseThumbnailDiskCache(useThumbnailDiskCache);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ShellListView::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/)
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoInsertColumns);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("AutoInsertColumns"), &propertyValue))) {
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
	propertyValue.lVal = properties.columnEnumerationTimeout;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ColumnEnumerationTimeout"), &propertyValue))) {
		return hr;
	}
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
	propertyValue.lVal = properties.displayFileTypeOverlays;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisplayFileTypeOverlays"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.displayThumbnailAdornments;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisplayThumbnailAdornments"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayThumbnails);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisplayThumbnails"), &propertyValue))) {
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
	propertyValue.lVal = properties.initialSortOrder;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("InitialSortOrder"), &propertyValue))) {
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
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.persistColumnSettingsAcrossNamespaces;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("PersistColumnSettingsAcrossNamespaces"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.preselectBasenameOnLabelEdit);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("PreselectBasenameOnLabelEdit"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processShellNotifications);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ProcessShellNotifications"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.selectSortColumn);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SelectSortColumn"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.setSortArrows);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SetSortArrows"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.sortOnHeaderClick);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SortOnHeaderClick"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.thumbnailSize;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ThumbnailSize"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useColumnResizability);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseColumnResizability"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useFastButImpreciseItemHandling);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseFastButImpreciseItemHandling"), &propertyValue))) {
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useThreadingForSlowColumns);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseThreadingForSlowColumns"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useThumbnailDiskCache);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseThumbnailDiskCache"), &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}

	return S_OK;
}

STDMETHODIMP ShellListView::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(pSize == NULL) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 15 VT_I4 properties...
	pSize->QuadPart += 15 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 17 VT_BOOL properties...
	pSize->QuadPart += 17 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

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

STDMETHODIMP ShellListView::Load(LPSTREAM pStream)
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
	if(version > 0x0102) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x01010101/*4x 0x01 (-> ShellListView)*/) {
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
	VARIANT_BOOL autoInsertColumns = propertyValue.boolVal;
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
	LONG columnEnumerationTimeout = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ShLvwManagedItemPropertiesConstants defaultManagedItemProperties = static_cast<ShLvwManagedItemPropertiesConstants>(propertyValue.lVal);

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
	ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays = static_cast<ShLvwDisplayFileTypeOverlaysConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments = static_cast<ShLvwDisplayThumbnailAdornmentsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayThumbnails = propertyValue.boolVal;
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
	SortOrderConstants initialSortOrder = static_cast<SortOrderConstants>(propertyValue.lVal);
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL selectSortColumn = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL setSortArrows = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL sortOnHeaderClick = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG thumbnailSize = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useColumnResizability = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useFastButImpreciseItemHandling = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	UseSystemImageListConstants useSystemImageList = static_cast<UseSystemImageListConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useThreadingForSlowColumns = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useThumbnailDiskCache = propertyValue.boolVal;
	UseGenericIconsConstants useGenericIcons = ugiOnlyForSlowItems;
	ShLvwPersistColumnSettingsAcrossNamespacesConstants persistColumnSettingsAcrossNamespaces = slpcsanDontPersist;
	if(version >= 0x0101) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		useGenericIcons = static_cast<UseGenericIconsConstants>(propertyValue.lVal);
		if(version >= 0x0102) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			persistColumnSettingsAcrossNamespaces = static_cast<ShLvwPersistColumnSettingsAcrossNamespacesConstants>(propertyValue.lVal);
		}
	}

	hr = put_AutoEditNewItems(autoEditNewItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoInsertColumns(autoInsertColumns);
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
	hr = put_ColumnEnumerationTimeout(columnEnumerationTimeout);
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
	hr = put_DisplayFileTypeOverlays(displayFileTypeOverlays);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayThumbnailAdornments(displayThumbnailAdornments);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayThumbnails(displayThumbnails);
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
	hr = put_InitialSortOrder(initialSortOrder);
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
	hr = put_PersistColumnSettingsAcrossNamespaces(persistColumnSettingsAcrossNamespaces);
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
	hr = put_SelectSortColumn(selectSortColumn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SetSortArrows(setSortArrows);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SortOnHeaderClick(sortOnHeaderClick);
	if(FAILED(hr)) {
		return hr;
	}
	// TODO: Does this lead to a second cache flush after put_DisplayThumbnails?
	hr = put_ThumbnailSize(thumbnailSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseColumnResizability(useColumnResizability);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseFastButImpreciseItemHandling(useFastButImpreciseItemHandling);
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
	hr = put_UseThreadingForSlowColumns(useThreadingForSlowColumns);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseThumbnailDiskCache(useThumbnailDiskCache);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ShellListView::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x01010101/*4x 0x01 (-> ShellListView)*/;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoInsertColumns);
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
	propertyValue.lVal = properties.columnEnumerationTimeout;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
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
	propertyValue.lVal = properties.displayFileTypeOverlays;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.displayThumbnailAdornments;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayThumbnails);
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
	propertyValue.lVal = properties.initialSortOrder;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.selectSortColumn);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.setSortArrows);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.sortOnHeaderClick);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.thumbnailSize;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useColumnResizability);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useFastButImpreciseItemHandling);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.useSystemImageList;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useThreadingForSlowColumns);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useThumbnailDiskCache);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0101 starts here
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.useGenericIcons;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.lVal = properties.persistColumnSettingsAcrossNamespaces;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


ShellListView::Properties::Properties()
{
	VariantInit(&emptyVariant);
	VariantClear(&emptyVariant);

	ZeroMemory(&itemDataForFastInsertion, sizeof(itemDataForFastInsertion));
	itemDataForFastInsertion.structSize = sizeof(EXLVMADDITEMDATA);
	itemDataForFastInsertion.iconIndex = -1;
	itemDataForFastInsertion.groupID = -2;

	ZeroMemory(&columnDataForFastInsertion, sizeof(columnDataForFastInsertion));
	columnDataForFastInsertion.structSize = sizeof(EXLVMADDCOLUMNDATA);
	columnDataForFastInsertion.insertAt = -1;

	hWndShellUIParentWindow = NULL;
	InitializeCriticalSectionAndSpinCount(&backgroundItemEnumResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&backgroundColumnEnumResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&slowColumnResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&subItemControlResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&infoTipResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&backgroundThumbnailResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&backgroundTileViewResultsCritSection, 400);
	InitializeCriticalSectionAndSpinCount(&tasksToEnqueueCritSection, 400);
	pDefaultNamespaceEnumSettings = NULL;

	pListItems = NULL;
	CComObject<ShellListViewItems>::CreateInstance(&pListItems);
	ATLASSUME(pListItems);
	pListItems->AddRef();

	pNamespaces = NULL;
	CComObject<ShellListViewNamespaces>::CreateInstance(&pNamespaces);
	ATLASSUME(pNamespaces);
	pNamespaces->AddRef();

	pColumns = NULL;
	CComObject<ShellListViewColumns>::CreateInstance(&pColumns);
	ATLASSUME(pColumns);
	pColumns->AddRef();

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

ShellListView::Properties::~Properties()
{
	if(pEnumTaskScheduler) {
		ATLASSERT(pEnumTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(pNormalTaskScheduler) {
		ATLASSERT(pNormalTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(pThumbnailTaskScheduler) {
		ATLASSERT(pThumbnailTaskScheduler->CountTasks(TOID_NULL) == 0);
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
	if(pListItems) {
		pListItems->Release();
		pListItems = NULL;
	}
	if(pColumns) {
		pColumns->Release();
		pColumns = NULL;
	}

	#ifdef USE_STL
		while(!shellItemDataOfInsertedItems.empty()) {
			delete shellItemDataOfInsertedItems.top();
			shellItemDataOfInsertedItems.pop();
		}
		/*while(!shellColumnIndexesOfInsertedColumns.empty()) {
			shellColumnIndexesOfInsertedColumns.pop();
		}*/
		EnterCriticalSection(&backgroundItemEnumResultsCritSection);
		while(!backgroundItemEnumResults.empty()) {
			LPSHLVWBACKGROUNDITEMENUMINFO pEnumResult = backgroundItemEnumResults.front();
			backgroundItemEnumResults.pop();
			if(pEnumResult) {
				if(pEnumResult->hPIDLBuffer) {
					DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundItemEnumResultsCritSection);
		EnterCriticalSection(&backgroundColumnEnumResultsCritSection);
		while(!backgroundColumnEnumResults.empty()) {
			LPSHLVWBACKGROUNDCOLUMNENUMINFO pEnumResult = backgroundColumnEnumResults.front();
			backgroundColumnEnumResults.pop();
			if(pEnumResult) {
				if(pEnumResult->hColumnBuffer) {
					DPA_DestroyCallback(pEnumResult->hColumnBuffer, FreeDPAColumnElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundColumnEnumResultsCritSection);
		EnterCriticalSection(&slowColumnResultsCritSection);
		while(!slowColumnResults.empty()) {
			LPSHLVWBACKGROUNDCOLUMNINFO pResult = slowColumnResults.front();
			slowColumnResults.pop();
			if(pResult) {
				delete pResult;
			}
		}
		LeaveCriticalSection(&slowColumnResultsCritSection);
		EnterCriticalSection(&subItemControlResultsCritSection);
		while(!subItemControlResults.empty()) {
			LPSHLVWBACKGROUNDCOLUMNINFO pResult = subItemControlResults.front();
			subItemControlResults.pop();
			if(pResult) {
				delete pResult;
			}
		}
		LeaveCriticalSection(&subItemControlResultsCritSection);
		EnterCriticalSection(&infoTipResultsCritSection);
		while(!infoTipResults.empty()) {
			LPSHLVWBACKGROUNDINFOTIPINFO pResult = infoTipResults.front();
			infoTipResults.pop();
			if(pResult) {
				if(pResult->pInfoTipText) {
					CoTaskMemFree(pResult->pInfoTipText);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&infoTipResultsCritSection);
		EnterCriticalSection(&backgroundThumbnailResultsCritSection);
		while(!backgroundThumbnailResults.empty()) {
			LPSHLVWBACKGROUNDTHUMBNAILINFO pResult = backgroundThumbnailResults.front();
			backgroundThumbnailResults.pop();
			if(pResult) {
				if(pResult->hThumbnailOrIcon) {
					DeleteObject(pResult->hThumbnailOrIcon);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&backgroundThumbnailResultsCritSection);
		EnterCriticalSection(&backgroundTileViewResultsCritSection);
		while(!backgroundTileViewResults.empty()) {
			LPSHLVWBACKGROUNDTILEVIEWINFO pResult = backgroundTileViewResults.front();
			backgroundTileViewResults.pop();
			if(pResult) {
				if(pResult->hColumnBuffer) {
					DPA_Destroy(pResult->hColumnBuffer);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&backgroundTileViewResultsCritSection);
		EnterCriticalSection(&tasksToEnqueueCritSection);
		while(!tasksToEnqueue.empty()) {
			LPSHCTLSBACKGROUNDTASKINFO pResult = tasksToEnqueue.front();
			tasksToEnqueue.pop();
			if(pResult) {
				if(pResult->pTask) {
					pResult->pTask->Release();
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&tasksToEnqueueCritSection);
	#else
		while(!shellItemDataOfInsertedItems.IsEmpty()) {
			delete shellItemDataOfInsertedItems.RemoveHead();
		}
		shellColumnIndexesOfInsertedColumns.RemoveAll();
		EnterCriticalSection(&backgroundItemEnumResultsCritSection);
		while(!backgroundItemEnumResults.IsEmpty()) {
			LPSHLVWBACKGROUNDITEMENUMINFO pEnumResult = backgroundItemEnumResults.RemoveHead();
			if(pEnumResult) {
				if(pEnumResult->hPIDLBuffer) {
					DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundItemEnumResultsCritSection);
		EnterCriticalSection(&backgroundColumnEnumResultsCritSection);
		while(!backgroundColumnEnumResults.IsEmpty()) {
			LPSHLVWBACKGROUNDCOLUMNENUMINFO pEnumResult = backgroundColumnEnumResults.RemoveHead();
			if(pEnumResult) {
				if(pEnumResult->hColumnBuffer) {
					DPA_DestroyCallback(pEnumResult->hColumnBuffer, FreeDPAColumnElement, NULL);
				}
				delete pEnumResult;
			}
		}
		LeaveCriticalSection(&backgroundColumnEnumResultsCritSection);
		EnterCriticalSection(&slowColumnResultsCritSection);
		while(!slowColumnResults.IsEmpty()) {
			LPSHLVWBACKGROUNDCOLUMNINFO pResult = slowColumnResults.RemoveHead();
			if(pResult) {
				delete pResult;
			}
		}
		LeaveCriticalSection(&slowColumnResultsCritSection);
		EnterCriticalSection(&subItemControlResultsCritSection);
		while(!subItemControlResults.IsEmpty()) {
			LPSHLVWBACKGROUNDCOLUMNINFO pResult = subItemControlResults.RemoveHead();
			if(pResult) {
				delete pResult;
			}
		}
		LeaveCriticalSection(&subItemControlResultsCritSection);
		EnterCriticalSection(&infoTipResultsCritSection);
		while(!infoTipResults.IsEmpty()) {
			LPSHLVWBACKGROUNDINFOTIPINFO pResult = infoTipResults.RemoveHead();
			if(pResult) {
				if(pResult->pInfoTipText) {
					CoTaskMemFree(pResult->pInfoTipText);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&infoTipResultsCritSection);
		EnterCriticalSection(&backgroundThumbnailResultsCritSection);
		while(!backgroundThumbnailResults.IsEmpty()) {
			LPSHLVWBACKGROUNDTHUMBNAILINFO pResult = backgroundThumbnailResults.RemoveHead();
			if(pResult) {
				if(pResult->hThumbnailOrIcon) {
					DeleteObject(pResult->hThumbnailOrIcon);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&backgroundThumbnailResultsCritSection);
		EnterCriticalSection(&backgroundTileViewResultsCritSection);
		while(!backgroundTileViewResults.IsEmpty()) {
			LPSHLVWBACKGROUNDTILEVIEWINFO pResult = backgroundTileViewResults.RemoveHead();
			if(pResult) {
				if(pResult->hColumnBuffer) {
					DPA_Destroy(pResult->hColumnBuffer);
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&backgroundTileViewResultsCritSection);
		EnterCriticalSection(&tasksToEnqueueCritSection);
		while(!tasksToEnqueue.IsEmpty()) {
			LPSHCTLSBACKGROUNDTASKINFO pResult = tasksToEnqueue.RemoveHead();
			if(pResult) {
				if(pResult->pTask) {
					pResult->pTask->Release();
				}
				delete pResult;
			}
		}
		LeaveCriticalSection(&tasksToEnqueueCritSection);
	#endif

	DeleteCriticalSection(&backgroundItemEnumResultsCritSection);
	DeleteCriticalSection(&backgroundColumnEnumResultsCritSection);
	DeleteCriticalSection(&slowColumnResultsCritSection);
	DeleteCriticalSection(&subItemControlResultsCritSection);
	DeleteCriticalSection(&infoTipResultsCritSection);
	DeleteCriticalSection(&backgroundThumbnailResultsCritSection);
	DeleteCriticalSection(&backgroundTileViewResultsCritSection);
	DeleteCriticalSection(&tasksToEnqueueCritSection);
}


HRESULT ShellListView::OnDraw(ATL_DRAWINFO& drawInfo)
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

STDMETHODIMP ShellListView::IOleObject_SetExtent(DWORD /*drawAspect*/, SIZEL* /*pSize*/)
{
	return E_FAIL;
}

HWND ShellListView::GethWndShellUIParentWindow(void)
{
	return properties.hWndShellUIParentWindow;
}

//////////////////////////////////////////////////////////////////////
// implementation of IContextMenuCB
STDMETHODIMP ShellListView::CallBack(LPSHELLFOLDER pShellFolder, HWND hWnd, LPDATAOBJECT pDataObject, UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: Check whether we've to handle e. g. DFM_WM_DRAWITEM to make owner-drawn menus work.
	//ATLASSERT(message != DFM_WM_MEASUREITEM);
	//ATLASSERT(message != DFM_WM_DRAWITEM);
	//ATLASSERT(message != DFM_WM_INITMENUPOPUP);
	//ATLASSERT(message != DFM_INVOKECOMMANDEX);
	ATLTRACE2(shlvwTraceContextMenus, 3, TEXT("Callback was called with pShellFolder=0x%X, hWnd=0x%X, pDataObject=0x%X, message=0x%X, wParam=0x%X, lParam=0x%X\n"), pShellFolder, hWnd, pDataObject, message, wParam, lParam);
	return ContextMenuCallback(pShellFolder, hWnd, pDataObject, message, wParam, lParam);
}
// implementation of IContextMenuCB
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ShellListView::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Thumbnails:
			*pName = GetResString(IDPC_THUMBNAILS).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP ShellListView::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_SHLVW_COLORCOMPRESSEDITEMS:
		case DISPID_SHLVW_COLORENCRYPTEDITEMS:
		case DISPID_SHLVW_HWNDSHELLUIPARENTWINDOW:
		case DISPID_SHLVW_SELECTSORTCOLUMN:
		case DISPID_SHLVW_SETSORTARROWS:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_SHLVW_AUTOEDITNEWITEMS:
		case DISPID_SHLVW_AUTOINSERTCOLUMNS:
		case DISPID_SHLVW_COLUMNENUMERATIONTIMEOUT:
		case DISPID_SHLVW_DEFAULTMANAGEDITEMPROPERTIES:
		case DISPID_SHLVW_DEFAULTNAMESPACEENUMSETTINGS:
		case DISPID_SHLVW_DISABLEDEVENTS:
		case DISPID_SHLVW_DISPLAYELEVATIONSHIELDOVERLAYS:
		case DISPID_SHLVW_HANDLEOLEDRAGDROP:
		case DISPID_SHLVW_HIDDENITEMSSTYLE:
		case DISPID_SHLVW_INFOTIPFLAGS:
		case DISPID_SHLVW_INITIALSORTORDER:
		case DISPID_SHLVW_ITEMENUMERATIONTIMEOUT:
		case DISPID_SHLVW_ITEMTYPESORTORDER:
		case DISPID_SHLVW_LIMITLABELEDITINPUT:
		case DISPID_SHLVW_LOADOVERLAYSONDEMAND:
		case DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES:
		case DISPID_SHLVW_PRESELECTBASENAMEONLABELEDIT:
		case DISPID_SHLVW_PROCESSSHELLNOTIFICATIONS:
		case DISPID_SHLVW_SORTCOLUMNINDEX:
		case DISPID_SHLVW_SORTONHEADERCLICK:
		case DISPID_SHLVW_USECOLUMNRESIZABILITY:
		case DISPID_SHLVW_USEFASTBUTIMPRECISEITEMHANDLING:
		case DISPID_SHLVW_USEGENERICICONS:
		case DISPID_SHLVW_USESYSTEMIMAGELIST:
		case DISPID_SHLVW_USETHREADINGFORSLOWCOLUMNS:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_SHLVW_APPID:
		case DISPID_SHLVW_APPNAME:
		case DISPID_SHLVW_APPSHORTNAME:
		case DISPID_SHLVW_BUILD:
		case DISPID_SHLVW_CHARSET:
		case DISPID_SHLVW_HIMAGELIST:
		case DISPID_SHLVW_HWND:
		case DISPID_SHLVW_ISRELEASE:
		case DISPID_SHLVW_PROGRAMMER:
		case DISPID_SHLVW_TESTER:
		case DISPID_SHLVW_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_SHLVW_COLUMNS:
		case DISPID_SHLVW_LISTITEMS:
		case DISPID_SHLVW_NAMESPACES:
		#ifdef ACTIVATE_SECZONE_FEATURES
			case DISPID_SHLVW_SECURITYZONES:
		#endif
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS:
		case DISPID_SHLVW_DISPLAYTHUMBNAILADORNMENTS:
		case DISPID_SHLVW_DISPLAYTHUMBNAILS:
		case DISPID_SHLVW_THUMBNAILSIZE:
			*pCategory = PROPCAT_Thumbnails;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ShellListView::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ShellListView::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ShellListView::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ShellBrowserControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ShellListView::GetSpecialThanks(void)
{
	return TEXT("Jim Barry, Geoff Chappell, Wine Headquarters");
}

CAtlString ShellListView::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ShellListView::GetVersion(void)
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
STDMETHODIMP ShellListView::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS:
			hr = GetDisplayStringForSetting(property, properties.displayFileTypeOverlays, description);
			break;
		case DISPID_SHLVW_HANDLEOLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, properties.handleOLEDragDrop, description);
			break;
		case DISPID_SHLVW_HIDDENITEMSSTYLE:
			hr = GetDisplayStringForSetting(property, properties.hiddenItemsStyle, description);
			break;
		case DISPID_SHLVW_INITIALSORTORDER:
			hr = GetDisplayStringForSetting(property, properties.initialSortOrder, description);
			break;
		case DISPID_SHLVW_ITEMTYPESORTORDER:
			hr = GetDisplayStringForSetting(property, properties.itemTypeSortOrder, description);
			break;
		case DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES:
			hr = GetDisplayStringForSetting(property, properties.persistColumnSettingsAcrossNamespaces, description);
			break;
		case DISPID_SHLVW_USEGENERICICONS:
			hr = GetDisplayStringForSetting(property, properties.useGenericIcons, description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellListView>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ShellListView::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_SHLVW_INITIALSORTORDER:
		case DISPID_SHLVW_ITEMTYPESORTORDER:
		case DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES:
			c = 2;
			break;
		case DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS:
		case DISPID_SHLVW_HIDDENITEMSSTYLE:
		case DISPID_SHLVW_USEGENERICICONS:
			c = 3;
			break;
		case DISPID_SHLVW_HANDLEOLEDRAGDROP:
			c = 6;
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellListView>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(property == DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS) {
			propertyValue--;
		} else if(property == DISPID_SHLVW_HANDLEOLEDRAGDROP && iDescription >= 4) {
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

STDMETHODIMP ShellListView::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS:
		case DISPID_SHLVW_HANDLEOLEDRAGDROP:
		case DISPID_SHLVW_HIDDENITEMSSTYLE:
		case DISPID_SHLVW_INITIALSORTORDER:
		case DISPID_SHLVW_ITEMTYPESORTORDER:
		case DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES:
		case DISPID_SHLVW_USEGENERICICONS:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ShellListView>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ShellListView::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	if(!pPropertyPage) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_SHLVW_DEFAULTMANAGEDITEMPROPERTIES:
			*pPropertyPage = CLSID_ShLvwDefaultManagedItemPropertiesProperties;
			return S_OK;
			break;
		case DISPID_SHLVW_DEFAULTNAMESPACEENUMSETTINGS:
			*pPropertyPage = CLSID_NamespaceEnumSettingsProperties;
			return S_OK;
			break;
		case DISPID_SHLVW_DISABLEDEVENTS:
		case DISPID_SHLVW_INFOTIPFLAGS:
		case DISPID_SHLVW_USESYSTEMIMAGELIST:
			*pPropertyPage = CLSID_CommonProperties;
			return S_OK;
			break;
		case DISPID_SHLVW_DISPLAYTHUMBNAILADORNMENTS:
			*pPropertyPage = CLSID_ThumbnailsProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<ShellListView>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ShellListView::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS:
			switch(cookie) {
				case sldftoFollowSystemSettings:
					description = GetResStringWithNumber(IDP_DISPLAYFILETYPEOVERLAYSFOLLOWSYSTEMSETTINGS, sldftoFollowSystemSettings);
					break;
				case sldftoNone:
					description = GetResStringWithNumber(IDP_DISPLAYFILETYPEOVERLAYSNONE, sldftoNone);
					break;
				case sldftoExecutableIcon:
					description = GetResStringWithNumber(IDP_DISPLAYFILETYPEOVERLAYSEXECUTABLEICON, sldftoExecutableIcon);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHLVW_HANDLEOLEDRAGDROP:
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
		case DISPID_SHLVW_HIDDENITEMSSTYLE:
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
		case DISPID_SHLVW_INITIALSORTORDER:
			switch(cookie) {
				case soAscending:
					description = GetResStringWithNumber(IDP_INITIALSORTORDERASCENDING, soAscending);
					break;
				case soDescending:
					description = GetResStringWithNumber(IDP_INITIALSORTORDERDESCENDING, soDescending);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHLVW_ITEMTYPESORTORDER:
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
		case DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES:
			switch(cookie) {
				case slpcsanDontPersist:
					description = GetResStringWithNumber(IDP_PERSISTCOLUMNSETTINGSACROSSNAMESPACESDONTPERSIST, slpcsanDontPersist);
					break;
				case slpcsanPersist:
					description = GetResStringWithNumber(IDP_PERSISTCOLUMNSETTINGSACROSSNAMESPACESPERSIST, slpcsanPersist);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SHLVW_USEGENERICICONS:
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
STDMETHODIMP ShellListView::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_ShLvwDefaultManagedItemPropertiesProperties;
		pPropertyPages->pElems[2] = CLSID_NamespaceEnumSettingsProperties;
		pPropertyPages->pElems[3] = CLSID_ThumbnailsProperties;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ShellListView::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ShellListView>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ShellListView::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[2] = {
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 2, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

HRESULT ShellListView::DoVerbAbout(HWND hWndParent)
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


STDMETHODIMP ShellListView::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 6;
	return S_OK;
}

STDMETHODIMP ShellListView::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ShellBrowserControls");
	return S_OK;
}

STDMETHODIMP ShellListView::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ShBrowserCtls");
	return S_OK;
}

STDMETHODIMP ShellListView::get_AutoEditNewItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoEditNewItems);
	return S_OK;
}

STDMETHODIMP ShellListView::put_AutoEditNewItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_AUTOEDITNEWITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoEditNewItems != b) {
		properties.autoEditNewItems = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_AUTOEDITNEWITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_AutoInsertColumns(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoInsertColumns);
	return S_OK;
}

STDMETHODIMP ShellListView::put_AutoInsertColumns(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_AUTOINSERTCOLUMNS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoInsertColumns != b) {
		properties.autoInsertColumns = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_AUTOINSERTCOLUMNS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ShellListView::get_CharSet(BSTR* pValue)
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

STDMETHODIMP ShellListView::get_ColorCompressedItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.colorCompressedItems);
	return S_OK;
}

STDMETHODIMP ShellListView::put_ColorCompressedItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_COLORCOMPRESSEDITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.colorCompressedItems != b) {
		properties.colorCompressedItems = b;
		SetDirty(TRUE);

		if(attachedControl.IsWindow()) {
			attachedControl.Invalidate();
		}
		FireOnChanged(DISPID_SHLVW_COLORCOMPRESSEDITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ColorEncryptedItems(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.colorEncryptedItems);
	return S_OK;
}

STDMETHODIMP ShellListView::put_ColorEncryptedItems(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_COLORENCRYPTEDITEMS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.colorEncryptedItems != b) {
		properties.colorEncryptedItems = b;
		SetDirty(TRUE);

		if(attachedControl.IsWindow()) {
			attachedControl.Invalidate();
		}
		FireOnChanged(DISPID_SHLVW_COLORENCRYPTEDITEMS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ColumnEnumerationTimeout(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.columnEnumerationTimeout;
	return S_OK;
}

STDMETHODIMP ShellListView::put_ColumnEnumerationTimeout(LONG newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_COLUMNENUMERATIONTIMEOUT);
	if((newValue < 1000) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.columnEnumerationTimeout != newValue) {
		properties.columnEnumerationTimeout = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_COLUMNENUMERATIONTIMEOUT);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_Columns(IShListViewColumns** ppColumns)
{
	ATLASSERT_POINTER(ppColumns, IShListViewColumns*);
	if(!ppColumns) {
		return E_POINTER;
	}

	if(attachedControl.IsWindow()) {
		return properties.pColumns->QueryInterface(IID_IShListViewColumns, reinterpret_cast<LPVOID*>(ppColumns));
	}
	return E_FAIL;
}

STDMETHODIMP ShellListView::get_DefaultManagedItemProperties(ShLvwManagedItemPropertiesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwManagedItemPropertiesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.defaultManagedItemProperties;
	return S_OK;
}

STDMETHODIMP ShellListView::put_DefaultManagedItemProperties(ShLvwManagedItemPropertiesConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DEFAULTMANAGEDITEMPROPERTIES);
	if(properties.defaultManagedItemProperties != newValue) {
		properties.defaultManagedItemProperties = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_DEFAULTMANAGEDITEMPROPERTIES);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_DefaultNamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings)
{
	ATLASSERT_POINTER(ppEnumSettings, INamespaceEnumSettings*);
	if(!ppEnumSettings) {
		return E_POINTER;
	}

	*ppEnumSettings = NULL;
	if(properties.pDefaultNamespaceEnumSettings) {
		return properties.pDefaultNamespaceEnumSettings->QueryInterface(IID_INamespaceEnumSettings, reinterpret_cast<LPVOID*>(ppEnumSettings));
	}
	return S_OK;
}

STDMETHODIMP ShellListView::putref_DefaultNamespaceEnumSettings(INamespaceEnumSettings* pEnumSettings)
{
	PUTPROPPROLOG(DISPID_SHLVW_DEFAULTNAMESPACEENUMSETTINGS);
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
	FireOnChanged(DISPID_SHLVW_DEFAULTNAMESPACEENUMSETTINGS);
	return S_OK;
}

STDMETHODIMP ShellListView::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ShellListView::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_DisplayElevationShieldOverlays(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayElevationShieldOverlays);
	return S_OK;
}

STDMETHODIMP ShellListView::put_DisplayElevationShieldOverlays(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DISPLAYELEVATIONSHIELDOVERLAYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayElevationShieldOverlays != b) {
		properties.displayElevationShieldOverlays = b;
		SetDirty(TRUE);

		BOOL invalidate = FALSE;
		if(properties.thumbnailsStatus.pThumbnailImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			pImgLstPriv->SetOptions(ILOF_DISPLAYELEVATIONSHIELDS, properties.displayElevationShieldOverlays ? ILOF_DISPLAYELEVATIONSHIELDS : 0);
			invalidate = TRUE;
		}
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

		FireOnChanged(DISPID_SHLVW_DISPLAYELEVATIONSHIELDOVERLAYS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_DisplayFileTypeOverlays(ShLvwDisplayFileTypeOverlaysConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwDisplayFileTypeOverlaysConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.displayFileTypeOverlays;
	return S_OK;
}

STDMETHODIMP ShellListView::put_DisplayFileTypeOverlays(ShLvwDisplayFileTypeOverlaysConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS);
	if(properties.displayFileTypeOverlays != newValue) {
		properties.displayFileTypeOverlays = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener && properties.displayThumbnails) {
			// TODO: Optimize this so that we don't update the thumbnails themselves.
			FlushIcons(-1, TRUE, TRUE);
		}
		FireOnChanged(DISPID_SHLVW_DISPLAYFILETYPEOVERLAYS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_DisplayThumbnailAdornments(ShLvwDisplayThumbnailAdornmentsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwDisplayThumbnailAdornmentsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.displayThumbnailAdornments;
	return S_OK;
}

STDMETHODIMP ShellListView::put_DisplayThumbnailAdornments(ShLvwDisplayThumbnailAdornmentsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DISPLAYTHUMBNAILADORNMENTS);
	if(properties.displayThumbnailAdornments != newValue) {
		properties.displayThumbnailAdornments = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener && properties.displayThumbnails) {
			// TODO: Optimize this so that we don't update the thumbnails themselves.
			FlushIcons(-1, TRUE, TRUE);
		}
		FireOnChanged(DISPID_SHLVW_DISPLAYTHUMBNAILADORNMENTS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_DisplayThumbnails(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayThumbnails);
	return S_OK;
}

STDMETHODIMP ShellListView::put_DisplayThumbnails(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_DISPLAYTHUMBNAILS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayThumbnails != b) {
		properties.displayThumbnails = b;
		SetDirty(TRUE);

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener) {
			SetSystemImageLists();
			FlushIcons(-1, TRUE, TRUE);
		}
		FireOnChanged(DISPID_SHLVW_DISPLAYTHUMBNAILS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_HandleOLEDragDrop(HandleOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HandleOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.handleOLEDragDrop;
	return S_OK;
}

STDMETHODIMP ShellListView::put_HandleOLEDragDrop(HandleOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_HANDLEOLEDRAGDROP);
	if(properties.handleOLEDragDrop != newValue) {
		properties.handleOLEDragDrop = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_HANDLEOLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_HiddenItemsStyle(HiddenItemsStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HiddenItemsStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hiddenItemsStyle;
	return S_OK;
}

STDMETHODIMP ShellListView::put_HiddenItemsStyle(HiddenItemsStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_HIDDENITEMSSTYLE);
	if(properties.hiddenItemsStyle != newValue) {
		properties.hiddenItemsStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_HIDDENITEMSSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(attachedControl));
	return S_OK;
}

STDMETHODIMP ShellListView::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
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

STDMETHODIMP ShellListView::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_HIMAGELIST);

	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilNonShellItems:
			properties.hImageList[0] = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			if(properties.thumbnailsStatus.pThumbnailImageList) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
			}
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

	FireOnChanged(DISPID_SHLVW_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_hWndShellUIParentWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hWndShellUIParentWindow);
	return S_OK;
}

STDMETHODIMP ShellListView::put_hWndShellUIParentWindow(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_HWNDSHELLUIPARENTWINDOW);
	if(properties.hWndShellUIParentWindow != static_cast<HWND>(LongToHandle(newValue))) {
		properties.hWndShellUIParentWindow = static_cast<HWND>(LongToHandle(newValue));
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_HWNDSHELLUIPARENTWINDOW);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_InfoTipFlags(InfoTipFlagsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, InfoTipFlagsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.infoTipFlags;
	return S_OK;
}

STDMETHODIMP ShellListView::put_InfoTipFlags(InfoTipFlagsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_INFOTIPFLAGS);
	if(properties.infoTipFlags != newValue) {
		properties.infoTipFlags = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_INFOTIPFLAGS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_InitialSortOrder(SortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.initialSortOrder;
	return S_OK;
}

STDMETHODIMP ShellListView::put_InitialSortOrder(SortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_INITIALSORTORDER);
	if(properties.initialSortOrder != newValue) {
		properties.initialSortOrder = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_INITIALSORTORDER);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP ShellListView::get_ItemEnumerationTimeout(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemEnumerationTimeout;
	return S_OK;
}

STDMETHODIMP ShellListView::put_ItemEnumerationTimeout(LONG newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_ITEMENUMERATIONTIMEOUT);
	if((newValue < 1000) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.itemEnumerationTimeout != newValue) {
		properties.itemEnumerationTimeout = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_ITEMENUMERATIONTIMEOUT);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ItemTypeSortOrder(ItemTypeSortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemTypeSortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemTypeSortOrder;
	return S_OK;
}

STDMETHODIMP ShellListView::put_ItemTypeSortOrder(ItemTypeSortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_ITEMTYPESORTORDER);
	if(properties.itemTypeSortOrder != newValue) {
		properties.itemTypeSortOrder = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_ITEMTYPESORTORDER);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_LimitLabelEditInput(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.limitLabelEditInput);
	return S_OK;
}

STDMETHODIMP ShellListView::put_LimitLabelEditInput(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_LIMITLABELEDITINPUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.limitLabelEditInput != b) {
		properties.limitLabelEditInput = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_LIMITLABELEDITINPUT);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ListItems(IShListViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IShListViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	if(attachedControl.IsWindow()) {
		return properties.pListItems->QueryInterface(IID_IShListViewItems, reinterpret_cast<LPVOID*>(ppItems));
	}
	return E_FAIL;
}

STDMETHODIMP ShellListView::get_LoadOverlaysOnDemand(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.loadOverlaysOnDemand);
	return S_OK;
}

STDMETHODIMP ShellListView::put_LoadOverlaysOnDemand(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_LOADOVERLAYSONDEMAND);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.loadOverlaysOnDemand != b) {
		properties.loadOverlaysOnDemand = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_LOADOVERLAYSONDEMAND);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_Namespaces(IShListViewNamespaces** ppNamespaces)
{
	ATLASSERT_POINTER(ppNamespaces, IShListViewNamespaces*);
	if(!ppNamespaces) {
		return E_POINTER;
	}

	if(attachedControl.IsWindow()) {
		return properties.pNamespaces->QueryInterface(IID_IShListViewNamespaces, reinterpret_cast<LPVOID*>(ppNamespaces));
	}
	return E_FAIL;
}

STDMETHODIMP ShellListView::get_PersistColumnSettingsAcrossNamespaces(ShLvwPersistColumnSettingsAcrossNamespacesConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ShLvwPersistColumnSettingsAcrossNamespacesConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.persistColumnSettingsAcrossNamespaces;
	return S_OK;
}

STDMETHODIMP ShellListView::put_PersistColumnSettingsAcrossNamespaces(ShLvwPersistColumnSettingsAcrossNamespacesConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES);
	if(properties.persistColumnSettingsAcrossNamespaces != newValue) {
		properties.persistColumnSettingsAcrossNamespaces = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_PERSISTCOLUMNSETTINGSACROSSNAMESPACES);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_PreselectBasenameOnLabelEdit(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.preselectBasenameOnLabelEdit);
	return S_OK;
}

STDMETHODIMP ShellListView::put_PreselectBasenameOnLabelEdit(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_PRESELECTBASENAMEONLABELEDIT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.preselectBasenameOnLabelEdit != b) {
		properties.preselectBasenameOnLabelEdit = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_PRESELECTBASENAMEONLABELEDIT);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ProcessShellNotifications(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processShellNotifications);
	return S_OK;
}

STDMETHODIMP ShellListView::put_ProcessShellNotifications(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_PROCESSSHELLNOTIFICATIONS);
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
		FireOnChanged(DISPID_SHLVW_PROCESSSHELLNOTIFICATIONS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

#ifdef ACTIVATE_SECZONE_FEATURES
	STDMETHODIMP ShellListView::get_SecurityZones(ISecurityZones** ppZones)
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

STDMETHODIMP ShellListView::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ShellListView::get_SelectSortColumn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.selectSortColumn);
	return S_OK;
}

STDMETHODIMP ShellListView::put_SelectSortColumn(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_SELECTSORTCOLUMN);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.selectSortColumn != b) {
		properties.selectSortColumn = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_SELECTSORTCOLUMN);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_SetSortArrows(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.setSortArrows);
	return S_OK;
}

STDMETHODIMP ShellListView::put_SetSortArrows(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_SETSORTARROWS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.setSortArrows != b) {
		properties.setSortArrows = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_SETSORTARROWS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_SortColumnIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(sortingSettings.columnIndex == -1) {
		*pValue = properties.columnsStatus.defaultSortColumn;
	} else {
		*pValue = sortingSettings.columnIndex;
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_SortOnHeaderClick(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.sortOnHeaderClick);
	return S_OK;
}

STDMETHODIMP ShellListView::put_SortOnHeaderClick(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_SORTONHEADERCLICK);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.sortOnHeaderClick != b) {
		properties.sortOnHeaderClick = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_SORTONHEADERCLICK);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_ThumbnailSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = (IsInDesignMode() || !properties.pAttachedInternalMessageListener ? properties.thumbnailSize : properties.thumbnailsStatus.thumbnailSize.cx);
	return S_OK;
}

STDMETHODIMP ShellListView::put_ThumbnailSize(LONG newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_THUMBNAILSIZE);
	if(newValue < -5 || (newValue >= 0 && newValue < 16) || newValue > 256) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.thumbnailSize != newValue) {
		properties.thumbnailSize = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener) {
			if(properties.displayThumbnails) {
				// resetting all icons and overlays is done in SetSystemImageLists
				SetSystemImageLists();
			}
		}
		FireOnChanged(DISPID_SHLVW_THUMBNAILSIZE);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseColumnResizability(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useColumnResizability);
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseColumnResizability(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USECOLUMNRESIZABILITY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useColumnResizability != b) {
		properties.useColumnResizability = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_USECOLUMNRESIZABILITY);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseFastButImpreciseItemHandling(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useFastButImpreciseItemHandling);
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseFastButImpreciseItemHandling(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USEFASTBUTIMPRECISEITEMHANDLING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useFastButImpreciseItemHandling != b) {
		properties.useFastButImpreciseItemHandling = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_USEFASTBUTIMPRECISEITEMHANDLING);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseGenericIcons(UseGenericIconsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UseGenericIconsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.useGenericIcons;
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseGenericIcons(UseGenericIconsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USEGENERICICONS);
	if(properties.useGenericIcons != newValue) {
		properties.useGenericIcons = newValue;

		if(!IsInDesignMode() && properties.pAttachedInternalMessageListener) {
			FlushIcons(-1, FALSE, TRUE);
		}
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_USEGENERICICONS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseSystemImageList(UseSystemImageListConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UseSystemImageListConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.useSystemImageList;
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseSystemImageList(UseSystemImageListConstants newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USESYSTEMIMAGELIST);
	if(properties.useSystemImageList != newValue) {
		properties.useSystemImageList = newValue;
		SetDirty(TRUE);
		SetSystemImageLists();
		FireOnChanged(DISPID_SHLVW_USESYSTEMIMAGELIST);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseThreadingForSlowColumns(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useThreadingForSlowColumns);
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseThreadingForSlowColumns(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USETHREADINGFORSLOWCOLUMNS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useThreadingForSlowColumns != b) {
		properties.useThreadingForSlowColumns = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_USETHREADINGFORSLOWCOLUMNS);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_UseThumbnailDiskCache(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useThumbnailDiskCache);
	return S_OK;
}

STDMETHODIMP ShellListView::put_UseThumbnailDiskCache(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SHLVW_USETHUMBNAILDISKCACHE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useThumbnailDiskCache != b) {
		properties.useThumbnailDiskCache = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SHLVW_USETHUMBNAILDISKCACHE);
	}
	return S_OK;
}

STDMETHODIMP ShellListView::get_Version(BSTR* pValue)
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


STDMETHODIMP ShellListView::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ShellListView::Attach(OLE_HANDLE hWnd)
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

	ATLASSUME(properties.pListItems);
	properties.pListItems->SetOwner(this);
	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(this);
	ATLASSUME(properties.pColumns);
	properties.pColumns->SetOwner(this);

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

	HWND hWndLvw = static_cast<HWND>(LongToHandle(hWnd));
	if(SendMessage(hWndLvw, SHBM_HANDSHAKE, 0, reinterpret_cast<LPARAM>(&handshakeDetails))) {
		ATLTRACE2(shlvwTraceCtlCommunication, 3, TEXT("Handshake with 0x%X succeeded\n"), hWndLvw);
		attachedControl.Attach(hWndLvw);
		AddRef();
		ConfigureControl();
		return S_OK;
	} else {
		if(handshakeDetails.processed) {
			ATLTRACE2(shlvwTraceCtlCommunication, 0, TEXT("Handshake with 0x%X failed (0x%X)\n"), hWndLvw, handshakeDetails.errorCode);
			return handshakeDetails.errorCode;
		} else {
			ATLTRACE2(shlvwTraceCtlCommunication, 0, TEXT("Handshake with 0x%X failed (no response)\n"), hWndLvw);
		}
	}
	ATLASSUME(properties.pListItems);
	properties.pListItems->SetOwner(NULL);
	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(NULL);
	ATLASSUME(properties.pColumns);
	properties.pColumns->SetOwner(NULL);
	return E_FAIL;
}

STDMETHODIMP ShellListView::CompareItems(IShListViewItem* pFirstItem, IShListViewItem* pSecondItem, LONG sortColumnIndex/* = -1*/, LONG* pResult/* = NULL*/)
{
	ATLASSERT_POINTER(pFirstItem, IShListViewItem);
	ATLASSERT_POINTER(pSecondItem, IShListViewItem);
	ATLASSERT_POINTER(pResult, LONG);
	if(!pFirstItem || !pSecondItem || !pResult) {
		return E_POINTER;
	}

	LONG itemID1 = -1;
	pFirstItem->get_ID(&itemID1);
	LONG itemID2 = -1;
	pSecondItem->get_ID(&itemID2);

	PCIDLIST_ABSOLUTE pIDL1 = ItemIDToPIDL(itemID1);
	PCIDLIST_ABSOLUTE pIDL2 = ItemIDToPIDL(itemID2);
	if(!pIDL1 || !pIDL2) {
		return E_INVALIDARG;
	}

	if(properties.useFastButImpreciseItemHandling) {
		// try to use the column namespace's IShellFolder interface
		if(properties.columnsStatus.pNamespaceSHF) {
			*pResult = static_cast<short>(HRESULT_CODE(properties.columnsStatus.pNamespaceSHF->CompareIDs(MAKELPARAM(max(sortColumnIndex, properties.columnsStatus.defaultSortColumn), 0), ILFindLastID(pIDL1), ILFindLastID(pIDL2))));
			return S_OK;
		}
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
		*pResult = static_cast<short>(HRESULT_CODE(pParentISF->CompareIDs(MAKELPARAM(max(sortColumnIndex, properties.columnsStatus.defaultSortColumn), 0), ILFindLastID(pIDL1), ILFindLastID(pIDL2))));
		return S_OK;
	}
	ILFree(pIDLParent1);

	/* TODO: On Vista, try to use SHCompareIDsFull (but check with "My Computer" whether sorting is still
	         right then!) */
	// fallback to the Desktop's IShellFolder interface
	*pResult = static_cast<short>(HRESULT_CODE(properties.pDesktopISF->CompareIDs(MAKELPARAM(max(sortColumnIndex, properties.columnsStatus.defaultSortColumn), 0), pIDL1, pIDL2)));
	return S_OK;
}

STDMETHODIMP ShellListView::CreateHeaderContextMenu(OLE_HANDLE* pMenu)
{
	ATLASSERT_POINTER(pMenu, OLE_HANDLE);
	if(!pMenu) {
		return E_POINTER;
	}

	ULONG defaultDisplayColumn = 0;
	// Windows Explorer always prevents column 0 from being removed
	/*ULONG dummy;
	if(properties.columnsStatus.pNamespaceSHF2) {
		properties.columnsStatus.pNamespaceSHF2->GetDefaultColumn(0, &dummy, &defaultDisplayColumn);
	}*/

	HMENU hMenu = NULL;
	HRESULT hr = CreateHeaderContextMenu(defaultDisplayColumn, hMenu);
	*pMenu = HandleToLong(hMenu);
	return hr;
}

STDMETHODIMP ShellListView::CreateShellContextMenu(VARIANT items, OLE_HANDLE* pMenu)
{
	ATLASSERT_POINTER(pMenu, OLE_HANDLE);
	if(!pMenu) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	if(items.vt == VT_DISPATCH && items.pdispVal) {
		CComQIPtr<IShListViewNamespace> pShLvwNamespace = items.pdispVal;
		if(pShLvwNamespace) {
			// a single ShellListViewNamespace object
			LONG pIDLNamespace = 0;
			hr = pShLvwNamespace->get_FullyQualifiedPIDL(&pIDLNamespace);
			if(SUCCEEDED(hr)) {
				HMENU hMenu = NULL;
				hr = CreateShellContextMenu(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDLNamespace), hMenu);
				*pMenu = HandleToLong(hMenu);
			}
			return hr;
		}
	}

	UINT itemCount = 0;
	PLONG pItems = NULL;
	hr = VariantToItemIDs(items, pItems, itemCount);
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

STDMETHODIMP ShellListView::DestroyHeaderContextMenu(void)
{
	HRESULT hr = E_FAIL;
	if(contextMenuData.currentHeaderContextMenu.IsMenu()) {
		Raise_DestroyingHeaderContextMenu(HandleToLong(static_cast<HMENU>(contextMenuData.currentHeaderContextMenu)));
		hr = S_OK;
		contextMenuData.currentHeaderContextMenu.DestroyMenu();
	}
	return hr;
}

STDMETHODIMP ShellListView::DestroyShellContextMenu(void)
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

STDMETHODIMP ShellListView::Detach(void)
{
	ATLASSERT(attachedControl.IsWindow());

	flags.detaching = TRUE;
	if(properties.columnsStatus.hColumnsReadyEvent) {
		// there might a task be waiting for this event
		SetEvent(properties.columnsStatus.hColumnsReadyEvent);
	}
	flags.acceptNewTasks = FALSE;
	if(attachedControl.m_hWnd) {
		DeregisterForShellNotifications();
	}
	if(properties.pEnumTaskScheduler) {
		properties.pEnumTaskScheduler->RemoveTasks(TASKID_ShLvwBackgroundItemEnum, NULL, TRUE);
		properties.pEnumTaskScheduler->RemoveTasks(TASKID_ShLvwBackgroundColumnEnum, NULL, TRUE);
		ATLASSERT(properties.pEnumTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(properties.pNormalTaskScheduler) {
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwAutoUpdate, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ElevationShield, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwIcon, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwOverlay, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwSlowColumn, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwSubItemControl, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwTileView, NULL, TRUE);
		properties.pNormalTaskScheduler->RemoveTasks(TASKID_ShLvwInfoTip, NULL, TRUE);
		ATLASSERT(properties.pNormalTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(properties.pThumbnailTaskScheduler) {
		properties.pThumbnailTaskScheduler->RemoveTasks(TASKID_ShLvwThumbnail, NULL, TRUE);
		properties.pThumbnailTaskScheduler->RemoveTasks(TASKID_ShLvwThumbnailTestCache, NULL, TRUE);
		properties.pThumbnailTaskScheduler->RemoveTasks(TASKID_ShLvwExtractThumbnail, NULL, TRUE);
		properties.pThumbnailTaskScheduler->RemoveTasks(TASKID_ShLvwExtractThumbnailFromDiskCache, NULL, TRUE);
		ATLASSERT(properties.pThumbnailTaskScheduler->CountTasks(TOID_NULL) == 0);
	}
	if(properties.displayThumbnails && properties.thumbnailsStatus.pThumbnailImageList) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, NULL)));
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, NULL)));
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, NULL)));
	} else {
		if(properties.useSystemImageList & usilSmallImageList) {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, NULL)));
		}
		if(properties.useSystemImageList & usilLargeImageList) {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, NULL)));
		}
		if(properties.useSystemImageList & usilExtraLargeImageList) {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, NULL)));
		}
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
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			delete iter->second;
		}
		properties.items.clear();

		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
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
		properties.shellColumnIndexesOfInsertedColumns.empty();
		properties.visibleColumns.clear();
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			delete properties.items.GetValueAt(p);
			properties.items.GetNextValue(p);
		}
		properties.items.RemoveAll();

		p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair1 = properties.namespaces.GetAt(p);
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
		properties.shellColumnIndexesOfInsertedColumns.RemoveAll();
		properties.visibleColumns.RemoveAll();
	#endif

	if(directoryWatchState.hThread) {
		InterlockedExchange(&directoryWatchState.terminate, TRUE);
		if(InterlockedCompareExchange(&directoryWatchState.isSuspended, TRUE, TRUE)) {
			ResumeThread(directoryWatchState.hThread);
		} else {
			// make sure it isn't blocked
			#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
				APIWrapper::CancelSynchronousIo(directoryWatchState.hThread, NULL);
			#else
				HANDLE hChangeEvent = NULL;
				InterlockedExchangePointer(&hChangeEvent, directoryWatchState.hChangeEvent);
				if(hChangeEvent != INVALID_HANDLE_VALUE) {
					SetEvent(hChangeEvent);
				}
			#endif
		}
		WaitForSingleObject(directoryWatchState.hThread, INFINITE);
		CloseHandle(directoryWatchState.hThread);
		directoryWatchState.hThread = NULL;
	}
	InterlockedExchange(&directoryWatchState.terminate, FALSE);

	DestroyShellContextMenu();
	contextMenuData.ResetToDefaults();
	dragDropStatus.ResetToDefaults();
	properties.columnsStatus.ResetToDefaults();
	properties.thumbnailsStatus.ResetToDefaults();
	properties.pUnifiedImageList = NULL;
	flags.toggleSortOrder = FALSE;
	cachedSettings.viewMode = -1;
	if(pThreadObject) {
		pThreadObject->Release();
		if(threadReferenceCounter != -1) {
			// we were the ones that have set the object
			APIWrapper::SHSetThreadRef(NULL, NULL);
		}
		pThreadObject = NULL;
	}
	flags.detaching = FALSE;
	ATLASSUME(properties.pListItems);
	properties.pListItems->SetOwner(NULL);
	ATLASSUME(properties.pNamespaces);
	properties.pNamespaces->SetOwner(NULL);
	ATLASSUME(properties.pColumns);
	properties.pColumns->SetOwner(NULL);
	Release();
	return S_OK;
}

STDMETHODIMP ShellListView::DisplayHeaderContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	POINT position = {x, y};
	attachedControl.ClientToScreen(&position);
	return DisplayHeaderContextMenu(position);
}

STDMETHODIMP ShellListView::DisplayShellContextMenu(VARIANT items, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	POINT position = {x, y};
	attachedControl.ClientToScreen(&position);

	HRESULT hr = S_OK;
	if(items.vt == VT_DISPATCH && items.pdispVal) {
		CComQIPtr<IShListViewNamespace> pShLvwNamespace = items.pdispVal;
		if(pShLvwNamespace) {
			// a single ShellListViewNamespace object
			LONG pIDLNamespace = 0;
			hr = pShLvwNamespace->get_FullyQualifiedPIDL(&pIDLNamespace);
			if(SUCCEEDED(hr)) {
				hr = DisplayShellContextMenu(*reinterpret_cast<PCIDLIST_ABSOLUTE*>(&pIDLNamespace), position);
			}
			return hr;
		}
	}

	UINT itemCount = 0;
	PLONG pItems = NULL;
	hr = VariantToItemIDs(items, pItems, itemCount);
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

STDMETHODIMP ShellListView::EnsureShellColumnsAreLoaded(VARIANT_BOOL waitForLastColumn/* = VARIANT_TRUE*/)
{
	return EnsureShellColumnsAreLoaded(waitForLastColumn == VARIANT_FALSE ? tmTimeOutThreading : tmNoThreading);
}

STDMETHODIMP ShellListView::FlushManagedIcons(VARIANT_BOOL includeOverlays/* = VARIANT_TRUE*/)
{
	FlushIcons(-1, VARIANTBOOL2BOOL(includeOverlays), FALSE);
	return S_OK;
}

STDMETHODIMP ShellListView::GetHeaderContextMenuItemString(LONG commandID, VARIANT* pItemCaption/* = NULL*/, VARIANT_BOOL* pCommandIDWasValid/* = NULL*/)
{
	ATLASSERT_POINTER(pCommandIDWasValid, VARIANT_BOOL);
	if(!pCommandIDWasValid) {
		return E_POINTER;
	}
	ATLASSERT(contextMenuData.currentHeaderContextMenu.IsMenu());
	if(!contextMenuData.currentHeaderContextMenu.IsMenu()) {
		return E_FAIL;
	}

	if(pItemCaption) {
		VariantInit(pItemCaption);
	}

	*pCommandIDWasValid = VARIANT_FALSE;
	if(commandID < 0 || static_cast<UINT>(commandID) > properties.columnsStatus.numberOfAllColumns) {
		return S_OK;
	}
	//commandID -= 0;

	if(pItemCaption) {
		TCHAR pBuffer[512];
		ZeroMemory(pBuffer, 512 * sizeof(TCHAR));
		CMenuItemInfo menuItemInfo;
		menuItemInfo.fMask = MIIM_ID | MIIM_STRING;
		menuItemInfo.wID = commandID;
		menuItemInfo.dwTypeData = pBuffer;
		menuItemInfo.cch = 512;
		if(contextMenuData.currentHeaderContextMenu.GetMenuItemInfo(commandID, FALSE, &menuItemInfo)) {
			*pCommandIDWasValid = VARIANT_TRUE;
			pItemCaption->vt = VT_BSTR;
			CT2OLE converter(pBuffer);
			LPOLESTR convertedString = converter;
			pItemCaption->bstrVal = SysAllocString(convertedString);
		}
	}
	return S_OK;
}

STDMETHODIMP ShellListView::GetShellContextMenuItemString(LONG commandID, VARIANT* pItemDescription/* = NULL*/, VARIANT* pItemVerb/* = NULL*/, VARIANT_BOOL* pCommandIDWasValid/* = NULL*/)
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

STDMETHODIMP ShellListView::InvokeDefaultShellContextMenuCommand(VARIANT items)
{
	HRESULT hr = S_OK;

	UINT itemCount = 0;
	PLONG pItems = NULL;
	hr = VariantToItemIDs(items, pItems, itemCount);
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

STDMETHODIMP ShellListView::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ShellListView::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ShellListView::SortItems(LONG shellColumnIndex/* = -1*/)
{
	// used to try to save a call to RealColumnIndexToIndex()
	int columnIndexToTransmit = -1;
	BOOL callRealColumnIndexToIndex = TRUE;

	int sortOrderToUse = properties.initialSortOrder;
	if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist && (properties.columnsStatus.previousSortColumn.fmtid != GUID_NULL || properties.columnsStatus.previousSortColumn.pid != 0)) {
		shellColumnIndex = PropertyKeyToRealIndex(properties.columnsStatus.previousSortColumn);
		ZeroMemory(&properties.columnsStatus.previousSortColumn, sizeof(properties.columnsStatus.previousSortColumn));
		sortOrderToUse = properties.columnsStatus.previousSortOrder;
		properties.columnsStatus.previousSortOrder = 0;
	}
	if(shellColumnIndex == -1) {
		shellColumnIndex = properties.columnsStatus.defaultSortColumn;
		sortOrderToUse = properties.initialSortOrder;
	}

	int buffer = sortingSettings.columnIndex;
	if(shellColumnIndex != sortingSettings.columnIndex) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_ChangingSortColumn(sortingSettings.columnIndex, shellColumnIndex, &cancel);
		if(cancel != VARIANT_FALSE) {
			return S_OK;
		}
	}

	if(properties.selectSortColumn || properties.setSortArrows) {
		// set sort arrow
		int columnIndex = RealColumnIndexToIndex(sortingSettings.columnIndex);
		if(columnIndex != -1) {
			if(shellColumnIndex == sortingSettings.columnIndex) {
				if(properties.setSortArrows) {
					int sortOrder = sortOrderToUse;
					if(sortingSettings.firstSortItemsCall) {
						// reset to ascending order
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTORDER, 0, sortOrder)));
						// the change might have been canceled
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
					} else {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
						if(flags.toggleSortOrder) {
							// toggle sort order
							sortOrder = (sortOrder == 0/*soAscending*/ ? 1/*soDescending*/ : 0/*soAscending*/);
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTORDER, 0, sortOrder)));
							// the change might have been canceled
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
						}
					}

					if(RunTimeHelper::IsCommCtrl6()) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, sortOrder == 0/*soAscending*/ ? 2/*saUp*/ : 1/*saDown*/)));
					} else {
						sortingSettings.FreeBitmapHandle();
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, sortOrder == 0/*soAscending*/ ? -2 : -1)));
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNBITMAP, columnIndex, reinterpret_cast<LPARAM>(&sortingSettings.headerBitmapHandle))));
					}
				}
				columnIndexToTransmit = columnIndex;
				callRealColumnIndexToIndex = FALSE;
			} else if(properties.setSortArrows) {
				// remove sort arrow from old sort column
				if(RunTimeHelper::IsCommCtrl6()) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, 0/*saNone*/)));
				} else {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, NULL)));
					sortingSettings.FreeBitmapHandle();
				}
			}
		}

		if(shellColumnIndex != sortingSettings.columnIndex) {
			// set new sort column
			columnIndex = RealColumnIndexToIndex(shellColumnIndex);
			if(columnIndex != -1) {
				int sortOrder = sortOrderToUse;
				if(properties.setSortArrows) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTORDER, 0, sortOrder)));
					// the change might have been canceled
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
				}
				if(RunTimeHelper::IsCommCtrl6()) {
					if(properties.setSortArrows) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, sortOrder == 0/*soAscending*/ ? 2/*saUp*/ : 1/*saDown*/)));
					}
					if(properties.selectSortColumn) {
						// select column
						attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, columnIndex, 0);
					}
				} else {
					if(properties.setSortArrows) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, sortOrder == 0/*soAscending*/ ? -2 : -1)));
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNBITMAP, columnIndex, reinterpret_cast<LPARAM>(&sortingSettings.headerBitmapHandle))));
					}
				}
			}
			columnIndexToTransmit = columnIndex;
			callRealColumnIndexToIndex = FALSE;
		}
	}

	sortingSettings.firstSortItemsCall = FALSE;
	sortingSettings.columnIndex = shellColumnIndex;

	int column = -1;
	if(shellColumnIndex > 0) {
		// the column will become relevant if the 2nd or 3rd criterion is used
		if(callRealColumnIndexToIndex) {
			columnIndexToTransmit = RealColumnIndexToIndex(shellColumnIndex);
		}
		column = columnIndexToTransmit;
	}
	HRESULT hr = E_FAIL;
	if(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SORTITEMS, column, 0))) {
		hr = S_OK;
	}
	if(shellColumnIndex != buffer) {
		Raise_ChangedSortColumn(buffer, sortingSettings.columnIndex);
	}
	return hr;
}


//////////////////////////////////////////////////////////////////////
// implementation of IMessageListener
HRESULT ShellListView::PostMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, DWORD cookie, BOOL eatenMessage)
{
	if(!pResult) {
		return E_POINTER;
	}

	BOOL wasHandled = TRUE;
	if(hWnd == attachedControl) {
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
			case LVM_GETITEM:
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

HRESULT ShellListView::PreMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, LPDWORD pCookie)
{
	if(!pResult || !pCookie) {
		return E_POINTER;
	}

	BOOL eatMessage = FALSE;
	if(hWnd == attachedControl) {
		BOOL wasHandled = TRUE;
		LRESULT lr = *pResult;
		__try {
			switch(message) {
				case LVM_SETITEM:
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
				case WM_TRIGGER_COLUMNENUMCOMPLETE:
					eatMessage = TRUE;
					lr = OnTriggerColumnEnumComplete(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_SETELEVATIONSHIELD:
					eatMessage = TRUE;
					lr = OnTriggerSetElevationShield(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_SETINFOTIP:
					eatMessage = TRUE;
					lr = OnTriggerSetInfoTip(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATEICON:
					eatMessage = TRUE;
					lr = OnTriggerUpdateIcon(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATEOVERLAY:
					eatMessage = TRUE;
					lr = OnTriggerUpdateOverlay(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATETEXT:
					eatMessage = TRUE;
					lr = OnTriggerUpdateText(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_ISLVWITEMVISIBLE:
					eatMessage = TRUE;
					if(properties.pAttachedInternalMessageListener) {
						lr = static_cast<LRESULT>(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ISITEMVISIBLE, wParam, lParam));
					} else {
						lr = static_cast<LRESULT>(E_FAIL);
					}
					break;
				case WM_TRIGGER_UPDATETHUMBNAIL:
					eatMessage = TRUE;
					lr = OnTriggerUpdateThumbnail(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATETILEVIEWCOLUMNS:
					eatMessage = TRUE;
					lr = OnTriggerUpdateTileViewColumns(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_UPDATESUBITEMCONTROL:
					eatMessage = TRUE;
					lr = OnTriggerUpdateSubItemControl(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_TRIGGER_ENQUEUETASK:
					eatMessage = TRUE;
					lr = OnTriggerEnqueueTask(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
					break;
				case WM_REPORT_THUMBNAILDISKCACHEACCESS:
					eatMessage = TRUE;
					lr = OnReportThumbnailDiskCacheAccess(message, wParam, lParam, wasHandled, *pCookie, eatMessage, TRUE);
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
HRESULT ShellListView::ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch(message) {
		case SHLVM_ALLOCATEMEMORY:
			return OnInternalAllocateMemory(message, wParam, lParam);
			break;
		case SHLVM_FREEMEMORY:
			return OnInternalFreeMemory(message, wParam, lParam);
			break;
		case SHLVM_INSERTEDCOLUMN:
			return OnInternalInsertedColumn(message, wParam, lParam);
			break;
		case SHLVM_FREECOLUMN:
			return OnInternalFreeColumn(message, wParam, lParam);
			break;
		case SHLVM_GETSHLISTVIEWCOLUMN:
			return OnInternalGetShListViewColumn(message, wParam, lParam);
			break;
		case SHLVM_HEADERCLICK:
			return OnInternalHeaderClick(message, wParam, lParam);
			break;
		case SHLVM_HEADERDROPDOWNMENU:
			ATLASSERT(FALSE && "TODO: Handle SHLVM_HEADERDROPDOWNMENU");
			break;
		case SHLVM_INSERTEDITEM:
			return OnInternalInsertedItem(message, wParam, lParam);
			break;
		case SHLVM_FREEITEM:
			return OnInternalFreeItem(message, wParam, lParam);
			break;
		case SHLVM_GETSHLISTVIEWITEM:
			return OnInternalGetShListViewItem(message, wParam, lParam);
			break;
		case SHLVM_RENAMEDITEM:
			return OnInternalRenamedItem(message, wParam, lParam);
			break;
		case SHLVM_COMPAREITEMS:
			return OnInternalCompareItems(message, wParam, lParam);
			break;
		case SHLVM_BEGINLABELEDITA:
		case SHLVM_BEGINLABELEDITW:
			return OnInternalBeginLabelEdit(message, wParam, lParam);
			break;
		case SHLVM_BEGINDRAG:
			return OnInternalBeginDrag(message, wParam, lParam);
			break;
		case SHLVM_HANDLEDRAGEVENT:
			return OnInternalHandleDragEvent(message, wParam, lParam);
			break;
		case SHLVM_CONTEXTMENU:
			return OnInternalContextMenu(message, wParam, lParam);
			break;
		case SHLVM_GETINFOTIP:
			return OnInternalGetInfoTip(message, wParam, lParam);
			break;
		case SHLVM_CUSTOMDRAW:
			return OnInternalCustomDraw(message, wParam, lParam);
			break;
		case SHLVM_CHANGEDVIEW:
			return OnInternalViewChanged(message, wParam, lParam);
			break;
		case SHLVM_GETINFOTIPEX:
			return OnInternalGetInfoTipEx(message, wParam, lParam);
			break;
		case SHLVM_GETSUBITEMCONTROL:
			return OnInternalGetSubItemControl(message, wParam, lParam);
			break;
		/*case SHLVM_ENDSUBITEMEDIT:
			return OnInternalEndSubItemEdit(message, wParam, lParam);
			break;*/
	}
	return E_NOTIMPL;
}
// implementation of IInternalMessageListener
//////////////////////////////////////////////////////////////////////


LRESULT ShellListView::OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
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
		BOOL isShellContextMenu = contextMenuData.currentShellContextMenu.IsMenu();
		BOOL isSystemMenuItem = FALSE;
		if(isShellContextMenu) {
			isSystemMenuItem = ((commandID >= MIN_CONTEXTMENUEXTENSION_CMDID) && (commandID <= MAX_CONTEXTMENUEXTENSION_CMDID));
		} else {
			isSystemMenuItem = ((commandID >= 0) && (commandID <= properties.columnsStatus.numberOfAllColumns));
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
			Raise_ChangedContextMenuSelection(HandleToLong(reinterpret_cast<HMENU>(lParam)), BOOL2VARIANTBOOL(isShellContextMenu), commandID, BOOL2VARIANTBOOL(!isSystemMenuItem));
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

LRESULT ShellListView::OnGetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	LPLVITEM pItemData = reinterpret_cast<LPLVITEM>(lParam);
	if(!pItemData) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pItemData->mask & LVIF_IMAGE) {
		LONG itemID = ItemIndexToID(pItemData->iItem);
		if(!IsShellItem(itemID)) {
			// get real icon index
			if(properties.thumbnailsStatus.pThumbnailImageList) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_NONSHELLITEMICON, &pItemData->iImage, NULL, NULL, NULL)));
			} else if(properties.pUnifiedImageList) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_NONSHELLITEMICON, &pItemData->iImage, NULL, NULL, NULL)));
			}
		}
	}
	return TRUE;
}

LRESULT ShellListView::OnSetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	LPLVITEM pItemData = reinterpret_cast<LPLVITEM>(lParam);
	if(!pItemData) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pItemData->mask & LVIF_IMAGE) {
		LONG itemID = ItemIndexToID(pItemData->iItem);
		if(!IsShellItem(itemID)) {
			// update icon index
			BOOL setID = FALSE;
			if(properties.thumbnailsStatus.pThumbnailImageList) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_NONSHELLITEMICON, NULL, pItemData->iImage, 0, NULL, 0)));
				setID = TRUE;
			}
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
				attachedControl.PostMessage(LVM_REDRAWITEMS, pItemData->iItem, pItemData->iItem);
			}
		}
	}
	return TRUE;
}

LRESULT ShellListView::OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& cookie, BOOL& /*eatenMessage*/, BOOL preProcessing)
{
	LPNMHDR pDetails = reinterpret_cast<LPNMHDR>(lParam);
	switch(pDetails->code) {
		case LVN_GETDISPINFOA:
		case LVN_GETDISPINFOW:
			if(preProcessing) {
				return OnGetDispInfoNotification(pDetails->code, pDetails, wasHandled, cookie, preProcessing);
			}
			break;
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ShellListView::OnReportThumbnailDiskCacheAccess(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	/* NOTE: Actually this message should never be sent if the code below fails, but it's better to
	 *       double-check instead of ending up with a cache that never gets closed.
	 */
	//ReplyMessage(TRUE);

	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(static_cast<LONG>(wParam));
	if(pItemDetails && pItemDetails->pIDLNamespace) {
		CComObject<ShellListViewNamespace>* pNamespace = NamespacePIDLToNamespace(pItemDetails->pIDLNamespace, TRUE);
		if(pNamespace) {
			if(pNamespace->properties.pThumbnailDiskCache) {
				pNamespace->properties.lastThumbnailDiskCacheAccessTime = lParam;
				// the cache has been accessed, i. e. currently it is opened
				pNamespace->properties.thumbnailDiskCacheIsOpened = TRUE;
				if(!flags.thumbnailDiskCacheTimerIsActive) {
					flags.thumbnailDiskCacheTimerIsActive = TRUE;
					attachedControl.SetTimer(timers.ID_LAZYCLOSETHUMBNAILDISKCACHE, timers.INT_LAZYCLOSETHUMBNAILDISKCACHE);
				}
				return TRUE;
			}
		}
	}
	return FALSE;
}

LRESULT ShellListView::OnSHChangeNotify(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
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
		ATLTRACE2(shlvwTraceAutoUpdate, 2, TEXT("Received notification 0x%X\n"), eventID);

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
			case SHCNE_FREESPACE:
				ret = OnSHChangeNotify_FREESPACE(ppIDLs);
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

LRESULT ShellListView::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& eatenMessage, BOOL /*preProcessing*/)
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

			long timeout = -1;
			switch(pBackgroundEnum->type) {
				case BET_ITEMS:
					timeout = properties.itemEnumerationTimeout;
					break;
				case BET_COLUMNS:
					timeout = properties.columnEnumerationTimeout;
					break;
			}
			if(diff >= static_cast<DWORD>(timeout) && !pBackgroundEnum->timedOut) {
				if(timeout != -1) {
					pBackgroundEnum->timedOut = TRUE;
					if(pBackgroundEnum->raiseEvents) {
						CComQIPtr<IShListViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
						switch(pBackgroundEnum->type) {
							case BET_ITEMS:
								Raise_ItemEnumerationTimedOut(pNamespace);
								break;
							case BET_COLUMNS:
								Raise_ColumnEnumerationTimedOut(pNamespace);
								break;
						}
					}
				}
			}
		}
	} else if(wParam == timers.ID_LAZYCLOSETHUMBNAILDISKCACHE) {
		eatenMessage = TRUE;
		LazyCloseThumbnailDiskCaches();
	} else {
		wasHandled = FALSE;
	}
	return 0;
}

LRESULT ShellListView::OnTriggerEnqueueTask(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHCTLSBACKGROUNDTASKINFO pTaskInfos = NULL;
		EnterCriticalSection(&properties.tasksToEnqueueCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.tasksToEnqueue.empty()) {
				pTaskInfos = properties.tasksToEnqueue.front();
				properties.tasksToEnqueue.pop();
				empty = FALSE;
			}
		#else
			if(!properties.tasksToEnqueue.IsEmpty()) {
				pTaskInfos = properties.tasksToEnqueue.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.tasksToEnqueueCritSection);

		if(empty) {
			break;
		} else if(!pTaskInfos) {
			continue;
		}

		ATLASSUME(pTaskInfos->pTask);
		if(pTaskInfos->pTask) {
			ATLVERIFY(SUCCEEDED(EnqueueTask(pTaskInfos->pTask, pTaskInfos->taskID, 0, pTaskInfos->taskPriority)));
			pTaskInfos->pTask->Release();
		}
		delete pTaskInfos;
	}
	return 0;
}

LRESULT ShellListView::OnTriggerItemEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDITEMENUMINFO pEnumResult = NULL;
		EnterCriticalSection(&properties.backgroundItemEnumResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.backgroundItemEnumResults.empty()) {
				pEnumResult = properties.backgroundItemEnumResults.front();
				properties.backgroundItemEnumResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.backgroundItemEnumResults.IsEmpty()) {
				pEnumResult = properties.backgroundItemEnumResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.backgroundItemEnumResultsCritSection);

		if(empty) {
			break;
		} else if(!pEnumResult) {
			continue;
		}

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
			if(namespaceHasBeenRemoved) {
				ATLTRACE2(shlvwTraceThreading, 3, TEXT("Ignoring insertion of %i items due to removed namespace\n"), DPA_GetPtrCount(pEnumResult->hPIDLBuffer));
				DPA_DestroyCallback(pEnumResult->hPIDLBuffer, FreeDPAPIDLElement, NULL);
			} else {
				DroppedPIDLOffsets* pOffsets = NULL;
				#ifdef USE_STL
					for(std::vector<DroppedPIDLOffsets*>::const_iterator it = dragDropStatus.bufferedPIDLOffsets.cbegin(); it != dragDropStatus.bufferedPIDLOffsets.cend(); ++it) {
						if((*it)->pTargetNamespace == pEnumResult->namespacePIDLToSet) {
							pOffsets = *it;
						}
					}
				#else
					for(size_t i = 0; i < dragDropStatus.bufferedPIDLOffsets.GetCount(); ++i) {
						if(ILIsEqual(dragDropStatus.bufferedPIDLOffsets[i]->pTargetNamespace, pEnumResult->namespacePIDLToSet)) {
							pOffsets = dragDropStatus.bufferedPIDLOffsets[i];
						}
					}
				#endif
				#ifdef USE_STL
					//std::unordered_map<LONG, PCIDLIST_ABSOLUTE> itemCache;
					std::vector<std::pair<LONG, PCIDLIST_ABSOLUTE>> itemCache;
				#else
					CAtlMap<LONG, PCIDLIST_ABSOLUTE> itemCache;
				#endif

				int offsetIndex = -1;
				LONG itemToEdit = -1;
				int itemCount = 0;
				for(int i = 0; TRUE; ++i) {
					PIDLIST_ABSOLUTE pIDLAbsolute = reinterpret_cast<PIDLIST_ABSOLUTE>(DPA_GetPtr(pEnumResult->hPIDLBuffer, i));
					if(!pIDLAbsolute) {
						break;
					}

					if(pEnumResult->checkForDuplicates) {
						if(i == 0) {
							#ifdef USE_STL
								//itemCache = properties.items;
								for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter2 = properties.items.cbegin(); iter2 != properties.items.cend(); ++iter2) {
									itemCache.push_back(std::pair<LONG, PCIDLIST_ABSOLUTE>(iter2->first, iter2->second->pIDL));
								}
							#else
								POSITION p = properties.items.GetStartPosition();
								while(p) {
									CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair2 = properties.items.GetAt(p);
									itemCache[pPair2->m_key] = pPair2->m_value->pIDL;
									properties.items.GetNext(p);
								}
							#endif
						}

						LONG existingItemID = -1;  //FindItem(pIDLAbsolute, FALSE);
						#ifdef USE_STL
							/*for(std::unordered_map<LONG, PCIDLIST_ABSOLUTE>::const_iterator iter = itemCache.cbegin(); iter != itemCache.cend(); ++iter) {
								if(ILIsEqual(iter->second, pIDLAbsolute)) {
									existingItemID = iter->first;
									itemCache.erase(iter);
									break;
								}
							}*/
							std::vector<std::pair<LONG, PCIDLIST_ABSOLUTE>>::const_iterator foundEntry;
							BOOL done = FALSE;
							#pragma omp parallel
							{
								int this_thread = omp_get_thread_num();
								int num_threads = omp_get_num_threads();
								int size = itemCache.size();
								int begin = (this_thread + 0) * size / num_threads;
								int end = (this_thread + 1) * size / num_threads;

								for(int j = begin; j < end; ++j) {
									#pragma omp flush(done)
									if(done) {
										break;
									}
									if(ILIsEqual(itemCache[j].second, pIDLAbsolute)) {
										done = TRUE;
										#pragma omp flush(done)
										existingItemID = itemCache[j].first;
										foundEntry = itemCache.cbegin();
										std::advance(foundEntry, j);
									}
								}
							}
							if(existingItemID != -1) {
								itemCache.erase(foundEntry);
							}
						#else
							POSITION p = itemCache.GetStartPosition();
							while(p) {
								CAtlMap<LONG, PCIDLIST_ABSOLUTE>::CPair* pPair2 = itemCache.GetAt(p);
								if(ILIsEqual(pPair2->m_value, pIDLAbsolute)) {
									existingItemID = pPair2->m_key;
									itemCache.RemoveAtPos(p);
									break;
								}
								itemCache.GetNext(p);
							}
						#endif
						if(existingItemID != -1) {
							// TODO: Should we do something about pOffsets here?
							ATLTRACE2(shlvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X because it's a duplicate of item %i\n"), pIDLAbsolute, existingItemID);
							RemoveBufferedShellItemNamespace(pIDLAbsolute);
							ILFree(pIDLAbsolute);
							//DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
							continue;
						}
					}
					if(!(properties.disabledEvents & deItemInsertionEvents)) {
						VARIANT_BOOL cancel = VARIANT_FALSE;
						Raise_InsertingItem(*reinterpret_cast<LONG*>(&pIDLAbsolute), &cancel);
						if(cancel != VARIANT_FALSE) {
							// TODO: Should we do something about pOffsets here?
							ATLTRACE2(shlvwTraceThreading, 3, TEXT("Skipping insertion of pIDL 0x%X as requested by client app\n"), pIDLAbsolute);
							RemoveBufferedShellItemNamespace(pIDLAbsolute);
							ILFree(pIDLAbsolute);
							//DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
							continue;
						}
					}

					if(pEnumResult->namespacePIDLToSet) {
						BufferShellItemNamespace(pIDLAbsolute, pEnumResult->namespacePIDLToSet);
					}
					LONG insertedItemID = FastInsertShellItem(pIDLAbsolute, pEnumResult->insertAt);
					if(pEnumResult->insertAt != -1) {
						++pEnumResult->insertAt;
					}
					ATLASSERT(insertedItemID != -1);
					if(insertedItemID == -1) {
						// TODO: Should we do anything about pOffsets here?
						RemoveBufferedShellItemNamespace(pIDLAbsolute);
						ILFree(pIDLAbsolute);
						//DPA_SetPtr(pEnumResult->hPIDLBuffer, i, NULL);
					} else {
						#ifdef DEBUG
							LONG itemID = PIDLToItemID(pIDLAbsolute, TRUE);
							ATLASSERT(itemID == insertedItemID);
							++itemCount;
						#endif
						if(pIDLAbsolute == pEnumResult->pIDLToLabelEdit) {
							itemToEdit = insertedItemID;
						} else if(pOffsets) {
							#define HIDA_GetPIDLItem(p, i) reinterpret_cast<PCIDLIST_RELATIVE>(reinterpret_cast<LPBYTE>(p) + p->aoffset[i + 1])
							switch(pOffsets->performedEffect) {
								case DROPEFFECT_COPY:
								case DROPEFFECT_MOVE:
									/* TODO: This will probably fail if the drop results in the dropped item being renamed,
													 e. g. because an item with the same name already exists in the target folder. */
									for(UINT i2 = 0; i2 < pOffsets->pDroppedPIDLs->cidl; ++i2) {
										if(ILIsEqual(reinterpret_cast<PCIDLIST_ABSOLUTE>(HIDA_GetPIDLItem(pOffsets->pDroppedPIDLs, i2)), reinterpret_cast<PCIDLIST_ABSOLUTE>(ILFindLastID(pIDLAbsolute)))) {
											offsetIndex = static_cast<int>(i2);
											break;
										}
									}
									break;
								case DROPEFFECT_LINK:
									// check the target of the link
									PIDLIST_ABSOLUTE pIDLTarget = NULL;
									if(SUCCEEDED(::GetLinkTarget(NULL, pIDLAbsolute, &pIDLTarget))) {
										for(UINT i2 = 0; i2 < pOffsets->pDroppedPIDLs->cidl; ++i2) {
											if(ILIsEqual(reinterpret_cast<PCIDLIST_ABSOLUTE>(HIDA_GetPIDLItem(pOffsets->pDroppedPIDLs, i2)), reinterpret_cast<PCIDLIST_ABSOLUTE>(ILFindLastID(pIDLTarget)))) {
												offsetIndex = static_cast<int>(i2);
												break;
											}
										}
										if(pIDLTarget) {
											ILFree(pIDLTarget);
										}
									}
									break;
							}

							if(offsetIndex != -1) {
								int itemIndex = ItemIDToIndex(insertedItemID);
								if(itemIndex >= 0) {
									if(cachedSettings.viewMode == -1) {
										ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
									}

									POINT itemPosition = {0};
									switch(cachedSettings.viewMode) {
										case 0/*vIcons*/:
											itemPosition.x = pOffsets->pOffsets[offsetIndex + 1].x + pOffsets->cursorPosition.x;
											itemPosition.y = pOffsets->pOffsets[offsetIndex + 1].y + pOffsets->cursorPosition.y;
											break;
										case 4/*vTiles*/:
										case 5/*vExtendedTiles*/:
											// TODO: Probably wrong for vExtendedTiles.
											// TODO: Don't use hardcoded values!
											itemPosition.x = 24 * pOffsets->pOffsets[offsetIndex + 1].x / 10 + pOffsets->cursorPosition.x;
											itemPosition.y = pOffsets->pOffsets[offsetIndex + 1].y + pOffsets->cursorPosition.y;
											break;
										case 1/*vSmallIcons*/:
											DWORD iconSpacing = static_cast<DWORD>(attachedControl.SendMessage(LVM_GETITEMSPACING, FALSE, 0));
											DWORD smallIconSpacing = static_cast<DWORD>(attachedControl.SendMessage(LVM_GETITEMSPACING, TRUE, 0));
											itemPosition.x = (pOffsets->pOffsets[offsetIndex + 1].x * LOWORD(smallIconSpacing) / LOWORD(iconSpacing)) + pOffsets->cursorPosition.x;
											itemPosition.y = (pOffsets->pOffsets[offsetIndex + 1].y * HIWORD(smallIconSpacing) / HIWORD(iconSpacing)) + pOffsets->cursorPosition.y;
											break;
									}
									if(cachedSettings.viewMode != 3/*vDetails*/ && cachedSettings.viewMode != 2/*vList*/) {
										attachedControl.SendMessage(LVM_SETITEMPOSITION32, itemIndex, reinterpret_cast<LPARAM>(&itemPosition));
									}
								}
							}
						}
					}
				}
				itemCount;
				ATLTRACE2(shlvwTraceThreading, 3, TEXT("Added %i out of %i items\n"), itemCount, DPA_GetPtrCount(pEnumResult->hPIDLBuffer));
				#ifdef USE_STL
					itemCache.clear();
				#else
					itemCache.RemoveAll();
				#endif

				if(!attachedControl.SendMessage(LVM_GETEDITCONTROL, 0, 0)) {
					CComQIPtr<IShListViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
					if(pNamespace) {
						AutoSortItemsConstants autoSort = asiNoAutoSort;
						pNamespace->get_AutoSortItems(&autoSort);
						switch(autoSort) {
							case asiAutoSortOnce:
								if(!pEnumResult->forceAutoSort) {
									break;
								}
							case asiAutoSort:
								SortItems(sortingSettings.columnIndex);
								break;
						}
					}
				}

				DPA_DeleteAllPtrs(pEnumResult->hPIDLBuffer);
				DPA_Destroy(pEnumResult->hPIDLBuffer);

				if(pBackgroundEnum->pTargetObject) {
					if(pBackgroundEnum->raiseEvents) {
						CComQIPtr<IShListViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
						if(pNamespace) {
							Raise_ItemEnumerationCompleted(pNamespace, VARIANT_FALSE);
						}
					}
					pBackgroundEnum->pTargetObject->Release();
				}
				delete pBackgroundEnum;
				pBackgroundEnum = NULL;

				if(itemToEdit != -1) {
					int itemIndex = ItemIDToIndex(itemToEdit);
					if(itemIndex >= 0) {
						if(properties.contextMenuPosition.x != -1 && properties.contextMenuPosition.y != -1) {
							attachedControl.SendMessage(LVM_SETITEMPOSITION32, itemIndex, reinterpret_cast<LPARAM>(&properties.contextMenuPosition));
							properties.contextMenuPosition.x = -1;
							properties.contextMenuPosition.y = -1;
						}
						if(properties.autoEditNewItems) {
							attachedControl.SendMessage(LVM_EDITLABEL, itemIndex, 0);
						}
					}
				}
			}
		}

		if(pBackgroundEnum) {
			if(pBackgroundEnum->pTargetObject) {
				pBackgroundEnum->pTargetObject->Release();
			}
			delete pBackgroundEnum;
		}
		if(pEnumResult) {
			delete pEnumResult;
		}
	}
	return 0;
}

LRESULT ShellListView::OnTriggerColumnEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDCOLUMNENUMINFO pEnumResult = NULL;
		EnterCriticalSection(&properties.backgroundColumnEnumResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.backgroundColumnEnumResults.empty()) {
				pEnumResult = properties.backgroundColumnEnumResults.front();
				properties.backgroundColumnEnumResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.backgroundColumnEnumResults.IsEmpty()) {
				pEnumResult = properties.backgroundColumnEnumResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.backgroundColumnEnumResultsCritSection);

		if(empty) {
			break;
		} else if(!pEnumResult) {
			continue;
		}

		LPBACKGROUNDENUMERATION pBackgroundEnum = NULL;
		BOOL namespaceHasBeenRemoved = FALSE;
		ATLASSERT(pEnumResult->taskID);
		#ifdef USE_STL
			std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION>::const_iterator iter = properties.backgroundEnums.find(pEnumResult->taskID);
			namespaceHasBeenRemoved = (iter == properties.backgroundEnums.cend());
			if(!namespaceHasBeenRemoved) {
				pBackgroundEnum = iter->second;
			}
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				if(!pEnumResult->moreColumnsToCome)
			#endif
			properties.backgroundEnums.erase(pEnumResult->taskID);
		#else
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION>::CPair* pPair = properties.backgroundEnums.Lookup(pEnumResult->taskID);
			namespaceHasBeenRemoved = !pPair;
			if(!namespaceHasBeenRemoved) {
				pBackgroundEnum = pPair->m_value;
			}
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				if(!pEnumResult->moreColumnsToCome)
			#endif
			properties.backgroundEnums.RemoveKey(pEnumResult->taskID);
		#endif

		if(pEnumResult->hColumnBuffer) {
			// insert the columns
			int columnCount = DPA_GetPtrCount(pEnumResult->hColumnBuffer);
			int realColumnIndex = pEnumResult->baseRealColumnIndex;

			if(namespaceHasBeenRemoved) {
				ATLTRACE2(shlvwTraceThreading, 3, TEXT("Ignoring insertion of %i columns due to removed namespace\n"), columnCount);
				DPA_DestroyCallback(pEnumResult->hColumnBuffer, FreeDPAColumnElement, NULL);
			} else {
				if(!properties.columnsStatus.pAllColumns) {
					properties.columnsStatus.pAllColumns = static_cast<LPSHELLLISTVIEWCOLUMNDATA>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, columnCount * sizeof(SHELLLISTVIEWCOLUMNDATA)));
					ATLASSERT(properties.columnsStatus.pAllColumns);
					properties.columnsStatus.currentArraySize = columnCount * sizeof(SHELLLISTVIEWCOLUMNDATA);
				}
				if((properties.columnsStatus.numberOfAllColumns + columnCount) * sizeof(SHELLLISTVIEWCOLUMNDATA) > properties.columnsStatus.currentArraySize) {
					// too small, so realloc
					LPSHELLLISTVIEWCOLUMNDATA p = static_cast<LPSHELLLISTVIEWCOLUMNDATA>(HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, properties.columnsStatus.pAllColumns, (properties.columnsStatus.numberOfAllColumns + columnCount) * sizeof(SHELLLISTVIEWCOLUMNDATA)));
					if(!p) {
						ATLASSERT(p);
						columnCount = 0;
					} else {
						properties.columnsStatus.currentArraySize = (properties.columnsStatus.numberOfAllColumns + columnCount) * sizeof(SHELLLISTVIEWCOLUMNDATA);
						properties.columnsStatus.pAllColumns = p;
					}
				}
				ATLASSERT(properties.columnsStatus.pAllColumns);

				for(int i = 0; i < columnCount; ++i) {
					LPSHELLLISTVIEWCOLUMNDATA pSourceColumnData = reinterpret_cast<LPSHELLLISTVIEWCOLUMNDATA>(DPA_GetPtr(pEnumResult->hColumnBuffer, i));
					if(!pSourceColumnData) {
						break;
					}
					int visibleInPreviousNamespace = -1;
					if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist && properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
						// copy data from the previous namespace
						LPSHELLLISTVIEWCOLUMNDATA pBackupColumnData = NULL;
						PROPERTYKEY sourcePropertyKey;
						if(properties.columnsStatus.pNamespaceSHF2 && SUCCEEDED(properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, &sourcePropertyKey))) {
							for(UINT j = 0; j < properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace; ++j) {
								PROPERTYKEY backupPropertyKey = properties.columnsStatus.pAllColumnsOfPreviousNamespace[j].propertyKey;
								if(sourcePropertyKey.fmtid == backupPropertyKey.fmtid && sourcePropertyKey.pid == backupPropertyKey.pid) {
									pBackupColumnData = &properties.columnsStatus.pAllColumnsOfPreviousNamespace[j];
									break;
								}
							}
						}
						if(pBackupColumnData) {
							pSourceColumnData->width = pBackupColumnData->width;
							lstrcpyn(pSourceColumnData->pText, pBackupColumnData->pText, MAX_LVCOLUMNTEXTLENGTH);
							pSourceColumnData->format &= ~LVCFMT_JUSTIFYMASK;
							pSourceColumnData->format |= (pBackupColumnData->format & LVCFMT_JUSTIFYMASK);
							pSourceColumnData->visibilityHasBeenChangedExplicitly = pBackupColumnData->visibilityHasBeenChangedExplicitly;
							if(pBackupColumnData->visibilityHasBeenChangedExplicitly) {
								visibleInPreviousNamespace = (pBackupColumnData->visibleInDetailsView ? 1 : 0);
							}
						}
					}

					LPSHELLLISTVIEWCOLUMNDATA pTargetColumnData = &properties.columnsStatus.pAllColumns[realColumnIndex];
					pTargetColumnData->columnID = pSourceColumnData->columnID;
					#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
						pTargetColumnData->pPropertyDescription = NULL;
						pTargetColumnData->displayType = pSourceColumnData->displayType;
						pTargetColumnData->typeFlags = pSourceColumnData->typeFlags;
					#endif
					pTargetColumnData->visibleInDetailsView = pSourceColumnData->visibleInDetailsView;
					pTargetColumnData->visibilityHasBeenChangedExplicitly = pSourceColumnData->visibilityHasBeenChangedExplicitly;
					pTargetColumnData->lastIndex_DetailsView = pSourceColumnData->lastIndex_DetailsView;
					pTargetColumnData->lastIndex_TilesView = pSourceColumnData->lastIndex_TilesView;
					lstrcpyn(pTargetColumnData->pText, pSourceColumnData->pText, MAX_LVCOLUMNTEXTLENGTH);
					pTargetColumnData->width = pSourceColumnData->width;
					pTargetColumnData->format = pSourceColumnData->format;
					pTargetColumnData->state = pSourceColumnData->state;
					++properties.columnsStatus.numberOfAllColumns;

					BOOL makeVisible = FALSE;
					if((pTargetColumnData->state & SHCOLSTATE_HIDDEN) == 0) {
						makeVisible = ((pTargetColumnData->state & SHCOLSTATE_ONBYDEFAULT) == SHCOLSTATE_ONBYDEFAULT);
						if(visibleInPreviousNamespace != -1) {
							makeVisible = (visibleInPreviousNamespace != 0);
						}
					}
					// always show the first column
					makeVisible = makeVisible || (realColumnIndex == 0);
					if(!(properties.disabledEvents & deColumnLoadingEvents)) {
						VARIANT_BOOL visible = BOOL2VARIANTBOOL(makeVisible);
						CComPtr<IShListViewColumn> pColumn;
						ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumn));
						Raise_LoadedColumn(pColumn, &visible);
						makeVisible = VARIANTBOOL2BOOL(visible);
					}

					if(makeVisible) {
						ATLVERIFY(ChangeColumnVisibility(realColumnIndex, TRUE, 0) != -1);
					}
					// ChangeColumnVisibility will set 'visibleInDetailsView' only if we're in Details view mode
					pTargetColumnData->visibleInDetailsView = makeVisible;
					++realColumnIndex;
					delete pSourceColumnData;
					//DPA_SetPtr(pEnumResult->hColumnBuffer, i, NULL);
				}

				DPA_DeleteAllPtrs(pEnumResult->hColumnBuffer);
				DPA_Destroy(pEnumResult->hColumnBuffer);
				#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
					if(!pEnumResult->moreColumnsToCome) {
				#endif
				if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist) {
					if(properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace) {
						// copy column order from the previous namespace
						HWND hWndHeader = reinterpret_cast<HWND>(attachedControl.SendMessage(LVM_GETHEADER, 0, 0));
						ATLASSERT(::IsWindow(hWndHeader));
						if(::IsWindow(hWndHeader)) {
							UINT columns = static_cast<UINT>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
							if(columns > 0) {
								PINT pColumns = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, columns * sizeof(int)));
								if(pColumns) {
									if(attachedControl.SendMessage(LVM_GETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pColumns))) {
										for(UINT i = 0; i < columns; ++i) {
											int realColumnIndex2 = ColumnIndexToRealIndex(pColumns[i]);
											if(realColumnIndex2 >= 0) {
												PROPERTYKEY propertyKey = {0};
												GetColumnPropertyKey(realColumnIndex2, &propertyKey);
												BOOL found = FALSE;
												UINT j = 0;
												for(j = 0; j < properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace; ++j) {
													if(propertyKey == properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace[j]) {
														found = TRUE;
														break;
													}
												}

												if(found && i != j && j < columns) {
													int tmp = pColumns[i];
													pColumns[i] = pColumns[j];
													pColumns[j] = tmp;
													i--;
												}
											}
										}
										if(attachedControl.SendMessage(LVM_SETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pColumns))) {
											attachedControl.InvalidateRect(NULL, TRUE);
										}
									}
									HeapFree(GetProcessHeap(), 0, pColumns);
								}
							}
						}
					}
				}
				properties.columnsStatus.loaded = SCLS_LOADED;
				SetEvent(properties.columnsStatus.hColumnsReadyEvent);

				if(pBackgroundEnum->pTargetObject) {
					if(pBackgroundEnum->raiseEvents) {
						CComQIPtr<IShListViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
						if(pNamespace) {
							Raise_ColumnEnumerationCompleted(pNamespace, VARIANT_FALSE);
						}
					}
					pBackgroundEnum->pTargetObject->Release();
				}
				delete pBackgroundEnum;
				#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
					}
				#endif
				pBackgroundEnum = NULL;
			}
		} else {
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				if(!pEnumResult->moreColumnsToCome) {
			#endif
			// we received an invalid column HDPA - handle this as if the columns are loaded now
			properties.columnsStatus.loaded = SCLS_LOADED;
			SetEvent(properties.columnsStatus.hColumnsReadyEvent);
			if(pBackgroundEnum) {
				if(pBackgroundEnum->pTargetObject) {
					pBackgroundEnum->pTargetObject->Release();
				}
				delete pBackgroundEnum;
			}
			#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
				}
			#endif
		}
		if(pEnumResult) {
			delete pEnumResult;
		}
	}

	if(properties.columnsStatus.loaded == SCLS_LOADED) {
		if(properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
			// the data of the previous namespace isn't needed anymore
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				UINT oldNumberOfAllColumnsOfPreviousNamespace = properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace;
			#endif
			properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace = 0;
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				for(UINT i = 0; i < oldNumberOfAllColumnsOfPreviousNamespace; ++i) {
					if(properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription) {
						properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription->Release();
						properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription = NULL;
					}
				}
			#endif
			HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pAllColumnsOfPreviousNamespace);
			properties.columnsStatus.pAllColumnsOfPreviousNamespace = NULL;
		}
		if(properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace) {
			// the data of the previous namespace isn't needed anymore
			properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace = 0;
			HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace);
			properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace = NULL;
		}
	}
	return 0;
}

LRESULT ShellListView::OnTriggerSetElevationShield(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	LONG itemID = static_cast<LONG>(wParam);
	BOOL redraw = FALSE;
	if(properties.displayThumbnails && properties.thumbnailsStatus.pThumbnailImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);

		UINT iconFlags = 0;
		ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &iconFlags)));
		if(lParam) {
			iconFlags |= AII_DISPLAYELEVATIONOVERLAY;
		} else {
			iconFlags &= ~AII_DISPLAYELEVATIONOVERLAY;
		}
		ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, iconFlags)));
		redraw = TRUE;
	}
	if(properties.pUnifiedImageList) {
		CComPtr<IImageListPrivate> pImgLstPriv;
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
		ATLASSUME(pImgLstPriv);

		UINT iconFlags = 0;
		ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &iconFlags)));
		if(lParam) {
			iconFlags |= AII_DISPLAYELEVATIONOVERLAY;
		} else {
			iconFlags &= ~AII_DISPLAYELEVATIONOVERLAY;
		}
		ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, iconFlags)));
		redraw = TRUE;
	}
	if(redraw) {
		int itemIndex = ItemIDToIndex(itemID);
		attachedControl.SendMessage(LVM_REDRAWITEMS, itemIndex, itemIndex);
	}
	return 0;
}

LRESULT ShellListView::OnTriggerSetInfoTip(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDINFOTIPINFO pResult = NULL;
		EnterCriticalSection(&properties.infoTipResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.infoTipResults.empty()) {
				pResult = properties.infoTipResults.front();
				properties.infoTipResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.infoTipResults.IsEmpty()) {
				pResult = properties.infoTipResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.infoTipResultsCritSection);

		if(empty) {
			break;
		} else if(!pResult) {
			continue;
		}

		int itemIndex = ItemIDToIndex(pResult->itemID);
		if(itemIndex >= 0) {
			LVSETINFOTIP infoTip = {0};
			infoTip.cbSize = sizeof(LVSETINFOTIP);
			infoTip.iItem = itemIndex;
			infoTip.pszText = pResult->pInfoTipText;
			attachedControl.SendMessage(LVM_SETINFOTIP, 0, reinterpret_cast<LPARAM>(&infoTip));
			ATLTRACE2(shlvwTraceThreading, 3, TEXT("Set info tip text of item %i\n"), itemIndex);
		}
		if(pResult->pInfoTipText) {
			CoTaskMemFree(pResult->pInfoTipText);
		}
		delete pResult;
	}
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	LONG itemID = static_cast<LONG>(wParam);
	LONG itemIndex = ItemIDToIndex(itemID);

	if(itemIndex >= 0) {
		BOOL useFallback = TRUE;
		if(properties.displayThumbnails && properties.thumbnailsStatus.pThumbnailImageList) {
			useFallback = FALSE;
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, static_cast<int>(lParam), 0, NULL, 0)));

			LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
			if(pItemDetails) {
				ATLVERIFY(SUCCEEDED(GetThumbnailAsync(pItemDetails->pIDL, itemID, properties.thumbnailsStatus.thumbnailSize)));
			}
			// NOTE: We don't call GetElevationShieldAsync, because ShLvwThumbnailTask already checks the elevation requirements.
		}
		if(properties.pUnifiedImageList) {
			useFallback = FALSE;
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, static_cast<int>(lParam), 0, NULL, 0)));

			if(!RunTimeHelper::IsCommCtrl6()) {
				UINT iconFlags = 0;
				ATLVERIFY(SUCCEEDED(pImgLstPriv->GetIconInfo(itemID, SIIF_FLAGS, NULL, NULL, NULL, &iconFlags)));
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_FLAGS, NULL, 0, 0, NULL, iconFlags | AII_USELEGACYDISPLAYCODE)));
			}
			if(properties.displayElevationShieldOverlays) {
				PCIDLIST_ABSOLUTE pIDL = ItemIDToPIDL(itemID);
				if(pIDL) {
					ATLVERIFY(SUCCEEDED(GetElevationShieldAsync(itemID, pIDL)));
				}
			}
		}
		if(useFallback) {
			LVITEM item = {0};
			item.iItem = itemIndex;
			item.iImage = static_cast<int>(lParam);
			item.mask = LVIF_IMAGE;
			attachedControl.SendMessage(LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
		} else {
			attachedControl.SendMessage(LVM_REDRAWITEMS, itemIndex, itemIndex);
		}
	}
	ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated icon of item %i to %i\n"), itemIndex, lParam);
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateOverlay(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	LONG itemID = static_cast<LONG>(wParam);
	LONG itemIndex = ItemIDToIndex(itemID);

	if(itemIndex >= 0) {
		LVITEM item = {0};
		item.state = INDEXTOOVERLAYMASK(static_cast<int>(lParam));
		item.stateMask = LVIS_OVERLAYMASK;
		attachedControl.SendMessage(LVM_SETITEMSTATE, itemIndex, reinterpret_cast<LPARAM>(&item));
	}
	ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated overlay of item %i to %i\n"), itemIndex, lParam);
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateSubItemControl(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		if(flags.detaching) {
			return 0;
		}

		while(TRUE) {
			LPSHLVWBACKGROUNDCOLUMNINFO pUpdateControlResult = NULL;
			EnterCriticalSection(&properties.subItemControlResultsCritSection);
			BOOL empty = TRUE;
			#ifdef USE_STL
				if(!properties.subItemControlResults.empty()) {
					pUpdateControlResult = properties.subItemControlResults.front();
					properties.subItemControlResults.pop();
					empty = FALSE;
				}
			#else
				if(!properties.subItemControlResults.IsEmpty()) {
					pUpdateControlResult = properties.subItemControlResults.RemoveHead();
					empty = FALSE;
				}
			#endif
			LeaveCriticalSection(&properties.subItemControlResultsCritSection);

			if(empty) {
				break;
			} else if(!pUpdateControlResult) {
				continue;
			}

			int itemIndex = ItemIDToIndex(pUpdateControlResult->itemID);
			int subItemIndex = ColumnIDToIndex(pUpdateControlResult->columnID);
			if(itemIndex != -1 && subItemIndex != -1) {
				if(pUpdateControlResult->pPropertyValue) {
					LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pUpdateControlResult->itemID);
					if(pItemDetails && (pItemDetails->managedProperties & slmipSubItemControls)) {
						LPSHELLLISTVIEWSUBITEMDATA pSubItem = NULL;
						#ifdef USE_STL
							std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = pItemDetails->subItems.find(pUpdateControlResult->columnID);
							if(it != pItemDetails->subItems.cend()) {
								pSubItem = it->second;
							}
						#else
							CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA>::CPair* pPair = pItemDetails->subItems.Lookup(pUpdateControlResult->columnID);
							if(pPair) {
								pSubItem = pPair->m_value;
							}
						#endif
						if(pSubItem) {
							LPPROPVARIANT pPropValue = pSubItem->pPropertyValue;
							pSubItem->pPropertyValue = pUpdateControlResult->pPropertyValue;
							attachedControl.SendMessage(LVM_REDRAWITEMS, itemIndex, itemIndex);
							ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated sub-item control value of item %i, sub-item %i\n"), itemIndex, subItemIndex);
							pUpdateControlResult->pPropertyValue = NULL;
							if(pPropValue) {
								PropVariantClear(pPropValue);
								delete pPropValue;
							}
						}
					}
					if(pUpdateControlResult->pPropertyValue) {
						PropVariantClear(pUpdateControlResult->pPropertyValue);
						delete pUpdateControlResult->pPropertyValue;
						pUpdateControlResult->pPropertyValue = NULL;
					}
				}
			}
			delete pUpdateControlResult;
		}
	#endif
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateText(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDCOLUMNINFO pUpdateTextResult = NULL;
		EnterCriticalSection(&properties.slowColumnResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.slowColumnResults.empty()) {
				pUpdateTextResult = properties.slowColumnResults.front();
				properties.slowColumnResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.slowColumnResults.IsEmpty()) {
				pUpdateTextResult = properties.slowColumnResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.slowColumnResultsCritSection);

		if(empty) {
			break;
		} else if(!pUpdateTextResult) {
			continue;
		}

		int itemIndex = ItemIDToIndex(pUpdateTextResult->itemID);
		int subItemIndex = ColumnIDToIndex(pUpdateTextResult->columnID);
		if(itemIndex != -1 && subItemIndex != -1) {
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				if(pUpdateTextResult->pPropertyValue) {
					LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pUpdateTextResult->itemID);
					if(pItemDetails && (pItemDetails->managedProperties & slmipSubItemControls)) {
						LPSHELLLISTVIEWSUBITEMDATA pSubItem = NULL;
						#ifdef USE_STL
							std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = pItemDetails->subItems.find(pUpdateTextResult->columnID);
							if(it != pItemDetails->subItems.cend()) {
								pSubItem = it->second;
							}
						#else
							CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA>::CPair* pPair = pItemDetails->subItems.Lookup(pUpdateTextResult->columnID);
							if(pPair) {
								pSubItem = pPair->m_value;
							}
						#endif
						if(pSubItem) {
							LPPROPVARIANT pPropValue = pSubItem->pPropertyValue;
							pSubItem->pPropertyValue = pUpdateTextResult->pPropertyValue;
							pUpdateTextResult->pPropertyValue = NULL;
							if(pPropValue) {
								PropVariantClear(pPropValue);
								delete pPropValue;
							}
						}
					}
					if(pUpdateTextResult->pPropertyValue) {
						PropVariantClear(pUpdateTextResult->pPropertyValue);
						delete pUpdateTextResult->pPropertyValue;
						pUpdateTextResult->pPropertyValue = NULL;
					}
				}
			#endif
			LVITEM item = {0};
			item.iSubItem = subItemIndex;
			#ifdef UNICODE
				item.pszText = pUpdateTextResult->pText;
			#else
				CW2A converter(pUpdateTextResult->pText);
				item.pszText = converter;
			#endif
			attachedControl.SendMessage(LVM_SETITEMTEXT, itemIndex, reinterpret_cast<LPARAM>(&item));
			ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated text of item %i, sub-item %i to %s\n"), itemIndex, subItemIndex, item.pszText);
		}
		delete pUpdateTextResult;
	}
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateThumbnail(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return 0;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDTHUMBNAILINFO pThumbnailResult = NULL;
		EnterCriticalSection(&properties.backgroundThumbnailResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.backgroundThumbnailResults.empty()) {
				pThumbnailResult = properties.backgroundThumbnailResults.front();
				properties.backgroundThumbnailResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.backgroundThumbnailResults.IsEmpty()) {
				pThumbnailResult = properties.backgroundThumbnailResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.backgroundThumbnailResultsCritSection);

		if(empty) {
			break;
		} else if(!pThumbnailResult) {
			continue;
		}
		if(!(pThumbnailResult->flags & AII_IMAGERESOURCEOVERLAY) && pThumbnailResult->pOverlayImageResource) {
			CoTaskMemFree(pThumbnailResult->pOverlayImageResource);
			pThumbnailResult->pOverlayImageResource = NULL;
		}

		int itemIndex = ItemIDToIndex(pThumbnailResult->itemID);
		if(itemIndex >= 0) {
			if(pThumbnailResult->aborted) {
				if(pThumbnailResult->pOverlayImageResource) {
					CoTaskMemFree(pThumbnailResult->pOverlayImageResource);
				}
				LVITEM item = {0};
				item.iItem = itemIndex;
				item.iImage = I_IMAGECALLBACK;
				item.mask = LVIF_IMAGE;
				attachedControl.SendMessage(LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));

			} else if(properties.thumbnailsStatus.pThumbnailImageList && pThumbnailResult->targetThumbnailSize.cx == properties.thumbnailsStatus.thumbnailSize.cx && pThumbnailResult->targetThumbnailSize.cy == properties.thumbnailsStatus.thumbnailSize.cy) {
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				if(!RunTimeHelper::IsCommCtrl6()) {
					pThumbnailResult->flags |= AII_USELEGACYDISPLAYCODE;
				}
				ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(pThumbnailResult->itemID, pThumbnailResult->mask, pThumbnailResult->hThumbnailOrIcon, 0, pThumbnailResult->executableIconIndex, pThumbnailResult->pOverlayImageResource, pThumbnailResult->flags)));
				pThumbnailResult->hThumbnailOrIcon = NULL;
				if(pThumbnailResult->flags & AII_USELEGACYDISPLAYCODE) {
					/* TODO: Find a better solution for the problem that the space between the items isn't updated,
					 *       so that fragments of old, larger thumbnails remain.
					 */
					RECT rc = {LVIR_BOUNDS, 0, 0, 0};
					attachedControl.SendMessage(LVM_GETITEMRECT, itemIndex, reinterpret_cast<LPARAM>(&rc));
					InflateRect(&rc, 240, 240);
					attachedControl.InvalidateRect(&rc);
				} else {
					attachedControl.SendMessage(LVM_REDRAWITEMS, itemIndex, itemIndex);
				}
			} else if(pThumbnailResult->pOverlayImageResource) {
				CoTaskMemFree(pThumbnailResult->pOverlayImageResource);
			}
		}
		if(!pThumbnailResult->hThumbnailOrIcon) {
			DeleteObject(pThumbnailResult->hThumbnailOrIcon);
		}
		delete pThumbnailResult;
		ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated thumbnail of item %i\n"), itemIndex);
	}
	return 0;
}

LRESULT ShellListView::OnTriggerUpdateTileViewColumns(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/)
{
	if(flags.detaching) {
		return FALSE;
	}

	while(TRUE) {
		LPSHLVWBACKGROUNDTILEVIEWINFO pResult = NULL;
		EnterCriticalSection(&properties.backgroundTileViewResultsCritSection);
		BOOL empty = TRUE;
		#ifdef USE_STL
			if(!properties.backgroundTileViewResults.empty()) {
				pResult = properties.backgroundTileViewResults.front();
				properties.backgroundTileViewResults.pop();
				empty = FALSE;
			}
		#else
			if(!properties.backgroundTileViewResults.IsEmpty()) {
				pResult = properties.backgroundTileViewResults.RemoveHead();
				empty = FALSE;
			}
		#endif
		LeaveCriticalSection(&properties.backgroundTileViewResultsCritSection);

		if(empty) {
			break;
		} else if(!pResult) {
			continue;
		}

		int itemIndex = ItemIDToIndex(pResult->itemID);
		if(itemIndex >= 0) {
			ATLASSERT(properties.columnsStatus.loaded == SCLS_LOADED);
			// set the columns
			LVTILEINFO tileInfo = {0};
			tileInfo.cbSize = RunTimeHelper::SizeOf_LVTILEINFO();
			tileInfo.iItem = itemIndex;
			tileInfo.cColumns = DPA_GetPtrCount(pResult->hColumnBuffer);
			if(tileInfo.cColumns > 0) {
				tileInfo.puColumns = static_cast<PUINT>(HeapAlloc(GetProcessHeap(), 0, tileInfo.cColumns * sizeof(UINT)));
				ATLASSERT(tileInfo.puColumns);
				if(tileInfo.puColumns) {
					tileInfo.piColFmt = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, tileInfo.cColumns * sizeof(int)));
					ATLASSERT(tileInfo.piColFmt);
					if(tileInfo.piColFmt) {
						for(UINT i = 0; i < tileInfo.cColumns; ++i) {
							int realColumnIndex = reinterpret_cast<int>(DPA_GetPtr(pResult->hColumnBuffer, i));
							ChangeColumnVisibility(realColumnIndex, TRUE, 0);
							tileInfo.puColumns[i] = RealColumnIndexToIndex(realColumnIndex);
							LPSHELLLISTVIEWCOLUMNDATA pColumnDetails = GetColumnDetails(realColumnIndex);
							if(pColumnDetails) {
								tileInfo.piColFmt[i] = pColumnDetails->format;
							}
						}
					} else {
						for(UINT i = 0; i < tileInfo.cColumns; ++i) {
							int realColumnIndex = reinterpret_cast<int>(DPA_GetPtr(pResult->hColumnBuffer, i));
							ChangeColumnVisibility(realColumnIndex, TRUE, 0);
							tileInfo.puColumns[i] = RealColumnIndexToIndex(realColumnIndex);
						}
					}
				} else {
					tileInfo.cColumns = 0;
				}
			}
			if(attachedControl.SendMessage(LVM_SETTILEINFO, 0, reinterpret_cast<LPARAM>(&tileInfo))) {
				if(tileInfo.puColumns) {
					HeapFree(GetProcessHeap(), 0, tileInfo.puColumns);
				}
				if(tileInfo.piColFmt) {
					HeapFree(GetProcessHeap(), 0, tileInfo.piColFmt);
				}
			}
			ATLTRACE2(shlvwTraceThreading, 3, TEXT("Updated tile view columns of item %i\n"), itemIndex);
		}
		DPA_Destroy(pResult->hColumnBuffer);
		delete pResult;
	}
	return 0;
}


HRESULT ShellListView::OnInternalAllocateMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	ATLASSERT(wParam > 0);
	return reinterpret_cast<HRESULT>(HeapAlloc(GetProcessHeap(), 0, wParam));
}

HRESULT ShellListView::OnInternalFreeMemory(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	ATLASSERT(lParam);
	HeapFree(GetProcessHeap(), 0, reinterpret_cast<LPVOID>(lParam));
	return S_OK;
}

HRESULT ShellListView::OnInternalInsertedColumn(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	LONG columnID = static_cast<LONG>(wParam);
	#ifdef USE_STL
		int realColumnIndex = properties.shellColumnIndexesOfInsertedColumns.top();
		properties.shellColumnIndexesOfInsertedColumns.pop();
	#else
		int realColumnIndex = properties.shellColumnIndexesOfInsertedColumns.RemoveHead();
	#endif
	ChangedColumnVisibility(realColumnIndex, columnID, TRUE);
	return 0;
}

HRESULT ShellListView::OnInternalFreeColumn(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	ATLASSERT(wParam != -1);

	if(wParam == -1) {
		#ifdef USE_STL
			std::vector<int> columns;
			for(std::unordered_map<LONG, int>::const_iterator iter = properties.visibleColumns.cbegin(); iter != properties.visibleColumns.cend(); ++iter) {
				columns.push_back(iter->second);
			}
			for(std::vector<int>::const_iterator iter = columns.cbegin(); iter != columns.cend(); ++iter) {
				ChangedColumnVisibility(*iter, -1, FALSE);
			}
			ATLASSERT(properties.visibleColumns.size() == 0);
		#else
			CAtlArray<int> columns;
			POSITION p = properties.visibleColumns.GetStartPosition();
			while(p) {
				columns.Add(properties.visibleColumns.GetNextValue(p));
			}
			for(size_t i = 0; i < columns.GetCount(); ++i) {
				ChangedColumnVisibility(columns[i], -1, FALSE);
			}
			ATLASSERT(properties.visibleColumns.GetCount() == 0);
		#endif
	} else {
		LONG columnID = static_cast<LONG>(wParam);
		#ifdef USE_STL
			std::unordered_map<LONG, int>::const_iterator iter = properties.visibleColumns.find(columnID);
			if(iter != properties.visibleColumns.cend()) {
				int realColumnIndex = iter->second;
		#else
			CAtlMap<LONG, int>::CPair* pPair = properties.visibleColumns.Lookup(columnID);
			if(pPair) {
				int realColumnIndex = pPair->m_value;
		#endif
			ChangedColumnVisibility(realColumnIndex, -1, FALSE);
		}
	}
	properties.pColumns->Reset();

	return 0;
}

HRESULT ShellListView::OnInternalGetShListViewColumn(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPDISPATCH* ppColumn = reinterpret_cast<LPDISPATCH*>(lParam);
	ATLASSERT_POINTER(ppColumn, LPDISPATCH);
	if(!ppColumn) {
		return S_FALSE;
	}

	*ppColumn = NULL;

	LONG columnID = static_cast<LONG>(wParam);
	if(IsShellColumn(columnID)) {
		int realColumnIndex = ColumnIDToRealIndex(columnID);
		if(realColumnIndex >= 0) {
			ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(ppColumn));
		}
	}
	return S_OK;
}

HRESULT ShellListView::OnInternalHeaderClick(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	if(!properties.sortOnHeaderClick) {
		return 0;
	}

	int realColumnIndex = ColumnIDToRealIndex(static_cast<LONG>(wParam));
	if(realColumnIndex != -1) {
		flags.toggleSortOrder = TRUE;
		SortItems(realColumnIndex);
		flags.toggleSortOrder = FALSE;
	} else if(sortingSettings.columnIndex != -1) {
		// not a shell column - remove sort arrow from old sort column
		int columnIndex = RealColumnIndexToIndex(sortingSettings.columnIndex);
		if(columnIndex != -1) {
			if(RunTimeHelper::IsCommCtrl6()) {
				if(properties.setSortArrows) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, 0/*saNone*/)));
				}
				if(properties.selectSortColumn) {
					// deselect column
					attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, static_cast<WPARAM>(-1), 0);
				}
			} else {
				if(properties.setSortArrows) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, NULL)));
					sortingSettings.FreeBitmapHandle();
				}
			}
		}
	}
	return 0;
}

HRESULT ShellListView::OnInternalInsertedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LONG itemID = static_cast<LONG>(wParam);
	LPLVITEM pItemData = reinterpret_cast<LPLVITEM>(lParam);
	if(pItemData->mask & LVIF_ISASHELLITEM) {
		#ifdef USE_STL
			LPSHELLLISTVIEWITEMDATA pShellItemData = properties.shellItemDataOfInsertedItems.top();
			properties.shellItemDataOfInsertedItems.pop();
			AddItem(itemID, pShellItemData);
			std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter = properties.namespacesOfInsertedItems.find(pShellItemData->pIDL);
			if(iter != properties.namespacesOfInsertedItems.cend()) {
				pShellItemData->pIDLNamespace = iter->second;
				properties.namespacesOfInsertedItems.erase(iter);
			}
		#else
			LPSHELLLISTVIEWITEMDATA pShellItemData = properties.shellItemDataOfInsertedItems.RemoveHead();
			AddItem(itemID, pShellItemData);
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::CPair* pPair = properties.namespacesOfInsertedItems.Lookup(pShellItemData->pIDL);
			if(pPair) {
				pShellItemData->pIDLNamespace = pPair->m_value;
				properties.namespacesOfInsertedItems.RemoveKey(pPair->m_key);
			}
		#endif
	} else if(pItemData->mask & LVIF_IMAGE) {
		// store icon index
		BOOL setIconIndex = FALSE;
		if(properties.thumbnailsStatus.pThumbnailImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, pItemData->iImage, 0, NULL, AII_NONSHELLITEM | AII_NOADORNMENT | AII_NOFILETYPEOVERLAY)));
			setIconIndex = TRUE;
		}
		if(properties.pUnifiedImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_NONSHELLITEMICON | SIIF_FLAGS, NULL, pItemData->iImage, 0, NULL, AII_NONSHELLITEM | AII_NOADORNMENT | AII_NOFILETYPEOVERLAY)));
			setIconIndex = TRUE;
		}
		if(setIconIndex) {
			// set the item's icon index to its ID
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETITEMICONINDEX, itemID, itemID)));
		}
	}
	return 0;
}

HRESULT ShellListView::OnInternalFreeItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
{
	if(wParam == -1) {
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			#ifdef USE_STL
				if(properties.items.size() > 0) {
					Raise_RemovingItem(NULL);
				}
			#else
				if(properties.items.GetCount() > 0) {
					Raise_RemovingItem(NULL);
				}
			#endif
		}

		#ifdef USE_STL
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				delete iter->second;
			}
			properties.items.clear();
		#else
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				delete properties.items.GetValueAt(p);
				properties.items.GetNextValue(p);
			}
			properties.items.RemoveAll();
		#endif
		if(properties.thumbnailsStatus.pThumbnailImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->RemoveIconInfo(0xFFFFFFFF)));
		}
		if(properties.pUnifiedImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->RemoveIconInfo(0xFFFFFFFF)));
		}
	} else {
		LONG itemID = static_cast<LONG>(wParam);
		#ifdef USE_STL
			std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.find(itemID);
			if(iter != properties.items.cend()) {
		#else
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemID);
			if(pPair) {
		#endif
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				CComPtr<IShListViewItem> pItem;
				ClassFactory::InitShellListItem(itemID, this, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
				Raise_RemovingItem(pItem);
			}
		#ifdef USE_STL
				LPSHELLLISTVIEWITEMDATA pData = iter->second;
				properties.items.erase(itemID);
		#else
				LPSHELLLISTVIEWITEMDATA pData = pPair->m_value;
				properties.items.RemoveKey(itemID);
		#endif
			delete pData;
		}
		if(properties.thumbnailsStatus.pThumbnailImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->RemoveIconInfo(itemID)));
		}
		if(properties.pUnifiedImageList) {
			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);
			ATLVERIFY(SUCCEEDED(pImgLstPriv->RemoveIconInfo(itemID)));
		}
	}
	properties.pListItems->Reset();

	return 0;
}

HRESULT ShellListView::OnInternalGetShListViewItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPDISPATCH* ppItem = reinterpret_cast<LPDISPATCH*>(lParam);
	ATLASSERT_POINTER(ppItem, LPDISPATCH);
	if(!ppItem) {
		return S_FALSE;
	}

	if(IsShellItem(static_cast<LONG>(wParam))) {
		ClassFactory::InitShellListItem(static_cast<LONG>(wParam), this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(ppItem));
	} else {
		*ppItem = NULL;
	}
	return S_OK;
}

HRESULT ShellListView::OnInternalRenamedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	BOOL success = TRUE;

	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(static_cast<LONG>(wParam));
	if(pItemDetails) {
		if(pItemDetails->managedProperties & slmipRenaming) {
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

HRESULT ShellListView::OnInternalCompareItems(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	#define ResultFromShort(i) ResultFromScode(MAKE_SCODE(SEVERITY_SUCCESS, 0, static_cast<USHORT>(i)))

	ATLASSERT_ARRAYPOINTER(wParam, LONG, 2);
	ATLASSERT_POINTER(lParam, BOOL);

	LONG* pItems = reinterpret_cast<LONG*>(wParam);
	LPBOOL pWasHandled = reinterpret_cast<LPBOOL>(lParam);
	if(!wParam || !lParam) {
		return E_INVALIDARG;
	}

	PCIDLIST_ABSOLUTE pIDL1 = ItemIDToPIDL(pItems[0]);
	PCIDLIST_ABSOLUTE pIDL2 = ItemIDToPIDL(pItems[1]);
	if(pIDL1) {
		if(pIDL2) {
			// both items are shell items - let the shell compare them
			*pWasHandled = TRUE;
			if(properties.useFastButImpreciseItemHandling) {
				// try to use the column namespace's IShellFolder interface
				if(properties.columnsStatus.pNamespaceSHF) {
					return properties.columnsStatus.pNamespaceSHF->CompareIDs(MAKELPARAM(max(sortingSettings.columnIndex, properties.columnsStatus.defaultSortColumn), 0), ILFindLastID(pIDL1), ILFindLastID(pIDL2));
				}
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
				return pParentISF->CompareIDs(MAKELPARAM(max(sortingSettings.columnIndex, properties.columnsStatus.defaultSortColumn), 0), ILFindLastID(pIDL1), ILFindLastID(pIDL2));
			}
			ILFree(pIDLParent1);

			/* TODO: On Vista, try to use SHCompareIDsFull (but check with "My Computer" whether sorting is still
			         right then!) */
			// fallback to the Desktop's IShellFolder interface
			return properties.pDesktopISF->CompareIDs(MAKELPARAM(max(sortingSettings.columnIndex, properties.columnsStatus.defaultSortColumn), 0), pIDL1, pIDL2);
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

HRESULT ShellListView::OnInternalBeginLabelEdit(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	NMLVDISPINFO* pNotificationDetails = reinterpret_cast<NMLVDISPINFO*>(lParam);
	//ATLASSERT_POINTER(pNotificationDetails, NMLVDISPINFOA);
	ATLASSERT(pNotificationDetails);
	if(!pNotificationDetails) {
		return E_INVALIDARG;
	}
	BOOL cancel = static_cast<BOOL>(wParam);

	ATLASSERT(pNotificationDetails->item.mask & LVIF_PARAM);
	LONG itemID = static_cast<LONG>(pNotificationDetails->item.lParam);
	if(RunTimeHelper::IsCommCtrl6()) {
		itemID = ItemIndexToID(pNotificationDetails->item.iItem);
	}

	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	if(pItemDetails) {
		if(pItemDetails->managedProperties & slmipRenaming) {
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

					HWND hWndEdit = reinterpret_cast<HWND>(SendMessage(pNotificationDetails->hdr.hwndFrom, LVM_GETEDITCONTROL, 0, 0));
					if(hWndEdit) {
						if(properties.limitLabelEditInput) {
							APIWrapper::SHLimitInputEdit(hWndEdit, pParentISF, NULL);
						}

						if(pItemDetails->managedProperties & slmipText) {
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

HRESULT ShellListView::OnInternalBeginDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPNMLISTVIEW pNotificationDetails = reinterpret_cast<LPNMLISTVIEW>(lParam);
	ATLASSERT_POINTER(pNotificationDetails, NMLISTVIEW);
	ATLASSERT(pNotificationDetails);
	if(!pNotificationDetails) {
		return 0;
	}

	if(properties.handleOLEDragDrop & hoddSourcePart) {
		// get the IDs of all selected items
		UINT selectedItems = attachedControl.SendMessage(LVM_GETSELECTEDCOUNT, 0, 0);
		PLONG pItems = new LONG[selectedItems];
		if(pItems) {
			int itemIndex = -1;
			for(UINT i = 0; i < selectedItems; ++i) {
				itemIndex = static_cast<int>(attachedControl.SendMessage(LVM_GETNEXTITEM, itemIndex, MAKELPARAM(LVNI_ALL | LVNI_SELECTED, 0)));
				ATLASSERT(itemIndex >= 0);
				// try to save the call to ItemIndexToID
				pItems[i] = ItemIndexToID(itemIndex);
			}

			// now get the pIDLs of the shell items
			PIDLIST_ABSOLUTE* ppIDLs = NULL;
			UINT pIDLCount = ItemIDsToPIDLs(pItems, selectedItems, ppIDLs, FALSE);
			if(pIDLCount > 0) {
				ATLASSERT_ARRAYPOINTER(ppIDLs, PIDLIST_ABSOLUTE, pIDLCount);
				if(ppIDLs) {
					// NOTE: It really makes a difference which IShellFolder is queried for the data object!
					BOOL singleNamespace = HaveSameParent(ppIDLs, pIDLCount);

					LPSHELLFOLDER pParentISF = NULL;
					PUITEMID_CHILD* pRelativePIDLs = NULL;

					if(singleNamespace) {
						HRESULT hr = ::SHBindToParent(ppIDLs[0], IID_PPV_ARGS(&pParentISF), NULL);
						if(SUCCEEDED(hr)) {
							// get the relative pIDLs
							ATLASSUME(pParentISF);
							pRelativePIDLs = new PUITEMID_CHILD[pIDLCount];
							ATLASSERT_ARRAYPOINTER(pRelativePIDLs, PUITEMID_CHILD, pIDLCount);
							if(pRelativePIDLs) {
								for(UINT i = 0; i < pIDLCount; ++i) {
									pRelativePIDLs[i] = ILFindLastID(ppIDLs[i]);
									ATLASSERT_POINTER(pRelativePIDLs[i], ITEMID_CHILD);
								}
							}
						}
					} else {
						ATLVERIFY(SUCCEEDED(DesktopShellFolderHook::CreateInstance(&pParentISF)));
						ATLASSUME(pParentISF);
						pRelativePIDLs = reinterpret_cast<PUITEMID_CHILD*>(ppIDLs);
					}

					if(pRelativePIDLs) {
						// retrieve the items' data object
						LPDATAOBJECT pDataObject = NULL;
						HRESULT hr = E_FAIL;
						if(pParentISF) {
							hr = pParentISF->GetUIObjectOf(GethWndShellUIParentWindow(), pIDLCount, pRelativePIDLs, IID_IDataObject, NULL, reinterpret_cast<LPVOID*>(&pDataObject));
						}
						if(SUCCEEDED(hr)) {
							ATLASSUME(pDataObject);
							// retrieve the supported verbs
							SFGAOF attributes = SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE;
							hr = pParentISF->GetAttributesOf(pIDLCount, pRelativePIDLs, &attributes);
							if(SUCCEEDED(hr)) {
								int supportedEffects = attributes & (SFGAO_CANCOPY | SFGAO_CANLINK | SFGAO_CANMOVE);
								if(!(supportedEffects & 2/*odeMove*/)) {
									// always support move if the view mode supports it
									if(cachedSettings.viewMode == -1) {
										ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
									}
									if(cachedSettings.viewMode != 3/*vDetails*/ && cachedSettings.viewMode != 2/*vList*/) {
										supportedEffects = supportedEffects | 2/*odeMove*/;
									}
								}
								if(supportedEffects != 0/*odeNone*/) {
									// create the item container
									ATLASSUME(properties.pAttachedInternalMessageListener);
									EXLVMCREATEITEMCONTAINERDATA containerData = {0};
									containerData.structSize = sizeof(EXLVMCREATEITEMCONTAINERDATA);
									containerData.numberOfItems = selectedItems;
									containerData.pItems = pItems;
									hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_CREATEITEMCONTAINER, 0, reinterpret_cast<LPARAM>(&containerData));
									if(SUCCEEDED(hr)) {
										ATLASSUME(containerData.pContainer);

										POINT ptOrigin = {0};
										attachedControl.SendMessage(LVM_GETORIGIN, 0, reinterpret_cast<LPARAM>(&ptOrigin));
										dragDropStatus.positionOfDragStart.x = ptOrigin.x + pNotificationDetails->ptAction.x;
										dragDropStatus.positionOfDragStart.y = ptOrigin.y + pNotificationDetails->ptAction.y;

										// finally call OLEDrag
										// TODO: What shall we do with <performedEffects>?
										int performedEffects = 0/*odeNone*/;
										EXLVMOLEDRAGDATA oleDragData = {0};
										oleDragData.structSize = sizeof(EXLVMOLEDRAGDATA);
										oleDragData.useSHDoDragDrop = TRUE;
										oleDragData.useShellDropSource = FALSE; //TODO: ((properties.handleOLEDragDrop & hoddSourcePartDefaultDropSource) == hoddSourcePartDefaultDropSource);
										oleDragData.pDataObject = pDataObject;
										oleDragData.supportedEffects = supportedEffects;
										oleDragData.pDraggedItems = containerData.pContainer;
										oleDragData.performedEffects = performedEffects;
										hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_OLEDRAG, 0, reinterpret_cast<LPARAM>(&oleDragData));
										if(SUCCEEDED(hr)) {
											performedEffects = oleDragData.performedEffects;
										} else if(oleDragData.useSHDoDragDrop) {
											oleDragData.useSHDoDragDrop = FALSE;
											oleDragData.useShellDropSource = FALSE;
											hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_OLEDRAG, 0, reinterpret_cast<LPARAM>(&oleDragData));
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

						if(pRelativePIDLs && pRelativePIDLs != reinterpret_cast<PUITEMID_CHILD*>(ppIDLs)) {
							delete[] pRelativePIDLs;
						}
					}

					delete[] ppIDLs;
					if(pParentISF) {
						pParentISF->Release();
					}
				}
			}
			delete[] pItems;
		}
	}

	return 0;
}

HRESULT ShellListView::OnInternalHandleDragEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	if(!(properties.handleOLEDragDrop & hoddTargetPart)) {
		return S_FALSE;
	}

	LPSHLVMDRAGDROPEVENTDATA pEventDetails = reinterpret_cast<LPSHLVMDRAGDROPEVENTDATA>(lParam);
	ATLASSERT(pEventDetails);
	if(pEventDetails->structSize < sizeof(SHLVMDRAGDROPEVENTDATA)) {
		return S_FALSE;
	}
	ATLASSUME(pEventDetails->pDataObject);
	if(!pEventDetails->pDataObject) {
		return S_FALSE;
	}

	#ifdef USE_STL
		BOOL singleNamespace = (properties.namespaces.size() == 1);
	#else
		BOOL singleNamespace = (properties.namespaces.GetCount() == 1);
	#endif
	ATLASSUME(properties.pAttachedInternalMessageListener);
	if(cachedSettings.viewMode == -1) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	}
	BOOL weAreSource = (properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_CONTROLISDRAGSOURCE, 0, 0) == S_OK);
	BOOL allowBackgroundDrop = TRUE;
	//if(viewMode == 3/*vDetails*/ || viewMode == 2/*vList*/) {
	//	// allow background drops only if we're not the drag source
	//	allowBackgroundDrop = !weAreSource;
	//}
	BOOL useInsertionMark = RunTimeHelper::IsCommCtrl6() && ((attachedControl.GetStyle() & LVS_AUTOARRANGE) == LVS_AUTOARRANGE);
	// TRUE if we're doing a drop on the background only because the target item doesn't support IDropTarget

	if(wParam != OLEDRAGEVENT_DROP) {
		dragDropStatus.lastKeyState = pEventDetails->keyState;
	}
	CComQIPtr<IDataObject> pDataObject = pEventDetails->pDataObject;
	LONG idDropTargetBeforeEvent = -1;
	if(!pEventDetails->headerIsTarget) {
		idDropTargetBeforeEvent = pEventDetails->dropTarget;
	}
	if(idDropTargetBeforeEvent != -1) {
		//if(viewMode != 3/*vDetails*/ && viewMode != 2/*vList*/) {
			LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(idDropTargetBeforeEvent);
			if(pItemDetails) {
				CComPtr<IDropTarget> pDropTarget = NULL;
				if(FAILED(::GetDropTarget(GethWndShellUIParentWindow(), pItemDetails->pIDL, &pDropTarget, FALSE))) {
					// act as if the background would be the drop target
					idDropTargetBeforeEvent = -1;
				}
			}
		//}
	}
	BOOL droppingOnBackground = (idDropTargetBeforeEvent == -1);

	#ifdef DEBUG
		CListViewCtrl listview = attachedControl;
		CAtlString itemText;
		if(!droppingOnBackground) {
			listview.GetItemText(ItemIDToIndex(idDropTargetBeforeEvent), 0, itemText);
		}
		ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Handling drag event %i (drop target item: %i - %s)\n"), wParam, idDropTargetBeforeEvent, itemText);

		if(wParam == OLEDRAGEVENT_DRAGENTER) {
			ATLASSERT(dragDropStatus.pCurrentDropTarget == NULL);
			ATLASSERT(dragDropStatus.idCurrentDropTarget == -2);
		}
	#endif

	CComPtr<IDropTarget> pNewDropTarget = NULL;
	BOOL isSameDropTarget = (idDropTargetBeforeEvent == dragDropStatus.idCurrentDropTarget);
	PCIDLIST_ABSOLUTE pDropTargetPIDL = NULL;
	BOOL forceAllowMove = FALSE;
	if(droppingOnBackground) {
		if(allowBackgroundDrop) {
			forceAllowMove = weAreSource && (pEventDetails->draggingMouseButton != MK_RBUTTON);
			if(singleNamespace) {
				#ifdef USE_STL
					pDropTargetPIDL = properties.namespaces.cbegin()->first;
				#else
					POSITION p = properties.namespaces.GetStartPosition();
					ATLASSERT(p);
					pDropTargetPIDL = properties.namespaces.GetKeyAt(p);
				#endif
			}
		}
	} else {
		LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(idDropTargetBeforeEvent);
		if(pItemDetails) {
			pDropTargetPIDL = pItemDetails->pIDL;
		}
	}
	if(pDropTargetPIDL) {
		if(isSameDropTarget) {
			pNewDropTarget = dragDropStatus.pCurrentDropTarget;
		} else {
			::GetDropTarget(GethWndShellUIParentWindow(), pDropTargetPIDL, &pNewDropTarget, droppingOnBackground);
		}
	}
	if(!pNewDropTarget) {
		pEventDetails->effect = (forceAllowMove ? DROPEFFECT_MOVE : DROPEFFECT_NONE);
	}

	LPDROPTARGET pEnteredDropTarget = NULL;
	if(wParam != OLEDRAGEVENT_DRAGLEAVE) {
		// let the shell propose a drop effect
		if(pNewDropTarget) {
			if(isSameDropTarget) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), pNewDropTarget);
				ATLVERIFY(SUCCEEDED(pNewDropTarget->DragOver((wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &pEventDetails->effect)));
			} else {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Entering drop target 0x%X\n"), pNewDropTarget);
				pNewDropTarget->DragEnter(pDataObject, (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &pEventDetails->effect);
				pEnteredDropTarget = pNewDropTarget;
			}
			BOOL b = LOWORD(pEventDetails->effect) == DROPEFFECT_MOVE && weAreSource && droppingOnBackground && (cachedSettings.viewMode == 3/*vDetails*/ || cachedSettings.viewMode == 2/*vList*/);
			b = b || (LOWORD(pEventDetails->effect) != DROPEFFECT_NONE && (pEventDetails->draggingMouseButton == MK_RBUTTON));
			if(b) {
				DWORD keyState = (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState);
				if((keyState & (MK_ALT | MK_SHIFT | MK_CONTROL)) == 0) {
					pEventDetails->effect = (pEventDetails->draggingMouseButton != MK_RBUTTON ? DROPEFFECT_NONE : DROPEFFECT_COPY | DROPEFFECT_MOVE/* | DROPEFFECT_LINK*/);
				} else if(keyState & MK_SHIFT) {
					pEventDetails->effect = DROPEFFECT_COPY;
				}
			} else if(LOWORD(pEventDetails->effect) == DROPEFFECT_NONE && forceAllowMove) {
				pEventDetails->effect = DROPEFFECT_MOVE;
			} else if(LOWORD(pEventDetails->effect) == DROPEFFECT_LINK && forceAllowMove) {
				// prefer move over link
				DWORD keyState = (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState);
				if((keyState & (MK_ALT | MK_SHIFT | MK_CONTROL)) == 0) {
					pEventDetails->effect = DROPEFFECT_MOVE;
				}
			}
		}
	}

	// fire the event
	properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_FIREDRAGDROPEVENT, wParam, lParam);

	// handle the easy OLEDRAGEVENT_DRAGLEAVE case
	if(wParam == OLEDRAGEVENT_DRAGLEAVE) {
		if(dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragLeave()));
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragLeave()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragLeave()));
		}
		dragDropStatus.idCurrentDropTarget = -2;
		if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
			if(useInsertionMark) {
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, static_cast<WPARAM>(-1), 0/*impNowhere*/)));
			}
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, NULL)));
		}
		return S_OK;
	}

	// the other cases are a bit more complicated

	// Did the event handler change the drop target?
	LONG idDropTargetAfterEvent = -1;
	if(!pEventDetails->headerIsTarget) {
		idDropTargetAfterEvent = pEventDetails->dropTarget;
	}
	if(idDropTargetAfterEvent != -1) {
		//if(viewMode != 3/*vDetails*/ && viewMode != 2/*vList*/) {
			LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(idDropTargetAfterEvent);
			if(pItemDetails) {
				CComPtr<IDropTarget> pDropTarget = NULL;
				if(FAILED(::GetDropTarget(GethWndShellUIParentWindow(), pItemDetails->pIDL, &pDropTarget, FALSE))) {
					// act as if the background would be the drop target
					idDropTargetAfterEvent = -1;
				}
			}
		//}
	}
	droppingOnBackground = (idDropTargetAfterEvent == -1);
	BOOL eventHandlerHasChangedTarget = (idDropTargetAfterEvent != idDropTargetBeforeEvent);
	isSameDropTarget = (idDropTargetAfterEvent == dragDropStatus.idCurrentDropTarget);
	if(eventHandlerHasChangedTarget) {
		// update pNewDropTarget
		if(pNewDropTarget) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), pNewDropTarget);
			pNewDropTarget->DragLeave();
			pNewDropTarget = NULL;
		}
		if(droppingOnBackground) {
			if(singleNamespace && allowBackgroundDrop) {
				#ifdef USE_STL
					pDropTargetPIDL = properties.namespaces.cbegin()->first;
				#else
					POSITION p = properties.namespaces.GetStartPosition();
					ATLASSERT(p);
					pDropTargetPIDL = properties.namespaces.GetKeyAt(p);
				#endif
			} else {
				pDropTargetPIDL = NULL;
			}
		} else {
			if(!isSameDropTarget) {
				pDropTargetPIDL = NULL;
				LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(idDropTargetAfterEvent);
				if(pItemDetails) {
					pDropTargetPIDL = pItemDetails->pIDL;
				}
			}
		}
		if(isSameDropTarget) {
			pNewDropTarget = dragDropStatus.pCurrentDropTarget;
		} else if(pDropTargetPIDL) {
			::GetDropTarget(GethWndShellUIParentWindow(), pDropTargetPIDL, &pNewDropTarget, droppingOnBackground);
		}
	}

	// now we don't have to worry about hDropTargetBeforeEvent anymore
	if(isSameDropTarget) {
		if(wParam == OLEDRAGEVENT_DRAGOVER) {
			// just call DragOver
			if(eventHandlerHasChangedTarget) {
				// we entered this target on a previous call of this method, but did not yet call DragOver this time
				if(dragDropStatus.pCurrentDropTarget) {
					ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
					DWORD effect = pEventDetails->effect;
					ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver(pEventDetails->keyState, pEventDetails->cursorPosition, &effect)));
				}
			} else {
				// we already called DragOver on this target before firing the event
			}
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragOver()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragOver(reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
		} else {
			if(wParam == OLEDRAGEVENT_DRAGENTER) {
				if(pEventDetails->pDropTargetHelper) {
					ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragEnter()\n"));
					ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragEnter(attachedControl, pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
				}
			}
			// nothing else to do here (will be done below)
		}
	} else {
		// leave the current drop target
		if(dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			dragDropStatus.pCurrentDropTarget->DragLeave();
			dragDropStatus.pCurrentDropTarget = NULL;
		}
		if(wParam != OLEDRAGEVENT_DRAGENTER && pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragLeave()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragLeave()));
		}

		// enter the new drop target
		dragDropStatus.pCurrentDropTarget = pNewDropTarget;
		if(dragDropStatus.pCurrentDropTarget && dragDropStatus.pCurrentDropTarget != pEnteredDropTarget) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Entering drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			DWORD effect = pEventDetails->effect;
			dragDropStatus.pCurrentDropTarget->DragEnter(pDataObject, (wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &effect);
		}
		if(pEventDetails->pDropTargetHelper) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragEnter()\n"));
			ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragEnter(attachedControl, pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
		}
		if(wParam == OLEDRAGEVENT_DROP && dragDropStatus.pCurrentDropTarget) {
			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Moving over drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			DWORD effect = pEventDetails->effect;
			ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragOver((wParam != OLEDRAGEVENT_DROP ? pEventDetails->keyState : dragDropStatus.lastKeyState), pEventDetails->cursorPosition, &effect)));
		}
	}
	dragDropStatus.idCurrentDropTarget = idDropTargetAfterEvent;

	if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
		BOOL noInsertMark = FALSE;
		if(LOWORD(pEventDetails->effect) != DROPEFFECT_NONE && !droppingOnBackground) {
			ATLASSERT(!pEventDetails->headerIsTarget);
			noInsertMark = (pEventDetails->pDropTarget != NULL);
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, reinterpret_cast<LPARAM>(pEventDetails->pDropTarget))));
		} else {
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, NULL)));
		}

		if(useInsertionMark) {
			if(noInsertMark) {
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, static_cast<WPARAM>(-1), 0/*impNowhere*/)));
			} else {
				POINT cursorPosition = *reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition);
				attachedControl.ScreenToClient(&cursorPosition);
				int indexAndPosition[2] = {-1, 0/*impNowhere*/};
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCLOSESTINSERTMARKPOSITION, reinterpret_cast<WPARAM>(&cursorPosition), reinterpret_cast<LPARAM>(indexAndPosition))));
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, indexAndPosition[0], indexAndPosition[1])));
			}
		}
	}

	/* Now we have called DragEnter/DragOver for the correct drop target and left the old drop target if
	 * necessary. We've also made the necessary calls to IDropTargetHelper.
	 * Now we can handle OLEDRAGEVENT_DROP.
	 */

	if(wParam == OLEDRAGEVENT_DROP) {
		HRESULT hr;
		if(weAreSource && pEventDetails->effect == DROPEFFECT_MOVE && droppingOnBackground) {
			// just move the items within the control
			POINT cursorPosition = *reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition);
			attachedControl.ScreenToClient(&cursorPosition);
			POINT ptOrigin = {0};
			attachedControl.SendMessage(LVM_GETORIGIN, 0, reinterpret_cast<LPARAM>(&ptOrigin));
			SIZE movedDistance = {ptOrigin.x + cursorPosition.x - dragDropStatus.positionOfDragStart.x, ptOrigin.y + cursorPosition.y - dragDropStatus.positionOfDragStart.y};
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_MOVEDRAGGEDITEMS, 0, reinterpret_cast<LPARAM>(&movedDistance))));

			// handle this as a leave
			if(dragDropStatus.pCurrentDropTarget) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Leaving drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
				ATLVERIFY(SUCCEEDED(dragDropStatus.pCurrentDropTarget->DragLeave()));
				dragDropStatus.pCurrentDropTarget = NULL;
			}
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::DragLeave()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->DragLeave()));
			}
			dragDropStatus.idCurrentDropTarget = -2;
			if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
				if(useInsertionMark) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, static_cast<WPARAM>(-1), 0/*impNowhere*/)));
				}
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, NULL)));
			}
			return S_OK;
		}

		if(!(dragDropStatus.lastKeyState & MK_LBUTTON)) {
			// a popup menu is likely to be displayed, so make sure the drag image is already hidden
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::Drop()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
			//dragDropStatus.idCurrentDropTarget = -2;
			if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
				if(useInsertionMark) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, static_cast<WPARAM>(-1), 0/*impNowhere*/)));
				}
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, NULL)));
			}
		}
		if(dragDropStatus.pCurrentDropTarget) {
			DroppedPIDLOffsets* pOffsets = NULL;
			if(dragDropStatus.idCurrentDropTarget == -1) {
				pOffsets = new DroppedPIDLOffsets(pDataObject);
				if(!pOffsets->pDroppedPIDLs || !pOffsets->pOffsets) {
					delete pOffsets;
					pOffsets = NULL;
				}
				if(pOffsets) {
					pOffsets->cursorPosition = *reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition);
					attachedControl.ScreenToClient(&pOffsets->cursorPosition);
					POINT origin = {0};
					attachedControl.SendMessage(LVM_GETORIGIN, 0, reinterpret_cast<LPARAM>(&origin));
					pOffsets->cursorPosition.x += origin.x;
					pOffsets->cursorPosition.y += origin.y;

					pOffsets->performedEffect = pEventDetails->effect;
					if(singleNamespace) {
						pOffsets->pTargetNamespace = ILCloneFull(pDropTargetPIDL);
					}
					#ifdef USE_STL
						dragDropStatus.bufferedPIDLOffsets.push_back(pOffsets);
					#else
						dragDropStatus.bufferedPIDLOffsets.Add(pOffsets);
					#endif
				}
			}

			ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Dropping on drop target 0x%X\n"), dragDropStatus.pCurrentDropTarget);
			hr = dragDropStatus.pCurrentDropTarget->Drop(pDataObject, pEventDetails->keyState, pEventDetails->cursorPosition, &pEventDetails->effect);
			ATLASSERT(SUCCEEDED(hr));
			dragDropStatus.pCurrentDropTarget = NULL;

			if(dragDropStatus.idCurrentDropTarget == -1 && pOffsets) {
				if(SUCCEEDED(hr)) {
					// the effect might have changed
					pOffsets->performedEffect = pEventDetails->effect;
				} else {
					// remove the buffered struct again
					#ifdef USE_STL
						dragDropStatus.bufferedPIDLOffsets.pop_back();
					#else
						dragDropStatus.bufferedPIDLOffsets.RemoveAt(dragDropStatus.bufferedPIDLOffsets.GetCount() - 1);
					#endif
					delete pOffsets;
				}
			}
		}
		if(dragDropStatus.lastKeyState & MK_LBUTTON) {
			if(pEventDetails->pDropTargetHelper) {
				ATLTRACE2(shlvwTraceDragDrop, 3, TEXT("Calling IDropTargetHelper::Drop()\n"));
				ATLVERIFY(SUCCEEDED(pEventDetails->pDropTargetHelper->Drop(pDataObject, reinterpret_cast<LPPOINT>(&pEventDetails->cursorPosition), pEventDetails->effect)));
			}
			if(properties.handleOLEDragDrop & hoddTargetPartWithDropHilite) {
				if(useInsertionMark) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETINSERTMARK, static_cast<WPARAM>(-1), 0/*impNowhere*/)));
				}
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETDROPHILITEDITEM, 0, NULL)));
			}
		}
		dragDropStatus.idCurrentDropTarget = -2;
	}
	return S_OK;
}

HRESULT ShellListView::OnInternalContextMenu(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPPOINT pMenuPosition = reinterpret_cast<LPPOINT>(lParam);
	ATLASSERT_POINTER(pMenuPosition, POINT);
	if(!pMenuPosition) {
		return E_INVALIDARG;
	}

	if(wParam) {
		HWND hWndHeader = reinterpret_cast<HWND>(attachedControl.SendMessage(LVM_GETHEADER, 0, 0));
		ATLASSERT(::IsWindow(hWndHeader));
		if(::IsWindow(hWndHeader)) {
			::ClientToScreen(hWndHeader, pMenuPosition);
			return DisplayHeaderContextMenu(*pMenuPosition);
		}
		return E_FAIL;
	}

	int selectedItems = static_cast<int>(attachedControl.SendMessage(LVM_GETSELECTEDCOUNT, 0, 0));

	int contextMenuItemIndex = -1;
	LONG contextMenuItemID = -1;
	if((pMenuPosition->x == -1) && (pMenuPosition->y == -1)) {
		if(selectedItems > 0) {
			contextMenuItemIndex = static_cast<int>(attachedControl.SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_SELECTED, 0)));
			if(contextMenuItemIndex != -1) {
				CRect itemRectangle;
				itemRectangle.left = LVIR_LABEL;
				if(attachedControl.SendMessage(LVM_GETITEMRECT, contextMenuItemIndex, reinterpret_cast<LPARAM>(&itemRectangle))) {
					*pMenuPosition = itemRectangle.CenterPoint();
					contextMenuItemID = ItemIndexToID(contextMenuItemIndex);
				}
			}
		}
	} else {
		EXLVMHITTESTDATA hitTestData = {0};
		hitTestData.structSize = sizeof(EXLVMHITTESTDATA);
		hitTestData.pointToTest = *pMenuPosition;
		HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestData));
		ATLASSERT(SUCCEEDED(hr));
		if(FAILED(hr)) {
			return 0;
		}
		contextMenuItemID = hitTestData.hitItemID;
		if(contextMenuItemID != -1) {
			contextMenuItemIndex = ItemIDToIndex(contextMenuItemID);
		}
	}

	attachedControl.ClientToScreen(pMenuPosition);
	// the check for selectedItems > 0 is a sanity check
	if(contextMenuItemID != -1 && selectedItems > 0) {
		PLONG pItems = new LONG[selectedItems];
		if(pItems) {
			int itemIndex = -1;
			for(int i = 0; i < selectedItems; ++i) {
				itemIndex = static_cast<int>(attachedControl.SendMessage(LVM_GETNEXTITEM, itemIndex, MAKELPARAM(LVNI_ALL | LVNI_SELECTED, 0)));
				ATLASSERT(itemIndex >= 0);
				// try to save the call to ItemIndexToID
				pItems[i] = (itemIndex == contextMenuItemIndex ? contextMenuItemID : ItemIndexToID(itemIndex));
			}

			// bring the caret item to index 0
			int caretIndex = -1;
			for(int i = 0; i < selectedItems; ++i) {
				if(pItems[i] == contextMenuItemID) {
					caretIndex = i;
					break;
				}
			}
			//ATLASSERT(caretIndex != -1);
			if(caretIndex != -1) {
				for(int i = caretIndex; i > 0; --i) {
					// swap
					LONG buffer = pItems[i - 1];
					pItems[i - 1] = pItems[i];
					pItems[i] = buffer;
				}
			}

			DisplayShellContextMenu(pItems, selectedItems, *pMenuPosition);
			delete[] pItems;
		}
	} else {
		#ifdef USE_STL
			BOOL singleNamespace = (properties.namespaces.size() == 1);
		#else
			BOOL singleNamespace = (properties.namespaces.GetCount() == 1);
		#endif
		if(singleNamespace) {
			#ifdef USE_STL
				PCIDLIST_ABSOLUTE pIDLNamespace = properties.namespaces.cbegin()->first;
			#else
				POSITION p = properties.namespaces.GetStartPosition();
				ATLASSERT(p);
				PCIDLIST_ABSOLUTE pIDLNamespace = properties.namespaces.GetKeyAt(p);
			#endif
			ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);
			DisplayShellContextMenu(pIDLNamespace, *pMenuPosition);
		} else {
			// not supported
		}
	}
	return 0;
}

HRESULT ShellListView::OnInternalGetInfoTip(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPWSTR* ppInfoTip = reinterpret_cast<LPWSTR*>(lParam);
	ATLASSERT_POINTER(ppInfoTip, LPWSTR);

	if(ppInfoTip) {
		LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(static_cast<LONG>(wParam));
		if(pItemDetails) {
			if(pItemDetails->managedProperties & slmipInfoTip) {
				GetInfoTip(GethWndShellUIParentWindow(), pItemDetails->pIDL, properties.infoTipFlags, ppInfoTip);
			}
		}
	}
	return 0;
}

HRESULT ShellListView::OnInternalGetInfoTipEx(UINT /*message*/, WPARAM wParam, LPARAM lParam)
{
	LPSHLVMGETINFOTIPDATA pInfoTipDetails = reinterpret_cast<LPSHLVMGETINFOTIPDATA>(lParam);
	ATLASSERT(pInfoTipDetails);
	if(pInfoTipDetails->structSize < sizeof(SHLVMGETINFOTIPDATA)) {
		return S_FALSE;
	}

	if(pInfoTipDetails) {
		LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(static_cast<LONG>(wParam));
		if(pItemDetails) {
			if(pItemDetails->managedProperties & slmipInfoTip) {
				BOOL extractSynchronously = TRUE;
				if(RunTimeHelper::IsCommCtrl6()) {
					CComPtr<IRunnableTask> pTask;
					HRESULT hr = ShLvwInfoTipTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.infoTipResults, &properties.infoTipResultsCritSection, pItemDetails->pIDL, static_cast<LONG>(wParam), properties.infoTipFlags, pInfoTipDetails->pPrependedText, pInfoTipDetails->prependedTextSize, &pTask);
					if(SUCCEEDED(hr)) {
						hr = EnqueueTask(pTask, TASKID_ShLvwInfoTip, 0, TASK_PRIORITY_LV_GET_INFOTIP);
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							extractSynchronously = FALSE;
							pInfoTipDetails->cancelToolTip = TRUE;
						}
					}
				}
				if(extractSynchronously) {
					GetInfoTip(GethWndShellUIParentWindow(), pItemDetails->pIDL, properties.infoTipFlags, &pInfoTipDetails->pShellInfoTipText);
					if(pInfoTipDetails->pShellInfoTipText) {
						pInfoTipDetails->shellInfoTipTextSize = lstrlenW(pInfoTipDetails->pShellInfoTipText);
					}
				}
			}
		}
	}
	return 0;
}

HRESULT ShellListView::OnInternalCustomDraw(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPNMLVCUSTOMDRAW pNotificationDetails = reinterpret_cast<LPNMLVCUSTOMDRAW>(lParam);
	ATLASSERT_POINTER(pNotificationDetails, NMLVCUSTOMDRAW);
	if(!pNotificationDetails) {
		return CDRF_DODEFAULT;
	}

	DWORD itemType = LVCDI_ITEM;
	if(RunTimeHelper::IsCommCtrl6()) {
		if((pNotificationDetails->dwItemType == LVCDI_GROUP) || (pNotificationDetails->dwItemType == LVCDI_ITEM)) {
			itemType = pNotificationDetails->dwItemType;
		}
	}

	if(itemType == LVCDI_ITEM) {
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
					// if we don't set the text back color explicitly, it will always be white
					COLORREF color = static_cast<COLORREF>(attachedControl.SendMessage(LVM_GETTEXTBKCOLOR, 0, 0));
					if(color == CLR_NONE) {
						color = static_cast<COLORREF>(attachedControl.SendMessage(LVM_GETBKCOLOR, 0, 0));
					}
					if(color != CLR_NONE) {
						pNotificationDetails->clrTextBk = color;
					}

					LONG itemID = static_cast<LONG>(pNotificationDetails->nmcd.lItemlParam);
					if(RunTimeHelper::IsCommCtrl6()) {
						itemID = ItemIndexToID(static_cast<int>(pNotificationDetails->nmcd.dwItemSpec));
					}
					LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
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
					return CDRF_DODEFAULT | CDRF_NEWFONT;
					break;
				}
			}
		}
	}

	return CDRF_DODEFAULT;
}

HRESULT ShellListView::OnInternalViewChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	static BOOL recursiveCall = FALSE;
	if(recursiveCall) {
		return S_OK;
	}

	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));

	if(properties.pUnifiedImageList && !properties.displayThumbnails) {
		// set the image list
		recursiveCall = TRUE;
		SetupUnifiedImageListForCurrentView();
		recursiveCall = FALSE;
	}

	if(properties.autoInsertColumns && CurrentViewNeedsColumns()) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETTILEVIEWITEMLINES, 0, reinterpret_cast<LPARAM>(&properties.columnsStatus.maxTileViewColumnCount))));

		if(properties.columnsStatus.loaded == SCLS_NOTLOADED) {
			EnsureShellColumnsAreLoaded(tmTimeOutThreading);
		} else if(IsInDetailsView()) {
			// restore the column visibilities
			for(int realColumnIndex = 0; static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns; ++realColumnIndex) {
				ChangeColumnVisibility(realColumnIndex, properties.columnsStatus.pAllColumns[realColumnIndex].visibleInDetailsView, 0);
			}
		} else if(IsInTilesView()) {
			// the column indexes may have changed, so reset tile view info
			#ifdef USE_STL
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					LONG itemID = iter->first;
					LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
			#else
				POSITION p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					LONG itemID = pPair->m_key;
					LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(pItemDetails->managedProperties & slmipTileViewColumns) {
					int itemIndex = ItemIDToIndex(itemID);
					if(itemIndex != -1) {
						LVTILEINFO tileInfo = {0};
						tileInfo.cbSize = RunTimeHelper::SizeOf_LVTILEINFO();
						tileInfo.iItem = itemIndex;
						tileInfo.cColumns = I_COLUMNSCALLBACK;
						attachedControl.SendMessage(LVM_SETTILEINFO, 0, reinterpret_cast<LPARAM>(&tileInfo));
					}
				}
			}
		}

		if(IsInDetailsView()) {
			if(properties.selectSortColumn || properties.setSortArrows) {
				// set sort arrow
				int columnIndex = RealColumnIndexToIndex(sortingSettings.columnIndex);
				if(columnIndex != -1) {
					if(properties.setSortArrows) {
						int sortOrder = 0/*soAscending*/;
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
						if(RunTimeHelper::IsCommCtrl6()) {
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, sortOrder == 0/*soAscending*/ ? 2/*saUp*/ : 1/*saDown*/)));
						} else {
							sortingSettings.FreeBitmapHandle();
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, sortOrder == 0/*soAscending*/ ? -2 : -1)));
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNBITMAP, columnIndex, reinterpret_cast<LPARAM>(&sortingSettings.headerBitmapHandle))));
						}
					}
					if(properties.selectSortColumn && RunTimeHelper::IsCommCtrl6()) {
						// select column
						attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, columnIndex, 0);
					}
				}
			}
		}
	}
	return S_OK;
}

HRESULT ShellListView::OnInternalGetSubItemControl(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
{
	LPSHLVMGETSUBITEMCONTROLDATA pSubItemControlDetails = reinterpret_cast<LPSHLVMGETSUBITEMCONTROLDATA>(lParam);
	ATLASSERT(pSubItemControlDetails);
	if(!pSubItemControlDetails || pSubItemControlDetails->structSize < sizeof(SHLVMGETSUBITEMCONTROLDATA)) {
		return S_FALSE;
	}

	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		BOOL visualRepresentation = (pSubItemControlDetails->controlKind == 0/*sickVisualRepresentation*/);
		BOOL inPlaceEditing = (pSubItemControlDetails->controlKind == 1/*sickInPlaceEditing*/);
		if(visualRepresentation /*|| inPlaceEditing*/) {
			LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pSubItemControlDetails->itemID);
			if(pItemDetails && (pItemDetails->managedProperties & slmipSubItemControls)) {
				if(properties.columnsStatus.pNamespaceSHF2 && APIWrapper::IsSupported_PSGetPropertyDescription()) {
					UINT realColumnIndex = ColumnIDToRealIndex(pSubItemControlDetails->columnID);
					if(SUCCEEDED(properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, &pSubItemControlDetails->propertyKey))) {
						BOOL handleMessage = visualRepresentation;
						if(!handleMessage && inPlaceEditing) {
							CComPtr<IPropertyStore> pPropertyStore = NULL;
							PCUITEMID_CHILD pIDLRelative = ILFindLastID(pItemDetails->pIDL);
							if(SUCCEEDED(properties.columnsStatus.pNamespaceSHF2->GetUIObjectOf(NULL, 1, &pIDLRelative, IID_IPropertyStore, 0, reinterpret_cast<LPVOID*>(&pPropertyStore))) && pPropertyStore) {
								CComQIPtr<IPropertyStoreCapabilities> pPropertyStoreCaps = pPropertyStore;
								if(pPropertyStoreCaps) {
									handleMessage = (pPropertyStoreCaps->IsPropertyWritable(pSubItemControlDetails->propertyKey) == S_OK);
								}
							}
						}
						if(handleMessage) {
							LPSHELLLISTVIEWSUBITEMDATA pSubItem = NULL;
							#ifdef USE_STL
								std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = pItemDetails->subItems.find(pSubItemControlDetails->columnID);
								if(it != pItemDetails->subItems.cend()) {
									pSubItem = it->second;
								}
							#else
								CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA>::CPair* pPair = pItemDetails->subItems.Lookup(pSubItemControlDetails->columnID);
								if(pPair) {
									pSubItem = pPair->m_value;
								}
							#endif
							pSubItemControlDetails->hasSetValue = FALSE;
							if(pSubItem) {
								if(!pSubItem->pPropertyValue) {
									// property value is still being retrieved
									PropVariantInit(&pSubItemControlDetails->propertyValue);
									/* NOTE: It looks better if we provide a sub-item control in this case, too, because the
									 *       controls are visible immediately. But unfortunately this leeds to two new
									 *       problems:
									 *       1) In Tiles view mode, the width of the sub-item control seems to be calculated
									 *          from the width of the item text, i.e. for items with short names, the
									 *          sub-items will be cut after a few characters ("Timo"/"Sys...").
									 *       2) In Tiles view mode, not all columns that should be displayed actually are
									 *          displayed.
									 */
								} else {
									// we have a property value for this sub-item, so use it
									pSubItemControlDetails->useSubItemControl = TRUE;
									pSubItemControlDetails->hasSetValue = SUCCEEDED(PropVariantCopy(&pSubItemControlDetails->propertyValue, pSubItem->pPropertyValue));
									if(pSubItemControlDetails->hasSetValue) {
										LPSHELLLISTVIEWCOLUMNDATA pColumnData = GetColumnDetails(realColumnIndex);
										switch(pColumnData->displayType) {
											case PDDT_BOOLEAN:
												pSubItemControlDetails->subItemControl = 8/*sicBooleanCheckMark*/;
												break;
											case PDDT_DATETIME:
												pSubItemControlDetails->subItemControl = 9/*sicCalendar*/;
												break;
											case PDDT_ENUMERATED:
												if(pSubItemControlDetails->propertyKey == PKEY_Rating || pSubItemControlDetails->propertyKey == PKEY_SimpleRating) {
													pSubItemControlDetails->subItemControl = 5/*sicRating*/;
												} else {
													pSubItemControlDetails->subItemControl = 11/*sicDropList*/;
												}
												break;
											case PDDT_NUMBER:
												if(pSubItemControlDetails->propertyKey == PKEY_Rating || pSubItemControlDetails->propertyKey == PKEY_SimpleRating) {
													pSubItemControlDetails->subItemControl = 5/*sicRating*/;
												} else if(pSubItemControlDetails->propertyKey == PKEY_PercentFull) {
													pSubItemControlDetails->subItemControl = 3/*sicPercentBar*/;
												} else {
													pSubItemControlDetails->subItemControl = 6/*sicText*/;
												}
												break;
											case PDDT_STRING:
												if(pSubItemControlDetails->propertyKey == PKEY_Link_TargetUrl) {
													pSubItemControlDetails->subItemControl = 12/*sicHyperlink*/;
												} else {
													pSubItemControlDetails->subItemControl = 6/*sicText*/;
												}
												break;
											default:
												pSubItemControlDetails->subItemControl = 6/*sicText*/;
												break;
										}
										if(pSubItemControlDetails->subItemControl == 6/*sicText*/) {
											if(pColumnData->typeFlags & PDTF_MULTIPLEVALUES) {
												//TODO: Currently not supported by ExplorerListView
												//pSubItemControlDetails->subItemControl = 2/*sicMultiValueText*/;
											}
										}
										if(!pColumnData->pPropertyDescription) {
											APIWrapper::PSGetPropertyDescription(pSubItemControlDetails->propertyKey, IID_PPV_ARGS(&pColumnData->pPropertyDescription), NULL);
										}
										if(pColumnData->pPropertyDescription) {
											pColumnData->pPropertyDescription->QueryInterface(IID_PPV_ARGS(&pSubItemControlDetails->pPropertyDescription));
										}
									}
								}
							} else {
								// we don't have an entry for this sub-item, so we do not yet have started a background task to retrieve its value
								pSubItem = new SHELLLISTVIEWSUBITEMDATA();
								pItemDetails->subItems[pSubItemControlDetails->columnID] = pSubItem;

								BOOL isUnreachableNetDrive = FALSE;
								BOOL isSlowColumn = ((properties.columnsStatus.pAllColumns[realColumnIndex].state & SHCOLSTATE_SLOW) == SHCOLSTATE_SLOW);
								BOOL isSlowItem = IsSlowItem(properties.columnsStatus.pNamespaceSHF, ILFindLastID(pItemDetails->pIDL), const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), TRUE, TRUE, &isUnreachableNetDrive);

								if(!(isSlowColumn && isSlowItem && isUnreachableNetDrive)) {
									CComPtr<IRunnableTask> pTask;
									HRESULT hr = ShLvwSubItemControlTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.subItemControlResults, &properties.subItemControlResultsCritSection, pItemDetails->pIDL, pSubItemControlDetails->itemID, pSubItemControlDetails->columnID, realColumnIndex, properties.columnsStatus.pIDLNamespace, &pTask);
									if(SUCCEEDED(hr)) {
										hr = EnqueueTask(pTask, TASKID_ShLvwSubItemControl, 0, TASK_PRIORITY_LV_SUBITEMCONTROLS - (isSlowItem ? 2 : 0));
										ATLASSERT(SUCCEEDED(hr));
									}
								}
							}
						}
					}
				}
			}
		}

		if(!pSubItemControlDetails->hasSetValue) {
			pSubItemControlDetails->useSubItemControl = FALSE;
		}
	#endif
	return S_OK;
}

//HRESULT ShellListView::OnInternalEndSubItemEdit(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
//{
//	LPSHLVMGETSUBITEMCONTROLDATA pSubItemControlDetails = reinterpret_cast<LPSHLVMGETSUBITEMCONTROLDATA>(lParam);
//	ATLASSERT(pSubItemControlDetails);
//	if(!pSubItemControlDetails || pSubItemControlDetails->structSize < sizeof(SHLVMGETSUBITEMCONTROLDATA)) {
//		return S_FALSE;
//	}
//
//	HRESULT hr = S_OK;
//	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
//		LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pSubItemControlDetails->itemID);
//		// TODO: Separate slmip* flag?
//		if(pItemDetails && (pItemDetails->managedProperties & slmipSubItemControls)) {
//			if(properties.columnsStatus.pNamespaceSHF2 && APIWrapper::IsSupported_SHCreateItemFromIDList()) {
//				CComPtr<IShellItem2> pShellItem = NULL;
//				APIWrapper::SHCreateItemFromIDList(pItemDetails->pIDL, IID_PPV_ARGS(&pShellItem), &hr);
//				if(SUCCEEDED(hr) && pShellItem) {
//					CComPtr<IPropertyStore> pPropertyStore = NULL;
//					hr = pShellItem->GetPropertyStoreForKeys(&pSubItemControlDetails->propertyKey, 1, GPS_READWRITE, IID_PPV_ARGS(&pPropertyStore));
//					if(SUCCEEDED(hr) && pPropertyStore) {
//						hr = pPropertyStore->SetValue(pSubItemControlDetails->propertyKey, pSubItemControlDetails->propertyValue);
//						if(SUCCEEDED(hr)) {
//							hr = pPropertyStore->Commit();
//						}
//						if(SUCCEEDED(hr)) {
//							LPPROPVARIANT pPropValue = NULL;
//							#ifdef USE_STL
//								std::unordered_map<LONG, LPPROPVARIANT>::const_iterator it = pItemDetails->properties.find(pSubItemControlDetails->columnID);
//								if(it != pItemDetails->properties.cend()) {
//									pPropValue = it->second;
//								}
//							#else
//								CAtlMap<LONG, LPPROPVARIANT>::CPair* pPair = pItemDetails->properties.Lookup(pSubItemControlDetails->columnID);
//								if(pPair) {
//									pPropValue = pPair->m_value;
//								}
//							#endif
//							LPPROPVARIANT pNewPropValue = new PROPVARIANT;
//							PropVariantInit(pNewPropValue);
//							PropVariantCopy(pNewPropValue, &pSubItemControlDetails->propertyValue);
//							pItemDetails->properties[pSubItemControlDetails->columnID] = pNewPropValue;
//							if(pPropValue) {
//								PropVariantClear(pPropValue);
//								delete pPropValue;
//							}
//						}
//					}
//				}
//			}
//		}
//	#endif
//	return hr;
//}


LRESULT ShellListView::OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/)
{
	NMLVDISPINFO* pNotificationData = reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_GETDISPINFOA) {
			ATLASSERT_POINTER(pNotificationData, NMLVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pNotificationData, NMLVDISPINFOW);
		}
	#endif

	//ATLASSERT(pNotificationData->item.mask & LVIF_PARAM);
	LONG itemID = static_cast<LONG>(pNotificationData->item.lParam);
	if(RunTimeHelper::IsCommCtrl6()) {
		itemID = ItemIndexToID(pNotificationData->item.iItem);
	}
	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	if(!pItemDetails) {
		return 0;
	}
	LONG columnID = (properties.columnsStatus.loaded == SCLS_NOTLOADED ? -1 : ColumnIndexToID(pNotificationData->item.iSubItem));
	int realColumnIndex = -1;
	if(columnID != -1) {
		realColumnIndex = ColumnIDToRealIndex(columnID);
	}

	BOOL needsText = ((pNotificationData->item.mask & LVIF_TEXT) == LVIF_TEXT);
	if(pNotificationData->item.iSubItem == 0) {
		needsText = needsText && ((pItemDetails->managedProperties & slmipText) == slmipText);
		columnID = -1;
	} else {
		if(IsShellColumn(columnID)) {
			needsText = needsText && ((pItemDetails->managedProperties & slmipSubItemsText) == slmipSubItemsText);
		} else {
			columnID = -1;
			needsText = FALSE;
		}
	}
	BOOL needsTileColumns = (((pNotificationData->item.mask & LVIF_COLUMNS) == LVIF_COLUMNS) && ((pItemDetails->managedProperties & slmipTileViewColumns) == slmipTileViewColumns));
	BOOL needsImage = (((pNotificationData->item.mask & LVIF_IMAGE) == LVIF_IMAGE) && ((pItemDetails->managedProperties & slmipIconIndex) == slmipIconIndex));
	BOOL needsAnyImage = needsImage;
	BOOL needsOverlay = needsAnyImage && properties.loadOverlaysOnDemand && ((pItemDetails->managedProperties & slmipOverlayIndex) == slmipOverlayIndex);
	BOOL needsState = needsAnyImage && (properties.hiddenItemsStyle == hisGhostedOnDemand) && ((pItemDetails->managedProperties & slmipGhosted) == slmipGhosted);
	BOOL needsAnything = (needsText || needsTileColumns || needsAnyImage || needsOverlay || needsState);

	if(needsAnything) {
		BOOL dontSetItem = FALSE;
		pNotificationData->item.stateMask = 0;
		pNotificationData->item.state = 0;

		CComPtr<IShellFolder> pParentISF = NULL;
		PCUITEMID_CHILD pRelativePIDL = NULL;
		HRESULT hr = ::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pParentISF);
		ATLASSERT(pRelativePIDL);
		if(SUCCEEDED(hr)) {
			CComQIPtr<IShellIcon> pParentISI;
			CComQIPtr<IShellIconOverlay> pParentISIO;
			SFGAOF attributes = 0;
			if(needsAnyImage) {
				pParentISI = pParentISF;
				attributes |= SFGAO_FOLDER;
			}
			if(needsState) {
				attributes |= SFGAO_GHOSTED;
				pNotificationData->item.mask |= LVIF_STATE;
				pNotificationData->item.stateMask |= LVIS_CUT;
			}
			if(needsOverlay) {
				pParentISIO = pParentISF;
				if(!pParentISIO) {
					attributes |= SFGAO_LINK | SFGAO_SHARE;
				}
				pNotificationData->item.mask |= LVIF_STATE;
				pNotificationData->item.stateMask |= LVIS_OVERLAYMASK;
			}
			if(attributes != 0) {
				hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
				//ATLASSERT(SUCCEEDED(hr));
			}

			// finally retrieve the data
			if(needsText) {
				LPWSTR pDisplayName = NULL;
				if(columnID == -1) {
					hr = GetDisplayName(pItemDetails->pIDL, pParentISF, pRelativePIDL, dntDisplayName, FALSE, &pDisplayName);
					ATLASSERT(SUCCEEDED(hr));
				} else {
					hr = E_FAIL;

					ATLASSERT(properties.columnsStatus.loaded != SCLS_NOTLOADED);
					ATLASSERT(properties.columnsStatus.pAllColumns);
					ATLASSERT(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns);
					BOOL extractSynchronously = TRUE;
					/* TODO: If we ever sort items using sub-item texts, we should set <dontSetItem> to FALSE and
					         extract the text synchronously. */
					BOOL isUnreachableNetDrive = FALSE;
					BOOL isSlowColumn = ((properties.columnsStatus.pAllColumns[realColumnIndex].state & SHCOLSTATE_SLOW) == SHCOLSTATE_SLOW);
					BOOL isSlowItem = IsSlowItem(properties.columnsStatus.pNamespaceSHF, ILFindLastID(pItemDetails->pIDL), const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), TRUE, TRUE, &isUnreachableNetDrive);
					BOOL isSlow = isSlowColumn || isSlowItem;

					// always retrieve sub-item texts on demand, if it is not a slow column/item
					// not storing the data seems to cause high CPU load
					dontSetItem = FALSE;//!isSlow;

					if(properties.useThreadingForSlowColumns && isSlow) {
						// make sure we don't receive another LVN_GETDISPINFO notification for this sub-item
						LVITEM item = {0};
						item.iSubItem = pNotificationData->item.iSubItem;
						attachedControl.SendMessage(LVM_SETITEMTEXT, pNotificationData->item.iItem, reinterpret_cast<LPARAM>(&item));

						#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
							if(pItemDetails->managedProperties & slmipSubItemControls) {
								LPSHELLLISTVIEWSUBITEMDATA pSubItem = NULL;
								#ifdef USE_STL
									std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = pItemDetails->subItems.find(columnID);
									if(it != pItemDetails->subItems.cend()) {
										pSubItem = it->second;
									}
								#else
									CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA>::CPair* pPair = pItemDetails->subItems.Lookup(columnID);
									if(pPair) {
										pSubItem = pPair->m_value;
									}
								#endif
								if(pSubItem) {
									LPPROPVARIANT pOldPropValue = pSubItem->pPropertyValue;
									pSubItem->pPropertyValue = NULL;
									if(pOldPropValue) {
										PropVariantClear(pOldPropValue);
										delete pOldPropValue;
									}
								} else {
									pSubItem = new SHELLLISTVIEWSUBITEMDATA();
									pItemDetails->subItems[columnID] = pSubItem;
								}
							}
						#endif
						if(!(isSlowColumn && isSlowItem && isUnreachableNetDrive)) {
							CComPtr<IRunnableTask> pTask;
							hr = ShLvwSlowColumnTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.slowColumnResults, &properties.slowColumnResultsCritSection, pItemDetails->pIDL, itemID, columnID, realColumnIndex, properties.columnsStatus.pIDLNamespace, &pTask);
							if(SUCCEEDED(hr)) {
								hr = EnqueueTask(pTask, TASKID_ShLvwSlowColumn, 0, TASK_PRIORITY_LV_BACKGROUNDENUMCOLUMNS - (isSlowItem ? 2 : 0));
								ATLASSERT(SUCCEEDED(hr));
							}
						}
						extractSynchronously = FALSE;
					}
					if(extractSynchronously) {
						#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
							LPSHELLLISTVIEWSUBITEMDATA pSubItem = NULL;
							if(pItemDetails->managedProperties & slmipSubItemControls) {
								#ifdef USE_STL
									std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = pItemDetails->subItems.find(columnID);
									if(it != pItemDetails->subItems.cend()) {
										pSubItem = it->second;
									}
								#else
									CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA>::CPair* pPair = pItemDetails->subItems.Lookup(columnID);
									if(pPair) {
										pSubItem = pPair->m_value;
									}
								#endif
								if(pSubItem) {
									LPPROPVARIANT pOldPropValue = pSubItem->pPropertyValue;
									pSubItem->pPropertyValue = NULL;
									if(pOldPropValue) {
										PropVariantClear(pOldPropValue);
										delete pOldPropValue;
									}
								} else {
									pSubItem = new SHELLLISTVIEWSUBITEMDATA();
									pItemDetails->subItems[columnID] = pSubItem;
								}
							}
						#endif
						if(isSlowColumn && isSlowItem) {
							// leave this column empty
							LVITEM item = {0};
							item.iSubItem = pNotificationData->item.iSubItem;
							attachedControl.SendMessage(LVM_SETITEMTEXT, pNotificationData->item.iItem, reinterpret_cast<LPARAM>(&item));
						} else {
							// extract the sub-item text here and now
							SHELLDETAILS column = {0};
							if(properties.columnsStatus.pNamespaceSHF2) {
								hr = properties.columnsStatus.pNamespaceSHF2->GetDetailsOf(ILFindLastID(pItemDetails->pIDL), realColumnIndex, &column);
							}
							if(FAILED(hr) && properties.columnsStatus.pNamespaceSHD) {
								hr = properties.columnsStatus.pNamespaceSHD->GetDetailsOf(ILFindLastID(pItemDetails->pIDL), realColumnIndex, &column);
							}
							if(SUCCEEDED(hr)) {
								ATLVERIFY(SUCCEEDED(StrRetToStrW(&column.str, ILFindLastID(pItemDetails->pIDL), &pDisplayName)));
							}

							#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
								if(pSubItem) {
									CComPtr<IRunnableTask> pTask;
									hr = ShLvwSubItemControlTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.subItemControlResults, &properties.subItemControlResultsCritSection, pItemDetails->pIDL, itemID, columnID, realColumnIndex, properties.columnsStatus.pIDLNamespace, &pTask);
									if(SUCCEEDED(hr)) {
										hr = EnqueueTask(pTask, TASKID_ShLvwSubItemControl, 0, TASK_PRIORITY_LV_SUBITEMCONTROLS - (isSlowItem ? 2 : 0));
										ATLASSERT(SUCCEEDED(hr));
									}
								}
							#endif
						}
					}
				}
				if(pDisplayName) {
					// NOTE: We'll have a small problem here if pDisplayName is shorter and not null-terminated
					if(pNotificationDetails->code == LVN_GETDISPINFOA) {
						CW2A converter(pDisplayName);
						LPSTR pBuffer = converter;
						lstrcpynA(reinterpret_cast<LPNMLVDISPINFOA>(pNotificationData)->item.pszText, pBuffer, pNotificationData->item.cchTextMax);
					} else {
						lstrcpynW(reinterpret_cast<LPNMLVDISPINFOW>(pNotificationData)->item.pszText, pDisplayName, pNotificationData->item.cchTextMax);
					}
					CoTaskMemFree(pDisplayName);
				}
			}

			if(needsImage) {
				if(properties.pUnifiedImageList || properties.displayThumbnails) {
					pNotificationData->item.iImage = itemID;
				} else {
					if(attributes & SFGAO_FOLDER) {
						pNotificationData->item.iImage = properties.DEFAULTICON_FOLDER;
					} else {
						pNotificationData->item.iImage = properties.DEFAULTICON_FILE;
					}
				}
			}
			if(needsAnyImage) {
				if(properties.displayThumbnails && properties.thumbnailsStatus.pThumbnailImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, (attributes & SFGAO_FOLDER ? properties.DEFAULTICON_FOLDER : properties.DEFAULTICON_FILE), 0, NULL, 0)));
				}
				if(properties.pUnifiedImageList) {
					CComPtr<IImageListPrivate> pImgLstPriv;
					properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
					ATLASSUME(pImgLstPriv);
					ATLVERIFY(SUCCEEDED(pImgLstPriv->SetIconInfo(itemID, SIIF_SYSTEMICON, NULL, (attributes & SFGAO_FOLDER ? properties.DEFAULTICON_FOLDER : properties.DEFAULTICON_FILE), 0, NULL, 0)));
				}
				ATLVERIFY(SUCCEEDED(GetIconAsync(pItemDetails->pIDL, itemID, pParentISI)));
			}

			if(needsState) {
				if(attributes & SFGAO_GHOSTED) {
					pNotificationData->item.state |= LVIS_CUT;
				}
			}

			if(needsOverlay) {
				int overlayIndex = OI_ASYNC;
				if(pParentISIO) {
					hr = pParentISIO->GetOverlayIndex(pRelativePIDL, &overlayIndex);
					if(hr == E_PENDING) {
						ATLVERIFY(SUCCEEDED(GetOverlayAsync(pItemDetails->pIDL, itemID, pParentISIO)));
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

			if(needsTileColumns) {
				// start a background task that sets puColumns to the correct shell columns
				// use the TileViewItemLines property as maximum (cache it somewhere)
				CComPtr<IRunnableTask> pTask;
				if(RunTimeHelper::IsVista()) {
					if(cachedSettings.viewMode == 5/*vExtendedTiles*/) {
						hr = ShLvwTileViewTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.backgroundTileViewResults, &properties.backgroundTileViewResultsCritSection, properties.columnsStatus.hColumnsReadyEvent, pItemDetails->pIDL, itemID, properties.columnsStatus.pIDLNamespace, properties.columnsStatus.maxTileViewColumnCount, PKEY_PropList_ExtendedTileInfo, &pTask);
					} else {
						hr = ShLvwTileViewTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.backgroundTileViewResults, &properties.backgroundTileViewResultsCritSection, properties.columnsStatus.hColumnsReadyEvent, pItemDetails->pIDL, itemID, properties.columnsStatus.pIDLNamespace, properties.columnsStatus.maxTileViewColumnCount, PKEY_PropList_TileInfo, &pTask);
					}
				} else {
					hr = ShLvwLegacyTileViewTask::CreateInstance(attachedControl, GethWndShellUIParentWindow(), &properties.backgroundTileViewResults, &properties.backgroundTileViewResultsCritSection, properties.columnsStatus.hColumnsReadyEvent, pItemDetails->pIDL, itemID, properties.columnsStatus.pIDLNamespace, properties.columnsStatus.maxTileViewColumnCount, &pTask);
				}
				if(SUCCEEDED(hr)) {
					BOOL isSlowItem = IsSlowItem(properties.columnsStatus.pNamespaceSHF, ILFindLastID(pItemDetails->pIDL), const_cast<PIDLIST_ABSOLUTE>(pItemDetails->pIDL), TRUE, TRUE);
					hr = EnqueueTask(pTask, TASKID_ShLvwTileView, 0, TASK_PRIORITY_LV_GET_TILEVIEWCOLUMNS - (isSlowItem ? 2 : 0));
					ATLASSERT(SUCCEEDED(hr));
				}
				pNotificationData->item.cColumns = 0;
			}
		}

		if(!dontSetItem) {
			pNotificationData->item.mask |= LVIF_DI_SETITEM;
		}
	}

	return 0;
}


inline HRESULT ShellListView::Raise_ChangedColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL hasBecomeVisible)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedColumnVisibility(pColumn, hasBecomeVisible);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, VARIANT_BOOL isShellContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedContextMenuSelection(hContextMenu, isShellContextMenu, commandID, isCustomMenuItem);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangedItemPIDL(LONG itemID, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<IShListViewItem> pItem;
		ClassFactory::InitShellListItem(itemID, this, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		return Fire_ChangedItemPIDL(pItem, *reinterpret_cast<LONG*>(&previousPIDL), *reinterpret_cast<LONG*>(&newPIDL));
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangedNamespacePIDL(IShListViewNamespace* pNamespace, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedNamespacePIDL(pNamespace, *reinterpret_cast<LONG*>(&previousPIDL), *reinterpret_cast<LONG*>(&newPIDL));
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangedSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedSortColumn(previousSortColumnIndex, newSortColumnIndex);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangingColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL willBecomeVisible, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingColumnVisibility(pColumn, willBecomeVisible, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ChangingSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingSortColumn(previousSortColumnIndex, newSortColumnIndex, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ColumnEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnEnumerationCompleted(pNamespace, aborted);
	//}
	//return S_OK;
}

HRESULT ShellListView::Raise_ColumnEnumerationStarted(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnEnumerationStarted(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ColumnEnumerationTimedOut(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnEnumerationTimedOut(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_CreatedHeaderContextMenu(OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedHeaderContextMenu(hContextMenu, minimumCustomCommandID, pCancelMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_CreatedShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedShellContextMenu(pItems, hContextMenu, minimumCustomCommandID, pCancelMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_CreatingShellContextMenu(LPDISPATCH pItems, ShellContextMenuStyleConstants* pContextMenuStyle, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CreatingShellContextMenu(pItems, pContextMenuStyle, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_DestroyingHeaderContextMenu(OLE_HANDLE hContextMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyingHeaderContextMenu(hContextMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_DestroyingShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyingShellContextMenu(pItems, hContextMenu);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_InsertedItem(IShListViewItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_InsertedNamespace(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedNamespace(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_InsertingItem(LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(fullyQualifiedPIDL, pCancelInsertion);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_InvokedHeaderContextMenuCommand(OLE_HANDLE hContextMenu, LONG commandID)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InvokedHeaderContextMenuCommand(hContextMenu, commandID);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_InvokedShellContextMenuCommand(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants usedWindowMode, CommandInvocationFlagsConstants usedInvocationFlags)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InvokedShellContextMenuCommand(pItems, hContextMenu, commandID, usedWindowMode, usedInvocationFlags);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ItemEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationCompleted(pNamespace, aborted);
	//}
	//return S_OK;
}

HRESULT ShellListView::Raise_ItemEnumerationStarted(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationStarted(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_ItemEnumerationTimedOut(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemEnumerationTimedOut(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_LoadedColumn(IShListViewColumn* pColumn, VARIANT_BOOL* pMakeVisible)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_LoadedColumn(pColumn, pMakeVisible);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_RemovingItem(IShListViewItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_RemovingNamespace(IShListViewNamespace* pNamespace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingNamespace(pNamespace);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_SelectedHeaderContextMenuItem(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL* pExecuteCommand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedHeaderContextMenuItem(hContextMenu, commandID, pExecuteCommand);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_SelectedShellContextMenuItem(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants* pWindowModeToUse, CommandInvocationFlagsConstants* pInvocationFlagsToUse, VARIANT_BOOL* pExecuteCommand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedShellContextMenuItem(pItems, hContextMenu, commandID, pWindowModeToUse, pInvocationFlagsToUse, pExecuteCommand);
	//}
	//return S_OK;
}

inline HRESULT ShellListView::Raise_UnloadedColumn(IShListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_UnloadedColumn(pColumn);
	//}
	//return S_OK;
}


void ShellListView::ConfigureControl(void)
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

	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETTILEVIEWITEMLINES, 0, reinterpret_cast<LPARAM>(&properties.columnsStatus.maxTileViewColumnCount))));
	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	properties.columnsStatus.averageCharWidth = GetAverageColumnHeaderCharWidth();
	flags.acceptNewTasks = TRUE;
	SetSystemImageLists();
	RegisterForShellNotifications();
	attachedControl.SetTimer(timers.ID_BACKGROUNDENUMTIMEOUT, timers.INT_BACKGROUNDENUMTIMEOUT);
}

SIZE ShellListView::GetThumbnailsSize(long userSetSize)
{
	SIZE thumbnailSize = {userSetSize, userSetSize};
	BOOL square = (thumbnailSize.cx == thumbnailSize.cy);
	PLONG p = reinterpret_cast<PLONG>(&thumbnailSize);
	for(int i = 0; i < 2; ++i) {
		switch(p[i]) {
			case -1:
			{
				DWORD thumbsSize = 96;
				DWORD dataSize = sizeof(thumbsSize);
				SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer"), TEXT("ThumbnailSize"), NULL, &thumbsSize, &dataSize);
				p[i] = thumbsSize;
				break;
			}
			case -2:
			case -3:
			case -4:
			case -5:
			{
				HIMAGELIST hImageList = NULL;
				if(APIWrapper::IsSupported_SHGetImageList()) {
					APIWrapper::SHGetImageList(p[i] == -2 ? SHIL_SMALL : (p[i] == -3 ? SHIL_LARGE : (p[i] == -4 ? SHIL_EXTRALARGE : SHIL_JUMBO)), IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
				} else if(p[i] != -4 && p[i] != -5) {
					SHFILEINFO details = {0};
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | (p[i] == -2 ? SHGFI_SMALLICON : SHGFI_LARGEICON) | SHGFI_USEFILEATTRIBUTES));
				}
				SIZE sz;
				if(hImageList) {
					ImageList_GetIconSize(hImageList, reinterpret_cast<PINT>(&sz.cx), reinterpret_cast<PINT>(&sz.cy));
				} else {
					switch(p[i]) {
						case -2:
							sz.cx = sz.cy = 16;
							break;
						case -3:
							sz.cx = sz.cy = 32;
							break;
						case -4:
							sz.cx = sz.cy = 48;
							break;
						case -5:
							sz.cx = sz.cy = 256;
							break;
					}
				}
				p[i] = reinterpret_cast<PLONG>(&sz)[i];
				break;
			}
		}
		if(square && p[0] < 0) {
			p[1] = p[0];
			break;
		}
	}
	return thumbnailSize;
}

void ShellListView::SetSystemImageLists(void)
{
	if(properties.pAttachedInternalMessageListener) {
		if(properties.displayThumbnails) {
			BOOL flushIcons = TRUE;
			if(!properties.thumbnailsStatus.pThumbnailImageList) {
				AggregateImageList_CreateInstance(&properties.thumbnailsStatus.pThumbnailImageList);
				// TODO: Isn't the next line the reason why we need to call FlushIcons in put_DisplayThumbnails?
				flushIcons = FALSE;
			}
			ATLASSUME(properties.thumbnailsStatus.pThumbnailImageList);
			properties.thumbnailsStatus.thumbnailSize = GetThumbnailsSize(properties.thumbnailSize);
			HRESULT hr = properties.thumbnailsStatus.pThumbnailImageList->SetIconSize(properties.thumbnailsStatus.thumbnailSize.cx, properties.thumbnailsStatus.thumbnailSize.cy);
			if(hr == E_NOTIMPL && flushIcons) {
				// we tried to resize the icons in an existing image list and this is not supported on this system
				properties.thumbnailsStatus.pThumbnailImageList = NULL;
				AggregateImageList_CreateInstance(&properties.thumbnailsStatus.pThumbnailImageList);
				ATLASSUME(properties.thumbnailsStatus.pThumbnailImageList);
				properties.thumbnailsStatus.pThumbnailImageList->SetIconSize(properties.thumbnailsStatus.thumbnailSize.cx, properties.thumbnailsStatus.thumbnailSize.cy);
				// NOTE: We flush the icons nevertheless!
			}

			CComPtr<IImageListPrivate> pImgLstPriv;
			properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
			ATLASSUME(pImgLstPriv);

			if(properties.pUnifiedImageList) {
				// transfer the settings for any non-shell items
				CComPtr<IImageListPrivate> pImgLstPrivUnified;
				properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPrivUnified));
				ATLASSUME(pImgLstPrivUnified);
				pImgLstPrivUnified->TransferNonShellItems(pImgLstPriv);
				properties.pUnifiedImageList = NULL;
			}

			UINT imageListFlags = (properties.displayElevationShieldOverlays ? ILOF_DISPLAYELEVATIONSHIELDS : 0);
			if((attachedControl.SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_BORDERSELECT) || !RunTimeHelper::IsCommCtrl6()) {
				imageListFlags |= ILOF_IGNOREEXTRAALPHA;
			}
			pImgLstPriv->SetOptions(0, imageListFlags);

			if(APIWrapper::IsSupported_SHGetImageList()) {
				HIMAGELIST hLargerImageList = NULL;
				HIMAGELIST hSmallerImageList1 = NULL;
				HIMAGELIST hSmallerImageList2 = NULL;
				if(properties.thumbnailsStatus.thumbnailSize.cx <= 48) {
					if(properties.thumbnailsStatus.thumbnailSize.cx <= 16) {
						APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hLargerImageList), NULL);
						hSmallerImageList2 = hLargerImageList;
						hSmallerImageList1 = hSmallerImageList2;
					} else {
						APIWrapper::SHGetImageList(SHIL_LARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hSmallerImageList2), NULL);
						if(properties.thumbnailsStatus.thumbnailSize.cx <= 32) {
							hLargerImageList = hSmallerImageList2;
						} else {
							APIWrapper::SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hLargerImageList), NULL);
						}
						if(!RunTimeHelper::IsCommCtrl6()) {
							APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hSmallerImageList1), NULL);
						} else {
							hSmallerImageList1 = hSmallerImageList2;
						}
					}
				} else {
					APIWrapper::SHGetImageList(SHIL_JUMBO, IID_IImageList, reinterpret_cast<LPVOID*>(&hLargerImageList), NULL);
					APIWrapper::SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hSmallerImageList2), NULL);
					if(!hLargerImageList) {
						hLargerImageList = hSmallerImageList2;
					}
					if(!RunTimeHelper::IsCommCtrl6() && properties.thumbnailsStatus.thumbnailSize.cx <= 96) {
						APIWrapper::SHGetImageList(SHIL_LARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hSmallerImageList1), NULL);
					} else {
						hSmallerImageList1 = hSmallerImageList2;
					}
				}
				pImgLstPriv->SetImageList(AIL_SHELLITEMS, hLargerImageList, NULL);
				pImgLstPriv->SetImageList(AIL_EXECUTABLEOVERLAYS, hSmallerImageList1, NULL);
				pImgLstPriv->SetImageList(AIL_OVERLAYS, hSmallerImageList2, NULL);
				pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
			} else {
				SHFILEINFO details = {0};
				HIMAGELIST hImageList = NULL;
				if(properties.thumbnailsStatus.thumbnailSize.cx <= 16) {
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
				} else {
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES));
				}
				pImgLstPriv->SetImageList(AIL_SHELLITEMS, hImageList, NULL);
				pImgLstPriv->SetImageList(AIL_OVERLAYS, hImageList, NULL);
				if(properties.thumbnailsStatus.thumbnailSize.cx <= 48) {
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
				}
				pImgLstPriv->SetImageList(AIL_EXECUTABLEOVERLAYS, hImageList, NULL);
				pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
			}

			HIMAGELIST hImageList = IImageListToHIMAGELIST(properties.thumbnailsStatus.pThumbnailImageList);
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, reinterpret_cast<LPARAM>(hImageList))));
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, reinterpret_cast<LPARAM>(hImageList))));
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, reinterpret_cast<LPARAM>(hImageList))));

			if(!RunTimeHelper::IsCommCtrl6()) {
				// re-arrange so that the item positions are updated
				// TODO: Is this really a good idea?
				attachedControl.SendMessage(LVM_ARRANGE, LVA_SNAPTOGRID, 0);
			}

			if(flushIcons) {
				FlushIcons(-1, TRUE, TRUE);
			}
		} else {
			if(TRUE) {
				if(!properties.pUnifiedImageList) {
					UnifiedImageList_CreateInstance(&properties.pUnifiedImageList);
				}
				ATLASSUME(properties.pUnifiedImageList);
				CComPtr<IImageListPrivate> pImgLstPriv;
				properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
				ATLASSUME(pImgLstPriv);
				if(properties.thumbnailsStatus.pThumbnailImageList) {
					// transfer the settings for any non-shell items
					CComPtr<IImageListPrivate> pImgLstPrivThumbnails;
					properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPrivThumbnails));
					ATLASSUME(pImgLstPrivThumbnails);
					pImgLstPrivThumbnails->TransferNonShellItems(pImgLstPriv);
					properties.thumbnailsStatus.pThumbnailImageList = NULL;
				}

				UINT imageListFlags = (properties.displayElevationShieldOverlays ? ILOF_DISPLAYELEVATIONSHIELDS : 0);
				if(!RunTimeHelper::IsCommCtrl6()) {
					imageListFlags |= ILOF_IGNOREEXTRAALPHA;
				}
				pImgLstPriv->SetOptions(0, imageListFlags);
				pImgLstPriv->SetImageList(AIL_NONSHELLITEMS, properties.hImageList[0], NULL);
				ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
				SetupUnifiedImageListForCurrentView();

				HIMAGELIST hImageList = IImageListToHIMAGELIST(properties.pUnifiedImageList);
				if(properties.useSystemImageList & usilSmallImageList) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, reinterpret_cast<LPARAM>(hImageList))));
				}
				if(properties.useSystemImageList & usilLargeImageList) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, reinterpret_cast<LPARAM>(hImageList))));
				}
				if(properties.useSystemImageList & usilExtraLargeImageList) {
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, reinterpret_cast<LPARAM>(hImageList))));
				}
			} else {
				// don't use UnifiedImageList
				properties.thumbnailsStatus.pThumbnailImageList = NULL;
				properties.pUnifiedImageList = NULL;
				if(properties.useSystemImageList & usilSmallImageList) {
					HIMAGELIST hImageList = NULL;
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					} else {
						SHFILEINFO details = {0};
						hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
					}
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, reinterpret_cast<LPARAM>(hImageList))));
				}
				if(properties.useSystemImageList & usilLargeImageList) {
					HIMAGELIST hImageList = NULL;
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_LARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					} else {
						SHFILEINFO details = {0};
						hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES));
					}
					ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, reinterpret_cast<LPARAM>(hImageList))));
				}
				if(properties.useSystemImageList & usilExtraLargeImageList) {
					if(APIWrapper::IsSupported_SHGetImageList()) {
						HIMAGELIST hImageList = NULL;
						APIWrapper::SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, reinterpret_cast<LPARAM>(hImageList))));
					}
				}
			}
		}
	}
}

void ShellListView::SetupUnifiedImageListForCurrentView(void)
{
	CComPtr<IImageListPrivate> pImgLstPriv;
	properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv));
	ATLASSUME(pImgLstPriv);

	HIMAGELIST hImageList = NULL;
	switch(cachedSettings.viewMode) {
		case 0/*vIcons*/:
			if(properties.useSystemImageList & usilLargeImageList) {
				if(APIWrapper::IsSupported_SHGetImageList()) {
					APIWrapper::SHGetImageList(SHIL_LARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
				} else {
					SHFILEINFO details = {0};
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES));
				}
			}
			break;
		case 1/*vSmallIcons*/:
		case 2/*vList*/:
		case 3/*vDetails*/:
			if(properties.useSystemImageList & usilSmallImageList) {
				if(APIWrapper::IsSupported_SHGetImageList()) {
					APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
				} else {
					SHFILEINFO details = {0};
					hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
				}
			}
			break;
		case 4/*vTiles*/:
		case 5/*vExtendedTiles*/:
			if(properties.useSystemImageList & usilExtraLargeImageList) {
				if(APIWrapper::IsSupported_SHGetImageList()) {
					APIWrapper::SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
				}
			}
			break;
	}
	if(hImageList) {
		pImgLstPriv->SetImageList(AIL_SHELLITEMS, hImageList, NULL);
		// SysListView32 caches things like the icon size, so force a refresh
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 0/*ilCurrentView*/, NULL)));
		hImageList = IImageListToHIMAGELIST(properties.pUnifiedImageList);
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 0/*ilCurrentView*/, reinterpret_cast<LPARAM>(hImageList))));
	}
}


void ShellListView::AddItem(LONG itemID, LPSHELLLISTVIEWITEMDATA pItemData)
{
	ATLASSERT(itemID != -1);
	ATLASSERT_POINTER(pItemData, SHELLLISTVIEWITEMDATA);

	properties.items[itemID] = pItemData;

	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		CComPtr<IShListViewItem> pItem;
		ClassFactory::InitShellListItem(itemID, this, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
		Raise_InsertedItem(pItem);
	}
}

HRESULT ShellListView::ApplyManagedProperties(LONG itemID)
{
	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	if(!pItemDetails) {
		return E_INVALIDARG;
	}

	CComPtr<INamespaceEnumSettings> pEnumSettings;
	if(pItemDetails->pIDLNamespace) {
		// return the namespace's enum settings
		CComObject<ShellListViewNamespace>* pNamespaceObj = NamespacePIDLToNamespace(pItemDetails->pIDLNamespace, TRUE);
		ATLASSUME(pNamespaceObj);
		ATLVERIFY(SUCCEEDED(pNamespaceObj->get_NamespaceEnumSettings(&pEnumSettings)));
	} else {
		// return the default enum settings
		ATLVERIFY(SUCCEEDED(get_DefaultNamespaceEnumSettings(&pEnumSettings)));
	}
	CachedEnumSettings enumSettings = CacheEnumSettings(pEnumSettings);
	if(!ShouldShowItem(pItemDetails->pIDL, &enumSettings)) {
		// remove the item
		int itemIndex = ItemIDToIndex(itemID);
		if(itemIndex >= 0) {
			attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
		}
		return S_OK;
	}

	if(pItemDetails->managedProperties != 0) {
		// update the managed properties
		HRESULT hr = E_FAIL;

		LONG itemIndex = ItemIDToIndex(itemID);
		if(itemIndex != -1 && attachedControl.IsWindow()) {
			// retrieve the current settings
			LVITEM item = {0};
			item.iItem = itemIndex;

			BOOL retrieveStateNow = ((properties.hiddenItemsStyle == hisGhosted) && ((pItemDetails->managedProperties & slmipGhosted) == slmipGhosted));
			retrieveStateNow = retrieveStateNow || (!properties.loadOverlaysOnDemand && ((pItemDetails->managedProperties & slmipOverlayIndex) == slmipOverlayIndex));
			// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
			retrieveStateNow = retrieveStateNow || !(pItemDetails->managedProperties & slmipIconIndex);
			HIMAGELIST h = NULL;
			ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETIMAGELIST, 0/*ilCurrentView*/, reinterpret_cast<LPARAM>(&h))));
			retrieveStateNow = retrieveStateNow || !h;

			CComPtr<IShellFolder> pParentISF = NULL;
			PCUITEMID_CHILD pRelativePIDL = NULL;
			if(retrieveStateNow) {
				ATLVERIFY(SUCCEEDED(::SHBindToParent(pItemDetails->pIDL, IID_PPV_ARGS(&pParentISF), &pRelativePIDL)));
				ATLASSUME(pParentISF);
				ATLASSERT_POINTER(pRelativePIDL, ITEMID_CHILD);
			}

			if((pItemDetails->managedProperties & slmipGhosted) == slmipGhosted) {
				if(properties.hiddenItemsStyle == hisGhosted || (properties.hiddenItemsStyle == hisGhostedOnDemand && retrieveStateNow)) {
					// set the 'cut' state here and now
					item.mask |= LVIF_STATE;
					item.stateMask |= LVIS_CUT;

					if(pParentISF && pRelativePIDL) {
						SFGAOF attributes = SFGAO_GHOSTED;
						hr = pParentISF->GetAttributesOf(1, &pRelativePIDL, &attributes);
						ATLASSERT(SUCCEEDED(hr));
						if(SUCCEEDED(hr)) {
							if(attributes & SFGAO_GHOSTED) {
								item.state |= LVIS_CUT;
							}
						}
					}
				}
			}
			if(pItemDetails->managedProperties & slmipIconIndex) {
				item.iImage = I_IMAGECALLBACK;
				item.mask |= LVIF_IMAGE;
			}
			if(pItemDetails->managedProperties & slmipText) {
				item.pszText = LPSTR_TEXTCALLBACK;
				item.mask |= LVIF_TEXT;
			}
			if((!properties.loadOverlaysOnDemand || retrieveStateNow) && ((pItemDetails->managedProperties & slmipOverlayIndex) == slmipOverlayIndex)) {
				// load the overlay here and now
				item.stateMask |= LVIS_OVERLAYMASK;
				item.mask |= LVIF_STATE;

				if(pParentISF && pRelativePIDL) {
					int overlayIndex = GetOverlayIndex(pParentISF, pRelativePIDL);
					if(overlayIndex > 0) {
						item.state |= INDEXTOOVERLAYMASK(overlayIndex);
					}
				}
			}

			// apply the new settings
			if(attachedControl.SendMessage(LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
				hr = S_OK;
			}

			if(SUCCEEDED(hr)) {
				#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
					if(pItemDetails->managedProperties & slmipSubItemControls) {
						#ifdef USE_STL
							for(std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator iter = pItemDetails->subItems.cbegin(); iter != pItemDetails->subItems.cend(); ++iter) {
								if(iter->second) {
									delete iter->second;
								}
							}
							pItemDetails->subItems.clear();
						#else
							POSITION p = pItemDetails->subItems.GetStartPosition();
							while(p) {
								LPSHELLLISTVIEWSUBITEMDATA pSubItem = pItemDetails->subItems.GetValueAt(p);
								if(pSubItem) {
									delete pSubItem;
								}
								pItemDetails->subItems.GetNextValue(p);
							}
							pItemDetails->subItems.RemoveAll();
						#endif
					}
				#endif
			}

			if(SUCCEEDED(hr) && RunTimeHelper::IsCommCtrl6()) {
				if(pItemDetails->managedProperties & slmipTileViewColumns) {
					LVTILEINFO tileInfo = {0};
					tileInfo.cbSize = RunTimeHelper::SizeOf_LVTILEINFO();
					tileInfo.iItem = itemIndex;
					tileInfo.cColumns = I_COLUMNSCALLBACK;
					if(!attachedControl.SendMessage(LVM_SETTILEINFO, 0, reinterpret_cast<LPARAM>(&tileInfo))) {
						hr = E_FAIL;
					}
				}
			}
			if(SUCCEEDED(hr)) {
				if(pItemDetails->managedProperties & slmipSubItemsText) {
					item.mask = LVIF_TEXT;
					item.pszText = LPSTR_TEXTCALLBACK;
					item.cchTextMax = 0;
					#ifdef USE_STL
						for(std::unordered_map<LONG, int>::const_iterator iter = properties.visibleColumns.cbegin(); iter != properties.visibleColumns.cend(); ++iter) {
							LONG columnID = iter->first;
					#else
						POSITION p = properties.visibleColumns.GetStartPosition();
						while(p) {
							LONG columnID = properties.visibleColumns.GetNextKey(p);
					#endif
						item.iSubItem = ColumnIDToIndex(columnID);
						if(!attachedControl.SendMessage(LVM_SETITEMTEXT, itemIndex, reinterpret_cast<LPARAM>(&item))) {
							hr = E_FAIL;
							break;
						}
					}
				}
			}
		}
		return hr;
	}

	return S_OK;
}

void ShellListView::BufferShellItemData(LPSHELLLISTVIEWITEMDATA pItemData)
{
	ATLASSERT_POINTER(pItemData, SHELLLISTVIEWITEMDATA);

	#ifdef USE_STL
		properties.shellItemDataOfInsertedItems.push(pItemData);
	#else
		properties.shellItemDataOfInsertedItems.AddTail(pItemData);
	#endif
}

LONG ShellListView::FastInsertShellItem(PIDLIST_ABSOLUTE pIDL, LONG insertAt)
{
	properties.itemDataForFastInsertion.insertAt = insertAt;
	properties.itemDataForFastInsertion.isGhosted = VARIANT_FALSE;
	properties.itemDataForFastInsertion.overlayIndex = 0;

	HIMAGELIST h = NULL;
	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETIMAGELIST, 0/*ilCurrentView*/, reinterpret_cast<LPARAM>(&h))));
	BOOL noImages = !h;
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

	LPSHELLLISTVIEWITEMDATA pItemData = new SHELLLISTVIEWITEMDATA();
	pItemData->pIDL = pIDL;
	pItemData->managedProperties = properties.defaultManagedItemProperties;
	BufferShellItemData(pItemData);

	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ADDITEM, TRUE, reinterpret_cast<LPARAM>(&properties.itemDataForFastInsertion));
	ATLASSERT(SUCCEEDED(hr));
	if(SUCCEEDED(hr)) {
		return properties.itemDataForFastInsertion.insertedItemID;
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
	return -1;
}

LPSHELLLISTVIEWITEMDATA ShellListView::GetItemDetails(LONG itemID)
{
	LPSHELLLISTVIEWITEMDATA pItemData = NULL;
	//__try {
		#ifdef USE_STL
			std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.find(itemID);
			if(iter != properties.items.cend()) {
				pItemData = iter->second;
			}
		#else
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemID);
			if(pPair) {
				pItemData = pPair->m_value;
			}
		#endif
	//} __except(GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION ? EXCEPTION_EXECUTE_HANDLER : EXCEPTION_CONTINUE_SEARCH) {
	//	pItemData = NULL;
	//}
	return pItemData;
}

BOOL ShellListView::IsShellItem(LONG itemID)
{
	#ifdef USE_STL
		std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.find(itemID);
		if(iter != properties.items.cend()) {
			return TRUE;
		}
	#else
		CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemID);
		if(pPair) {
			return TRUE;
		}
	#endif
	return FALSE;
}

int ShellListView::ItemIDToIndex(LONG itemID)
{
	int itemIndex = -1;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ITEMIDTOINDEX, itemID, reinterpret_cast<LPARAM>(&itemIndex));
	return (SUCCEEDED(hr) ? itemIndex : -1);
}

LONG ShellListView::ItemIndexToID(int itemIndex)
{
	LONG itemID = -1;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ITEMINDEXTOID, itemIndex, reinterpret_cast<LPARAM>(&itemID));
	return (SUCCEEDED(hr) ? itemID : -1);
}

PCIDLIST_ABSOLUTE ShellListView::ItemIDToPIDL(LONG itemID)
{
	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	return (pItemDetails ? pItemDetails->pIDL : NULL);
}

UINT ShellListView::ItemIDsToPIDLs(PLONG pItemIDs, UINT itemCount, PIDLIST_ABSOLUTE*& ppIDLs, BOOL keepNonShellItems)
{
	ATLASSERT_ARRAYPOINTER(pItemIDs, LONG, itemCount);
	if(!pItemIDs) {
		ppIDLs = NULL;
		return 0;
	}

	ppIDLs = new PIDLIST_ABSOLUTE[itemCount];
	UINT pIDLCount = 0;
	for(UINT item = 0; item < itemCount; ++item) {
		LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pItemIDs[item]);
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

CComObject<ShellListViewNamespace>* ShellListView::ItemIDToNamespace(LONG itemID)
{
	return NamespacePIDLToNamespace(ItemIDToNamespacePIDL(itemID), TRUE);
}

PCIDLIST_ABSOLUTE ShellListView::ItemIDToNamespacePIDL(LONG itemID)
{
	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	return (pItemDetails ? pItemDetails->pIDLNamespace : NULL);
}

LONG ShellListView::PIDLToItemID(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	#ifdef USE_STL
		if(exactMatch) {
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(iter->second->pIDL == pIDL) {
					return iter->first;
				}
			}
		} else {
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				if(ILIsEqual(iter->second->pIDL, pIDL)) {
					return iter->first;
				}
			}
		}
	#else
		if(exactMatch) {
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(pPair->m_value->pIDL == pIDL) {
					return pPair->m_key;
				}
				properties.items.GetNext(p);
			}
		} else {
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				if(ILIsEqual(pPair->m_value->pIDL, pIDL)) {
					return pPair->m_key;
				}
				properties.items.GetNext(p);
			}
		}
	#endif
	return -1;
}

LONG ShellListView::ParsingNameToItemID(LPWSTR pParsingName)
{
	#ifdef USE_STL
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
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
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
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
	return -1;
}

HRESULT ShellListView::VariantToItemIDs(VARIANT& items, PLONG& pItemIDs, UINT& itemCount)
{
	HRESULT hr = S_OK;
	itemCount = 0;
	pItemIDs = NULL;
	switch(items.vt) {
		case VT_DISPATCH:
			// invalid arg - VB runtime error 380
			hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			if(items.pdispVal) {
				CComQIPtr<IShListViewItem> pShLvwItem = items.pdispVal;
				if(pShLvwItem) {
					// a single ShellListViewItem object
					pItemIDs = static_cast<PLONG>(HeapAlloc(GetProcessHeap(), 0, sizeof(LONG)));
					pItemIDs[0] = -1;
					hr = pShLvwItem->get_ID(&pItemIDs[0]);
					itemCount = 1;
				} else {
					CComQIPtr<IShListViewItems> pShLvwItems = items.pdispVal;
					if(pShLvwItems) {
						// a ShellListViewItems collection
						CComQIPtr<IEnumVARIANT> pEnumerator = pShLvwItems;
						LONG c = 0;
						pShLvwItems->Count(&c);
						itemCount = c;
						if(itemCount > 0 && pEnumerator) {
							pItemIDs = static_cast<PLONG>(HeapAlloc(GetProcessHeap(), 0, itemCount * sizeof(LONG)));
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							for(UINT i = 0; i < itemCount && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
								pItemIDs[i] = -1;
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pShLvwItem = item.pdispVal;
									if(pShLvwItem) {
										pShLvwItem->get_ID(&pItemIDs[i]);
									}
								}
								VariantClear(&item);
							}
							hr = S_OK;
						}
					} else {
						// let ExplorerListView try its luck
						EXLVMGETITEMIDSDATA data = {0};
						data.structSize = sizeof(EXLVMGETITEMIDSDATA);
						data.pVariant = &items;
						hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETITEMIDSFROMVARIANT, 0, reinterpret_cast<LPARAM>(&data));
						itemCount = data.numberOfItems;
						pItemIDs = data.pItemIDs;
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
						pItemIDs = static_cast<PLONG>(HeapAlloc(GetProcessHeap(), 0, itemCount * sizeof(LONG)));
						itemCount = 0;
						for(LONG i = lBound; (i <= uBound) && isValid; ++i) {
							pItemIDs[itemCount] = -1;
							if(SafeArrayGetElement(items.parray, &i, &element) == S_OK) {
								if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_I4))) {
									pItemIDs[itemCount++] = element.lVal;
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
				pItemIDs = static_cast<PLONG>(HeapAlloc(GetProcessHeap(), 0, sizeof(LONG)));
				pItemIDs[0] = items.lVal;
				itemCount = 1;
			} else {
				// invalid arg - raise VB runtime error 380
				hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
	}

	return hr;
}

HRESULT ShellListView::SHBindToParent(LONG itemID, REFIID requiredInterface, LPVOID* ppInterfaceImpl, PCUITEMID_CHILD* pPIDLLast, PCIDLIST_ABSOLUTE* pPIDL/* = NULL*/)
{
	LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
	if(pItemDetails) {
		if(pPIDL) {
			*pPIDL = pItemDetails->pIDL;
		}
		return ::SHBindToParent(pItemDetails->pIDL, requiredInterface, ppInterfaceImpl, pPIDLLast);
	}
	return E_INVALIDARG;
}

BOOL ShellListView::ValidateItem(LONG itemID, LPSHELLLISTVIEWITEMDATA pItemDetails/* = NULL*/)
{
	if(!pItemDetails) {
		pItemDetails = GetItemDetails(itemID);
	}
	if(pItemDetails) {
		return ValidatePIDL(pItemDetails->pIDL);
	}
	return FALSE;
}

void ShellListView::UpdateItemPIDL(LONG itemID, LPSHELLLISTVIEWITEMDATA pItemDetails, PIDLIST_ABSOLUTE pIDLNew)
{
	UINT itemIDsUpdatedItem = ILCount(pIDLNew);
	if(itemIDsUpdatedItem == ILCount(pItemDetails->pIDL)) {
		PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
		pItemDetails->pIDL = ILCloneFull(pIDLNew);
		if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
			Raise_ChangedItemPIDL(itemID, tmp, pItemDetails->pIDL);
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
			Raise_ChangedItemPIDL(itemID, tmp, pItemDetails->pIDL);
		}
		ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
	}
}


void ShellListView::AddNamespace(PCIDLIST_ABSOLUTE pIDL, CComObject<ShellListViewNamespace>* pNamespace)
{
	ATLASSERT_POINTER(pIDL, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pNamespace, CComObject<ShellListViewNamespace>);

	properties.namespaces[pIDL] = pNamespace;

	#ifdef USE_STL
		BOOL gotNewMainNamespace = (properties.namespaces.cbegin()->first == pIDL);
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		BOOL gotNewMainNamespace = (properties.namespaces.GetKeyAt(p) == pIDL);
	#endif

	if(gotNewMainNamespace) {
		SetMainNamespace(pIDL);
		if(properties.autoInsertColumns && CurrentViewNeedsColumns()) {
			// add new shell columns
			EnsureShellColumnsAreLoaded(tmTimeOutThreading);
		}
	}

	if(!(properties.disabledEvents & deNamespaceInsertionEvents)) {
		CComQIPtr<IShListViewNamespace> pNamespaceItf = pNamespace;
		Raise_InsertedNamespace(pNamespaceItf);
	}
}

void ShellListView::BufferShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem, PCIDLIST_ABSOLUTE pIDLNamespace)
{
	ATLASSERT_POINTER(pIDLItem, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);

	properties.namespacesOfInsertedItems[pIDLItem] = pIDLNamespace;
}

void ShellListView::RemoveBufferedShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem)
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

CComObject<ShellListViewNamespace>* ShellListView::IndexToNamespace(UINT index)
{
	#ifdef USE_STL
		if((index >= 0) && (index < properties.namespaces.size())) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin();
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

PCIDLIST_ABSOLUTE ShellListView::IndexToNamespacePIDL(UINT index)
{
	#ifdef USE_STL
		if((index >= 0) && (index < properties.namespaces.size())) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin();
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

LONG ShellListView::NamespacePIDLToIndex(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	LONG index = 0;
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDL);
			if(iter != properties.namespaces.cend()) {
				index = static_cast<LONG>(std::distance(properties.namespaces.cbegin(), iter));
				return index;
			}
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter, ++index) {
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

CComObject<ShellListViewNamespace>* ShellListView::NamespacePIDLToNamespace(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch)
{
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDL);
			if(iter != properties.namespaces.cend()) {
				return iter->second;
			}
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				if(ILIsEqual(iter->first, pIDL)) {
					return iter->second;
				}
			}
		}
	#else
		if(exactMatch) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.Lookup(pIDL);
			if(pPair) {
				return pPair->m_value;
			}
		} else {
			POSITION p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				if(ILIsEqual(pPair->m_key, pIDL)) {
					return pPair->m_value;
				}
				properties.namespaces.GetNext(p);
			}
		}
	#endif
	return NULL;
}

CComObject<ShellListViewNamespace>* ShellListView::NamespaceParsingNameToNamespace(LPWSTR pParsingName)
{
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
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
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
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

HRESULT ShellListView::RemoveNamespace(PCIDLIST_ABSOLUTE pIDLNamespace, BOOL exactMatch, BOOL removeItemsFromListView)
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
			CComQIPtr<IShListViewNamespace> pNamespace = pBackgroundEnum->pTargetObject;
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
						switch(pBackgroundEnum->type) {
							case BET_ITEMS:
								Raise_ItemEnumerationCompleted(pNamespace, VARIANT_TRUE);
								break;
							case BET_COLUMNS:
								Raise_ColumnEnumerationCompleted(pNamespace, VARIANT_TRUE);
								break;
						}
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

	// remove buffered dropped pIDL offsets
	#ifdef USE_STL
		for(std::vector<DroppedPIDLOffsets*>::iterator it = dragDropStatus.bufferedPIDLOffsets.begin(); it != dragDropStatus.bufferedPIDLOffsets.end();) {
			if((*it)->pTargetNamespace == pIDLNamespace) {
				DroppedPIDLOffsets* pTmp = *it;
				it = dragDropStatus.bufferedPIDLOffsets.erase(it);
				delete pTmp;
			}
		}
	#else
		for(size_t i = 0; i < dragDropStatus.bufferedPIDLOffsets.GetCount(); ++i) {
			if(dragDropStatus.bufferedPIDLOffsets[i]->pTargetNamespace == pIDLNamespace) {
				size_t i2 = i--;
				DroppedPIDLOffsets* pTmp = dragDropStatus.bufferedPIDLOffsets[i2];
				dragDropStatus.bufferedPIDLOffsets.RemoveAt(i2);
				delete pTmp;
			}
		}
	#endif

	// collect the affected items
	#ifdef USE_STL
		std::vector<LONG> itemsToRemove;
		if(!properties.useFastButImpreciseItemHandling) {
			if(exactMatch) {
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					if(iter->second->pIDLNamespace == pIDLNamespace) {
						itemsToRemove.push_back(iter->first);
					}
				}
			} else {
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					if(ILIsEqual(iter->second->pIDLNamespace, pIDLNamespace)) {
						itemsToRemove.push_back(iter->first);
					}
				}
			}
		}
	#else
		CAtlArray<LONG> itemsToRemove;
		if(!properties.useFastButImpreciseItemHandling) {
			p = properties.items.GetStartPosition();
			if(exactMatch) {
				while(p) {
					CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					if(pPair->m_value->pIDLNamespace == pIDLNamespace) {
						itemsToRemove.Add(pPair->m_key);
					}
					properties.items.GetNext(p);
				}
			} else {
				while(p) {
					CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					if(ILIsEqual(pPair->m_value->pIDLNamespace, pIDLNamespace)) {
						itemsToRemove.Add(pPair->m_key);
					}
					properties.items.GetNext(p);
				}
			}
		}
	#endif

	if(removeItemsFromListView) {
		// remove the items from the listview
		if(properties.useFastButImpreciseItemHandling) {
			attachedControl.SendMessage(LVM_DELETEALLITEMS, 0, 0);
			/* The control should have sent us a notification about the deletion. The notification handler
				 should have freed the pIDL. */
		} else {
			CWindowEx2(attachedControl).InternalSetRedraw(FALSE);
			#ifdef USE_STL
				for(size_t i = 0; i < itemsToRemove.size(); ++i) {
			#else
				for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
			#endif
				attachedControl.SendMessage(LVM_DELETEITEM, ItemIDToIndex(itemsToRemove[i]), 0);
				/* The control should have sent us a notification about the deletion. The notification handler
					 should have freed the pIDL. */
			}
			CWindowEx2(attachedControl).InternalSetRedraw(TRUE);
		}
		hr = S_OK;
	} else {
		// transfer the items to normal items - free the data
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToRemove.size(); ++i) {
				std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.find(itemsToRemove[i]);
				if(iter != properties.items.cend()) {
					LONG itemID = iter->first;
		#else
			for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.Lookup(itemsToRemove[i]);
				if(pPair) {
					LONG itemID = pPair->m_key;
		#endif
				if(!(properties.disabledEvents & deItemDeletionEvents)) {
					CComPtr<IShListViewItem> pItem;
					ClassFactory::InitShellListItem(itemID, this, IID_IShListViewItem, reinterpret_cast<LPUNKNOWN*>(&pItem));
					Raise_RemovingItem(pItem);
				}
		#ifdef USE_STL
					LPSHELLLISTVIEWITEMDATA pData = iter->second;
					properties.items.erase(iter);
		#else
					LPSHELLLISTVIEWITEMDATA pData = pPair->m_value;
					properties.items.RemoveKey(itemID);
		#endif
				delete pData;
			}
		}
		hr = S_OK;
	}

	// remove the namespace
	BOOL removedMainNamespace = TRUE;
	#ifdef USE_STL
		if(exactMatch) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.find(pIDLNamespace);
			removedMainNamespace = (iter == properties.namespaces.cbegin());
			if(iter != properties.namespaces.cend()) {
				if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
					CComQIPtr<IShListViewNamespace> pNamespaceItf = iter->second;
					Raise_RemovingNamespace(pNamespaceItf);
				}
				CComObject<ShellListViewNamespace>* pElement = iter->second;
				properties.namespaces.erase(iter);
				pElement->Release();
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDLNamespace));
		} else {
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				if(ILIsEqual(iter->first, pIDLNamespace)) {
					if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
						CComQIPtr<IShListViewNamespace> pNamespaceItf = iter->second;
						Raise_RemovingNamespace(pNamespaceItf);
					}
					CComObject<ShellListViewNamespace>* pElement = iter->second;
					PIDLIST_ABSOLUTE pIDL = const_cast<PIDLIST_ABSOLUTE>(iter->first);
					properties.namespaces.erase(iter);
					pElement->Release();
					ILFree(pIDL);
					break;
				}
				removedMainNamespace = FALSE;
			}
		}
	#else
		if(exactMatch) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = NULL;
			p = properties.namespaces.GetStartPosition();
			if(p) {
				pPair = properties.namespaces.GetAt(p);
				removedMainNamespace = (pPair->m_key == pIDLNamespace);
				if(pPair->m_key != pIDLNamespace) {
					pPair = properties.namespaces.Lookup(pIDLNamespace);
				}
			} else {
				pPair = properties.namespaces.Lookup(pIDLNamespace);
			}
			CComObject<ShellListViewNamespace>* pElement = NULL;
			if(pPair) {
				if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
					CComQIPtr<IShListViewNamespace> pNamespaceItf = pPair->m_value;
					Raise_RemovingNamespace(pNamespaceItf);
				}
				pElement = pPair->m_value;
			} else {
				removedMainNamespace = FALSE;
			}
			properties.namespaces.RemoveKey(pIDLNamespace);
			if(pElement) {
				pElement->Release();
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDLNamespace));
		} else {
			p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				if(ILIsEqual(pPair->m_key, pIDLNamespace)) {
					if(!(properties.disabledEvents & deNamespaceDeletionEvents)) {
						CComQIPtr<IShListViewNamespace> pNamespaceItf = pPair->m_value;
						Raise_RemovingNamespace(pNamespaceItf);
					}
					CComObject<ShellListViewNamespace>* pElement = pPair->m_value;
					PIDLIST_ABSOLUTE pIDL = const_cast<PIDLIST_ABSOLUTE>(pPair->m_key);
					properties.namespaces.RemoveAtPos(p);
					pElement->Release();
					ILFree(pIDL);
					break;
				}
				removedMainNamespace = FALSE;
				properties.namespaces.GetNext(p);
			}
		}
	#endif

	if(removedMainNamespace) {
		#ifdef USE_STL
			if(properties.namespaces.size() > 0) {
				SetMainNamespace(properties.namespaces.cbegin()->first);
		#else
			p = properties.namespaces.GetStartPosition();
			if(p) {
				SetMainNamespace(properties.namespaces.GetKeyAt(p));
		#endif
		} else {
			SetMainNamespace(NULL);
		}
		if(properties.autoInsertColumns && properties.columnsStatus.pIDLNamespace && CurrentViewNeedsColumns()) {
			// add new shell columns
			EnsureShellColumnsAreLoaded(tmTimeOutThreading);
		}
	}
	return hr;
}

HRESULT ShellListView::RemoveAllNamespaces(BOOL removeItemsFromListView)
{
	HRESULT hr = S_OK;
	#ifdef USE_STL
		while(properties.namespaces.size() > 0) {
			hr = RemoveNamespace(properties.namespaces.cbegin()->first, TRUE, removeItemsFromListView);
	#else
		while(properties.namespaces.GetCount() > 0) {
			hr = RemoveNamespace(properties.namespaces.GetKeyAt(properties.namespaces.GetStartPosition()), TRUE, removeItemsFromListView);
	#endif
		if(FAILED(hr)) {
			break;
		}
	}
	return hr;
}

void ShellListView::LazyCloseThumbnailDiskCaches(BOOL forceClose/* = FALSE*/)
{
	ATLASSERT(forceClose == FALSE && "LazyCloseThumbnailDiskCaches() has not been tested with forceClose == TRUE");

	BOOL anyOpenCaches = FALSE;
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			CComObject<ShellListViewNamespace>* pNamespace = iter->second;
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			CComObject<ShellListViewNamespace>* pNamespace = pPair->m_value;
			properties.namespaces.GetNext(p);
	#endif
		if(pNamespace->properties.pThumbnailDiskCache && pNamespace->properties.thumbnailDiskCacheIsOpened) {
			anyOpenCaches = TRUE;
			if(GetTickCount() - pNamespace->properties.lastThumbnailDiskCacheAccessTime > timers.INT_LAZYCLOSETHUMBNAILDISKCACHE) {
				// close it, if it is not locked (if it's locked, an access is taking place right now)
				DWORD dummy;
				if(forceClose || (pNamespace->properties.pThumbnailDiskCache->GetMode(&dummy) == S_OK && pNamespace->properties.pThumbnailDiskCache->IsLocked() == S_FALSE)) {
					pNamespace->properties.pThumbnailDiskCache->Close(NULL);
					pNamespace->properties.thumbnailDiskCacheIsOpened = FALSE;
					pNamespace->properties.lastThumbnailDiskCacheAccessTime = 0;
				}
			}
		}
	}
	if(!anyOpenCaches) {
		attachedControl.KillTimer(timers.ID_LAZYCLOSETHUMBNAILDISKCACHE);
		flags.thumbnailDiskCacheTimerIsActive = FALSE;
	}
}


BOOL ShellListView::CurrentViewNeedsColumns(void)
{
	int columnsVisibility = 0/*chvInvisible*/;
	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNHEADERVISIBILITY, 0, reinterpret_cast<LPARAM>(&columnsVisibility))));
	if(columnsVisibility == 2/*chvVisibleInAllViews*/) {
		return TRUE;
	}
	// NOTE: We handle chvInvisible like chvVisibleInDetailsView!

	if(cachedSettings.viewMode == -1) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	}
	return (cachedSettings.viewMode == 3/*vDetails*/ || cachedSettings.viewMode == 4/*vTiles*/ || cachedSettings.viewMode == 5/*vExtendedTiles*/);
}

BOOL ShellListView::IsInDetailsView(void)
{
	if(cachedSettings.viewMode == -1) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	}
	return (cachedSettings.viewMode == 3/*vDetails*/);
}

BOOL ShellListView::IsInSmallIconsView(void)
{
	if(cachedSettings.viewMode == -1) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	}
	return (cachedSettings.viewMode == 1/*vSmallIcons*/);
}

BOOL ShellListView::IsInTilesView(void)
{
	if(cachedSettings.viewMode == -1) {
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETVIEWMODE, 0, reinterpret_cast<LPARAM>(&cachedSettings.viewMode))));
	}
	return (cachedSettings.viewMode == 4/*vTiles*/ || cachedSettings.viewMode == 5/*vExtendedTiles*/);
}

HRESULT ShellListView::SetMainNamespace(PCIDLIST_ABSOLUTE pIDLNamespace)
{
	HRESULT hr = S_OK;
	if(pIDLNamespace != properties.columnsStatus.pIDLNamespace) {
		if(properties.autoInsertColumns) {
			// Don't touch the backup data if we're changing from NULL to NOT NULL!
			if(properties.columnsStatus.pIDLNamespace || !pIDLNamespace) {
				if(properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
					#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
						UINT oldNumberOfAllColumnsOfPreviousNamespace = properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace;
					#endif
					properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace = 0;
					#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
						for(UINT i = 0; i < oldNumberOfAllColumnsOfPreviousNamespace; ++i) {
							if(properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription) {
								properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription->Release();
								properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription = NULL;
							}
						}
					#endif
					HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pAllColumnsOfPreviousNamespace);
					properties.columnsStatus.pAllColumnsOfPreviousNamespace = NULL;
				}
				if(properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace) {
					properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace = 0;
					HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace);
					properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace = NULL;
				}
				ZeroMemory(&properties.columnsStatus.previousSortColumn, sizeof(properties.columnsStatus.previousSortColumn));
				properties.columnsStatus.previousSortOrder = 0;

				if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist) {
					if(properties.columnsStatus.pAllColumns) {
						// make a copy
						properties.columnsStatus.pAllColumnsOfPreviousNamespace = static_cast<LPSHELLLISTVIEWCOLUMNDATA>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, properties.columnsStatus.numberOfAllColumns * sizeof(SHELLLISTVIEWCOLUMNDATA)));
						ATLASSERT(properties.columnsStatus.pAllColumnsOfPreviousNamespace);
						CopyMemory(properties.columnsStatus.pAllColumnsOfPreviousNamespace, properties.columnsStatus.pAllColumns, properties.columnsStatus.numberOfAllColumns * sizeof(SHELLLISTVIEWCOLUMNDATA));
						properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace = properties.columnsStatus.numberOfAllColumns;
						for(UINT i = 0; i < properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace; ++i) {
							GetColumnPropertyKey(static_cast<int>(i), &properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].propertyKey);
							#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
								if(properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription) {
									properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription->AddRef();
								}
							#endif
						}
					}

					HWND hWndHeader = reinterpret_cast<HWND>(attachedControl.SendMessage(LVM_GETHEADER, 0, 0));
					ATLASSERT(::IsWindow(hWndHeader));
					if(::IsWindow(hWndHeader)) {
						int columns = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
						if(columns > 0) {
							PINT pColumns = new int[columns];
							ATLASSERT(pColumns);
							if(attachedControl.SendMessage(LVM_GETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pColumns))) {
								properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace = static_cast<PROPERTYKEY*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, columns * sizeof(PROPERTYKEY)));
								ATLASSERT(properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace);

								for(int i = 0; i < columns; ++i) {
									int realColumnIndex = ColumnIndexToRealIndex(pColumns[i]);
									if(realColumnIndex >= 0) {
										GetColumnPropertyKey(realColumnIndex, &properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace[i]);
									}
								}
								properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace = columns;
							}
							delete[] pColumns;
						}
					}
					if(sortingSettings.columnIndex >= 0) {
						int realColumnIndex = ColumnIndexToRealIndex(sortingSettings.columnIndex);
						if(realColumnIndex >= 0) {
							GetColumnPropertyKey(realColumnIndex, &properties.columnsStatus.previousSortColumn);
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&properties.columnsStatus.previousSortOrder))));
						}
					}
				}
			}

			hr = UnloadShellColumns();
		}

		ATLASSERT(SUCCEEDED(hr));
		if(SUCCEEDED(hr)) {
			/* NOTE: If we ever extend the directory watch feature so that we can watch multiple namespaces,
			*        we'd have to move directoryWatchState to the namespace object and synchronize each thread
			*        with the lifetime of the namespace object. */
			if(pIDLNamespace && !IsSlowItem(const_cast<PIDLIST_ABSOLUTE>(pIDLNamespace), TRUE, TRUE)) {
				// start directory watch thread (or make it watch another directory)
				InterlockedExchangePointer(&directoryWatchState.pIDLWatchedNamespace, const_cast<LPVOID>(reinterpret_cast<LPCVOID>(pIDLNamespace)));
				InterlockedExchange(&directoryWatchState.watchNewDirectory, TRUE);

				#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
					if(!directoryWatchState.hThread && APIWrapper::IsSupported_CancelSynchronousIo()) {
				#else
					if(!directoryWatchState.hThread) {
				#endif
					directoryWatchState.hThread = CreateThread(NULL, 0, DirWatchThreadProc, &directoryWatchState, CREATE_SUSPENDED, NULL);
					//ATLTRACE(TEXT("1 - Created dir watch thread: 0x%X\n"), directoryWatchState.hThread);
					InterlockedExchange(&directoryWatchState.isSuspended, TRUE);
				}
				if(directoryWatchState.hThread) {
					if(InterlockedCompareExchange(&directoryWatchState.isSuspended, TRUE, TRUE)) {
						ResumeThread(directoryWatchState.hThread);
					} else {
						// together with watchNewDirectory=TRUE this will make the thread switch to the new directory
						#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
							APIWrapper::CancelSynchronousIo(directoryWatchState.hThread, NULL);
						#else
							HANDLE hChangeEvent = NULL;
							InterlockedExchangePointer(&hChangeEvent, directoryWatchState.hChangeEvent);
							if(hChangeEvent != INVALID_HANDLE_VALUE) {
								SetEvent(hChangeEvent);
							}
						#endif
					}
				}
			} else {
				// suspend directory watch thread
				if(directoryWatchState.hThread) {
					InterlockedExchange(&directoryWatchState.suspend, TRUE);
					// raise thread priority so that it will suspend itself before we resume it
					SetThreadPriority(directoryWatchState.hThread, THREAD_PRIORITY_HIGHEST);
					// make sure it isn't blocked
					#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
						APIWrapper::CancelSynchronousIo(directoryWatchState.hThread, NULL);
					#else
						HANDLE hChangeEvent = NULL;
						InterlockedExchangePointer(&hChangeEvent, directoryWatchState.hChangeEvent);
						if(hChangeEvent != INVALID_HANDLE_VALUE) {
							SetEvent(hChangeEvent);
						}
					#endif
				}
				InterlockedExchangePointer(&directoryWatchState.pIDLWatchedNamespace, NULL);
			}

			if(properties.autoInsertColumns) {
				properties.columnsStatus.pIDLNamespace = pIDLNamespace;
				properties.columnsStatus.pNamespaceSHF = NULL;
				properties.columnsStatus.pNamespaceSHF2 = NULL;
				properties.columnsStatus.pNamespaceSHD = NULL;
				sortingSettings.firstSortItemsCall = TRUE;

				if(properties.columnsStatus.pIDLNamespace) {
					hr = BindToPIDL(properties.columnsStatus.pIDLNamespace, IID_PPV_ARGS(&properties.columnsStatus.pNamespaceSHF));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						ATLASSUME(properties.columnsStatus.pNamespaceSHF);
						properties.columnsStatus.pNamespaceSHF->QueryInterface(IID_PPV_ARGS(&properties.columnsStatus.pNamespaceSHF2));
						properties.columnsStatus.pNamespaceSHF->CreateViewObject(GethWndShellUIParentWindow(), IID_PPV_ARGS(&properties.columnsStatus.pNamespaceSHD));
						if(!properties.columnsStatus.pNamespaceSHF2 && !properties.columnsStatus.pNamespaceSHD) {
							// TODO: for filesystem folders provide a default implementation, for other folders try to use SFVM_GETDETAILSOF
							ATLASSERT(FALSE);
						}
						if(properties.columnsStatus.pNamespaceSHF2) {
							ULONG dummy = 0;
							properties.columnsStatus.defaultSortColumn = 0;
							properties.columnsStatus.pNamespaceSHF2->GetDefaultColumn(0, reinterpret_cast<PULONG>(&properties.columnsStatus.defaultSortColumn), &dummy);
						}
					}
					if(properties.columnsStatus.hColumnsReadyEvent) {
						ResetEvent(properties.columnsStatus.hColumnsReadyEvent);
					}
				}
			}
		}
	}
	return hr;
}

HRESULT ShellListView::EnsureShellColumnsAreLoaded(ThreadingMode threadingMode)
{
	#define INITIALNUMBEROFCOLUMNS	10
	#define COLUMNSGROW							10

	//ATLASSERT(properties.columnsStatus.pIDLNamespace);
	if(!properties.columnsStatus.pIDLNamespace) {
		return E_POINTER;
	}
	if(properties.columnsStatus.loaded != SCLS_NOTLOADED) {
		/* TODO: Shouldn't we wait here in case of SCLS_LOADING? However, we would have to wait in a way that
		 *       doesn't block the current thread, because it is the main thread. Maybe we should use a message
		 *       pump.
		 */
		return S_OK;
	}

	ResetEvent(properties.columnsStatus.hColumnsReadyEvent);
	CComObject<ShellListViewNamespace>* pNamespaceObj = NamespacePIDLToNamespace(properties.columnsStatus.pIDLNamespace, TRUE);
	ATLASSUME(pNamespaceObj);
	CComPtr<IShListViewNamespace> pNamespace = NULL;
	pNamespaceObj->QueryInterface(IID_PPV_ARGS(&pNamespace));
	ATLASSUME(pNamespace);
	Raise_ColumnEnumerationStarted(pNamespace);

	HRESULT hr = E_FAIL;
	SHELLDETAILS column = {0};
	int realColumnIndex = 0;
	properties.columnsStatus.numberOfAllColumns = 0;
	HANDLE hProcessHeap = GetProcessHeap();
	DWORD enumerationBegin = GetTickCount();
	BOOL completedSynchronously = TRUE;
	do {
		SHCOLSTATEF columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
		if(properties.columnsStatus.pNamespaceSHF2) {
			hr = properties.columnsStatus.pNamespaceSHF2->GetDetailsOf(NULL, realColumnIndex, &column);
			if(FAILED(properties.columnsStatus.pNamespaceSHF2->GetDefaultColumnState(realColumnIndex, &columnState))) {
				columnState = SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR;
			}
		} else {
			hr = E_FAIL;
		}
		if(FAILED(hr) && properties.columnsStatus.pNamespaceSHD) {
			hr = properties.columnsStatus.pNamespaceSHD->GetDetailsOf(NULL, realColumnIndex, &column);
		}

		if(SUCCEEDED(hr)) {
			// store the column
			if(!properties.columnsStatus.pAllColumns) {
				properties.columnsStatus.pAllColumns = static_cast<LPSHELLLISTVIEWCOLUMNDATA>(HeapAlloc(hProcessHeap, HEAP_ZERO_MEMORY, INITIALNUMBEROFCOLUMNS * sizeof(SHELLLISTVIEWCOLUMNDATA)));
				ATLASSERT(properties.columnsStatus.pAllColumns);
				properties.columnsStatus.currentArraySize = INITIALNUMBEROFCOLUMNS * sizeof(SHELLLISTVIEWCOLUMNDATA);
			}
			if((realColumnIndex + 1) * sizeof(SHELLLISTVIEWCOLUMNDATA) > properties.columnsStatus.currentArraySize) {
				// too small, so realloc
				LPSHELLLISTVIEWCOLUMNDATA p = static_cast<LPSHELLLISTVIEWCOLUMNDATA>(HeapReAlloc(hProcessHeap, HEAP_ZERO_MEMORY, properties.columnsStatus.pAllColumns, properties.columnsStatus.currentArraySize + COLUMNSGROW * sizeof(SHELLLISTVIEWCOLUMNDATA)));
				if(!p) {
					ATLASSERT(p);
					break;
				}
				properties.columnsStatus.currentArraySize += COLUMNSGROW * sizeof(SHELLLISTVIEWCOLUMNDATA);
				properties.columnsStatus.pAllColumns = p;
			}
			ATLASSERT(properties.columnsStatus.pAllColumns);
			LPSHELLLISTVIEWCOLUMNDATA pColumnData = &properties.columnsStatus.pAllColumns[realColumnIndex];
			pColumnData->columnID = -1;
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				pColumnData->pPropertyDescription = NULL;
				pColumnData->displayType = PDDT_STRING;
				pColumnData->typeFlags = PDTF_DEFAULT;
			#endif
			pColumnData->visibleInDetailsView = FALSE;
			pColumnData->visibilityHasBeenChangedExplicitly = FALSE;
			pColumnData->lastIndex_DetailsView = -1;
			pColumnData->lastIndex_TilesView = -1;
			ATLVERIFY(SUCCEEDED(StrRetToBuf(&column.str, NULL, pColumnData->pText, MAX_LVCOLUMNTEXTLENGTH)));
			if(column.cxChar <= 0) {
				pColumnData->width = 96;
			} else {
				pColumnData->width = column.cxChar * properties.columnsStatus.averageCharWidth;
			}
			pColumnData->format = column.fmt;
			PROPERTYKEY propertyKey = {0};
			BOOL hasPropertyKey = FALSE;
			if(properties.columnsStatus.pNamespaceSHF2 && SUCCEEDED(properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, &propertyKey))) {
				hasPropertyKey = TRUE;
				if(APIWrapper::IsSupported_PSGetPropertyDescription()) {
					IPropertyDescription* pPropertyDescription = NULL;
					#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
						pPropertyDescription = pColumnData->pPropertyDescription;
					#endif
					APIWrapper::PSGetPropertyDescription(propertyKey, IID_PPV_ARGS(&pPropertyDescription), NULL);
					if(pPropertyDescription) {
						UINT width = 0;
						if(SUCCEEDED(pPropertyDescription->GetDefaultColumnWidth(&width))) {
							pColumnData->width = static_cast<int>(width) * properties.columnsStatus.averageCharWidth;
						}

						PROPDESC_VIEW_FLAGS viewFlags;
						if(SUCCEEDED(pPropertyDescription->GetViewFlags(&viewFlags))) {
							// TODO: Is this correct?
							if(viewFlags & PDVF_BEGINNEWGROUP) {
								pColumnData->format |= LVCFMT_LINE_BREAK;
							}
							if(viewFlags & PDVF_FILLAREA) {
								pColumnData->format |= LVCFMT_FILL;
							}
							// TODO: Is this correct?
							if(viewFlags & PDVF_HIDELABEL) {
								pColumnData->format |= LVCFMT_NO_TITLE;
							}
							if(viewFlags & PDVF_CANWRAP) {
								pColumnData->format |= LVCFMT_WRAP;
							}
						}

						#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
							if(RunTimeHelper::IsVista()) {
								pPropertyDescription->GetDisplayType(&pColumnData->displayType);
								pPropertyDescription->GetTypeFlags(PDTF_MASK_ALL, &pColumnData->typeFlags);
							}
						#else
							pPropertyDescription->Release();
						#endif
					}
				}
			}
			pColumnData->state = columnState;

			int visibleInPreviousNamespace = -1;
			if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist && properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
				// copy data from the previous namespace
				LPSHELLLISTVIEWCOLUMNDATA pBackupColumnData = NULL;
				if(hasPropertyKey) {
					for(UINT i = 0; i < properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace; ++i) {
						PROPERTYKEY backupPropertyKey = properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].propertyKey;
						if(propertyKey.fmtid == backupPropertyKey.fmtid && propertyKey.pid == backupPropertyKey.pid) {
							pBackupColumnData = &properties.columnsStatus.pAllColumnsOfPreviousNamespace[i];
							break;
						}
					}
				}
				if(pBackupColumnData) {
					pColumnData->width = pBackupColumnData->width;
					lstrcpyn(pColumnData->pText, pBackupColumnData->pText, MAX_LVCOLUMNTEXTLENGTH);
					pColumnData->format &= ~LVCFMT_JUSTIFYMASK;
					pColumnData->format |= (pBackupColumnData->format & LVCFMT_JUSTIFYMASK);
					pColumnData->visibilityHasBeenChangedExplicitly = pBackupColumnData->visibilityHasBeenChangedExplicitly;
					if(pBackupColumnData->visibilityHasBeenChangedExplicitly) {
						visibleInPreviousNamespace = (pBackupColumnData->visibleInDetailsView ? 1 : 0);
					}
				}
			}
			BOOL makeVisible = FALSE;
			if((columnState & SHCOLSTATE_HIDDEN) == 0) {
				makeVisible = ((columnState & SHCOLSTATE_ONBYDEFAULT) == SHCOLSTATE_ONBYDEFAULT);
				if(visibleInPreviousNamespace != -1) {
					makeVisible = (visibleInPreviousNamespace != 0);
				}
			}
			// always show the first column
			makeVisible = makeVisible || (realColumnIndex == 0);

			++properties.columnsStatus.numberOfAllColumns;

			if(!(properties.disabledEvents & deColumnLoadingEvents)) {
				VARIANT_BOOL visible = BOOL2VARIANTBOOL(makeVisible);
				CComPtr<IShListViewColumn> pColumn;
				ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumn));
				Raise_LoadedColumn(pColumn, &visible);
				makeVisible = VARIANTBOOL2BOOL(visible);
			}

			if(makeVisible) {
				ATLVERIFY(ChangeColumnVisibility(realColumnIndex, TRUE, 0) != -1);
			}
			// ChangeColumnVisibility will set 'visibleInDetailsView' only if we're in Details view mode
			pColumnData->visibleInDetailsView = makeVisible;
			++realColumnIndex;

			if(threadingMode == tmTimeOutThreading && GetTickCount() - enumerationBegin >= 200) {
				// switch to background enumeration
				CComPtr<IRunnableTask> pTask;
				/* NOTE: Starting with revision 387 we don't marshal pNamespaceSHF2 and pNamespaceSHD anymore,
				 *       because this lead to deadlocks in calls to IShellFolder2 methods in
				 *       ShLvwBackgroundColumnEnumTask if the app was being closed.
				 */
				hr = ShLvwBackgroundColumnEnumTask::CreateInstance(attachedControl, &properties.backgroundColumnEnumResults, &properties.backgroundColumnEnumResultsCritSection, GethWndShellUIParentWindow(), properties.columnsStatus.pIDLNamespace, realColumnIndex, properties.columnsStatus.averageCharWidth, &pTask);
				if(SUCCEEDED(hr)) {
					INamespaceEnumTask* pEnumTask = NULL;
					pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask));
					ATLASSUME(pEnumTask);
					if(pEnumTask) {
						CComQIPtr<IDispatch> pTargetObject = pNamespace;
						ATLASSUME(pTargetObject);
						ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pTargetObject)));
						pEnumTask->Release();
					}
					ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShLvwBackgroundColumnEnum, 0, TASK_PRIORITY_LV_BACKGROUNDENUMCOLUMNS, enumerationBegin)));
					completedSynchronously = FALSE;
					break;
				}
			}
		}
	} while(hr == S_OK);
	if(properties.persistColumnSettingsAcrossNamespaces == slpcsanPersist && properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace) {
		// copy column order from the previous namespace
		HWND hWndHeader = reinterpret_cast<HWND>(attachedControl.SendMessage(LVM_GETHEADER, 0, 0));
		ATLASSERT(::IsWindow(hWndHeader));
		if(::IsWindow(hWndHeader)) {
			UINT columns = static_cast<UINT>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
			if(columns > 0) {
				PINT pColumns = static_cast<PINT>(HeapAlloc(GetProcessHeap(), 0, columns * sizeof(int)));
				if(pColumns) {
					if(attachedControl.SendMessage(LVM_GETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pColumns))) {
						for(UINT i = 0; i < columns; ++i) {
							int realColumnIndex2 = ColumnIndexToRealIndex(pColumns[i]);
							if(realColumnIndex2 >= 0) {
								PROPERTYKEY propertyKey = {0};
								GetColumnPropertyKey(realColumnIndex2, &propertyKey);
								BOOL found = FALSE;
								UINT j = 0;
								for(j = 0; j < properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace; ++j) {
									if(propertyKey == properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace[j]) {
										found = TRUE;
										break;
									}
								}

								if(found && i != j && j < columns) {
									int tmp = pColumns[i];
									pColumns[i] = pColumns[j];
									pColumns[j] = tmp;
									i--;
								}
							}
						}
						if(attachedControl.SendMessage(LVM_SETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pColumns))) {
							attachedControl.InvalidateRect(NULL, TRUE);
						}
					}
					HeapFree(GetProcessHeap(), 0, pColumns);
				}
			}
		}
	}
	properties.columnsStatus.loaded = (completedSynchronously ? SCLS_LOADED : SCLS_LOADING);

	if(completedSynchronously) {
		SetEvent(properties.columnsStatus.hColumnsReadyEvent);
		Raise_ColumnEnumerationCompleted(pNamespace, VARIANT_FALSE);
	}
	if(properties.columnsStatus.loaded == SCLS_LOADED) {
		if(properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
			// the data of the previous namespace isn't needed anymore
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				UINT oldNumberOfAllColumnsOfPreviousNamespace = properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace;
			#endif
			properties.columnsStatus.numberOfAllColumnsOfPreviousNamespace = 0;
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				for(UINT i = 0; i < oldNumberOfAllColumnsOfPreviousNamespace; ++i) {
					if(properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription) {
						properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription->Release();
						properties.columnsStatus.pAllColumnsOfPreviousNamespace[i].pPropertyDescription = NULL;
					}
				}
			#endif
			HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pAllColumnsOfPreviousNamespace);
			properties.columnsStatus.pAllColumnsOfPreviousNamespace = NULL;
		}
		if(properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace) {
			// the data of the previous namespace isn't needed anymore
			properties.columnsStatus.numberOfVisibleColumnsOfPreviousNamespace = 0;
			HeapFree(GetProcessHeap(), 0, properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace);
			properties.columnsStatus.pVisibleColumnOrderOfPreviousNamespace = NULL;
		}
	}
	return S_OK;
}

HRESULT ShellListView::UnloadShellColumns(void)
{
	if(RunTimeHelper::IsCommCtrl6()) {
		int selectedColumn = ColumnIndexToRealIndex(static_cast<int>(attachedControl.SendMessage(LVM_GETSELECTEDCOLUMN, 0, 0)));
		if(selectedColumn != -1) {
			// remove column selection
			attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, static_cast<WPARAM>(-1), 0);
		}
	}

	#ifdef USE_STL
		std::vector<int> columns;
		for(std::unordered_map<LONG, int>::const_iterator iter = properties.visibleColumns.cbegin(); iter != properties.visibleColumns.cend(); ++iter) {
			columns.push_back(iter->second);
		}
		for(std::vector<int>::const_iterator iter = columns.cbegin(); iter != columns.cend(); ++iter) {
			ChangeColumnVisibility(*iter, FALSE, CCVF_FORUNLOADSHELLCOLUMNS);
		}
		ATLASSERT(properties.visibleColumns.size() == 0);
	#else
		CAtlArray<int> columns;
		POSITION p = properties.visibleColumns.GetStartPosition();
		while(p) {
			columns.Add(properties.visibleColumns.GetNextValue(p));
		}
		for(size_t i = 0; i < columns.GetCount(); ++i) {
			ChangeColumnVisibility(columns[i], FALSE, CCVF_FORUNLOADSHELLCOLUMNS);
		}
		ATLASSERT(properties.visibleColumns.GetCount() == 0);
	#endif

	if(properties.columnsStatus.loaded != SCLS_NOTLOADED) {
		if(!(properties.disabledEvents & deColumnLoadingEvents)) {
			Raise_UnloadedColumn(NULL);
		}
	}
	properties.columnsStatus.ResetToDefaults(TRUE);
	sortingSettings.columnIndex = -1;
	return S_OK;
}

int ShellListView::GetAverageColumnHeaderCharWidth(void)
{
	ATLASSERT(attachedControl.IsWindow());

	int averageCharWidth = 6;
	CDCHandle dc = ::GetDC(HWND_DESKTOP);
	if(dc) {
		HFONT hPreviousFont = dc.SelectFont(attachedControl.GetFont());
		SIZE sz = {0};
		GetTextExtentPoint32(dc, TEXT("0"), 1, &sz);
		averageCharWidth = sz.cx;
		dc.SelectFont(hPreviousFont);
		::ReleaseDC(HWND_DESKTOP, dc);
	}
	return averageCharWidth;
}

void ShellListView::BufferShellColumnRealIndex(int realColumnIndex)
{
	ATLASSERT(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns);

	#ifdef USE_STL
		properties.shellColumnIndexesOfInsertedColumns.push(realColumnIndex);
	#else
		properties.shellColumnIndexesOfInsertedColumns.AddTail(realColumnIndex);
	#endif
}

LONG ShellListView::ChangeColumnVisibility(int realColumnIndex, BOOL makeVisible, DWORD changeColumnVisibilityflags)
{
	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		LPSHELLLISTVIEWCOLUMNDATA pColumn = properties.columnsStatus.pAllColumns;
		pColumn += realColumnIndex;
		if((changeColumnVisibilityflags & (CCVF_ISEXPLICITCHANGE | CCVF_ISEXPLICITCHANGEIFDIFFERENT)) == CCVF_ISEXPLICITCHANGE) {
			pColumn->visibilityHasBeenChangedExplicitly = TRUE;
		}
		if(makeVisible) {
			if(pColumn->columnID == -1) {
				// insert the column
				if((changeColumnVisibilityflags & CCVF_ISEXPLICITCHANGEIFDIFFERENT) == CCVF_ISEXPLICITCHANGEIFDIFFERENT) {
					pColumn->visibilityHasBeenChangedExplicitly = TRUE;
				}
				VARIANT_BOOL cancel = VARIANT_FALSE;
				if(!(properties.disabledEvents & deColumnVisibilityEvents)) {
					CComPtr<IShListViewColumn> pColumnObj;
					ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumnObj));
					Raise_ChangingColumnVisibility(pColumnObj, BOOL2VARIANTBOOL(makeVisible), &cancel);
				}
				if(cancel == VARIANT_FALSE) {
					properties.columnDataForFastInsertion.insertAt = (IsInDetailsView() ? pColumn->lastIndex_DetailsView : pColumn->lastIndex_TilesView);
					properties.columnDataForFastInsertion.pColumnText = pColumn->pText;
					properties.columnDataForFastInsertion.columnWidth = pColumn->width;
					properties.columnDataForFastInsertion.columnFormat = pColumn->format & ~LVCFMT_FIXED_WIDTH;
					if(properties.useColumnResizability != VARIANT_FALSE) {
						if(pColumn->state & SHCOLSTATE_FIXED_WIDTH) {
							properties.columnDataForFastInsertion.columnFormat |= LVCFMT_FIXED_WIDTH;
						}
					}
					properties.columnDataForFastInsertion.columnFormat |= LVCFMT_BITMAP_ON_RIGHT;
					properties.columnDataForFastInsertion.columnSubItemIndex = realColumnIndex;
					BufferShellColumnRealIndex(realColumnIndex);
					HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_ADDCOLUMN, TRUE, reinterpret_cast<LPARAM>(&properties.columnDataForFastInsertion));
					ATLASSERT(SUCCEEDED(hr));
					if(FAILED(hr)) {
						// clean-up buffer
						#ifdef USE_STL
							properties.shellColumnIndexesOfInsertedColumns.pop();
						#else
							properties.shellColumnIndexesOfInsertedColumns.RemoveHeadNoReturn();
						#endif
					}
				}
			}
		} else {
			if(pColumn->columnID != -1) {
				// remove the column
				if((changeColumnVisibilityflags & CCVF_ISEXPLICITCHANGEIFDIFFERENT) == CCVF_ISEXPLICITCHANGEIFDIFFERENT) {
					pColumn->visibilityHasBeenChangedExplicitly = TRUE;
				}
				LPSHELLLISTVIEWCOLUMNDATA pColumnBackup = NULL;
				if((changeColumnVisibilityflags & CCVF_FORUNLOADSHELLCOLUMNS) && properties.columnsStatus.pAllColumnsOfPreviousNamespace) {
					pColumnBackup = properties.columnsStatus.pAllColumnsOfPreviousNamespace;
					pColumnBackup += realColumnIndex;
					pColumnBackup->visibilityHasBeenChangedExplicitly = pColumn->visibilityHasBeenChangedExplicitly;
				}

				int selectedColumn = -1;
				if(properties.selectSortColumn && RunTimeHelper::IsCommCtrl6()) {
					selectedColumn = ColumnIndexToRealIndex(static_cast<int>(attachedControl.SendMessage(LVM_GETSELECTEDCOLUMN, 0, 0)));
				}
				if(IsInDetailsView()) {
					pColumn->lastIndex_DetailsView = RealColumnIndexToIndex(realColumnIndex);
					if(pColumnBackup) {
						pColumnBackup->lastIndex_DetailsView = pColumn->lastIndex_DetailsView;
					}
				} else {
					pColumn->lastIndex_TilesView = RealColumnIndexToIndex(realColumnIndex);
					if(pColumnBackup) {
						pColumnBackup->lastIndex_TilesView = pColumn->lastIndex_TilesView;
					}
				}
				VARIANT_BOOL cancel = VARIANT_FALSE;
				if(!(properties.disabledEvents & deColumnVisibilityEvents)) {
					CComPtr<IShListViewColumn> pColumnObj;
					ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumnObj));
					Raise_ChangingColumnVisibility(pColumnObj, BOOL2VARIANTBOOL(makeVisible), &cancel);
				}
				int columnIndex = ColumnIDToIndex(pColumn->columnID);
				if(columnIndex >= 0) {
					LVCOLUMN column = {0};
					column.mask = LVCF_WIDTH;
					if(attachedControl.SendMessage(LVM_GETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column))) {
						pColumn->width = column.cx;
						if(pColumnBackup) {
							pColumnBackup->width = pColumn->width;
						}
					}
				}
				if(cancel == VARIANT_FALSE) {
					if(properties.selectSortColumn && selectedColumn == realColumnIndex && RunTimeHelper::IsCommCtrl6()) {
						// remove column selection
						attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, static_cast<WPARAM>(-1), 0);
					}
					// remove the column from the listview
					if(columnIndex >= 0) {
						attachedControl.SendMessage(LVM_DELETECOLUMN, columnIndex, 0);
						/* The control should have sent us a notification about the deletion. The notification handler
						   should have updated pColumn. */
					}
				}
			}
		}
		return pColumn->columnID;
	}
	return -1;
}

void ShellListView::ChangedColumnVisibility(int realColumnIndex, LONG newColumnID, BOOL madeVisible)
{
	ATLASSERT((madeVisible && (newColumnID != -1)) || (!madeVisible && (newColumnID == -1)));

	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		LONG oldColumnID = properties.columnsStatus.pAllColumns[realColumnIndex].columnID;
		properties.columnsStatus.pAllColumns[realColumnIndex].columnID = newColumnID;

		BOOL realChange = ((madeVisible && oldColumnID == -1) || (!madeVisible && oldColumnID != -1));
		if(realChange) {
			if(madeVisible) {
				properties.visibleColumns[newColumnID] = realColumnIndex;
			} else {
				#ifdef USE_STL
					properties.visibleColumns.erase(oldColumnID);
				#else
					properties.visibleColumns.RemoveKey(oldColumnID);
				#endif
			}
			if(IsInDetailsView()) {
				properties.columnsStatus.pAllColumns[realColumnIndex].visibleInDetailsView = madeVisible;

				if(madeVisible && (properties.selectSortColumn || properties.setSortArrows) && realColumnIndex == sortingSettings.columnIndex) {
					// set sort arrow
					int columnIndex = RealColumnIndexToIndex(sortingSettings.columnIndex);
					if(columnIndex != -1) {
						if(properties.setSortArrows) {
							int sortOrder = 0/*soAscending*/;
							ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETSORTORDER, 0, reinterpret_cast<LPARAM>(&sortOrder))));
							if(RunTimeHelper::IsCommCtrl6()) {
								ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETSORTARROW, columnIndex, sortOrder == 0/*soAscending*/ ? 2/*saUp*/ : 1/*saDown*/)));
							} else {
								sortingSettings.FreeBitmapHandle();
								ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETCOLUMNBITMAP, columnIndex, sortOrder == 0/*soAscending*/ ? -2 : -1)));
								ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETCOLUMNBITMAP, columnIndex, reinterpret_cast<LPARAM>(&sortingSettings.headerBitmapHandle))));
							}
						}
						if(properties.selectSortColumn && RunTimeHelper::IsCommCtrl6()) {
							// select column
							attachedControl.SendMessage(LVM_SETSELECTEDCOLUMN, columnIndex, 0);
						}
					}
				}
			}

			if(!(properties.disabledEvents & deColumnVisibilityEvents)) {
				CComPtr<IShListViewColumn> pColumn;
				ClassFactory::InitShellListColumn(realColumnIndex, this, IID_IShListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pColumn));
				Raise_ChangedColumnVisibility(pColumn, BOOL2VARIANTBOOL(madeVisible));
			}
		}
	}
}

LPSHELLLISTVIEWCOLUMNDATA ShellListView::GetColumnDetails(int realColumnIndex)
{
	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		/* TODO: This is not very wise. The address could change due to EnsureShellColumnsAreLoaded() calling
		         HeapReAlloc() while the caller is accessing the data. */
		return &properties.columnsStatus.pAllColumns[realColumnIndex];
	}
	return NULL;
}

HRESULT ShellListView::GetColumnPropertyKey(int realColumnIndex, PROPERTYKEY* pPropertyKey)
{
	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		if(properties.columnsStatus.pNamespaceSHF2) {
			return properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, pPropertyKey);
		} else {
			return E_NOINTERFACE;
		}
	}
	return E_INVALIDARG;
}

HRESULT ShellListView::GetColumnPropertyDescription(int realColumnIndex, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			LPSHELLLISTVIEWCOLUMNDATA pColumnDetails = GetColumnDetails(realColumnIndex);
			if(!pColumnDetails || !pColumnDetails->pPropertyDescription) {
				if(properties.columnsStatus.pNamespaceSHF2) {
					PROPERTYKEY propertyKey;
					HRESULT hr = properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, &propertyKey);
					if(SUCCEEDED(hr)) {
						HRESULT hr2;
						hr = APIWrapper::PSGetPropertyDescription(propertyKey, requiredInterface, ppInterfaceImpl, &hr2);
						if(SUCCEEDED(hr)) {
							hr = hr2;
							if(pColumnDetails) {
								(*reinterpret_cast<IPropertyDescription**>(ppInterfaceImpl))->QueryInterface(IID_PPV_ARGS(&pColumnDetails->pPropertyDescription));
							}
						}
					}
					return hr;
				} else {
					return E_NOINTERFACE;
				}
			} else {
				return pColumnDetails->pPropertyDescription->QueryInterface(requiredInterface, ppInterfaceImpl);
			}
		#else
			if(properties.columnsStatus.pNamespaceSHF2) {
				PROPERTYKEY propertyKey;
				HRESULT hr = properties.columnsStatus.pNamespaceSHF2->MapColumnToSCID(realColumnIndex, &propertyKey);
				if(SUCCEEDED(hr)) {
					HRESULT hr2;
					hr = APIWrapper::PSGetPropertyDescription(propertyKey, requiredInterface, ppInterfaceImpl, &hr2);
					if(SUCCEEDED(hr)) {
						hr = hr2;
					}
				}
				return hr;
			} else {
				return E_NOINTERFACE;
			}
		#endif
	}
	return E_INVALIDARG;
}

BOOL ShellListView::IsShellColumn(LONG columnID)
{
	#ifdef USE_STL
		return (properties.visibleColumns.find(columnID) != properties.visibleColumns.cend());
	#else
		return (properties.visibleColumns.Lookup(columnID) != NULL);
	#endif
}

int ShellListView::ColumnIDToIndex(LONG columnID)
{
	int columnIndex = -1;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_COLUMNIDTOINDEX, columnID, reinterpret_cast<LPARAM>(&columnIndex));
	return (SUCCEEDED(hr) ? columnIndex : -1);
}

int ShellListView::ColumnIDToRealIndex(LONG columnID)
{
	#ifdef USE_STL
		std::unordered_map<LONG, int>::const_iterator iter = properties.visibleColumns.find(columnID);
		if(iter != properties.visibleColumns.cend()) {
			return iter->second;
		}
	#else
		CAtlMap<LONG, int>::CPair* pPair = properties.visibleColumns.Lookup(columnID);
		if(pPair) {
			return pPair->m_value;
		}
	#endif
	return -1;
}

LONG ShellListView::ColumnIndexToID(int columnIndex)
{
	LONG columnID = -1;
	HRESULT hr = properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_COLUMNINDEXTOID, columnIndex, reinterpret_cast<LPARAM>(&columnID));
	return (SUCCEEDED(hr) ? columnID : -1);
}

int ShellListView::ColumnIndexToRealIndex(int columnIndex)
{
	LONG columnID = ColumnIndexToID(columnIndex);
	if(columnID != -1) {
		return ColumnIDToRealIndex(columnID);
	}
	return -1;
}

int ShellListView::PropertyKeyToRealIndex(const PROPERTYKEY& propertyKey)
{
	PROPERTYKEY pk;
	for(int realColumnIndex = 0; static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns; ++realColumnIndex) {
		if(SUCCEEDED(GetColumnPropertyKey(realColumnIndex, &pk))) {
			if(pk.fmtid == propertyKey.fmtid && pk.pid == propertyKey.pid) {
				return realColumnIndex;
			}
		} else {
			return -1;
		}
	}
	return -1;
}

LONG ShellListView::RealColumnIndexToID(int realColumnIndex)
{
	if(realColumnIndex >= 0 && static_cast<UINT>(realColumnIndex) < properties.columnsStatus.numberOfAllColumns) {
		return properties.columnsStatus.pAllColumns[realColumnIndex].columnID;
	}
	return -1;
}

int ShellListView::RealColumnIndexToIndex(int realColumnIndex)
{
	LONG columnID = RealColumnIndexToID(realColumnIndex);
	if(columnID != -1) {
		return ColumnIDToIndex(columnID);
	}
	return -1;
}

BOOL ShellListView::SetColumnAlignment(int realColumnIndex, AlignmentConstants alignment)
{
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = GetColumnDetails(realColumnIndex);
	if(pColumnData) {
		DWORD alignmentFlags = 0;
		switch(alignment) {
			case alLeft:
				alignmentFlags |= LVCFMT_LEFT;
				break;
			case alCenter:
				alignmentFlags |= LVCFMT_CENTER;
				break;
			case alRight:
				alignmentFlags |= LVCFMT_RIGHT;
				break;
		}
		pColumnData->format &= ~LVCFMT_JUSTIFYMASK;
		pColumnData->format |= (alignmentFlags & LVCFMT_JUSTIFYMASK);
		if(pColumnData->columnID != -1) {
			int columnIndex = RealColumnIndexToIndex(realColumnIndex);
			if(columnIndex >= 0) {
				LVCOLUMN column = {0};
				column.mask = LVCF_FMT;
				attachedControl.SendMessage(LVM_GETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column));
				column.fmt &= ~LVCFMT_JUSTIFYMASK;
				column.fmt |= (alignmentFlags & LVCFMT_JUSTIFYMASK);
				if(attachedControl.SendMessage(LVM_SETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column))) {
					attachedControl.InvalidateRect(NULL, TRUE);
				}
			}
		}
		return TRUE;
	}
	return FALSE;
}

BOOL ShellListView::SetColumnCaption(int realColumnIndex, LPTSTR pText)
{
	ATLASSERT_POINTER(pText, TCHAR);

	LPSHELLLISTVIEWCOLUMNDATA pColumnData = GetColumnDetails(realColumnIndex);
	if(pColumnData) {
		lstrcpyn(pColumnData->pText, pText, MAX_LVCOLUMNTEXTLENGTH);
		if(pColumnData->columnID != -1) {
			int columnIndex = RealColumnIndexToIndex(realColumnIndex);
			if(columnIndex >= 0) {
				HWND hWndHeader = reinterpret_cast<HWND>(attachedControl.SendMessage(LVM_GETHEADER, 0, 0));
				ATLASSERT(::IsWindow(hWndHeader));
				if(::IsWindow(hWndHeader)) {
					HDITEM column = {0};
					column.mask = HDI_TEXT;
					column.pszText = pText;
					SendMessage(hWndHeader, HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
				}
			}
		}
		return TRUE;
	}
	return FALSE;
}

BOOL ShellListView::SetColumnDefaultWidth(int realColumnIndex, OLE_XSIZE_PIXELS width)
{
	LPSHELLLISTVIEWCOLUMNDATA pColumnData = GetColumnDetails(realColumnIndex);
	if(pColumnData) {
		pColumnData->width = width;
		if(pColumnData->columnID != -1) {
			int columnIndex = RealColumnIndexToIndex(realColumnIndex);
			if(columnIndex >= 0) {
				if(width < 0) {
					attachedControl.SendMessage(LVM_SETCOLUMNWIDTH, 0, MAKELPARAM(width, 0));
				} else {
					LVCOLUMN column = {0};
					column.cx = width;
					column.mask = LVCF_WIDTH;
					attachedControl.SendMessage(LVM_SETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column));
				}
			}
		}
		return TRUE;
	}
	return FALSE;
}


HRESULT ShellListView::CreateShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, HMENU& hMenu)
{
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);

	DestroyShellContextMenu();

	contextMenuData.bufferedCtrlDown = FALSE;
	contextMenuData.bufferedShiftDown = FALSE;
	UINT contextMenuFlags = CMF_NORMAL;
	if(GetKeyState(VK_SHIFT) & 0x8000) {
		contextMenuFlags |= CMF_EXTENDEDVERBS;
		contextMenuData.bufferedShiftDown = TRUE;
	}
	if(GetKeyState(VK_CONTROL) & 0x8000) {
		contextMenuData.bufferedCtrlDown = TRUE;
	}
	// TODO: Use CDefFolderMenu_Create2(NULL, NULL, 0, NULL, pSFParent, NULL, 0, NULL, (IContextMenu**)ppvOut);
	//       This might fix a few context menu extensions like Subversion.

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

		CComObject<ShellListViewNamespace>* pNamespace = NamespacePIDLToNamespace(pIDLNamespace, TRUE);
		ATLASSUME(pNamespace);
		pNamespace->QueryInterface<IDispatch>(&contextMenuData.pContextMenuItems);

		// allow customization of 'contextMenuFlags'
		ShellContextMenuStyleConstants contextMenuStyle = static_cast<ShellContextMenuStyleConstants>(contextMenuFlags);
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_CreatingShellContextMenu(contextMenuData.pContextMenuItems, &contextMenuStyle, &cancel);
		contextMenuFlags = static_cast<UINT>(contextMenuStyle);

		if((cancel == VARIANT_FALSE) && contextMenuData.currentShellContextMenu.CreatePopupMenu()) {
			// fill the menu
			if(contextMenuData.pIContextMenu) {
				hr = contextMenuData.pIContextMenu->QueryContextMenu(contextMenuData.currentShellContextMenu, 0, MIN_CONTEXTMENUEXTENSION_CMDID, MAX_CONTEXTMENUEXTENSION_CMDID, contextMenuFlags);
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

HRESULT ShellListView::CreateShellContextMenu(PLONG pItems, UINT itemCount, UINT predefinedFlags, HMENU& hMenu)
{
	ATLASSERT_ARRAYPOINTER(pItems, LONG, itemCount);
	if(!pItems) {
		return E_POINTER;
	}

	DestroyShellContextMenu();

	PIDLIST_ABSOLUTE* ppIDLs = NULL;
	UINT pIDLCount = ItemIDsToPIDLs(pItems, itemCount, ppIDLs, FALSE);
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
	UINT contextMenuFlags = predefinedFlags;
	if(GetKeyState(VK_SHIFT) & 0x8000) {
		contextMenuFlags |= CMF_EXTENDEDVERBS;
		contextMenuData.bufferedShiftDown = TRUE;
	}
	if(GetKeyState(VK_CONTROL) & 0x8000) {
		contextMenuData.bufferedCtrlDown = TRUE;
	}

	if(singleNamespace) {
		CComPtr<IShellFolder> pParentISF = NULL;
		hr = ::SHBindToParent(ppIDLs[0], IID_PPV_ARGS(&pParentISF), NULL);
		ATLTRACE2(shlvwTraceContextMenus, 3, TEXT("SHBindToParent returned 0x%X\n"), hr);
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

				if(attachedControl.GetStyle() & LVS_EDITLABELS) {
					SFGAOF attributes = SFGAO_CANRENAME;
					hr = pParentISF->GetAttributesOf(pIDLCount, pRelativePIDLs, &attributes);
					if(SUCCEEDED(hr)) {
						if(attributes & SFGAO_CANRENAME) {
							contextMenuFlags |= CMF_CANRENAME;
						}
					}
				}
				if(FAILED(pParentISF->GetUIObjectOf(GethWndShellUIParentWindow(), pIDLCount, pRelativePIDLs, IID_IContextMenu, NULL, reinterpret_cast<LPVOID*>(&contextMenuData.pIContextMenu)))) {
					contextMenuData.pIContextMenu = NULL;
				}
				if(pRelativePIDLs) {
					delete[] pRelativePIDLs;
				}
				hr = S_OK;
			} else {
				return E_OUTOFMEMORY;
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
				contextMenuFlags |= CMF_CANRENAME;
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
			contextMenuFlags |= CMF_ITEMMENU;
		}

		ATLASSUME(properties.pAttachedInternalMessageListener);
		EXLVMCREATEITEMCONTAINERDATA containerData = {0};
		containerData.structSize = sizeof(EXLVMCREATEITEMCONTAINERDATA);
		containerData.numberOfItems = itemCount;
		containerData.pItems = pItems;
		ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_CREATEITEMCONTAINER, 0, reinterpret_cast<LPARAM>(&containerData))));
		if(containerData.pContainer) {
			contextMenuData.pContextMenuItems = containerData.pContainer;
		}

		// allow customization of 'contextMenuFlags'
		ShellContextMenuStyleConstants contextMenuStyle = static_cast<ShellContextMenuStyleConstants>(contextMenuFlags);
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_CreatingShellContextMenu(contextMenuData.pContextMenuItems, &contextMenuStyle, &cancel);
		contextMenuFlags = static_cast<UINT>(contextMenuStyle);

		if((cancel == VARIANT_FALSE) && contextMenuData.currentShellContextMenu.CreatePopupMenu()) {
			// fill the menu
			if(contextMenuData.pIContextMenu) {
				hr = contextMenuData.pIContextMenu->QueryContextMenu(contextMenuData.currentShellContextMenu, 0, MIN_CONTEXTMENUEXTENSION_CMDID, MAX_CONTEXTMENUEXTENSION_CMDID, contextMenuFlags);
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

HRESULT ShellListView::CreateHeaderContextMenu(ULONG defaultDisplayColumn, HMENU& hMenu)
{
	DestroyHeaderContextMenu();

	if(properties.columnsStatus.loaded == SCLS_LOADED && contextMenuData.currentHeaderContextMenu.CreatePopupMenu()) {
		// fill the menu
		BOOL insertMoreItem = FALSE;
		#define MAXITEMS 25

		CMenuItemInfo menuItemInfo;
		menuItemInfo.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
		UINT insertedItems = 0;
		for(UINT i = 0; i < min(MAXITEMS, properties.columnsStatus.numberOfAllColumns); ++i) {
			SHCOLSTATEF columnState = properties.columnsStatus.pAllColumns[i].state;
			if(!(columnState & SHCOLSTATE_HIDDEN)) {
				if(columnState & SHCOLSTATE_SECONDARYUI) {
					insertMoreItem = TRUE;
				} else {
					menuItemInfo.fState = (i == defaultDisplayColumn ? MFS_DISABLED : MFS_ENABLED) | (properties.columnsStatus.pAllColumns[i].columnID != -1 ? MFS_CHECKED : MFS_UNCHECKED);
					menuItemInfo.wID = i;
					menuItemInfo.dwTypeData = properties.columnsStatus.pAllColumns[i].pText;
					menuItemInfo.cch = lstrlen(menuItemInfo.dwTypeData);
					contextMenuData.currentHeaderContextMenu.InsertMenuItem(i, TRUE, &menuItemInfo);
					++insertedItems;
				}
			}
		}
		insertMoreItem = insertMoreItem || (insertedItems < properties.columnsStatus.numberOfAllColumns);

		if(insertMoreItem) {
			#define IDS_MORE 0x1099

			TCHAR pBuffer[MAX_PATH];
			HINSTANCE hShell32 = LoadLibraryEx(TEXT("shell32.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
			if(hShell32) {
				LoadString(hShell32, IDS_MORE, pBuffer, MAX_PATH);
				FreeLibrary(hShell32);
			}

			menuItemInfo.fMask = MIIM_FTYPE;
			menuItemInfo.fType = MFT_SEPARATOR;
			contextMenuData.currentHeaderContextMenu.InsertMenuItem(properties.columnsStatus.numberOfAllColumns, TRUE, &menuItemInfo);
			menuItemInfo.fMask = MIIM_ID | MIIM_STRING;
			menuItemInfo.wID = properties.columnsStatus.numberOfAllColumns;
			menuItemInfo.dwTypeData = pBuffer;
			menuItemInfo.cch = lstrlen(menuItemInfo.dwTypeData);
			contextMenuData.currentHeaderContextMenu.InsertMenuItem(properties.columnsStatus.numberOfAllColumns + 1, TRUE, &menuItemInfo);
		}

		// allow custom items
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_CreatedHeaderContextMenu(HandleToLong(static_cast<HMENU>(contextMenuData.currentHeaderContextMenu)), properties.columnsStatus.numberOfAllColumns + 1, &cancel);
		if(cancel == VARIANT_FALSE) {
			PrettifyMenu(contextMenuData.currentHeaderContextMenu);
			hMenu = contextMenuData.currentHeaderContextMenu;
		}
		return S_OK;
	}
	return E_FAIL;
}

HRESULT ShellListView::DisplayShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, POINT& position)
{
	properties.contextMenuPosition.x = -1;
	properties.contextMenuPosition.y = -1;

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
						//if(properties.autoEditNewItems) {
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
						//}

						/* NOTE: In order to be able to position the created item at the clicked point, we need
						 *       everything that we're doing to automatically enter label-edit mode for this item.
						 *       So we check for properties.autoEditNewItems just when we would enter label-edit
						 *       mode in OnTriggerItemEnumComplete().
						 */
						properties.timeOfLastItemCreatorInvocation = 0;
						if(isSubItemOfNewMenu) {
							properties.timeOfLastItemCreatorInvocation = GetTickCount();

							properties.contextMenuPosition = position;
							attachedControl.ScreenToClient(&properties.contextMenuPosition);
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

HRESULT ShellListView::DisplayShellContextMenu(PLONG pItems, UINT itemCount, POINT& position)
{
	HMENU hMenu;
	HRESULT hr = CreateShellContextMenu(pItems, itemCount, CMF_NORMAL, hMenu);
	if(SUCCEEDED(hr) && contextMenuData.currentShellContextMenu.IsMenu()) {
		// display the menu
		int caretItemIndex = -1;
		LVITEM item = {0};
		caretItemIndex = ItemIDToIndex(pItems[0]);
		item.stateMask = LVIS_DROPHILITED;
		item.state = LVIS_DROPHILITED;
		attachedControl.SendMessage(LVM_SETITEMSTATE, caretItemIndex, reinterpret_cast<LPARAM>(&item));
		UINT chosenCommand = contextMenuData.currentShellContextMenu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, position.x, position.y, attachedControl);
		item.state = 0;
		attachedControl.SendMessage(LVM_SETITEMSTATE, caretItemIndex, reinterpret_cast<LPARAM>(&item));

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
							invoked = (attachedControl.SendMessage(LVM_EDITLABEL, caretItemIndex, 0) != NULL);
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
							LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pItems[0]);
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

HRESULT ShellListView::DisplayHeaderContextMenu(POINT& position)
{
	ULONG defaultDisplayColumn = 0;
	// Windows Explorer always prevents column 0 from being removed
	/*ULONG dummy;
	if(properties.columnsStatus.pNamespaceSHF2) {
		properties.columnsStatus.pNamespaceSHF2->GetDefaultColumn(0, &dummy, &defaultDisplayColumn);
	}*/

	HMENU hMenu;
	HRESULT hr = CreateHeaderContextMenu(defaultDisplayColumn, hMenu);
	if(SUCCEEDED(hr)) {
		// display the menu
		UINT chosenCommand = contextMenuData.currentHeaderContextMenu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, position.x, position.y, attachedControl);
		if(chosenCommand != 0) {
			VARIANT_BOOL executeCommand = VARIANT_TRUE;
			Raise_SelectedHeaderContextMenuItem(HandleToLong(static_cast<HMENU>(contextMenuData.currentHeaderContextMenu)), chosenCommand, &executeCommand);
			if(executeCommand != VARIANT_FALSE) {
				if((chosenCommand >= 0) && (chosenCommand <= properties.columnsStatus.numberOfAllColumns)) {
					BOOL invoked = FALSE;
					if(chosenCommand /*- 0*/ == properties.columnsStatus.numberOfAllColumns) {
						SetColumnsDlg dlg;
						dlg.ShowDialog(this, GethWndShellUIParentWindow(), attachedControl, properties.columnsStatus.numberOfAllColumns, properties.columnsStatus.pAllColumns, defaultDisplayColumn);
						invoked = TRUE;
					} else {
						CMenuItemInfo menuItemInfo;
						menuItemInfo.fMask = MIIM_STATE;
						contextMenuData.currentHeaderContextMenu.GetMenuItemInfo(chosenCommand, FALSE, &menuItemInfo);
						ChangeColumnVisibility(chosenCommand /*- 0*/, (menuItemInfo.fState & MFS_CHECKED) == 0, CCVF_ISEXPLICITCHANGE);
						invoked = TRUE;
					}

					if(invoked) {
						Raise_InvokedHeaderContextMenuCommand(HandleToLong(static_cast<HMENU>(contextMenuData.currentHeaderContextMenu)), chosenCommand);
					}
				}
			}
		}
		ATLVERIFY(SUCCEEDED(DestroyHeaderContextMenu()));
	}

	return hr;
}

HRESULT ShellListView::InvokeDefaultShellContextMenuCommand(PLONG pItems, UINT itemCount)
{
	HMENU hMenu;
	HRESULT hr = CreateShellContextMenu(pItems, itemCount, CMF_DEFAULTONLY, hMenu);
	if(SUCCEEDED(hr)) {
		ATLASSERT(contextMenuData.currentShellContextMenu.IsMenu());
		int caretItemIndex = -1;
		caretItemIndex = ItemIDToIndex(pItems[0]);
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
							invoked = (attachedControl.SendMessage(LVM_EDITLABEL, caretItemIndex, 0) != NULL);
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
							LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(pItems[0]);
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


HRESULT ShellListView::DeregisterForShellNotifications(void)
{
	if(!attachedControl.IsWindow() || !shellNotificationData.IsRegistered()) {
		return S_OK;
	}

	if(SHChangeNotifyDeregister(shellNotificationData.registrationID)) {
		ATLTRACE2(shlvwTraceAutoUpdate, 3, TEXT("Deregistered 0x%X\n"), shellNotificationData.registrationID);
		return S_OK;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Failed to deregister 0x%X\n"), shellNotificationData.registrationID);
	return E_FAIL;
}

HRESULT ShellListView::RegisterForShellNotifications(void)
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
	ATLTRACE2(shlvwTraceAutoUpdate, 3, TEXT("SHChangeNotifyRegister returned 0x%X\n"), shellNotificationData.registrationID);
	return (shellNotificationData.registrationID > 0 ? S_OK : E_FAIL);
}


LONG ShellListView::FindItem(PCIDLIST_ABSOLUTE pIDLToCheck, BOOL simplePIDL, BOOL exactMatch/* = FALSE*/)
{
	ATLASSERT_POINTER(pIDLToCheck, ITEMIDLIST_ABSOLUTE);

	BOOL freePIDL = FALSE;
	if(simplePIDL && pIDLToCheck) {
		pIDLToCheck = SimpleToRealPIDL(pIDLToCheck);
		freePIDL = TRUE;
	}
	LONG itemID = PIDLToItemID(pIDLToCheck, exactMatch);
	if(freePIDL) {
		ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDLToCheck));
	}
	return itemID;
}

void ShellListView::FlushIcons(int iconIndex, BOOL flushOverlays, BOOL dontRebuildSysImageList)
{
	BOOL fullFlush = (iconIndex == -1);
	ShLvwManagedItemPropertiesConstants flagsToCheckFor;
	if(fullFlush) {
		flagsToCheckFor = static_cast<ShLvwManagedItemPropertiesConstants>(slmipIconIndex | slmipOverlayIndex);
		if(!dontRebuildSysImageList) {
			if(properties.useSystemImageList & static_cast<UseSystemImageListConstants>(usilSmallImageList | usilLargeImageList | usilExtraLargeImageList)) {
				if(!RunTimeHelperEx::IsWin8()) {
					// flush the system imagelist
					if(properties.useSystemImageList & usilSmallImageList) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 1/*ilSmall*/, NULL)));
					}
					if(properties.useSystemImageList & usilLargeImageList) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 2/*ilLarge*/, NULL)));
					}
					if(properties.useSystemImageList & usilExtraLargeImageList) {
						ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_SETIMAGELIST, 3/*ilExtraLarge*/, NULL)));
					}
					FlushSystemImageList();
					SetSystemImageLists();
				}
			}

			SHFILEINFO fileInfo = {0};
			SHGetFileInfo(TEXT(".zyxwv12"), FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
			properties.DEFAULTICON_FILE = fileInfo.iIcon;
			SHGetFileInfo(TEXT("folder"), FILE_ATTRIBUTE_DIRECTORY, &fileInfo, sizeof(SHFILEINFO), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
			properties.DEFAULTICON_FOLDER = fileInfo.iIcon;
		}
	} else {
		flagsToCheckFor = slmipIconIndex;
		if(!dontRebuildSysImageList) {
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
	}
	HIMAGELIST hImageList = NULL;
	ATLVERIFY(SUCCEEDED(properties.pAttachedInternalMessageListener->ProcessMessage(EXLVM_GETIMAGELIST, 0/*ilCurrentView*/, reinterpret_cast<LPARAM>(&hImageList))));

	CComPtr<IImageListPrivate> pImgLstPriv1 = NULL;
	CComPtr<IImageListPrivate> pImgLstPriv2 = NULL;
	if(properties.displayThumbnails && properties.thumbnailsStatus.pThumbnailImageList) {
		properties.thumbnailsStatus.pThumbnailImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv1));
		ATLASSUME(pImgLstPriv1);
	}
	if(properties.pUnifiedImageList) {
		properties.pUnifiedImageList->QueryInterface(IID_IImageListPrivate, reinterpret_cast<LPVOID*>(&pImgLstPriv2));
		ATLASSUME(pImgLstPriv2);
	}

	LVITEM item = {0};
	#ifdef USE_STL
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			item.iItem = ItemIDToIndex(itemID);
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		POSITION p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			item.iItem = ItemIDToIndex(itemID);
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(item.iItem == -1) {
			// an error occurred
			continue;
		}

		if(pItemDetails->managedProperties & flagsToCheckFor) {
			item.mask = 0;
			item.stateMask = 0;
			item.state = 0;
			if(pItemDetails->managedProperties & slmipIconIndex) {
				item.mask |= LVIF_IMAGE;
			}
			if(flushOverlays && (pItemDetails->managedProperties & slmipOverlayIndex)) {
				item.mask |= LVIF_STATE;
				item.stateMask = LVIS_OVERLAYMASK;
			}

			BOOL update = fullFlush;
			if(fullFlush) {
				item.iImage = I_IMAGECALLBACK;
				if(flushOverlays && (pItemDetails->managedProperties & slmipOverlayIndex)) {
					// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
					BOOL retrieveOverlayNow = !properties.loadOverlaysOnDemand || !(pItemDetails->managedProperties & slmipIconIndex) || !hImageList;
					if(retrieveOverlayNow) {
						// load the overlay here and now
						int overlayIndex = GetOverlayIndex(pItemDetails->pIDL);
						if(overlayIndex > 0) {
							item.state = INDEXTOOVERLAYMASK(overlayIndex);
						}
					}
				}
			} else {
				if(pImgLstPriv1) {
					if(pItemDetails->managedProperties & slmipIconIndex) {
						int systemIconIndex = -1;
						int executableIconIndex = -1;
						// TODO - TODO what???
						if(SUCCEEDED(pImgLstPriv1->GetIconInfo(itemID, SIIF_SYSTEMICON | SIIF_OVERLAYIMAGE, &systemIconIndex, &executableIconIndex, NULL, NULL))) {
							if(systemIconIndex == iconIndex || executableIconIndex == iconIndex) {
								item.iImage = I_IMAGECALLBACK;
								update = TRUE;
							}
						}
					}
				} else if(pImgLstPriv2) {
					if(pItemDetails->managedProperties & slmipIconIndex) {
						int systemIconIndex = -1;
						// TODO - TODO what???
						if(SUCCEEDED(pImgLstPriv2->GetIconInfo(itemID, SIIF_SYSTEMICON, &systemIconIndex, NULL, NULL, NULL))) {
							if(systemIconIndex == iconIndex) {
								item.iImage = I_IMAGECALLBACK;
								update = TRUE;
							}
						}
					}
				} else {
					attachedControl.SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
					if((pItemDetails->managedProperties & slmipIconIndex) && (item.iImage == iconIndex)) {
						item.iImage = I_IMAGECALLBACK;
						update = TRUE;
					}
				}

				if(flushOverlays && (pItemDetails->managedProperties & slmipOverlayIndex)) {
					// we can delay-set the state only if an icon is delay-loaded and the control is displaying images
					BOOL retrieveOverlayNow = !properties.loadOverlaysOnDemand || !(pItemDetails->managedProperties & slmipIconIndex) || !hImageList;
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
				attachedControl.SendMessage(LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item));
			}
		}
	}
}


LRESULT ShellListView::OnSHChangeNotify_ASSOCCHANGED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}

	if(ILIsEmpty(ppIDLs[0])) {
		// changed icon
		ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED - flushing all icons\n"));
		FlushIcons(-1, FALSE, FALSE);
	} else {
		if(ppIDLs[0] && reinterpret_cast<LPSHChangeDWORDAsIDList>(const_cast<PIDLIST_ABSOLUTE>(ppIDLs[0]))->dwItem1 == 0) {
			// the desktop might have a new icon or lost one (happens on Windows 2000 if TweakUI is used)
			PIDLIST_ABSOLUTE pIDLDesktop = GetDesktopPIDL();
			// update namespaces
			#ifdef USE_STL
				std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
				for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
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
					CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
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
				std::vector<LONG> itemsToRemove;
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					LONG itemID = iter->first;
					LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
			#else
				CAtlArray<LONG> itemsToRemove;
				p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					LONG itemID = pPair->m_key;
					LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(ILIsParent(pIDLDesktop, pItemDetails->pIDL, FALSE)) {
					// validate the item
					// TODO: ValidatePIDL doesn't return FALSE if the shell item has been hidden
					if(!ValidatePIDL(pItemDetails->pIDL)) {
						// remove the item
						#ifdef USE_STL
							itemsToRemove.push_back(itemID);
						#else
							itemsToRemove.Add(itemID);
						#endif
					}
				}
			}
			#ifdef USE_STL
				for(size_t i = 0; i < itemsToRemove.size(); ++i) {
			#else
				for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
			#endif
				int itemIndex = ItemIDToIndex(itemsToRemove[i]);
				if(itemIndex != -1) {
					attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
					/* The control should have sent us a notification about the deletion. The notification handler
						 should have freed the pIDL. */
				}
			}
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
				for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
					LONG itemID = iter->first;
					LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
			#else
				POSITION p = properties.items.GetStartPosition();
				while(p) {
					CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
					LONG itemID = pPair->m_key;
					LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
					properties.items.GetNext(p);
			#endif
				if(ILIsEqual(pItemDetails->pIDL, pIDLMyDocs)) {
					PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
					pItemDetails->pIDL = ILCloneFull(pIDLMyDocs);
					if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
						Raise_ChangedItemPIDL(itemID, tmp, pItemDetails->pIDL);
					}
					ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				}
			}
			ILFree(pIDLMyDocs);

			// re-sort desktop namespace
			#ifdef USE_STL
				for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
					if(ILIsEmpty(iter->first)) {
						AutoSortItemsConstants autoSort = asiNoAutoSort;
						iter->second->get_AutoSortItems(&autoSort);
						if(autoSort == asiAutoSort) {
							SortItems(sortingSettings.columnIndex);
							break;
						}
					}
				}
			#else
				p = properties.namespaces.GetStartPosition();
				while(p) {
					CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
					if(ILIsEmpty(pPair->m_key)) {
						AutoSortItemsConstants autoSort = asiNoAutoSort;
						pPair->m_value->get_AutoSortItems(&autoSort);
						if(autoSort == asiAutoSort) {
							SortItems(sortingSettings.columnIndex);
							break;
						}
					}
					properties.namespaces.GetNext(p);
				}
			#endif
		} else {
			// changed file type link
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling SHCNE_ASSOCCHANGED - flushing all icons\n"));
			FlushIcons(-1, FALSE, FALSE);
		}
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellListView::OnSHChangeNotify_CHANGEDSHARE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling changed share ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling changed share pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsParent(ppIDLs[0], iter->first, FALSE)) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsParent(ppIDLs[0], pPair->m_key, FALSE)) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<LONG> itemsToUpdate;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(itemID);
			#else
				itemsToUpdate.Add(itemID);
			#endif
			// update the pIDL
			UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
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

LRESULT ShellListView::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item creation ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item creation pIDL=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
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
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	// update all namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && (ILIsEqual(pIDLRecycler, iter->first) || ILIsParent(pIDLRecycler, iter->first, FALSE))) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be equal or child of global recycler\n"), iter->first);
				iter->second->UpdateEnumeration();
			} else {
				iter->second->OnSHChangeNotify_CREATE(ppIDLs[0]);
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && (ILIsEqual(pIDLRecycler, pPair->m_key) || ILIsParent(pIDLRecycler, pPair->m_key, FALSE))) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be equal or child of global recycler\n"), pPair->m_key);
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

LRESULT ShellListView::OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir, PCIDLIST_ABSOLUTE pIDLNamespace, INamespaceEnumSettings* pEnumSettings)
{
	ATLASSERT_POINTER(simplePIDLAddedDir, ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(pIDLNamespace, ITEMIDLIST_ABSOLUTE);
	ATLASSUME(pEnumSettings);

	BOOL autoLabelEdit = FALSE;
	if(GetTickCount() - properties.timeOfLastItemCreatorInvocation <= 250) {
		autoLabelEdit = TRUE;
	} else {
		properties.contextMenuPosition.x = -1;
		properties.contextMenuPosition.y = -1;
	}

	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShLvwInsertSingleItemTask::CreateInstance(attachedControl, &properties.backgroundItemEnumResults, &properties.backgroundItemEnumResultsCritSection, -1, simplePIDLAddedDir, TRUE, pIDLNamespace, pEnumSettings, pIDLNamespace, TRUE, autoLabelEdit, &pTask);
	if(SUCCEEDED(hr)) {
		CComObject<ShellListViewNamespace>* pNamespaceObj = NamespacePIDLToNamespace(pIDLNamespace, TRUE);
		CComQIPtr<IDispatch> pNamespace = pNamespaceObj;
		INamespaceEnumTask* pEnumTask = NULL;
		ATLVERIFY(SUCCEEDED(pTask->QueryInterface(IID_INamespaceEnumTask, reinterpret_cast<LPVOID*>(&pEnumTask))));
		ATLASSUME(pEnumTask);
		if(pEnumTask) {
			ATLVERIFY(SUCCEEDED(pEnumTask->SetTarget(pNamespace)));
			pEnumTask->Release();
		}
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShLvwAutoUpdate, 0, TASK_PRIORITY_LV_BACKGROUNDENUM)));
	}
	return 0;
}

LRESULT ShellListView::OnSHChangeNotify_DELETE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item deletion ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, (ppIDLs[1] ? 2 : 1));
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item deletion pIDL1=0x%X (count: %i), pIDL2=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
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
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	// update all namespaces
	#ifdef USE_STL
		std::vector<CComObject<ShellListViewNamespace>*> namespacesToCheck;
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && ILIsParent(pIDLRecycler, iter->first, FALSE)) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), iter->first);
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
		for(std::vector<CComObject<ShellListViewNamespace>*>::const_iterator iter = namespacesToCheck.cbegin(); iter != namespacesToCheck.cend(); ++iter) {
			(*iter)->OnSHChangeNotify_DELETE(ppIDLs[0]);
		}
	#else
		CAtlArray<CComObject<ShellListViewNamespace>*> namespacesToCheck;
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pPair->m_key, FALSE)) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), pPair->m_key);
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
		std::vector<LONG> itemsToRemove;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToRemove;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// remove the item
			#ifdef USE_STL
				itemsToRemove.push_back(itemID);
			#else
				itemsToRemove.Add(itemID);
			#endif
		} else if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pItemDetails->pIDL, FALSE)) {
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found item %i to be child of global recycler\n"), itemID);
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(itemID);
				#else
					itemsToRemove.Add(itemID);
				#endif
			}
		}
	}
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToRemove.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
	#endif
		int itemIndex = ItemIDToIndex(itemsToRemove[i]);
		if(itemIndex != -1) {
			attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	}
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

LRESULT ShellListView::OnSHChangeNotify_DRIVEADD(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling new drive ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling new drive pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0]) || ILIsEqual(iter->first, pIDLMyComputer)) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0]) || ILIsEqual(pPair->m_key, pIDLMyComputer)) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<LONG> itemsToUpdate;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(itemID);
			#else
				itemsToUpdate.Add(itemID);
			#endif
			// update the pIDL
			UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
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

LRESULT ShellListView::OnSHChangeNotify_DRIVEREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling drive removal ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling drive removal pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();

	// update namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
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
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
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
		std::vector<LONG> itemsToRemove;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToRemove;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(itemID);
				#else
					itemsToRemove.Add(itemID);
				#endif
			}
		}
	}
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToRemove.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
	#endif
		int itemIndex = ItemIDToIndex(itemsToRemove[i]);
		if(itemIndex != -1) {
			attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	}
	ILFree(pIDLMyComputer);

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellListView::OnSHChangeNotify_FREESPACE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling change of free drive space ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling change of free drive space pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLMyComputer = GetMyComputerPIDL();

	if(CurrentViewNeedsColumns()) {
		// update the items
		#ifdef USE_STL
			std::vector<LONG> itemsToUpdate;
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				LONG itemID = iter->first;
				LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
		#else
			CAtlArray<LONG> itemsToUpdate;
			POSITION p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				LONG itemID = pPair->m_key;
				LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
				properties.items.GetNext(p);
		#endif
			if(ILIsParent(pIDLMyComputer, pItemDetails->pIDL, TRUE)) {
				// re-apply any managed properties
				#ifdef USE_STL
					itemsToUpdate.push_back(itemID);
				#else
					itemsToUpdate.Add(itemID);
				#endif
				// not necessary here
				//UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
			}
		}
		ILFree(pIDLMyComputer);

		// update the item's properties
		#ifdef USE_STL
			for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
		#else
			for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
		#endif
			ApplyManagedProperties(itemsToUpdate[i]);
		}
	}

	// NOTE: ppIDL[0] seems to be freed by the shell
	return 0;
}

LRESULT ShellListView::OnSHChangeNotify_MEDIAINSERTED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling media insertion ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling media insertion pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(ILIsEqual(iter->first, ppIDLs[0])) {
				iter->second->UpdateEnumeration();
			}
		}
	#else
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(ILIsEqual(pPair->m_key, ppIDLs[0])) {
				pPair->m_value->UpdateEnumeration();
			}
			properties.namespaces.GetNext(p);
		}
	#endif

	// update the items
	#ifdef USE_STL
		std::vector<LONG> itemsToUpdate;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(itemID);
			#else
				itemsToUpdate.Add(itemID);
			#endif
			UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
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

LRESULT ShellListView::OnSHChangeNotify_MEDIAREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling media removal ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling media removal pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
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
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
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
		std::vector<LONG> itemsToRemove;
		std::vector<LONG> itemsToUpdate;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToRemove;
		CAtlArray<LONG> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0])) {
			// re-apply any managed properties
			// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
			#ifdef USE_STL
				itemsToUpdate.push_back(itemID);
			#else
				itemsToUpdate.Add(itemID);
			#endif
			UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
		} else if(ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			#ifdef USE_STL
				itemsToRemove.push_back(itemID);
			#else
				itemsToRemove.Add(itemID);
			#endif
		}
	}
	ILFree(pIDLNew);

	#ifdef USE_STL
		for(size_t i = 0; i < itemsToRemove.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
	#endif
		int itemIndex = ItemIDToIndex(itemsToRemove[i]);
		if(itemIndex >= 0) {
			attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	}

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

LRESULT ShellListView::OnSHChangeNotify_RENAME(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item rename ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 2);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item rename pIDL1=0x%X (items: %i), pIDL2=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	ATLASSERT_POINTER(ppIDLs[1], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0] || !ppIDLs[1]) {
		// shouldn't happen
		return 0;
	}

	#ifdef USE_STL
		std::vector<LONG> itemsToUpdate;
		std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacePIDLsToReplace;
		std::vector<CComObject<ShellListViewNamespace>*> namespacesToUpdate;
	#else
		CAtlArray<LONG> itemsToUpdate;
		CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacePIDLsToReplace;
		CAtlArray<CComObject<ShellListViewNamespace>*> namespacesToUpdate;
		POSITION p;
	#endif

	PIDLIST_ABSOLUTE pIDLOld = ILCloneFull(ppIDLs[0]);
	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[1]);
	BOOL isMove = !HaveSameParent(ppIDLs, 2);
	if(isMove) {
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
			for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
				PCIDLIST_ABSOLUTE pIDLNamespaceOld = iter->first;
				CComObject<ShellListViewNamespace>* pNamespaceObj = iter->second;
		#else
			p = properties.namespaces.GetStartPosition();
			while(p) {
				CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
				PCIDLIST_ABSOLUTE pIDLNamespaceOld = pPair->m_key;
				CComObject<ShellListViewNamespace>* pNamespaceObj = pPair->m_value;
				properties.namespaces.GetNext(p);
		#endif
			if(ILIsParent(pIDLNamespaceOld, pIDLNew, TRUE)) {
				// maybe add the item
				/* NOTE: OnTriggerItemEnumComplete will sort the items. If this method isn't reached, it is because
				         ShLvwBackgroundItemEnumTask produced a serious error. */
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
			for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
				LONG itemID = iter->first;
				LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;

				PCIDLIST_ABSOLUTE pIDLNamespaceNew = NULL;
				std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter2 = namespacePIDLsToReplace.find(pItemDetails->pIDLNamespace);
				if(iter2 != namespacePIDLsToReplace.cend()) {
					pIDLNamespaceNew = iter2->second;
				}
		#else
			p = properties.items.GetStartPosition();
			while(p) {
				CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
				LONG itemID = pPair->m_key;
				LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
				properties.items.GetNext(p);

				PCIDLIST_ABSOLUTE pIDLNamespaceNew = NULL;
				namespacePIDLsToReplace.Lookup(pItemDetails->pIDLNamespace, pIDLNamespaceNew);
		#endif
			if(pIDLNamespaceNew) {
				// exchange the namespace pIDL
				pItemDetails->pIDLNamespace = pIDLNamespaceNew;
			}
			// update the item's pIDL
			if(ILIsEqual(pItemDetails->pIDL, pIDLOld)) {
				// exchange the pIDL and re-apply any managed properties
				PCIDLIST_ABSOLUTE tmp = pItemDetails->pIDL;
				pItemDetails->pIDL = ILCloneFull(pIDLNew);
				if(!(properties.disabledEvents & deItemPIDLChangeEvents)) {
					Raise_ChangedItemPIDL(itemID, tmp, pItemDetails->pIDL);
				}
				ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
				#ifdef USE_STL
					itemsToUpdate.push_back(itemID);
				#else
					itemsToUpdate.Add(itemID);
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
					Raise_ChangedItemPIDL(itemID, tmp, pItemDetails->pIDL);
				}
				ILFree(const_cast<PIDLIST_ABSOLUTE>(tmp));
				// ApplyManagedProperties will remove the item if it doesn't match the filters anymore
				#ifdef USE_STL
					itemsToUpdate.push_back(itemID);
				#else
					itemsToUpdate.Add(itemID);
				#endif
			}
		}
	}

	// store the new namespace pIDLs
	#ifdef USE_STL
		for(std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::const_iterator iter = namespacePIDLsToReplace.cbegin(); iter != namespacePIDLsToReplace.cend(); ++iter) {
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter2 = properties.namespaces.find(iter->first);
			if(iter2 != properties.namespaces.cend()) {
				iter2->second->properties.pIDL = iter->second;
				properties.namespaces[iter->second] = iter2->second;
				properties.namespaces.erase(iter->first);
				if(properties.columnsStatus.pIDLNamespace == iter->first) {
					SetMainNamespace(iter->second);
					if(properties.autoInsertColumns && CurrentViewNeedsColumns()) {
						// add new shell columns
						EnsureShellColumnsAreLoaded();
					}
				}
				if(!(properties.disabledEvents & deNamespacePIDLChangeEvents)) {
					CComQIPtr<IShListViewNamespace> pNamespaceItf = iter2->second;
					Raise_ChangedNamespacePIDL(pNamespaceItf, iter->first, iter->second);
				}
			}
			ILFree(const_cast<PIDLIST_ABSOLUTE>(iter->first));
		}
	#else
		p = namespacePIDLsToReplace.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE>::CPair* pPair = namespacePIDLsToReplace.GetAt(p);
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair2 = properties.namespaces.Lookup(pPair->m_key);
			pPair2->m_value->properties.pIDL = pPair->m_value;
			properties.namespaces[pPair->m_value] = pPair2->m_value;
			properties.namespaces.RemoveKey(pPair->m_key);
			if(properties.columnsStatus.pIDLNamespace == pPair->m_key) {
				SetMainNamespace(pPair->m_value);
				if(properties.autoInsertColumns && CurrentViewNeedsColumns()) {
					// add new shell columns
					EnsureShellColumnsAreLoaded();
				}
			}
			if(!(properties.disabledEvents & deNamespacePIDLChangeEvents)) {
				CComQIPtr<IShListViewNamespace> pNamespaceItf = pPair2->m_value;
				Raise_ChangedNamespacePIDL(pNamespaceItf, pPair->m_key, pPair->m_value);
			}

			ILFree(const_cast<PIDLIST_ABSOLUTE>(pPair->m_key));
			namespacePIDLsToReplace.GetNext(p);
		}
	#endif
	ILFree(pIDLOld);
	ILFree(pIDLNew);

	// update the item's properties
	#ifdef USE_STL
		for(size_t i = 0; i < itemsToUpdate.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToUpdate.GetCount(); ++i) {
	#endif
		ApplyManagedProperties(itemsToUpdate[i]);
	}

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

LRESULT ShellListView::OnSHChangeNotify_UPDATEDIR(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling directory update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling directory update pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
	ATLASSERT_POINTER(ppIDLs[0], ITEMIDLIST_ABSOLUTE);
	if(!ppIDLs[0]) {
		// shouldn't happen
		return 0;
	}
	// sanity check
	// Happens on Vista: ATLASSERT(!IsSubItemOfGlobalRecycleBin(ppIDLs[0]));
	BOOL updateGlobalRecycler/* = IsSubItemOfDriveRecycleBin(ppIDLs[0])*/;   // TODO: "IsSubItemOfGlobalRecycleBin(ppIDLs[0]) || "?
	#ifdef DEBUG
		updateGlobalRecycler = IsSubItemOfDriveRecycleBin(ppIDLs[0]);
		if(updateGlobalRecycler) {
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found pIDL 0x%X to be sub-item of a drive's recycler\n"), ppIDLs[0]);
		}
	#endif
	/* HACK: If many files are restored from the recycler at once, the shell often sends a SHCNE_UPDATEDIR
	         for the target directory only. */
	updateGlobalRecycler = TRUE;
	PIDLIST_ABSOLUTE pIDLRecycler = GetGlobalRecycleBinPIDL();

	PIDLIST_ABSOLUTE pIDLNew = SimpleToRealPIDL(ppIDLs[0]);

	// update all namespaces
	#ifdef USE_STL
		std::vector<PCIDLIST_ABSOLUTE> namespacesToRemove;
		std::vector<CComObject<ShellListViewNamespace>*> namespacesToUpdate;
		for(std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator iter = properties.namespaces.cbegin(); iter != properties.namespaces.cend(); ++iter) {
			if(updateGlobalRecycler && (ILIsEqual(iter->first, pIDLRecycler) || ILIsParent(pIDLRecycler, iter->first, FALSE))) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), iter->first);
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
		for(std::vector<CComObject<ShellListViewNamespace>*>::const_iterator iter = namespacesToUpdate.cbegin(); iter != namespacesToUpdate.cend(); ++iter) {
			(*iter)->UpdateEnumeration();
		}
	#else
		CAtlArray<PCIDLIST_ABSOLUTE> namespacesToRemove;
		CAtlArray<CComObject<ShellListViewNamespace>*> namespacesToUpdate;
		POSITION p = properties.namespaces.GetStartPosition();
		while(p) {
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::CPair* pPair = properties.namespaces.GetAt(p);
			if(updateGlobalRecycler && (ILIsEqual(pPair->m_key, pIDLRecycler) || ILIsParent(pIDLRecycler, pPair->m_key, FALSE))) {
				ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found namespace 0x%X to be child of global recycler\n"), pPair->m_key);
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
		std::vector<LONG> itemsToRemove;
		std::vector<LONG> itemsToUpdate;
		for(std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator iter = properties.items.cbegin(); iter != properties.items.cend(); ++iter) {
			LONG itemID = iter->first;
			LPSHELLLISTVIEWITEMDATA pItemDetails = iter->second;
	#else
		CAtlArray<LONG> itemsToRemove;
		CAtlArray<LONG> itemsToUpdate;
		p = properties.items.GetStartPosition();
		while(p) {
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA>::CPair* pPair = properties.items.GetAt(p);
			LONG itemID = pPair->m_key;
			LPSHELLLISTVIEWITEMDATA pItemDetails = pPair->m_value;
			properties.items.GetNext(p);
	#endif
		if(ILIsEqual(pItemDetails->pIDL, ppIDLs[0]) || ILIsParent(ppIDLs[0], pItemDetails->pIDL, FALSE)) {
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				#ifdef USE_STL
					itemsToRemove.push_back(itemID);
				#else
					itemsToRemove.Add(itemID);
				#endif
			} else {
				#ifdef USE_STL
					itemsToUpdate.push_back(itemID);
				#else
					itemsToUpdate.Add(itemID);
				#endif
				// update the pIDL
				UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
			}
		} else if(updateGlobalRecycler && ILIsParent(pIDLRecycler, pItemDetails->pIDL, FALSE)) {
			ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Found item %i to be child of global recycler\n"), itemID);
			// validate the item
			if(!ValidatePIDL(pItemDetails->pIDL)) {
				// remove the item
				#ifdef USE_STL
					itemsToRemove.push_back(itemID);
				#else
					itemsToRemove.Add(itemID);
				#endif
			} else {
				#ifdef USE_STL
					itemsToUpdate.push_back(itemID);
				#else
					itemsToUpdate.Add(itemID);
				#endif
				// update the pIDL
				// TODO: Is calling UpdateItemPIDL() necessary? Which pIDL should be passed?
				//UpdateItemPIDL(itemID, pItemDetails, pIDLNew);
			}
		}
	}
	ILFree(pIDLRecycler);
	ILFree(pIDLNew);

	#ifdef USE_STL
		for(size_t i = 0; i < itemsToRemove.size(); ++i) {
	#else
		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
	#endif
		int itemIndex = ItemIDToIndex(itemsToRemove[i]);
		if(itemIndex != -1) {
			attachedControl.SendMessage(LVM_DELETEITEM, itemIndex, 0);
			/* The control should have sent us a notification about the deletion. The notification handler
			   should have freed the pIDL. */
		}
	}

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

LRESULT ShellListView::OnSHChangeNotify_UPDATEIMAGE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling image update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, (ppIDLs[1] ? 2 : 1));
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling image update pIDL1=0x%X (count: %i), pIDL2=0x%X (count: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]), ppIDLs[1], ILCount(ppIDLs[1]));
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
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling update of image %i\n"), oldIconIndex);

	// update all items
	FlushIcons(oldIconIndex, oldIconIndex == -1, FALSE);

	// NOTE: ppIDL[0] and ppIDL[1] seem to be freed by the shell
	return 0;
}

LRESULT ShellListView::OnSHChangeNotify_UPDATEITEM(PCIDLIST_ABSOLUTE_ARRAY ppIDLs)
{
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item update ppIDLs=0x%X\n"), ppIDLs);
	ATLASSERT_ARRAYPOINTER(ppIDLs, PCIDLIST_ABSOLUTE, 1);
	if(!ppIDLs) {
		// shouldn't happen
		return 0;
	}
	ATLTRACE2(shlvwTraceAutoUpdate, 1, TEXT("Handling item update pIDL1=0x%X (items: %i)\n"), ppIDLs[0], ILCount(ppIDLs[0]));
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


HRESULT ShellListView::EnqueueTask(IRunnableTask* pTask, REFTASKOWNERID taskGroupID, DWORD_PTR lParam, DWORD priority, DWORD operationStart/* = 0*/)
{
	if(!flags.acceptNewTasks) {
		return E_FAIL;
	}

	HRESULT hr;

	ATLASSUME(pTask);
	ATLASSERT(priority > 0);

	IShellTaskScheduler* pTaskSchedulerToUse = NULL;

	if(taskGroupID == TASKID_ShLvwBackgroundItemEnum || taskGroupID == TASKID_ShLvwBackgroundColumnEnum) {
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
	} else if(taskGroupID == TASKID_ShLvwThumbnail || taskGroupID == TASKID_ShLvwThumbnailTestCache || taskGroupID == TASKID_ShLvwExtractThumbnail || taskGroupID == TASKID_ShLvwExtractThumbnailFromDiskCache) {
		if(!properties.pThumbnailTaskScheduler) {
			// create the scheduler
			hr = CoCreateInstance(CLSID_ShellTaskScheduler, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&properties.pThumbnailTaskScheduler));
			ATLASSERT(SUCCEEDED(hr));
			ATLASSUME(properties.pThumbnailTaskScheduler);
			if(FAILED(hr)) {
				return hr;
			}

			// immediately stop execution of the current task when the scheduler is released
			properties.pThumbnailTaskScheduler->Status(ITSSFLAG_KILL_ON_DESTROY | ITSSFLAG_THREAD_TERMINATE_TIMEOUT, THREAD_TIMEOUT_DEFAULT);
		}
		pTaskSchedulerToUse = properties.pThumbnailTaskScheduler;
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

		if(taskGroupID == TASKID_ShLvwBackgroundItemEnum || taskGroupID == TASKID_ShLvwAutoUpdate) {
			pBackgroundEnum = new BACKGROUNDENUMERATION(BET_ITEMS, pTarget, (operationStart != 0), operationStart);
		} else if(taskGroupID == TASKID_ShLvwBackgroundColumnEnum) {
			pBackgroundEnum = new BACKGROUNDENUMERATION(BET_COLUMNS, pTarget, (operationStart != 0), operationStart);
		}
		ATLASSERT_POINTER(pBackgroundEnum, BACKGROUNDENUMERATION);
		properties.backgroundEnums[taskID] = pBackgroundEnum;
	}

	hr = pTaskSchedulerToUse->AddTask(pTask, taskGroupID, lParam, priority);
	ATLASSERT(SUCCEEDED(hr));
	#ifdef DEBUG
		if(taskGroupID == TASKID_ShLvwBackgroundColumnEnum) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added background column enum task\n"));
		} else if(taskGroupID == TASKID_ShLvwBackgroundItemEnum) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added background item enum task\n"));
		} else if(taskGroupID == TASKID_ShLvwSlowColumn) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added slow column task\n"));
		} else if(taskGroupID == TASKID_ShLvwSubItemControl) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added sub-item control task\n"));
		} else if(taskGroupID == TASKID_ShLvwTileView) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added tile view task\n"));
		} else if(taskGroupID == TASKID_ShLvwIcon) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added icon task\n"));
		} else if(taskGroupID == TASKID_ShLvwInfoTip) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added info tip task\n"));
		} else if(taskGroupID == TASKID_ShLvwOverlay) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added overlay task\n"));
		} else if(taskGroupID == TASKID_ElevationShield) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added elevation shield task\n"));
		} else if(taskGroupID == TASKID_ShLvwThumbnail) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added thumbnail task\n"));
		} else if(taskGroupID == TASKID_ShLvwThumbnailTestCache) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added thumbnail cache testing task\n"));
		} else if(taskGroupID == TASKID_ShLvwExtractThumbnail) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added thumbnail extraction task\n"));
		} else if(taskGroupID == TASKID_ShLvwExtractThumbnailFromDiskCache) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added thumbnail cache extraction task\n"));
		} else if(taskGroupID == TASKID_ShLvwAutoUpdate) {
			ATLTRACE2(shlvwTraceThreading, 4, TEXT("Added AutoUpdate related task\n"));
		}
	#endif

	if(pEnumTask) {
		pEnumTask->Release();
	}
	return hr;
}

HRESULT ShellListView::GetIconAsync(PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShLvwIconTask::CreateInstance(attachedControl, pIDL, itemID, pParentISI, properties.useGenericIcons, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShLvwIcon, 0, TASK_PRIORITY_LV_GET_ICON)));
	}
	return hr;
}

HRESULT ShellListView::GetOverlayAsync(PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIconOverlay* pParentISIO)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ShLvwOverlayTask::CreateInstance(attachedControl, pIDL, itemID, pParentISIO, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShLvwOverlay, 0, TASK_PRIORITY_LV_GET_OVERLAY)));
	}
	return hr;
}

HRESULT ShellListView::GetThumbnailAsync(PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE& thumbnailSize)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = E_FAIL;
	if(RunTimeHelper::IsVista()) {
		hr = ShLvwThumbnailTask::CreateInstance(GethWndShellUIParentWindow(), attachedControl, &properties.backgroundThumbnailResults, &properties.backgroundThumbnailResultsCritSection, pIDL, itemID, thumbnailSize, properties.displayThumbnailAdornments, properties.displayFileTypeOverlays, RunTimeHelper::IsCommCtrl6(), &pTask);
	} else {
		// use the disk cache only if the thumbnail size matches the system settings
		BOOL useDiskCache = properties.useThumbnailDiskCache;
		if(useDiskCache && properties.thumbnailSize != -1) {
			DWORD thumbsSize = 96;
			DWORD dataSize = sizeof(thumbsSize);
			SHGetValue(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer"), TEXT("ThumbnailSize"), NULL, &thumbsSize, &dataSize);
			useDiskCache = useDiskCache && (properties.thumbnailSize == static_cast<long>(thumbsSize));
		}
		LPSHELLIMAGESTORE pThumbnailDiskCache = NULL;
		if(useDiskCache) {
			// if the item is part of a namespace, use this namespace's IShellImageStore object
			LPSHELLLISTVIEWITEMDATA pItemDetails = GetItemDetails(itemID);
			if(pItemDetails && pItemDetails->pIDLNamespace) {
				CComObject<ShellListViewNamespace>* pNamespace = NamespacePIDLToNamespace(pItemDetails->pIDLNamespace, TRUE);
				if(pNamespace) {
					if(!pNamespace->properties.pThumbnailDiskCache) {
						hr = CoCreateInstance(CLSID_ShellThumbnailDiskCache, NULL, CLSCTX_INPROC, IID_IShellImageStore, reinterpret_cast<LPVOID*>(&pNamespace->properties.pThumbnailDiskCache));
						if(SUCCEEDED(hr)) {
							CComQIPtr<IPersistFolder, &IID_IPersistFolder> pPersistFolder = pNamespace->properties.pThumbnailDiskCache;
							if(pPersistFolder) {
								if(FAILED(pPersistFolder->Initialize(pItemDetails->pIDLNamespace))) {
									pNamespace->properties.pThumbnailDiskCache->Release();
									pNamespace->properties.pThumbnailDiskCache = NULL;
								}
							} else {
								pNamespace->properties.pThumbnailDiskCache->Release();
								pNamespace->properties.pThumbnailDiskCache = NULL;
							}
						}
					}
					pThumbnailDiskCache = pNamespace->properties.pThumbnailDiskCache;
					// NOTE: pThumbnailDiskCache will be AddRef'ed by ShLvwLegacyThumbnailTask
				}
			}
		}
		hr = ShLvwLegacyThumbnailTask::CreateInstance(GethWndShellUIParentWindow(), attachedControl, &properties.backgroundThumbnailResults, &properties.backgroundThumbnailResultsCritSection, &properties.tasksToEnqueue, &properties.tasksToEnqueueCritSection, pIDL, itemID, useDiskCache, pThumbnailDiskCache, thumbnailSize, properties.displayFileTypeOverlays, &pTask);
	}
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ShLvwThumbnail, 0, TASK_PRIORITY_LV_GET_ICON)));
	}
	return hr;
}

HRESULT ShellListView::GetElevationShieldAsync(LONG itemID, PCIDLIST_ABSOLUTE pIDL)
{
	CComPtr<IRunnableTask> pTask;
	HRESULT hr = ElevationShieldTask::CreateInstance(GethWndShellUIParentWindow(), attachedControl, pIDL, itemID, &pTask);
	if(SUCCEEDED(hr)) {
		ATLVERIFY(SUCCEEDED(EnqueueTask(pTask, TASKID_ElevationShield, 0, TASK_PRIORITY_GET_ELEVATIONSHIELD)));
	}
	return hr;
}

DWORD WINAPI ShellListView::DirWatchThreadProc(LPVOID lpParameter)
{
	DirectoryWatchState* pThreadState = reinterpret_cast<DirectoryWatchState*>(lpParameter);
	if(pThreadState) {
		HANDLE hWatchedDir = INVALID_HANDLE_VALUE;
		WCHAR pPath[1024];
		pPath[0] = L'\0';
		InterlockedExchange(&pThreadState->isSuspended, FALSE);
		
		const int BUFFERSIZE = 4096;
		LPVOID pBuffer = HeapAlloc(GetProcessHeap(), 0, BUFFERSIZE);
		ATLASSERT(pBuffer);
		if(pBuffer) {
			#ifndef SYNCHRONOUS_READDIRECTORYCHANGES
				if(pThreadState->hChangeEvent != INVALID_HANDLE_VALUE) {
					CloseHandle(pThreadState->hChangeEvent);
					pThreadState->hChangeEvent = INVALID_HANDLE_VALUE;
				}
				pThreadState->hChangeEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
				if(pThreadState->hChangeEvent == INVALID_HANDLE_VALUE) {
					HeapFree(GetProcessHeap(), 0, pBuffer);
					return 0;
				}
				OVERLAPPED overlapped;
				RtlZeroMemory(&overlapped, sizeof(overlapped));
				overlapped.hEvent = pThreadState->hChangeEvent;
			#endif

			while(!InterlockedCompareExchange(&pThreadState->terminate, TRUE, TRUE)) {
				if(InterlockedCompareExchange(&pThreadState->suspend, TRUE, TRUE)) {
					InterlockedExchange(&pThreadState->isSuspended, TRUE);
					SuspendThread(GetCurrentThread());
					InterlockedExchange(&pThreadState->isSuspended, FALSE);
					InterlockedExchange(&pThreadState->suspend, FALSE);
					SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_NORMAL);
					continue;
				}
				if(InterlockedCompareExchange(&pThreadState->watchNewDirectory, TRUE, TRUE)) {
					InterlockedExchange(&pThreadState->watchNewDirectory, FALSE);
					if(hWatchedDir != INVALID_HANDLE_VALUE) {
						CloseHandle(hWatchedDir);
						hWatchedDir = INVALID_HANDLE_VALUE;
					}
					PIDLIST_ABSOLUTE pIDL = reinterpret_cast<PIDLIST_ABSOLUTE>(InterlockedCompareExchangePointer(&pThreadState->pIDLWatchedNamespace, NULL, NULL));
					ATLASSERT(pIDL);
					if(pIDL) {
						if(SHGetPathFromIDListW(pIDL, pPath)) {
							#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
								hWatchedDir = CreateFileW(pPath, FILE_LIST_DIRECTORY | GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
							#else
								hWatchedDir = CreateFileW(pPath, FILE_LIST_DIRECTORY | GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, NULL);
							#endif
						}
					}
					//ATLASSERT(hWatchedDir != INVALID_HANDLE_VALUE);
					if(hWatchedDir == INVALID_HANDLE_VALUE) {
						InterlockedExchange(&pThreadState->suspend, TRUE);
						continue;
					}
				}
				
				BOOL continueWatch;
				DWORD bytesRead;
				do {
					bytesRead = 0;
					ZeroMemory(pBuffer, BUFFERSIZE);
					#ifdef SYNCHRONOUS_READDIRECTORYCHANGES
						continueWatch = ReadDirectoryChangesW(hWatchedDir, pBuffer, BUFFERSIZE, TRUE, FILE_NOTIFY_CHANGE_FILE_NAME, &bytesRead, NULL, NULL);
					#else
						continueWatch = ReadDirectoryChangesW(hWatchedDir, pBuffer, BUFFERSIZE, TRUE, FILE_NOTIFY_CHANGE_FILE_NAME, &bytesRead, &overlapped, NULL);
						if(continueWatch) {
							switch(WaitForSingleObject(pThreadState->hChangeEvent, INFINITE)) {
								case WAIT_OBJECT_0:
									continueWatch = !InterlockedCompareExchange(&pThreadState->suspend, TRUE, TRUE) && !InterlockedCompareExchange(&pThreadState->watchNewDirectory, TRUE, TRUE) && !InterlockedCompareExchange(&pThreadState->terminate, TRUE, TRUE);
									break;
								default:
									continueWatch = FALSE;
									break;
							}
						}
					#endif
					if(continueWatch) {
						PFILE_NOTIFY_INFORMATION pDetails = reinterpret_cast<PFILE_NOTIFY_INFORMATION>(pBuffer);
						if(pDetails->Action == FILE_ACTION_ADDED) {
							pDetails->FileName[pDetails->FileNameLength] = L'\0';
							WCHAR pFullPath[2048];
							lstrcpynW(pFullPath, pPath, 1024);
							PathAddBackslashW(pFullPath);
							StringCchCatW(pFullPath, 2048, pDetails->FileName);
							PIDLIST_ABSOLUTE pIDL = APIWrapper::ILCreateFromPathW_LFN(pFullPath);
							if(pIDL) {
								if(PathIsDirectoryW(pFullPath)) {
									SHChangeNotify(SHCNE_MKDIR, SHCNF_IDLIST | SHCNF_FLUSH, pIDL, NULL);
								} else {
									SHChangeNotify(SHCNE_CREATE, SHCNF_IDLIST | SHCNF_FLUSH, pIDL, NULL);
								}
								ILFree(pIDL);
							}
						}
					}
					if(continueWatch) {
						continueWatch = continueWatch && !InterlockedCompareExchange(&pThreadState->suspend, TRUE, TRUE) && !InterlockedCompareExchange(&pThreadState->watchNewDirectory, TRUE, TRUE) && !InterlockedCompareExchange(&pThreadState->terminate, TRUE, TRUE);
					}
					#ifndef SYNCHRONOUS_READDIRECTORYCHANGES
						if(continueWatch) {
							ResetEvent(pThreadState->hChangeEvent);
						}
					#endif
				} while(continueWatch);
			}
			HeapFree(GetProcessHeap(), 0, pBuffer);
		}
		if(hWatchedDir != INVALID_HANDLE_VALUE) {
			CloseHandle(hWatchedDir);
			hWatchedDir = INVALID_HANDLE_VALUE;
		}
	}
	return 0;
}


BOOL ShellListView::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}