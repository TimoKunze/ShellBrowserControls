// NamespaceEnumSettingsProperties.cpp: The controls' "Default Namespace Enum Settings" property page

#include "stdafx.h"
#include "NamespaceEnumSettingsProperties.h"


NamespaceEnumSettingsProperties::NamespaceEnumSettingsProperties()
{
	m_dwTitleID = IDS_TITLENAMESPACEENUMSETTINGSPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGNAMESPACEENUMSETTINGSPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP NamespaceEnumSettingsProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettingsProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<NamespaceEnumSettingsProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.enumerationFlagsList.SubclassWindow(GetDlgItem(IDC_ENUMERATIONFLAGSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.enumerationFlagsList.GetImageList(LVSIL_STATE));
	controls.enumerationFlagsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.enumerationFlagsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.enumerationFlagsList.AddColumn(TEXT(""), 0);

	controls.attributesList.SubclassWindow(GetDlgItem(IDC_ATTRIBUTESBOX));
	hStateImageList = SetupStateImageList(controls.attributesList.GetImageList(LVSIL_STATE));
	controls.attributesList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.attributesList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.attributesList.AddColumn(TEXT(""), 0);

	controls.inExcludedAttributesSelector.Attach(GetDlgItem(IDC_INEXCLUDEDATTRIBUTESBOX));
	controls.inExcludedAttributesSelector.AddString(TEXT("Excluded"));
	controls.inExcludedAttributesSelector.AddString(TEXT("Included"));
	controls.inExcludedAttributesSelector.SetCurSel(0);

	controls.itemTypeSelector.Attach(GetDlgItem(IDC_ITEMTYPEBOX));
	controls.itemTypeSelector.AddString(TEXT("File item"));
	controls.itemTypeSelector.AddString(TEXT("Folder item"));
	controls.itemTypeSelector.SetCurSel(0);

	controls.attributesTypeSelector.Attach(GetDlgItem(IDC_ATTRIBUTESTYPEBOX));
	controls.attributesTypeSelector.AddString(TEXT("File attributes"));
	controls.attributesTypeSelector.AddString(TEXT("Shell attributes"));
	controls.attributesTypeSelector.SetCurSel(0);

	// setup the toolbar
	CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP NamespaceEnumSettingsProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<NamespaceEnumSettingsProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT NamespaceEnumSettingsProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.enumerationFlagsList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes items that are folders."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes items that are no folders (e. g. files)."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes items that are hidden."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration may include items that are hidden following the system settings."))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes network printer objects."))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes items that can be shared."))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration includes items that are storages (i. e. can be accessed via 'IStorage') or contain storages."))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Set this flag if you want an enumeration suitable for a navigation pane.\nRequires Windows 7 or newer."))));
				break;
			case 8:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Set this flag if you want items that can be enumerated quickly.\nRequires Windows Vista or newer."))));
				break;
			case 9:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumeration retrieves the items as a simple list even if the namespace is not structured in that way.\nRequires Windows Vista or newer."))));
				break;
			case 10:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The item enumerator may return only a few items and report the other items via asynchronous shell change notifications.\nRequires Windows Vista or newer."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
	}

	delete[] pBuffer;
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}

			if(pNotificationDetails->hwndFrom == controls.enumerationFlagsList) {
				switch(pDetails->iItem) {
					case 2/*snseIncludeHiddenItems*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 2) {
							// uncheck next item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(1);
							item.stateMask = LVIS_STATEIMAGEMASK;
							::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem + 1, reinterpret_cast<LPARAM>(&item));
						}
						break;
					case 3/*snseMayIncludeHiddenItems*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 2) {
							// uncheck previous item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(1);
							item.stateMask = LVIS_STATEIMAGEMASK;
							::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem - 1, reinterpret_cast<LPARAM>(&item));
						}
						break;
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnComboBoxSelectionChanged(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	BOOL isDirty = m_bDirty;

	BufferAttributeSettings();

	properties.currentAttributesSelection = 0;
	if(controls.inExcludedAttributesSelector.GetCurSel() == 1) {
		properties.currentAttributesSelection |= 0x1;
	}
	if(controls.itemTypeSelector.GetCurSel() == 1) {
		properties.currentAttributesSelection |= 0x2;
	}
	if(controls.attributesTypeSelector.GetCurSel() == 1) {
		properties.currentAttributesSelection |= 0x4;
	}

	RestoreAttributeSettings();

	SetDirty(isDirty);
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IShListView*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IShTreeView*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT NamespaceEnumSettingsProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ShellListView Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IShListView*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ShellTreeView Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IShTreeView*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}


void NamespaceEnumSettingsProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComPtr<INamespaceEnumSettings> pEnumSettings;
		IUnknown* pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			reinterpret_cast<IShListView*>(pControl)->get_DefaultNamespaceEnumSettings(&pEnumSettings);
			pControl->Release();
		} else if(m_ppUnk[object]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			reinterpret_cast<IShTreeView*>(pControl)->get_DefaultNamespaceEnumSettings(&pEnumSettings);
			pControl->Release();
		}

		if(pEnumSettings) {
			ShNamespaceEnumerationConstants enumerationFlags = static_cast<ShNamespaceEnumerationConstants>(0);
			pEnumSettings->get_EnumerationFlags(&enumerationFlags);
			LONG l = static_cast<LONG>(enumerationFlags);
			SetBit(controls.enumerationFlagsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, snseIncludeFolders);
			SetBit(controls.enumerationFlagsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, snseIncludeNonFolders);
			// snseIncludeHiddenItems and snseMayIncludeHiddenItems are mutual exclusive
			l &= ~snseMayIncludeHiddenItems;
			SetBit(controls.enumerationFlagsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, snseIncludeHiddenItems);
			if(!(l & snseIncludeHiddenItems)) {
				SetBit(controls.enumerationFlagsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, snseMayIncludeHiddenItems);
			}
			SetBit(controls.enumerationFlagsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, snseIncludeNetPrinters);
			SetBit(controls.enumerationFlagsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, snseIncludeShareableItems);
			SetBit(controls.enumerationFlagsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, snseIncludeStoragesAndAncestors);
			SetBit(controls.enumerationFlagsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, snseEnumForNavigationPane);
			SetBit(controls.enumerationFlagsList.GetItemState(8, LVIS_STATEIMAGEMASK), l, snseLookingForFastItems);
			SetBit(controls.enumerationFlagsList.GetItemState(9, LVIS_STATEIMAGEMASK), l, snseFlatList);
			SetBit(controls.enumerationFlagsList.GetItemState(10, LVIS_STATEIMAGEMASK), l, snseUseShellNotifications);
			pEnumSettings->put_EnumerationFlags(static_cast<ShNamespaceEnumerationConstants>(l));

			BufferAttributeSettings();

			pEnumSettings->put_ExcludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(properties.pAttributes[0][object]));
			pEnumSettings->put_IncludedFileItemFileAttributes(static_cast<ItemFileAttributesConstants>(properties.pAttributes[1][object]));
			pEnumSettings->put_ExcludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(properties.pAttributes[2][object]));
			pEnumSettings->put_IncludedFolderItemFileAttributes(static_cast<ItemFileAttributesConstants>(properties.pAttributes[3][object]));
			pEnumSettings->put_ExcludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(properties.pAttributes[4][object]));
			pEnumSettings->put_IncludedFileItemShellAttributes(static_cast<ItemShellAttributesConstants>(properties.pAttributes[5][object]));
			pEnumSettings->put_ExcludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(properties.pAttributes[6][object]));
			pEnumSettings->put_IncludedFolderItemShellAttributes(static_cast<ItemShellAttributesConstants>(properties.pAttributes[7][object]));
		}
	}

	SetDirty(FALSE);
}

void NamespaceEnumSettingsProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	ShNamespaceEnumerationConstants* pEnumerationFlags = new ShNamespaceEnumerationConstants[m_nObjects];
	ZeroMemory(pEnumerationFlags, m_nObjects * sizeof(ShNamespaceEnumerationConstants));
	for(int i = 0; i < 8; ++i) {
		if(properties.pAttributes[i]) {
			delete[] properties.pAttributes[i];
		}
		properties.pAttributes[i] = new LONG[m_nObjects];
		ZeroMemory(properties.pAttributes[i], m_nObjects * sizeof(LONG));
	}
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComPtr<INamespaceEnumSettings> pEnumSettings;
		IUnknown* pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			reinterpret_cast<IShListView*>(pControl)->get_DefaultNamespaceEnumSettings(&pEnumSettings);
			pControl->Release();
		} else if(m_ppUnk[object]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			reinterpret_cast<IShTreeView*>(pControl)->get_DefaultNamespaceEnumSettings(&pEnumSettings);
			pControl->Release();
		}

		if(pEnumSettings) {
			pEnumSettings->get_EnumerationFlags(&pEnumerationFlags[object]);

			pEnumSettings->get_ExcludedFileItemFileAttributes(reinterpret_cast<ItemFileAttributesConstants*>(&properties.pAttributes[0][object]));
			pEnumSettings->get_IncludedFileItemFileAttributes(reinterpret_cast<ItemFileAttributesConstants*>(&properties.pAttributes[1][object]));
			pEnumSettings->get_ExcludedFolderItemFileAttributes(reinterpret_cast<ItemFileAttributesConstants*>(&properties.pAttributes[2][object]));
			pEnumSettings->get_IncludedFolderItemFileAttributes(reinterpret_cast<ItemFileAttributesConstants*>(&properties.pAttributes[3][object]));
			pEnumSettings->get_ExcludedFileItemShellAttributes(reinterpret_cast<ItemShellAttributesConstants*>(&properties.pAttributes[4][object]));
			pEnumSettings->get_IncludedFileItemShellAttributes(reinterpret_cast<ItemShellAttributesConstants*>(&properties.pAttributes[5][object]));
			pEnumSettings->get_ExcludedFolderItemShellAttributes(reinterpret_cast<ItemShellAttributesConstants*>(&properties.pAttributes[6][object]));
			pEnumSettings->get_IncludedFolderItemShellAttributes(reinterpret_cast<ItemShellAttributesConstants*>(&properties.pAttributes[7][object]));
		}
	}

	// fill the listboxes
	LONG* pl = reinterpret_cast<LONG*>(pEnumerationFlags);
	properties.hWndCheckMarksAreSetFor = controls.enumerationFlagsList;
	controls.enumerationFlagsList.DeleteAllItems();
	controls.enumerationFlagsList.AddItem(0, 0, TEXT("Include folders"));
	controls.enumerationFlagsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, snseIncludeFolders), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(1, 0, TEXT("Include non-folder items"));
	controls.enumerationFlagsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, snseIncludeNonFolders), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(2, 0, TEXT("Include hidden items"));
	controls.enumerationFlagsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, snseIncludeHiddenItems), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(3, 0, TEXT("Include hidden items (follow system settings)"));
	controls.enumerationFlagsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, snseMayIncludeHiddenItems), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(4, 0, TEXT("Include network printers"));
	controls.enumerationFlagsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, snseIncludeNetPrinters), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(5, 0, TEXT("Include shareable items"));
	controls.enumerationFlagsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, snseIncludeShareableItems), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(6, 0, TEXT("Include storage items and containers"));
	controls.enumerationFlagsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, snseIncludeStoragesAndAncestors), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(7, 0, TEXT("Enumeration is for navigation pane"));
	controls.enumerationFlagsList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, snseEnumForNavigationPane), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(8, 0, TEXT("Skip slow items"));
	controls.enumerationFlagsList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, snseLookingForFastItems), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(9, 0, TEXT("Flat list"));
	controls.enumerationFlagsList.SetItemState(9, CalculateStateImageMask(m_nObjects, pl, snseFlatList), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.AddItem(10, 0, TEXT("Use shell notifications"));
	controls.enumerationFlagsList.SetItemState(10, CalculateStateImageMask(m_nObjects, pl, snseUseShellNotifications), LVIS_STATEIMAGEMASK);
	controls.enumerationFlagsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

	properties.hWndCheckMarksAreSetFor = NULL;
	delete[] pEnumerationFlags;

	RestoreAttributeSettings();
	SetDirty(FALSE);
}

void NamespaceEnumSettingsProperties::RestoreAttributeSettings(void)
{
	// fill the listboxes
	LONG* pl = properties.pAttributes[properties.currentAttributesSelection];
	properties.hWndCheckMarksAreSetFor = controls.attributesList;
	controls.attributesList.DeleteAllItems();
	if(properties.currentAttributesSelection & 0x4) {
		// edit shell attributes
		controls.attributesList.AddItem(0, 0, TEXT("Can be copied"));
		controls.attributesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, isaCanBeCopied), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(1, 0, TEXT("Can be moved"));
		controls.attributesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, isaCanBeMoved), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(2, 0, TEXT("Can be linked"));
		controls.attributesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, isaCanBeLinked), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(3, 0, TEXT("Provides storage"));
		controls.attributesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, isaProvidesStorage), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(4, 0, TEXT("Can be renamed"));
		controls.attributesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, isaCanBeRenamed), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(5, 0, TEXT("Can be deleted"));
		controls.attributesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, isaCanBeDeleted), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(6, 0, TEXT("Has property sheets"));
		controls.attributesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, isaHasPropertySheets), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(7, 0, TEXT("Can be drop target & supports paste"));
		controls.attributesList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, isaAcceptsDrops), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(8, 0, TEXT("Encrypted"));
		controls.attributesList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, isaIsEncrypted), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(9, 0, TEXT("Slow on access"));
		controls.attributesList.SetItemState(9, CalculateStateImageMask(m_nObjects, pl, isaIsSlow), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(10, 0, TEXT("Icon should be ghosted"));
		controls.attributesList.SetItemState(10, CalculateStateImageMask(m_nObjects, pl, isaIsGhosted), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(11, 0, TEXT("Link"));
		controls.attributesList.SetItemState(11, CalculateStateImageMask(m_nObjects, pl, isaIsLink), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(12, 0, TEXT("Shared"));
		controls.attributesList.SetItemState(12, CalculateStateImageMask(m_nObjects, pl, isaIsShared), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(13, 0, TEXT("Read-only"));
		controls.attributesList.SetItemState(13, CalculateStateImageMask(m_nObjects, pl, isaIsReadOnly), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(14, 0, TEXT("Hidden"));
		controls.attributesList.SetItemState(14, CalculateStateImageMask(m_nObjects, pl, isaIsHidden), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(15, 0, TEXT("Not enumerated"));
		controls.attributesList.SetItemState(15, CalculateStateImageMask(m_nObjects, pl, isaIsNonEnumerated), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(16, 0, TEXT("New content"));
		controls.attributesList.SetItemState(16, CalculateStateImageMask(m_nObjects, pl, isaIsNewContent), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(17, 0, TEXT("Provides stream"));
		controls.attributesList.SetItemState(17, CalculateStateImageMask(m_nObjects, pl, isaProvidesStream), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(18, 0, TEXT("Provides stream or storage"));
		controls.attributesList.SetItemState(18, CalculateStateImageMask(m_nObjects, pl, isaContainsStreamsOrStorages), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(19, 0, TEXT("Is on removable media"));
		controls.attributesList.SetItemState(19, CalculateStateImageMask(m_nObjects, pl, isaIsRemovable), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(20, 0, TEXT("Compressed"));
		controls.attributesList.SetItemState(20, CalculateStateImageMask(m_nObjects, pl, isaIsCompressed), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(21, 0, TEXT("Supports IShellFolder, but is no folder"));
		controls.attributesList.SetItemState(21, CalculateStateImageMask(m_nObjects, pl, isaBrowsableInPlace), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(22, 0, TEXT("Contains file system objects"));
		controls.attributesList.SetItemState(22, CalculateStateImageMask(m_nObjects, pl, isaContainsFileSystemItems), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(23, 0, TEXT("Folder"));
		controls.attributesList.SetItemState(23, CalculateStateImageMask(m_nObjects, pl, isaIsFolder), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(24, 0, TEXT("File system object"));
		controls.attributesList.SetItemState(24, CalculateStateImageMask(m_nObjects, pl, isaIsPartOfFileSystem), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(25, 0, TEXT("Contains folders"));
		controls.attributesList.SetItemState(25, CalculateStateImageMask(m_nObjects, pl, isaContainsSubFolders), LVIS_STATEIMAGEMASK);
	} else {
		// edit file attributes
		controls.attributesList.AddItem(0, 0, TEXT("Read-only"));
		controls.attributesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, ifaReadOnly), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(1, 0, TEXT("Hidden"));
		controls.attributesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, ifaHidden), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(2, 0, TEXT("System"));
		controls.attributesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, ifaSystem), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(3, 0, TEXT("Directory"));
		controls.attributesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, ifaDirectory), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(4, 0, TEXT("Archive"));
		controls.attributesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, ifaArchive), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(5, 0, TEXT("Normal (not hidden, read-only etc)"));
		controls.attributesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, ifaNormal), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(6, 0, TEXT("Used for temporary storage"));
		controls.attributesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, ifaTemporary), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(7, 0, TEXT("Sparse file"));
		controls.attributesList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, ifaSparseFile), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(8, 0, TEXT("Symbolic link or has a reparse point"));
		controls.attributesList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, ifaReparsePoint), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(9, 0, TEXT("Compressed"));
		controls.attributesList.SetItemState(9, CalculateStateImageMask(m_nObjects, pl, ifaCompressed), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(10, 0, TEXT("Not available immediately"));
		controls.attributesList.SetItemState(10, CalculateStateImageMask(m_nObjects, pl, ifaOffline), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(11, 0, TEXT("Not indexed by content indexing service"));
		controls.attributesList.SetItemState(11, CalculateStateImageMask(m_nObjects, pl, ifaNotContentIndexed), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(12, 0, TEXT("Encrypted"));
		controls.attributesList.SetItemState(12, CalculateStateImageMask(m_nObjects, pl, ifaEncrypted), LVIS_STATEIMAGEMASK);
		controls.attributesList.AddItem(13, 0, TEXT("Virtual file"));
		controls.attributesList.SetItemState(13, CalculateStateImageMask(m_nObjects, pl, ifaVirtual), LVIS_STATEIMAGEMASK);
	}
	controls.attributesList.SetColumnWidth(0, LVSCW_AUTOSIZE);
	properties.hWndCheckMarksAreSetFor = NULL;
}

void NamespaceEnumSettingsProperties::BufferAttributeSettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LONG l = properties.pAttributes[properties.currentAttributesSelection][object];
		if(properties.currentAttributesSelection & 0x4) {
			// edit shell attributes
			SetBit(controls.attributesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, isaCanBeCopied);
			SetBit(controls.attributesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, isaCanBeMoved);
			SetBit(controls.attributesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, isaCanBeLinked);
			SetBit(controls.attributesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, isaProvidesStorage);
			SetBit(controls.attributesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, isaCanBeRenamed);
			SetBit(controls.attributesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, isaCanBeDeleted);
			SetBit(controls.attributesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, isaHasPropertySheets);
			SetBit(controls.attributesList.GetItemState(7, LVIS_STATEIMAGEMASK), l, isaAcceptsDrops);
			SetBit(controls.attributesList.GetItemState(8, LVIS_STATEIMAGEMASK), l, isaIsEncrypted);
			SetBit(controls.attributesList.GetItemState(9, LVIS_STATEIMAGEMASK), l, isaIsSlow);
			SetBit(controls.attributesList.GetItemState(10, LVIS_STATEIMAGEMASK), l, isaIsGhosted);
			SetBit(controls.attributesList.GetItemState(11, LVIS_STATEIMAGEMASK), l, isaIsLink);
			SetBit(controls.attributesList.GetItemState(12, LVIS_STATEIMAGEMASK), l, isaIsShared);
			SetBit(controls.attributesList.GetItemState(13, LVIS_STATEIMAGEMASK), l, isaIsReadOnly);
			SetBit(controls.attributesList.GetItemState(14, LVIS_STATEIMAGEMASK), l, isaIsHidden);
			SetBit(controls.attributesList.GetItemState(15, LVIS_STATEIMAGEMASK), l, isaIsNonEnumerated);
			SetBit(controls.attributesList.GetItemState(16, LVIS_STATEIMAGEMASK), l, isaIsNewContent);
			SetBit(controls.attributesList.GetItemState(17, LVIS_STATEIMAGEMASK), l, isaProvidesStream);
			SetBit(controls.attributesList.GetItemState(18, LVIS_STATEIMAGEMASK), l, isaContainsStreamsOrStorages);
			SetBit(controls.attributesList.GetItemState(19, LVIS_STATEIMAGEMASK), l, isaIsRemovable);
			SetBit(controls.attributesList.GetItemState(20, LVIS_STATEIMAGEMASK), l, isaIsCompressed);
			SetBit(controls.attributesList.GetItemState(21, LVIS_STATEIMAGEMASK), l, isaBrowsableInPlace);
			SetBit(controls.attributesList.GetItemState(22, LVIS_STATEIMAGEMASK), l, isaContainsFileSystemItems);
			SetBit(controls.attributesList.GetItemState(23, LVIS_STATEIMAGEMASK), l, isaIsFolder);
			SetBit(controls.attributesList.GetItemState(24, LVIS_STATEIMAGEMASK), l, isaIsPartOfFileSystem);
			SetBit(controls.attributesList.GetItemState(25, LVIS_STATEIMAGEMASK), l, isaContainsSubFolders);
		} else {
			// edit file attributes
			SetBit(controls.attributesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, ifaReadOnly);
			SetBit(controls.attributesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, ifaHidden);
			SetBit(controls.attributesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, ifaSystem);
			SetBit(controls.attributesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, ifaDirectory);
			SetBit(controls.attributesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, ifaArchive);
			SetBit(controls.attributesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, ifaNormal);
			SetBit(controls.attributesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, ifaTemporary);
			SetBit(controls.attributesList.GetItemState(7, LVIS_STATEIMAGEMASK), l, ifaSparseFile);
			SetBit(controls.attributesList.GetItemState(8, LVIS_STATEIMAGEMASK), l, ifaReparsePoint);
			SetBit(controls.attributesList.GetItemState(9, LVIS_STATEIMAGEMASK), l, ifaCompressed);
			SetBit(controls.attributesList.GetItemState(10, LVIS_STATEIMAGEMASK), l, ifaOffline);
			SetBit(controls.attributesList.GetItemState(11, LVIS_STATEIMAGEMASK), l, ifaNotContentIndexed);
			SetBit(controls.attributesList.GetItemState(12, LVIS_STATEIMAGEMASK), l, ifaEncrypted);
			SetBit(controls.attributesList.GetItemState(13, LVIS_STATEIMAGEMASK), l, ifaVirtual);
		}
		properties.pAttributes[properties.currentAttributesSelection][object] = l;
	}
}

int NamespaceEnumSettingsProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if((pArray[object] & bitsToCheckFor) == bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void NamespaceEnumSettingsProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}