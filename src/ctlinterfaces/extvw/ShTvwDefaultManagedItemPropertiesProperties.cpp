// ShTvwDefaultManagedItemPropertiesProperties.cpp: The ShellTreeView control's "Default Managed Properties" property page

#include "stdafx.h"
#include "ShTvwDefaultManagedItemPropertiesProperties.h"


ShTvwDefaultManagedItemPropertiesProperties::ShTvwDefaultManagedItemPropertiesProperties()
{
	m_dwTitleID = IDS_TITLEDEFAULTMANAGEDITEMPROPERTIESPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGDEFAULTMANAGEDITEMPROPERTIESPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP ShTvwDefaultManagedItemPropertiesProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP ShTvwDefaultManagedItemPropertiesProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}

	IPropertyPageImpl<ShTvwDefaultManagedItemPropertiesProperties>::SetObjects(objects, ppControls);

	if(properties.weAreInitialized) {
		LoadSettings();
	}
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// attach to the controls
	controls.managedPropertiesList.SubclassWindow(GetDlgItem(IDC_MANAGEDPROPERTIESBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.managedPropertiesList.GetImageList(LVSIL_STATE));
	controls.managedPropertiesList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.managedPropertiesList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.managedPropertiesList.AddColumn(TEXT(""), 0);

	// setup the toolbar
	WTL::CRect toolbarRect;
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

	properties.weAreInitialized = TRUE;
	return TRUE;
}


LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.managedPropertiesList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's 'Ghosted' property."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's 'IconIndex' property."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's 'OverlayIndex' property."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's 'SelectedIconIndex' property."))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control inserts the shell item's sub-items as the treeview item's sub-items automatically."))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control inserts the shell item's sub-items as the treeview item's sub-items automatically and also sorts them."))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's 'Text' property."))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control handles label-editing of the treeview item automatically."))));
				break;
			case 8:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the treeview item's info tip."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
	}

	delete[] pBuffer;
	return 0;
}

LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

			if(pNotificationDetails->hwndFrom == controls.managedPropertiesList) {
				switch(pDetails->iItem) {
					case 4/*stmipSubItems*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 1) {
							// uncheck next item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(1);
							item.stateMask = LVIS_STATEIMAGEMASK;
							::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem + 1, reinterpret_cast<LPARAM>(&item));
						}
						break;
					case 5/*stmipSubItemsSorting*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 2) {
							// check previous item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(2);
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

LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IShTreeView, &IID_IShTreeView> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}

LRESULT ShTvwDefaultManagedItemPropertiesProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IShTreeView, &IID_IShTreeView> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(FALSE, NULL, TEXT("ShellTreeView Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}


void ShTvwDefaultManagedItemPropertiesProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IShTreeView, &IID_IShTreeView> pControl(m_ppUnk[object]);
		if(pControl) {
			ShTvwManagedItemPropertiesConstants managedProperties = static_cast<ShTvwManagedItemPropertiesConstants>(0);
			pControl->get_DefaultManagedItemProperties(&managedProperties);
			LONG l = static_cast<LONG>(managedProperties);
			SetBit(controls.managedPropertiesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, stmipGhosted);
			SetBit(controls.managedPropertiesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, stmipIconIndex);
			SetBit(controls.managedPropertiesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, stmipOverlayIndex);
			SetBit(controls.managedPropertiesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, stmipSelectedIconIndex);
			SetBit(controls.managedPropertiesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, stmipSubItems);
			SetBit(controls.managedPropertiesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, stmipSubItemsSorting);
			SetBit(controls.managedPropertiesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, stmipText);
			SetBit(controls.managedPropertiesList.GetItemState(7, LVIS_STATEIMAGEMASK), l, stmipRenaming);
			SetBit(controls.managedPropertiesList.GetItemState(8, LVIS_STATEIMAGEMASK), l, stmipInfoTip);
			pControl->put_DefaultManagedItemProperties(static_cast<ShTvwManagedItemPropertiesConstants>(l));
		}
	}

	SetDirty(FALSE);
}

void ShTvwDefaultManagedItemPropertiesProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	ShTvwManagedItemPropertiesConstants* pManagedProperties = static_cast<ShTvwManagedItemPropertiesConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(ShTvwManagedItemPropertiesConstants)));
	if(pManagedProperties) {
		for(UINT object = 0; object < m_nObjects; ++object) {
			CComQIPtr<IShTreeView, &IID_IShTreeView> pControl(m_ppUnk[object]);
			if(pControl) {
				pControl->get_DefaultManagedItemProperties(&pManagedProperties[object]);
			}
		}

		// fill the listboxes
		LONG* pl = reinterpret_cast<LONG*>(pManagedProperties);
		properties.hWndCheckMarksAreSetFor = controls.managedPropertiesList;
		controls.managedPropertiesList.DeleteAllItems();
		controls.managedPropertiesList.AddItem(0, 0, TEXT("'Ghosted' property"));
		controls.managedPropertiesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, stmipGhosted), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(1, 0, TEXT("'IconIndex' property"));
		controls.managedPropertiesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, stmipIconIndex), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(2, 0, TEXT("'OverlayIndex' property"));
		controls.managedPropertiesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, stmipOverlayIndex), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(3, 0, TEXT("'SelectedIconIndex' property"));
		controls.managedPropertiesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, stmipSelectedIconIndex), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(4, 0, TEXT("'SubItems' property"));
		controls.managedPropertiesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, stmipSubItems), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(5, 0, TEXT("'SubItems' property with sorting"));
		controls.managedPropertiesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, stmipSubItemsSorting), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(6, 0, TEXT("'Text' property"));
		controls.managedPropertiesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, stmipText), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(7, 0, TEXT("Item renaming"));
		controls.managedPropertiesList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, stmipRenaming), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(8, 0, TEXT("Info tips"));
		controls.managedPropertiesList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, stmipInfoTip), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;

		HeapFree(GetProcessHeap(), 0, pManagedProperties);
	}

	SetDirty(FALSE);
}

int ShTvwDefaultManagedItemPropertiesProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
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

void ShTvwDefaultManagedItemPropertiesProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
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