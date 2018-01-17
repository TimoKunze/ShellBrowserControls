// ShLvwDefaultManagedItemPropertiesProperties.cpp: The ShellListView control's "Default Managed Properties" property page

#include "stdafx.h"
#include "ShLvwDefaultManagedItemPropertiesProperties.h"


ShLvwDefaultManagedItemPropertiesProperties::ShLvwDefaultManagedItemPropertiesProperties()
{
	m_dwTitleID = IDS_TITLEDEFAULTMANAGEDITEMPROPERTIESPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGDEFAULTMANAGEDITEMPROPERTIESPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP ShLvwDefaultManagedItemPropertiesProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP ShLvwDefaultManagedItemPropertiesProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}

	IPropertyPageImpl<ShLvwDefaultManagedItemPropertiesProperties>::SetObjects(objects, ppControls);

	if(properties.weAreInitialized) {
		LoadSettings();
	}
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnInitDialog(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
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


LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.managedPropertiesList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's 'Ghosted' property."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's 'IconIndex' property."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's 'OverlayIndex' property."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's 'Text' property."))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview sub-item's 'Text' property."))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's 'TileViewColumns' property."))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control handles label-editing of the listview item automatically."))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's info tip."))));
				break;
			case 8:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("The control manages the listview item's sub-item controls."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
	}

	delete[] pBuffer;
	return 0;
}

LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IShListView, &IID_IShListView> pControl(m_ppUnk[0]);
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

LRESULT ShLvwDefaultManagedItemPropertiesProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IShListView, &IID_IShListView> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(FALSE, NULL, TEXT("ShellListView Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
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


void ShLvwDefaultManagedItemPropertiesProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IShListView, &IID_IShListView> pControl(m_ppUnk[object]);
		if(pControl) {
			ShLvwManagedItemPropertiesConstants managedProperties = static_cast<ShLvwManagedItemPropertiesConstants>(0);
			pControl->get_DefaultManagedItemProperties(&managedProperties);
			LONG l = static_cast<LONG>(managedProperties);
			SetBit(controls.managedPropertiesList.GetItemState(0, LVIS_STATEIMAGEMASK), l, slmipGhosted);
			SetBit(controls.managedPropertiesList.GetItemState(1, LVIS_STATEIMAGEMASK), l, slmipIconIndex);
			SetBit(controls.managedPropertiesList.GetItemState(2, LVIS_STATEIMAGEMASK), l, slmipOverlayIndex);
			SetBit(controls.managedPropertiesList.GetItemState(3, LVIS_STATEIMAGEMASK), l, slmipText);
			SetBit(controls.managedPropertiesList.GetItemState(4, LVIS_STATEIMAGEMASK), l, slmipSubItemsText);
			SetBit(controls.managedPropertiesList.GetItemState(5, LVIS_STATEIMAGEMASK), l, slmipTileViewColumns);
			SetBit(controls.managedPropertiesList.GetItemState(6, LVIS_STATEIMAGEMASK), l, slmipRenaming);
			SetBit(controls.managedPropertiesList.GetItemState(7, LVIS_STATEIMAGEMASK), l, slmipInfoTip);
			SetBit(controls.managedPropertiesList.GetItemState(8, LVIS_STATEIMAGEMASK), l, slmipSubItemControls);
			pControl->put_DefaultManagedItemProperties(static_cast<ShLvwManagedItemPropertiesConstants>(l));
		}
	}

	SetDirty(FALSE);
}

void ShLvwDefaultManagedItemPropertiesProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	ShLvwManagedItemPropertiesConstants* pManagedProperties = static_cast<ShLvwManagedItemPropertiesConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(ShLvwManagedItemPropertiesConstants)));
	if(pManagedProperties) {
		for(UINT object = 0; object < m_nObjects; ++object) {
			CComQIPtr<IShListView, &IID_IShListView> pControl(m_ppUnk[object]);
			if(pControl) {
				pControl->get_DefaultManagedItemProperties(&pManagedProperties[object]);
			}
		}

		// fill the listboxes
		LONG* pl = reinterpret_cast<LONG*>(pManagedProperties);
		properties.hWndCheckMarksAreSetFor = controls.managedPropertiesList;
		controls.managedPropertiesList.DeleteAllItems();
		controls.managedPropertiesList.AddItem(0, 0, TEXT("'Ghosted' property"));
		controls.managedPropertiesList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, slmipGhosted), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(1, 0, TEXT("'IconIndex' property"));
		controls.managedPropertiesList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, slmipIconIndex), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(2, 0, TEXT("'OverlayIndex' property"));
		controls.managedPropertiesList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, slmipOverlayIndex), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(3, 0, TEXT("'Text' property"));
		controls.managedPropertiesList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, slmipText), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(4, 0, TEXT("'Text' property (sub-items)"));
		controls.managedPropertiesList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, slmipSubItemsText), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(5, 0, TEXT("'TileViewColumns' property"));
		controls.managedPropertiesList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, slmipTileViewColumns), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(6, 0, TEXT("Item renaming"));
		controls.managedPropertiesList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, slmipRenaming), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(7, 0, TEXT("Info tips"));
		controls.managedPropertiesList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, slmipInfoTip), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.AddItem(8, 0, TEXT("Sub-item controls"));
		controls.managedPropertiesList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, slmipSubItemControls), LVIS_STATEIMAGEMASK);
		controls.managedPropertiesList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;

		HeapFree(GetProcessHeap(), 0, pManagedProperties);
	}

	SetDirty(FALSE);
}

int ShLvwDefaultManagedItemPropertiesProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
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

void ShLvwDefaultManagedItemPropertiesProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
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