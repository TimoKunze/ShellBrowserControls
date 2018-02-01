// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	controls.infoTipFlagsList.SubclassWindow(GetDlgItem(IDC_INFOTIPFLAGSBOX));
	hStateImageList = SetupStateImageList(controls.infoTipFlagsList.GetImageList(LVSIL_STATE));
	controls.infoTipFlagsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.infoTipFlagsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.infoTipFlagsList.AddColumn(TEXT(""), 0);

	controls.useSystemImageListList.SubclassWindow(GetDlgItem(IDC_USESYSTEMIMAGELISTBOX));
	hStateImageList = SetupStateImageList(controls.useSystemImageListList.GetImageList(LVSIL_STATE));
	controls.useSystemImageListList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.useSystemImageListList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.useSystemImageListList.AddColumn(TEXT(""), 0);

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
	if(controls.useSystemImageListList.GetItemCount() > 1) {
		controls.useSystemImageListList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected view modes"));
	}
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.disabledEventsList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingItem, InsertedItem"))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ChangedItemPIDL"))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingItem"))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertedNamespace"))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ChangedNamespacePIDL"))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingNamespace"))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("LoadedColumn, UnloadedColumn"))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ChangingColumnVisibility, ChangedColumnVisibility"))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	} else if(pNotificationDetails->hwndFrom == controls.infoTipFlagsList) {
		switch(pDetails->iItem) {
			case 0:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Retrieve the item's display name if it is a link."))));
				break;
			case 1:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Exclude the link's target info tip text from the info tip, if the item is a link."))));
				break;
			case 2:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Retrieve the path to the link's target if the item is a link."))));
				break;
			case 3:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Allow the shell to perform time-consuming tasks to retrieve the item's info tip text.\nRequires Windows XP or newer."))));
				break;
			case 4:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Allow or disallow (following the system settings) the shell to perform time-consuming tasks to retrieve the item's info tip text.\nRequires Windows XP or newer."))));
				break;
			case 5:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Retrieve the info tip text in a single line.\nRequires Windows Vista or newer."))));
				break;
			case 6:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Don't display an info tip."))));
				break;
			case 7:
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Follow the system settings when deciding whether to display an info tip."))));
				break;
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	} else if(pNotificationDetails->hwndFrom == controls.useSystemImageListList) {
		if(controls.useSystemImageListList.GetItemCount() > 1) {
			switch(pDetails->iItem) {
				case 0:
					ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Small Icons, List, Details"))));
					break;
				case 1:
					ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Icons"))));
					break;
				case 2:
					ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Tiles"))));
					break;
			}
			ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
		}
	}

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

			if(pNotificationDetails->hwndFrom == controls.infoTipFlagsList) {
				switch(pDetails->iItem) {
					case 3/*itfAllowSlowInfoTip*/:
					case 6/*itfNoInfoTip*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 2) {
							// uncheck next item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(1);
							item.stateMask = LVIS_STATEIMAGEMASK;
							::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem + 1, reinterpret_cast<LPARAM>(&item));
						}
						break;
					case 4/*itfAllowSlowInfoTipFollowSysSettings*/:
					case 7/*itfNoInfoTipFollowSystemSettings*/:
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

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	LPUNKNOWN pControl = NULL;
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

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	LPUNKNOWN pControl = NULL;
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


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IShListView*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deItemPIDLChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deNamespaceInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deNamespacePIDLChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, deNamespaceDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, deColumnLoadingEvents);
			SetBit(controls.disabledEventsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, deColumnVisibilityEvents);
			reinterpret_cast<IShListView*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			InfoTipFlagsConstants infoTipFlags = static_cast<InfoTipFlagsConstants>(0);
			reinterpret_cast<IShListView*>(pControl)->get_InfoTipFlags(&infoTipFlags);
			l = static_cast<LONG>(infoTipFlags);
			SetBit(controls.infoTipFlagsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, itfLinkName);
			SetBit(controls.infoTipFlagsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, itfNoLinkTarget);
			SetBit(controls.infoTipFlagsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, itfLinkTarget);
			SetBit(controls.infoTipFlagsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, itfAllowSlowInfoTip);
			SetBit(controls.infoTipFlagsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, itfAllowSlowInfoTipFollowSysSettings);
			SetBit(controls.infoTipFlagsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, itfSingleLine);
			SetBit(controls.infoTipFlagsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, itfNoInfoTip);
			SetBit(controls.infoTipFlagsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, itfNoInfoTipFollowSystemSettings);
			reinterpret_cast<IShListView*>(pControl)->put_InfoTipFlags(static_cast<InfoTipFlagsConstants>(l));

			UseSystemImageListConstants useSystemImageList = static_cast<UseSystemImageListConstants>(0);
			reinterpret_cast<IShListView*>(pControl)->get_UseSystemImageList(&useSystemImageList);
			l = static_cast<LONG>(useSystemImageList);
			SetBit(controls.useSystemImageListList.GetItemState(0, LVIS_STATEIMAGEMASK), l, usilSmallImageList);
			SetBit(controls.useSystemImageListList.GetItemState(1, LVIS_STATEIMAGEMASK), l, usilLargeImageList);
			SetBit(controls.useSystemImageListList.GetItemState(2, LVIS_STATEIMAGEMASK), l, usilExtraLargeImageList);
			reinterpret_cast<IShListView*>(pControl)->put_UseSystemImageList(static_cast<UseSystemImageListConstants>(l));
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IShTreeView*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deItemPIDLChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deNamespaceInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deNamespacePIDLChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, deNamespaceDeletionEvents);
			reinterpret_cast<IShTreeView*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			InfoTipFlagsConstants infoTipFlags = static_cast<InfoTipFlagsConstants>(0);
			reinterpret_cast<IShTreeView*>(pControl)->get_InfoTipFlags(&infoTipFlags);
			l = static_cast<LONG>(infoTipFlags);
			SetBit(controls.infoTipFlagsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, itfLinkName);
			SetBit(controls.infoTipFlagsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, itfNoLinkTarget);
			SetBit(controls.infoTipFlagsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, itfLinkTarget);
			SetBit(controls.infoTipFlagsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, itfAllowSlowInfoTip);
			SetBit(controls.infoTipFlagsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, itfAllowSlowInfoTipFollowSysSettings);
			SetBit(controls.infoTipFlagsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, itfSingleLine);
			SetBit(controls.infoTipFlagsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, itfNoInfoTip);
			SetBit(controls.infoTipFlagsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, itfNoInfoTipFollowSystemSettings);
			reinterpret_cast<IShTreeView*>(pControl)->put_InfoTipFlags(static_cast<InfoTipFlagsConstants>(l));

			UseSystemImageListConstants useSystemImageList = static_cast<UseSystemImageListConstants>(0);
			reinterpret_cast<IShTreeView*>(pControl)->get_UseSystemImageList(&useSystemImageList);
			l = static_cast<LONG>(useSystemImageList);
			SetBit(controls.useSystemImageListList.GetItemState(0, LVIS_STATEIMAGEMASK), l, usilSmallImageList);
			reinterpret_cast<IShTreeView*>(pControl)->put_UseSystemImageList(static_cast<UseSystemImageListConstants>(l));
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	int numberOfShellListViews = 0;
	int numberOfShellTreeViews = 0;
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		InfoTipFlagsConstants* pInfoTipFlags = static_cast<InfoTipFlagsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(InfoTipFlagsConstants)));
		if(pInfoTipFlags) {
			ZeroMemory(pInfoTipFlags, m_nObjects * sizeof(InfoTipFlagsConstants));
			UseSystemImageListConstants* pUseSystemImageList = static_cast<UseSystemImageListConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(UseSystemImageListConstants)));
			if(pUseSystemImageList) {
				ZeroMemory(pUseSystemImageList, m_nObjects * sizeof(UseSystemImageListConstants));
				for(UINT object = 0; object < m_nObjects; ++object) {
					LPUNKNOWN pControl = NULL;
					if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfShellListViews;
						reinterpret_cast<IShListView*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						reinterpret_cast<IShListView*>(pControl)->get_InfoTipFlags(&pInfoTipFlags[object]);
						reinterpret_cast<IShListView*>(pControl)->get_UseSystemImageList(&pUseSystemImageList[object]);
						pControl->Release();

					} else if(m_ppUnk[object]->QueryInterface(IID_IShTreeView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
						++numberOfShellTreeViews;
						reinterpret_cast<IShTreeView*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
						reinterpret_cast<IShTreeView*>(pControl)->get_InfoTipFlags(&pInfoTipFlags[object]);
						reinterpret_cast<IShTreeView*>(pControl)->get_UseSystemImageList(&pUseSystemImageList[object]);
						pControl->Release();
					}
				}

				// fill the listboxes
				LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
				properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
				controls.disabledEventsList.DeleteAllItems();
				controls.disabledEventsList.AddItem(0, 0, TEXT("Item insertions"));
				controls.disabledEventsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, deItemInsertionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(1, 0, TEXT("Item pIDL changes"));
				controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deItemPIDLChangeEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(2, 0, TEXT("Item deletions"));
				controls.disabledEventsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, deItemDeletionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(3, 0, TEXT("Namespace insertions"));
				controls.disabledEventsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, deNamespaceInsertionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(4, 0, TEXT("Namespace pIDL changes"));
				controls.disabledEventsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, deNamespacePIDLChangeEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(5, 0, TEXT("Namespace deletions"));
				controls.disabledEventsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, deNamespaceDeletionEvents), LVIS_STATEIMAGEMASK);
				if(numberOfShellListViews > 0) {
					controls.disabledEventsList.AddItem(6, 0, TEXT("Column loading events"));
					controls.disabledEventsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, deColumnLoadingEvents), LVIS_STATEIMAGEMASK);
					controls.disabledEventsList.AddItem(7, 0, TEXT("Column visibility changes"));
					controls.disabledEventsList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, deColumnVisibilityEvents), LVIS_STATEIMAGEMASK);
				}
				controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pInfoTipFlags);
				properties.hWndCheckMarksAreSetFor = controls.infoTipFlagsList;
				controls.infoTipFlagsList.DeleteAllItems();
				controls.infoTipFlagsList.AddItem(0, 0, TEXT("Retrieve link display name"));
				controls.infoTipFlagsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, itfLinkName), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(1, 0, TEXT("Exclude link target"));
				controls.infoTipFlagsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, itfNoLinkTarget), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(2, 0, TEXT("Retrieve link target"));
				controls.infoTipFlagsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, itfLinkTarget), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(3, 0, TEXT("Allow slow info tip"));
				controls.infoTipFlagsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, itfAllowSlowInfoTip), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(4, 0, TEXT("Allow slow info tip (follow system settings)"));
				controls.infoTipFlagsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, itfAllowSlowInfoTipFollowSysSettings), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(5, 0, TEXT("Single line"));
				controls.infoTipFlagsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, itfSingleLine), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(6, 0, TEXT("No info tip"));
				controls.infoTipFlagsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, itfNoInfoTip), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.AddItem(7, 0, TEXT("No info tip (follow system settings)"));
				controls.infoTipFlagsList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, itfNoInfoTipFollowSystemSettings), LVIS_STATEIMAGEMASK);
				controls.infoTipFlagsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pUseSystemImageList);
				properties.hWndCheckMarksAreSetFor = controls.useSystemImageListList;
				controls.useSystemImageListList.DeleteAllItems();
				controls.useSystemImageListList.AddItem(0, 0, TEXT("Small"));
				controls.useSystemImageListList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, usilSmallImageList), LVIS_STATEIMAGEMASK);
				if(numberOfShellListViews > 0) {
					controls.useSystemImageListList.AddItem(1, 0, TEXT("Large"));
					controls.useSystemImageListList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, usilLargeImageList), LVIS_STATEIMAGEMASK);
					controls.useSystemImageListList.AddItem(2, 0, TEXT("Extra-large"));
					controls.useSystemImageListList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, usilExtraLargeImageList), LVIS_STATEIMAGEMASK);
				}
				controls.useSystemImageListList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				properties.hWndCheckMarksAreSetFor = NULL;

				HeapFree(GetProcessHeap(), 0, pUseSystemImageList);
			}
			HeapFree(GetProcessHeap(), 0, pInfoTipFlags);
		}
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
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

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
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