// ThumbnailsProperties.cpp: The controls' "Thumbnails" property page

#include "stdafx.h"
#include "ThumbnailsProperties.h"


ThumbnailsProperties::ThumbnailsProperties()
{
	m_dwTitleID = IDS_TITLETHUMBNAILSPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGTHUMBNAILSPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP ThumbnailsProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP ThumbnailsProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<ThumbnailsProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.displayThumbnailAdornmentsList.SubclassWindow(GetDlgItem(IDC_DISPLAYTHUMBNAILADORNMENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.displayThumbnailAdornmentsList.GetImageList(LVSIL_STATE));
	controls.displayThumbnailAdornmentsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.displayThumbnailAdornmentsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.displayThumbnailAdornmentsList.AddColumn(TEXT(""), 0);

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

STDMETHODIMP ThumbnailsProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<ThumbnailsProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT ThumbnailsProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

			if(pNotificationDetails->hwndFrom == controls.displayThumbnailAdornmentsList && pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
				switch(pDetails->iItem) {
					case 0/*sldtaAny*/:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 2) {
							// check all other items
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(2);
							item.stateMask = LVIS_STATEIMAGEMASK;
							for(int i = 1; i < ::SendMessage(pNotificationDetails->hwndFrom, LVM_GETITEMCOUNT, 0, 0); i++) {
								::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, i, reinterpret_cast<LPARAM>(&item));
							}
						}
						break;
					default:
						if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 1) {
							// uncheck 'Any' item
							LVITEM item = {0};
							item.state = INDEXTOSTATEIMAGEMASK(1);
							item.stateMask = LVIS_STATEIMAGEMASK;
							::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, 0/*sldtaAny*/, reinterpret_cast<LPARAM>(&item));
						}
						break;
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT ThumbnailsProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ThumbnailsProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT ThumbnailsProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
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
	}
	return 0;
}

LRESULT ThumbnailsProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
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
	}
	return 0;
}


void ThumbnailsProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments = sldtaAny;
			reinterpret_cast<IShListView*>(pControl)->get_DisplayThumbnailAdornments(&displayThumbnailAdornments);
			LONG l = static_cast<LONG>(displayThumbnailAdornments);
			SetBit(controls.displayThumbnailAdornmentsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, sldtaAny);
			SetBit(controls.displayThumbnailAdornmentsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, sldtaDropShadow);
			SetBit(controls.displayThumbnailAdornmentsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, sldtaPhotoBorder);
			SetBit(controls.displayThumbnailAdornmentsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, sldtaVideoSprocket);
			reinterpret_cast<IShListView*>(pControl)->put_DisplayThumbnailAdornments(static_cast<ShLvwDisplayThumbnailAdornmentsConstants>(l));
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void ThumbnailsProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	ShLvwDisplayThumbnailAdornmentsConstants* pDisplayThumbnailAdornments = static_cast<ShLvwDisplayThumbnailAdornmentsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(ShLvwDisplayThumbnailAdornmentsConstants)));
	if(pDisplayThumbnailAdornments) {
		ZeroMemory(pDisplayThumbnailAdornments, m_nObjects * sizeof(ShLvwDisplayThumbnailAdornmentsConstants));
		for(UINT object = 0; object < m_nObjects; ++object) {
			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[object]->QueryInterface(IID_IShListView, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				reinterpret_cast<IShListView*>(pControl)->get_DisplayThumbnailAdornments(&pDisplayThumbnailAdornments[object]);
				pControl->Release();
			}
		}

		// fill the listboxes
		LONG* pl = reinterpret_cast<LONG*>(pDisplayThumbnailAdornments);
		properties.hWndCheckMarksAreSetFor = controls.displayThumbnailAdornmentsList;
		controls.displayThumbnailAdornmentsList.DeleteAllItems();
		controls.displayThumbnailAdornmentsList.AddItem(0, 0, TEXT("Any"));
		controls.displayThumbnailAdornmentsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, sldtaAny, TRUE), LVIS_STATEIMAGEMASK);
		controls.displayThumbnailAdornmentsList.AddItem(1, 0, TEXT("Drop Shadow"));
		controls.displayThumbnailAdornmentsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, sldtaDropShadow), LVIS_STATEIMAGEMASK);
		controls.displayThumbnailAdornmentsList.AddItem(2, 0, TEXT("Photo Border"));
		controls.displayThumbnailAdornmentsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, sldtaPhotoBorder), LVIS_STATEIMAGEMASK);
		controls.displayThumbnailAdornmentsList.AddItem(3, 0, TEXT("Video Sprocket"));
		controls.displayThumbnailAdornmentsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, sldtaVideoSprocket), LVIS_STATEIMAGEMASK);
		controls.displayThumbnailAdornmentsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;
		HeapFree(GetProcessHeap(), 0, pDisplayThumbnailAdornments);
	}

	SetDirty(FALSE);
}

int ThumbnailsProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor, BOOL checkForEquality/* = FALSE*/)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		BOOL isSet = (checkForEquality ? pArray[object] == bitsToCheckFor : pArray[object] & bitsToCheckFor);
		if(isSet) {
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

void ThumbnailsProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
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