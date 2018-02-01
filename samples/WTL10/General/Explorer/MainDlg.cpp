// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::IsComctl32Version600OrNewer(void)
{
	BOOL ret = FALSE;

	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		DLLGETVERSIONPROC pDllGetVersion = reinterpret_cast<DLLGETVERSIONPROC>(GetProcAddress(hMod, "DllGetVersion"));
		if(pDllGetVersion) {
			DLLVERSIONINFO versionInfo = {0};
			versionInfo.cbSize = sizeof(versionInfo);
			HRESULT hr = (*pDllGetVersion)(&versionInfo);
			if(SUCCEEDED(hr)) {
				ret = (versionInfo.dwMajorVersion >= 6);
			}
		}
		FreeLibrary(hMod);
	}

	return ret;
}


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	if((pMsg->message == WM_KEYDOWN) && ((pMsg->wParam == VK_ESCAPE) || (pMsg->wParam == VK_RETURN))) {
		// otherwise label-edit wouldn't work
		return FALSE;
	} else {
		return CWindow::IsDialogMessage(pMsg);
	}
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnCtlColorStatic(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(reinterpret_cast<HWND>(lParam) == loadingAnimation) {
		return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW));
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(controls.shtvw) {
		controls.shtvw->Detach();
	}
	if(controls.shlvw) {
		controls.shlvw->Detach();
	}
	splitter.SetSplitterPanes(NULL, NULL);

	if(controls.extvw) {
		IDispEventImpl<IDC_EXTVW, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>::DispEventUnadvise(controls.extvw);
		controls.extvw.Release();
	}
	if(controls.shtvw) {
		IDispEventImpl<IDC_SHTVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventUnadvise(controls.shtvw);
		controls.shtvw.Release();
	}
	if(controls.exlvw) {
		IDispEventImpl<IDC_EXLVW, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 0>::DispEventUnadvise(controls.exlvw);
		controls.exlvw.Release();
	}
	if(controls.shlvw) {
		IDispEventImpl<IDC_SHLVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShListViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventUnadvise(controls.shlvw);
		controls.shlvw.Release();
	}

	HICON hIcon = statusBar.GetIcon(1);
	if(hIcon) {
		DestroyIcon(hIcon);
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	wasHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnHScroll(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	int newPosition = viewModeSlider.GetPos();
	switch(newPosition) {
		case 1:
			ChangeViewMode(svmTiles);
			break;
		case 2:
			ChangeViewMode(svmDetails);
			break;
		case 3:
			ChangeViewMode(svmList);
			break;
		case 4:
			ChangeViewMode(svmSmallIcons);
			break;
		case 5:
			ChangeViewMode(svmIcons);
			break;
		default:
			if(newPosition >= 6) {
				// thumbnail size between 32 and 256 in steps of 8 pixels
				ChangeViewMode(svmThumbnails);
			}
			break;
	}
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	NONCLIENTMETRICS nonClientMetrics = {0};
	nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
	SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
	statusFont.CreateFontIndirect(&nonClientMetrics.lfStatusFont);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	if(IsComctl32Version600OrNewer()) {
		imageLists[0].Create(16, 16, ILC_COLOR32, 2, 0);
		HICON hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP161632), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
		imageLists[0].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND161632), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
		imageLists[0].AddIcon(hIcon);
		DestroyIcon(hIcon);

		imageLists[1].Create(32, 32, ILC_COLOR32, 2, 0);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP323232), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR));
		imageLists[1].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND323232), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR));
		imageLists[1].AddIcon(hIcon);
		DestroyIcon(hIcon);

		imageLists[2].Create(48, 48, ILC_COLOR32, 2, 0);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP484832), IMAGE_ICON, 48, 48, LR_DEFAULTCOLOR));
		imageLists[2].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND484832), IMAGE_ICON, 48, 48, LR_DEFAULTCOLOR));
		imageLists[2].AddIcon(hIcon);
		DestroyIcon(hIcon);
	} else {
		imageLists[0].Create(16, 16, ILC_COLOR24, 2, 0);
		HICON hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP16168), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
		imageLists[0].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND16168), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
		imageLists[0].AddIcon(hIcon);
		DestroyIcon(hIcon);

		imageLists[1].Create(32, 32, ILC_COLOR24, 2, 0);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP32328), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR));
		imageLists[1].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND32328), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR));
		imageLists[1].AddIcon(hIcon);
		DestroyIcon(hIcon);

		imageLists[2].Create(48, 48, ILC_COLOR24, 2, 0);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_UP48488), IMAGE_ICON, 48, 48, LR_DEFAULTCOLOR));
		imageLists[2].AddIcon(hIcon);
		DestroyIcon(hIcon);
		hIcon = reinterpret_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_FIND48488), IMAGE_ICON, 48, 48, LR_DEFAULTCOLOR));
		imageLists[2].AddIcon(hIcon);
		DestroyIcon(hIcon);
	}

	viewModeSlider.SubclassWindow(GetDlgItem(IDC_VIEWMODESLIDER));
	viewModeSlider.SetRange(1, 34);
	viewModeSlider.SetTicFreq(1);
	loadingAnimation.Attach(GetDlgItem(IDC_LOADANIMATION));

	// create the status bar
	statusBar.Create(*this, rcDefault, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | SBARS_SIZEGRIP);
	statusBar.SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(statusFont)), MAKELPARAM(TRUE, 0));

	extvwWnd.SubclassWindow(GetDlgItem(IDC_EXTVW));
	extvwWnd.QueryControl(__uuidof(ExTVwLibU::IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvw));
	if(controls.extvw) {
		IDispEventImpl<IDC_EXTVW, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>::DispEventAdvise(controls.extvw);
	}
	CAxWindow wnd;
	wnd.Attach(GetDlgItem(IDC_SHTVW));
	wnd.QueryControl(__uuidof(ShBrowserCtlsLibU::IShTreeView), reinterpret_cast<LPVOID*>(&controls.shtvw));
	wnd.Detach();
	if(controls.shtvw) {
		IDispEventImpl<IDC_SHTVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventAdvise(controls.shtvw);
	}
	exlvwWnd.SubclassWindow(GetDlgItem(IDC_EXLVW));
	exlvwWnd.QueryControl(__uuidof(ExLVwLibU::IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvw));
	if(controls.exlvw) {
		IDispEventImpl<IDC_EXLVW, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 0>::DispEventAdvise(controls.exlvw);
	}
	wnd.Attach(GetDlgItem(IDC_SHLVW));
	wnd.QueryControl(__uuidof(ShBrowserCtlsLibU::IShListView), reinterpret_cast<LPVOID*>(&controls.shlvw));
	wnd.Detach();
	if(controls.shlvw) {
		IDispEventImpl<IDC_SHLVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShListViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventAdvise(controls.shlvw);
	}

	// create the splitter window that will host the tree view and the list view
	splitter.Create(*this, rcDefault, NULL, WS_CHILDWINDOW | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN);

	// setup the splitter
	extvwWnd.SetParent(splitter);
	exlvwWnd.SetParent(splitter);
	splitter.SetSplitterPanes(extvwWnd, exlvwWnd);
	UpdateLayout();
	splitter.SetSplitterPosPct(25);

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.extvw->GethWnd())), L"explorer", NULL);
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvw->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	UISetCheck(ID_VIEW_STATUS_BAR, 1);

	SetupListView();
	SetupTreeView();
	ChangeViewMode(ExLvwViewModeToShellViewMode(controls.exlvw->GetView()));

	return TRUE;
}

LRESULT CMainDlg::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	UpdateLayout();
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*wasHandled*/)
{
	if(controls.shlvw) {
		controls.shlvw->About();
	}
	return 0;
}

LRESULT CMainDlg::OnToolTipGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFO pDetails = reinterpret_cast<LPNMTTDISPINFO>(pNotificationDetails);
	ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(TCHAR));

	TCHAR pBuffer[MAX_PATH];
	lstrcpyn(pBuffer, TEXT("View Mode: "), MAX_PATH);
	switch(viewModeSlider.GetPos()) {
		case 1:
			lstrcat(pBuffer, TEXT("Tiles"));
			break;
		case 2:
			lstrcat(pBuffer, TEXT("Details"));
			break;
		case 3:
			lstrcat(pBuffer, TEXT("List"));
			break;
		case 4:
			lstrcat(pBuffer, TEXT("Small Icons"));
			break;
		case 5:
			lstrcat(pBuffer, TEXT("Icons"));
			break;
		default:
			if(viewModeSlider.GetPos() >= 6) {
				// thumbnail size between 32 and 256 in steps of 8 pixels
				int viewMode = viewModeSlider.GetPos();
				viewMode = 24 + ((viewMode - 5) << 3);
				CAtlString s;
				s.Format(TEXT("Thumbnails (%ix%i)"), viewMode, viewMode);
				lstrcat(pBuffer, s);
			}
	}

	lstrcpyn(reinterpret_cast<LPTSTR>(pToolTipBuffer), pBuffer, MAX_PATH);
	pDetails->lpszText = reinterpret_cast<LPTSTR>(pToolTipBuffer);
	return 0;
}

LRESULT CMainDlg::OnViewStatusBar(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	BOOL isVisible = !statusBar.IsWindowVisible();
	statusBar.ShowWindow(isVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_STATUS_BAR, isVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainDlg::OnChangeViewMode(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	switch(controlID) {
		case ID_VIEW_FILMSTRIP:
			viewModeSlider.SetPos(((96 - 24) >> 3) + 5);
			ChangeViewMode(svmFilmstrip);
			break;
		case ID_VIEW_THUMBNAILS:
			viewModeSlider.SetPos(((96 - 24) >> 3) + 5);
			ChangeViewMode(svmThumbnails);
			break;
		case ID_VIEW_TILES:
			ChangeViewMode(svmTiles);
			break;
		case ID_VIEW_EXTENDEDTILES:
			ChangeViewMode(svmExtendedTiles);
			break;
		case ID_VIEW_ICONS:
			ChangeViewMode(svmIcons);
			break;
		case ID_VIEW_SMALLICONS:
			ChangeViewMode(svmSmallIcons);
			break;
		case ID_VIEW_LIST:
			ChangeViewMode(svmList);
			break;
		case ID_VIEW_DETAILS:
			ChangeViewMode(svmDetails);
			break;
	}
	return 0;
}

LRESULT CMainDlg::OnOrganizeFavorites(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	CFavoritesDlg dlgFavorites;
	dlgFavorites.DoModal();
	return 0;
}


void CMainDlg::ChangeViewMode(CMainDlg::ShellViewModeConstants newView)
{
	// configure the listview
	controls.exlvw->PutFullRowSelect(RunTimeHelper::IsVista() && newView == svmDetails ? ExLVwLibU::frsExtendedMode : ExLVwLibU::frsDisabled);
	controls.exlvw->PutSingleRow(newView == svmFilmstrip ? VARIANT_TRUE : VARIANT_FALSE);
	controls.exlvw->PutBorderSelect(!RunTimeHelper::IsVista() && (newView == svmFilmstrip || newView == svmThumbnails) ? VARIANT_TRUE : VARIANT_FALSE);
	switch(newView) {
		case svmFilmstrip:
		case svmThumbnails:
			controls.shlvw->PuthImageList(ShBrowserCtlsLibU::ilNonShellItems, HandleToLong(imageLists[2].m_hImageList));
			controls.shlvw->PutThumbnailSize(viewModeSlider.GetPos() >= 6 ? 24 + (viewModeSlider.GetPos() - 5) * 8 : -1);
			controls.shlvw->PutDisplayThumbnails(VARIANT_TRUE);
			break;
		case svmTiles:
			controls.shlvw->PuthImageList(ShBrowserCtlsLibU::ilNonShellItems, HandleToLong(imageLists[2].m_hImageList));
			controls.exlvw->PutTileViewItemLines(2);
			if(RunTimeHelper::IsVista()) {
				controls.shlvw->PutThumbnailSize(-4);
				controls.shlvw->PutDisplayThumbnails(VARIANT_TRUE);
			} else {
				controls.shlvw->PutDisplayThumbnails(VARIANT_FALSE);
			}
			break;
		case svmExtendedTiles:
			controls.shlvw->PuthImageList(ShBrowserCtlsLibU::ilNonShellItems, HandleToLong(imageLists[2].m_hImageList));
			controls.exlvw->PutTileViewItemLines(5);
			if(RunTimeHelper::IsVista()) {
				controls.shlvw->PutThumbnailSize(-4);
				controls.shlvw->PutDisplayThumbnails(VARIANT_TRUE);
			} else {
				controls.shlvw->PutDisplayThumbnails(VARIANT_FALSE);
			}
			break;
		case svmIcons:
			controls.shlvw->PuthImageList(ShBrowserCtlsLibU::ilNonShellItems, HandleToLong(imageLists[1].m_hImageList));
			if(RunTimeHelper::IsVista()) {
				controls.shlvw->PutThumbnailSize(-3);
				controls.shlvw->PutDisplayThumbnails(VARIANT_TRUE);
			} else {
				controls.shlvw->PutDisplayThumbnails(VARIANT_FALSE);
			}
			break;
		case svmSmallIcons:
		case svmList:
		case svmDetails:
			controls.shlvw->PuthImageList(ShBrowserCtlsLibU::ilNonShellItems, HandleToLong(imageLists[0].m_hImageList));
			controls.shlvw->PutDisplayThumbnails(VARIANT_FALSE);
			break;
	}
	controls.exlvw->PutView(ShellViewModeToExLvwViewMode(newView));

	shellViewMode = newView;
	if(shellViewMode == svmTiles || shellViewMode == svmExtendedTiles) {
		// on older systems these view modes may be missing, so make sure the menu is in sync
		shellViewMode = ExLvwViewModeToShellViewMode(controls.exlvw->GetView());
	}
	UINT viewModeMenuItemID = ID_VIEW_FIRST;
	switch(newView) {
		case svmFilmstrip:
			viewModeMenuItemID = ID_VIEW_FILMSTRIP;
			break;
		case svmThumbnails:
			viewModeMenuItemID = ID_VIEW_THUMBNAILS;
			break;
		case svmTiles:
			viewModeMenuItemID = ID_VIEW_TILES;
			viewModeSlider.SetPos(1);
			break;
		case svmExtendedTiles:
			viewModeMenuItemID = ID_VIEW_EXTENDEDTILES;
			viewModeSlider.SetPos(1);
			break;
		case svmIcons:
			viewModeMenuItemID = ID_VIEW_ICONS;
			viewModeSlider.SetPos(5);
			break;
		case svmSmallIcons:
			viewModeMenuItemID = ID_VIEW_SMALLICONS;
			viewModeSlider.SetPos(4);
			break;
		case svmList:
			viewModeMenuItemID = ID_VIEW_LIST;
			viewModeSlider.SetPos(3);
			break;
		case svmDetails:
			viewModeMenuItemID = ID_VIEW_DETAILS;
			viewModeSlider.SetPos(2);
			break;
	}
	HMENU hSubMenu = GetSubMenu(GetMenu(), 0);
	CheckMenuRadioItem(hSubMenu, ID_VIEW_FIRST, ID_VIEW_LAST, viewModeMenuItemID, MF_BYCOMMAND);
}

CMainDlg::ShellViewModeConstants CMainDlg::ExLvwViewModeToShellViewMode(ExLVwLibU::ViewConstants exlvwViewMode)
{
	switch(exlvwViewMode) {
		case ExLVwLibU::vTiles:
			return svmTiles;
		case ExLVwLibU::vExtendedTiles:
			return svmExtendedTiles;
		case ExLVwLibU::vIcons:
			return svmIcons;
		case ExLVwLibU::vSmallIcons:
			return svmSmallIcons;
		case ExLVwLibU::vList:
			return svmList;
		case ExLVwLibU::vDetails:
			return svmDetails;
	}
	return svmIcons;
}

void CMainDlg::InsertCustomBackgroundMenuCommands(HMENU hMenu, int minCmdID)
{
	// calculate command IDs
	COMMANDID_VIEW_FILMSTRIP = minCmdID + svmFilmstrip;
	COMMANDID_VIEW_THUMBNAILS = minCmdID + svmThumbnails;
	COMMANDID_VIEW_TILES = minCmdID + svmTiles;
	COMMANDID_VIEW_EXTENDEDTILES = minCmdID + svmExtendedTiles;
	COMMANDID_VIEW_ICONS = minCmdID + svmIcons;
	COMMANDID_VIEW_SMALLICONS = minCmdID + svmSmallIcons;
	COMMANDID_VIEW_LIST = minCmdID + svmList;
	COMMANDID_VIEW_DETAILS = minCmdID + svmDetails;

	COMMANDID_VIEW_FIRST = minCmdID;
	COMMANDID_VIEW_LAST = COMMANDID_VIEW_FIRST;
	COMMANDID_VIEW_LAST = max(max(COMMANDID_VIEW_FILMSTRIP, COMMANDID_VIEW_THUMBNAILS), max(max(COMMANDID_VIEW_TILES, COMMANDID_VIEW_EXTENDEDTILES), max(max(COMMANDID_VIEW_ICONS, COMMANDID_VIEW_SMALLICONS), max(COMMANDID_VIEW_LIST, COMMANDID_VIEW_DETAILS))));
	COMMANDID_VIEW = COMMANDID_VIEW_LAST + 1;
	COMMANDID_CUSTOMIZE = COMMANDID_VIEW + 1;

	MENUITEMINFO menuItem = {0};
	menuItem.cbSize = sizeof(menuItem);

	// insert a separator which will separate our items from the others
	menuItem.fMask = MIIM_FTYPE;
	menuItem.fType = MFT_SEPARATOR;
	InsertMenuItem(hMenu, 0, TRUE, &menuItem);

	// insert a view sub menu
	menuItem.fMask = MIIM_ID | MIIM_FTYPE | MIIM_STATE | MIIM_STRING | MIIM_SUBMENU;
	menuItem.fType = MFT_STRING;
	menuItem.fState = MFS_ENABLED;
	menuItem.wID = COMMANDID_VIEW;
	menuItem.dwTypeData = TEXT("View");
	HMENU hSubMenu = CreatePopupMenu();
	menuItem.hSubMenu = hSubMenu;
	InsertMenuItem(hMenu, 0, TRUE, &menuItem);

	// fill the view sub menu
	menuItem.fMask = MIIM_ID | MIIM_FTYPE | MIIM_STATE | MIIM_STRING;
	menuItem.wID = COMMANDID_VIEW_FILMSTRIP;
	menuItem.dwTypeData = TEXT("Filmstrip");
	InsertMenuItem(hSubMenu, 0, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_THUMBNAILS;
	menuItem.dwTypeData = TEXT("Thumbnails");
	InsertMenuItem(hSubMenu, 0, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_TILES;
	menuItem.dwTypeData = TEXT("Tiles");
	InsertMenuItem(hSubMenu, 0, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_EXTENDEDTILES;
	menuItem.dwTypeData = TEXT("Extended Tiles");
	InsertMenuItem(hSubMenu, 1, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_ICONS;
	menuItem.dwTypeData = TEXT("Icons");
	InsertMenuItem(hSubMenu, 2, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_SMALLICONS;
	menuItem.dwTypeData = TEXT("Small Icons");
	InsertMenuItem(hSubMenu, 3, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_LIST;
	menuItem.dwTypeData = TEXT("List");
	InsertMenuItem(hSubMenu, 4, TRUE, &menuItem);
	menuItem.wID = COMMANDID_VIEW_DETAILS;
	menuItem.dwTypeData = TEXT("Details");
	InsertMenuItem(hSubMenu, 5, TRUE, &menuItem);

	if(controls.shlvw->GetNamespaces()->GetItem(0, ShBrowserCtlsLibU::slnsitIndex)->GetCustomizable()) {
		// insert the customize item
		menuItem.wID = COMMANDID_CUSTOMIZE;
		menuItem.dwTypeData = TEXT("Customize this folder...");
		InsertMenuItem(hMenu, 1, TRUE, &menuItem);
	}

	// check the current view mode
	CheckMenuRadioItem(hSubMenu, COMMANDID_VIEW_FIRST, COMMANDID_VIEW_LAST, minCmdID + shellViewMode, MF_BYCOMMAND);
}

void CMainDlg::SetupListView(void)
{
	if(controls.shlvw) {
		controls.shlvw->Attach(controls.exlvw->GethWnd());
		controls.shlvw->PuthWndShellUIParentWindow(HandleToLong(static_cast<HWND>(*this)));
	}
}

void CMainDlg::SetupTreeView(void)
{
	if(controls.shtvw) {
		controls.shtvw->Attach(controls.extvw->GethWnd());
		controls.shtvw->PuthWndShellUIParentWindow(HandleToLong(static_cast<HWND>(*this)));
	}

	LPITEMIDLIST pIDLDesktop = NULL;
	SHGetFolderLocation(*this, CSIDL_DESKTOP, NULL, 0, &pIDLDesktop);
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = *reinterpret_cast<LONG*>(&pIDLDesktop);
	BSTR bstr = NULL;
	CComPtr<ShBrowserCtlsLibU::IShTreeViewItem> pItem = controls.shtvw->GetTreeItems()->Add(v, NULL, static_cast<OLE_HANDLE>(ExTVwLibU::iaFirst), static_cast<ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants>(-1), bstr, ExTVwLibU::heYes, -2, -2, -2, 0, VARIANT_FALSE, 1);
	VariantClear(&v);
	if(pItem) {
		CComQIPtr<ExTVwLibU::ITreeViewItem> pCaret = pItem->GetTreeViewItemObject();
		controls.extvw->putref_CaretItem(VARIANT_TRUE, pCaret);
		pCaret->Expand(VARIANT_FALSE);
	}
}

ExLVwLibU::ViewConstants CMainDlg::ShellViewModeToExLvwViewMode(CMainDlg::ShellViewModeConstants shViewMode)
{
	switch(shViewMode) {
		case svmFilmstrip:
			return ExLVwLibU::vIcons;
		case svmThumbnails:
			return ExLVwLibU::vIcons;
		case svmTiles:
			return ExLVwLibU::vTiles;
		case svmExtendedTiles:
			return ExLVwLibU::vExtendedTiles;
		case svmIcons:
			return ExLVwLibU::vIcons;
		case svmSmallIcons:
			return ExLVwLibU::vSmallIcons;
		case svmList:
			return ExLVwLibU::vList;
		case svmDetails:
			return ExLVwLibU::vDetails;
	}
	return ExLVwLibU::vIcons;
}

void CMainDlg::UpdateLayout(void)
{
	RECT clientRectangle = {0};
	GetClientRect(&clientRectangle);
	if(statusBar.IsWindowVisible()) {
		statusBar.SendMessage(WM_SIZE, 0, 0);
		RECT rc = {0};
		statusBar.GetWindowRect(&rc);
		clientRectangle.bottom -= rc.bottom - rc.top;
		int parts[] = {clientRectangle.right - clientRectangle.left - 100, clientRectangle.right - clientRectangle.left};
		statusBar.SetParts(2, parts);
	}
	if(viewModeSlider.IsWindow()) {
		RECT currentSliderRect = {0};
		viewModeSlider.GetWindowRect(&currentSliderRect);
		RECT sliderRect = clientRectangle;
		sliderRect.top = sliderRect.bottom - (currentSliderRect.bottom - currentSliderRect.top);
		viewModeSlider.SetWindowPos(NULL, &sliderRect, SWP_NOZORDER | SWP_NOACTIVATE);
		clientRectangle.bottom -= sliderRect.bottom - sliderRect.top;
	}
	if(splitter.IsWindow()) {
		splitter.SetWindowPos(NULL, &clientRectangle, SWP_NOZORDER | SWP_NOACTIVATE);
	}
	RECT exlvwRect = {0};
	exlvwWnd.GetWindowRect(&exlvwRect);
	ScreenToClient(&exlvwRect);
	loadingAnimation.SetWindowPos(NULL, &exlvwRect, SWP_NOZORDER | SWP_NOACTIVATE);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}


void __stdcall CMainDlg::CaretChangedExtvw(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long /*caretChangeReason*/)
{
	CComQIPtr<ExTVwLibU::ITreeViewItem> pNewCaret = newCaretItem;
	controls.shlvw->GetNamespaces()->RemoveAll(VARIANT_TRUE);
	if(pNewCaret->GetParentItem()) {
		controls.exlvw->GetListItems()->Add(bstr_t(".."), -1, 0, 0, 0, -2, _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR));
	}
	CComQIPtr<ShBrowserCtlsLibU::IShTreeViewItem> pNewShellCaret = pNewCaret->GetShellTreeViewItemObject();
	if(pNewShellCaret) {
		LONG l = pNewShellCaret->GetFullyQualifiedPIDL();
		LPITEMIDLIST pIDLNamespace = ILClone(*reinterpret_cast<LPITEMIDLIST*>(&l));
		VARIANT v;
		VariantInit(&v);
		v.vt = VT_I4;
		v.lVal = *reinterpret_cast<LONG*>(&pIDLNamespace);
		CComPtr<ShBrowserCtlsLibU::IShListViewNamespace> pNamespace = controls.shlvw->GetNamespaces()->Add(v, -1, NULL, ShBrowserCtlsLibU::asiAutoSortOnce);
		if(!pNamespace) {
			CComQIPtr<ExTVwLibU::ITreeViewItem> pPrevCaret = previousCaretItem;
			if(pPrevCaret) {
				controls.extvw->PutRefCaretItem(VARIANT_TRUE, pPrevCaret);
			}
		}
		VariantClear(&v);

		CComPtr<ExTVwLibU::ITreeViewItem> pCaretItem = controls.extvw->GetCaretItem(VARIANT_TRUE);
		CComQIPtr<ShBrowserCtlsLibU::IShTreeViewItem> pShellCaretItem = pCaretItem->GetShellTreeViewItemObject();
		_bstr_t s = TEXT("Address: ") + pShellCaretItem->GetDisplayName(ShBrowserCtlsLibU::dntAddressBarNameFollowSysSettings, VARIANT_FALSE);
		statusBar.SetText(0, s);

		CComPtr<ShBrowserCtlsLibU::ISecurityZone> pSecZone = NULL;
		pShellCaretItem->get_SecurityZone(&pSecZone);
		HICON hOldIcon = statusBar.GetIcon(1);
		HICON hIcon = NULL;
		if(pSecZone) {
			s = pSecZone->GetDisplayName();
			hIcon = static_cast<HICON>(LongToHandle(pSecZone->GethIcon()));
		} else {
			s = TEXT("");
		}
		statusBar.SetText(1, s);
		statusBar.SetIcon(1, hIcon);
		if(hOldIcon) {
			DestroyIcon(hOldIcon);
		}
	}
}

void __stdcall CMainDlg::DestroyedControlWindowExtvw(long /*hWnd*/)
{
	controls.shtvw->Detach();
}

void __stdcall CMainDlg::RecreatedControlWindowExtvw(long /*hWnd*/)
{
	SetupTreeView();
}

void __stdcall CMainDlg::ChangedContextMenuSelectionShtvw(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
{
	statusBar.SetSimple(hContextMenu != NULL);
	if(hContextMenu) {
		CAtlString helpText = TEXT("");
		if(isCustomMenuItem == VARIANT_FALSE) {
			VARIANT v;
			VariantInit(&v);
			controls.shtvw->GetShellContextMenuItemString(commandID, &v, NULL);
			helpText = v.bstrVal;
			VariantClear(&v);
		}
		statusBar.SetText(SB_SIMPLEID, helpText);
	}
}

void __stdcall CMainDlg::ItemEnumerationCompletedShtvw(LPDISPATCH Namespace, VARIANT_BOOL /*aborted*/)
{
	if(Namespace) {
		CComQIPtr<ShBrowserCtlsLibU::IShTreeViewItem> pShellTreeItem = Namespace;
		CComQIPtr<ExTVwLibU::ITreeViewItem> pTreeItem = pShellTreeItem->GetTreeViewItemObject();
		CComPtr<ExTVwLibU::ITreeViewItems> pSubItems = pTreeItem->GetSubItems();
		pSubItems->PutFilterType(ExTVwLibU::fpItemData, ExTVwLibU::ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray<VARIANT> arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(1));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerTreeView expects an array of VARIANTs!
		pSubItems->PutFilter(ExTVwLibU::fpItemData, filter);
		pSubItems->RemoveAll();
	}
}

void __stdcall CMainDlg::ItemEnumerationTimedOutShtvw(LPDISPATCH Namespace)
{
	if(Namespace) {
		CComQIPtr<ShBrowserCtlsLibU::IShTreeViewItem> pShellTreeItem = Namespace;
		CComQIPtr<ExTVwLibU::ITreeViewItem> pTreeItem = pShellTreeItem->GetTreeViewItemObject();
		pTreeItem->GetSubItems()->Add(_bstr_t(TEXT("Loading")), NULL, _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), ExTVwLibU::heNo, 1, -2, -2, 1, VARIANT_FALSE, 1)->PutBold(VARIANT_TRUE);
	}
}

void __stdcall CMainDlg::CaretChangedExlvw(LPDISPATCH /*previousCaretItem*/, LPDISPATCH newCaretItem)
{
	CComQIPtr<ExLVwLibU::IListViewItem> pNewCaret = newCaretItem;
	if(pNewCaret) {
		CComQIPtr<ShBrowserCtlsLibU::IShListViewItem> pNewShellCaret = pNewCaret->GetShellListViewItemObject();
		if(pNewShellCaret) {
			if(pNewShellCaret->GetShellAttributes(static_cast<ShBrowserCtlsLibU::ItemShellAttributesConstants>(0xFEFFE17F), VARIANT_FALSE) & ShBrowserCtlsLibU::isaIsLink) {
				_bstr_t s = TEXT("Target: ") + pNewShellCaret->GetLinkTarget();
				statusBar.SetText(0, s);
			}
		}
	}
}

void __stdcall CMainDlg::DestroyedControlWindowExlvw(long /*hWnd*/)
{
	controls.shlvw->Detach();
}

void __stdcall CMainDlg::ItemActivateExlvw(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*shift*/, long /*x*/, long /*y*/)
{
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(!pItem->GetShellListViewItemObject()) {
		// must be our custom ".." item
		LONG l = controls.shlvw->GetNamespaces()->GetItem(0, ShBrowserCtlsLibU::slnsitIndex)->GetFullyQualifiedPIDL();
		LPITEMIDLIST pIDLToOpen = ILClone(*reinterpret_cast<LPITEMIDLIST*>(&l));
		ILRemoveLastID(pIDLToOpen);
		VARIANT v;
		VariantInit(&v);
		v.vt = VT_I4;
		v.lVal = *reinterpret_cast<LONG*>(&pIDLToOpen);

		if(pIDLToOpen) {
			controls.shtvw->EnsureItemIsLoaded(v);
			CComPtr<ShBrowserCtlsLibU::IShTreeViewItem> pShellTvwItem = controls.shtvw->GetTreeItems()->GetItem(v, ShBrowserCtlsLibU::stiitEqualPIDL);
			if(pShellTvwItem) {
				CComQIPtr<ExTVwLibU::ITreeViewItem> pTvwItem = pShellTvwItem->GetTreeViewItemObject();
				controls.extvw->PutRefCaretItem(VARIANT_TRUE, pTvwItem);
			}
		}
		VariantClear(&v);
		ILFree(pIDLToOpen);
	} else {
		CComPtr<ExLVwLibU::IListViewItems> pCollection = controls.exlvw->GetListItems();
		pCollection->PutFilterType(ExLVwLibU::fpSelected, ExLVwLibU::ftIncluding);
		_variant_t filter;
		filter.Clear();
		CComSafeArray<VARIANT> arr;
		arr.Create(1, 1);
		arr.SetAt(1, _variant_t(true));
		filter.parray = arr.Detach();
		filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
		pCollection->PutFilter(ExLVwLibU::fpSelected, filter);

		controls.shlvw->InvokeDefaultShellContextMenuCommand(_variant_t(pCollection));
	}
}

void __stdcall CMainDlg::ItemGetInfoTipTextExlvw(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/)
{
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(!pItem->GetShellListViewItemObject()) {
		*infoTipText = SysAllocString(OLESTR("Go to parent folder"));
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvw(long /*hWnd*/)
{
	SetupListView();
}

void __stdcall CMainDlg::StartingLabelEditingExlvw(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing)
{
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	*cancelEditing = (pItem->GetShellListViewItemObject() ? VARIANT_FALSE : VARIANT_TRUE);
}

void __stdcall CMainDlg::ChangedContextMenuSelectionShlvw(OLE_HANDLE hContextMenu, VARIANT_BOOL isShellContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
{
	if(isShellContextMenu != VARIANT_FALSE) {
		statusBar.SetSimple(hContextMenu != NULL);
		if(hContextMenu) {
			CAtlString helpText = TEXT("");
			if(isCustomMenuItem == VARIANT_FALSE) {
				VARIANT v;
				VariantInit(&v);
				controls.shlvw->GetShellContextMenuItemString(commandID, &v, NULL);
				helpText = v.bstrVal;
				VariantClear(&v);
			}
			statusBar.SetText(SB_SIMPLEID, helpText);
		}
	}
}

void __stdcall CMainDlg::CreatedShellContextMenuShlvw(LPDISPATCH Items, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* /*cancelMenu*/)
{
	if(Items) {
		CComQIPtr<ShBrowserCtlsLibU::IShListViewNamespace> pShellListNamespace = Items;
		if(pShellListNamespace) {
			InsertCustomBackgroundMenuCommands(static_cast<HMENU>(LongToHandle(hContextMenu)), minimumCustomCommandID);
		}
	}
}

void __stdcall CMainDlg::ItemEnumerationCompletedShlvw(LPDISPATCH /*Namespace*/, VARIANT_BOOL /*aborted*/)
{
	if(!RunTimeHelper::IsVista()) {
		loadingAnimation.ShowWindow(SW_HIDE);
		loadingAnimation.Stop();
	}
}

void __stdcall CMainDlg::ItemEnumerationTimedOutShlvw(LPDISPATCH /*Namespace*/)
{
	if(!RunTimeHelper::IsVista()) {
		loadingAnimation.ShowWindow(SW_SHOW);
		loadingAnimation.BringWindowToTop();

		HMODULE hMod = LoadLibraryEx(TEXT("shell32.dll"), NULL, LOAD_LIBRARY_AS_DATAFILE);
		if(hMod) {
			loadingAnimation.SendMessage(ACM_OPEN, reinterpret_cast<WPARAM>(hMod), reinterpret_cast<LPARAM>(MAKEINTRESOURCE(150)));
			FreeLibrary(hMod);
		}
	}
}

void __stdcall CMainDlg::SelectedShellContextMenuItemShlvw(LPDISPATCH Items, OLE_HANDLE /*hContextMenu*/, LONG commandID, ShBrowserCtlsLibU::WindowModeConstants* /*windowModeToUse*/, ShBrowserCtlsLibU::CommandInvocationFlagsConstants* /*invocationFlagsToUse*/, VARIANT_BOOL* executeCommand)
{
	if(Items) {
		CComQIPtr<ExLVwLibU::IListViewItemContainer> pItemContainer = Items;
		if(pItemContainer) {
			VARIANT verb;
			VariantInit(&verb);
			controls.shlvw->GetShellContextMenuItemString(commandID, NULL, &verb);
			if(SUCCEEDED(VariantChangeType(&verb, &verb, 0, VT_BSTR))) {
				CComBSTR bstrVerb = verb.bstrVal;
				bstrVerb.ToLower();
				if(bstrVerb == OLESTR("open") || bstrVerb == OLESTR("explore")) {
					if(pItemContainer->Count() == 1) {
						PIDLIST_ABSOLUTE pIDLToOpen = NULL;
						CComQIPtr<IEnumVARIANT> pEnum = pItemContainer->Get_NewEnum();
						VARIANT vItem;
						VariantInit(&vItem);
						while(pEnum->Next(1, &vItem, NULL) == S_OK) {
							if(vItem.vt == VT_DISPATCH) {
								CComQIPtr<ExLVwLibU::IListViewItem> pLvwItem = vItem.pdispVal;
								if(pLvwItem) {
									CComQIPtr<ShBrowserCtlsLibU::IShListViewItem> pShellItem = pLvwItem->GetShellListViewItemObject();
									if(pShellItem->GetShellAttributes(static_cast<ShBrowserCtlsLibU::ItemShellAttributesConstants>(0xFEFFE17F), VARIANT_FALSE) & ShBrowserCtlsLibU::isaIsFolder) {
										LONG p = pShellItem->GetFullyQualifiedPIDL();
										pIDLToOpen = ILClone(*reinterpret_cast<LPCITEMIDLIST*>(&p));
										*executeCommand = VARIANT_FALSE;
									} else if(pShellItem->GetShellAttributes(static_cast<ShBrowserCtlsLibU::ItemShellAttributesConstants>(0xFEFFE17F), VARIANT_FALSE) & ShBrowserCtlsLibU::isaIsLink) {
										LONG p = pShellItem->GetLinkTargetPIDL();
										pIDLToOpen = ILClone(*reinterpret_cast<LPCITEMIDLIST*>(&p));
										*executeCommand = VARIANT_FALSE;
									}
								}
							}
							VariantClear(&vItem);
						}

						if(pIDLToOpen) {
							// check whether the target is a filesystem file
							TCHAR pBuffer[1024];
							SHGetPathFromIDList(pIDLToOpen, pBuffer);
							if(lstrlen(pBuffer) > 0 && !PathIsDirectory(pBuffer) && bstrVerb != OLESTR("explore")) {
								// it is a file, so don't try to browse it (except the user has selected "explore")
								ILFree(pIDLToOpen);
								pIDLToOpen = NULL;
								*executeCommand = VARIANT_TRUE;
							}
						}

						if(pIDLToOpen) {
							CWaitCursor wait;
							VARIANT v;
							VariantInit(&v);
							v.vt = VT_I4;
							v.lVal = *reinterpret_cast<LONG*>(&pIDLToOpen);
							controls.shtvw->EnsureItemIsLoaded(v);
							CComQIPtr<ShBrowserCtlsLibU::IShTreeViewItem> pShellItem = controls.shtvw->GetTreeItems()->GetItem(v.lVal, ShBrowserCtlsLibU::stiitEqualPIDL);
							if(pShellItem) {
								CComQIPtr<ExTVwLibU::ITreeViewItem> pTreeItem = pShellItem->GetTreeViewItemObject();
								controls.extvw->putref_CaretItem(VARIANT_TRUE, pTreeItem);
							}
							VariantClear(&v);
							ILFree(pIDLToOpen);
						}
					}
				}
			}
			VariantClear(&verb);
		} else {
			CComQIPtr<ShBrowserCtlsLibU::IShListViewNamespace> pNamespace = Items;
			if(commandID >= COMMANDID_VIEW_FIRST && commandID <= COMMANDID_VIEW_LAST) {
				// the user wants to change the view mode
				if(commandID == COMMANDID_VIEW_FILMSTRIP || commandID == COMMANDID_VIEW_THUMBNAILS) {
					viewModeSlider.SetPos(((96 - 24) >> 3) + 5);
				}
				ChangeViewMode(static_cast<ShellViewModeConstants>(commandID - COMMANDID_VIEW_FIRST));
			} else if(commandID == COMMANDID_CUSTOMIZE) {
				pNamespace->Customize();
			}
		}
	}
}
