// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include "FavoritesDlg.h"

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CUpdateUI<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EXTVW, CMainDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>,
    public IDispEventImpl<IDC_SHTVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_EXLVW, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 0>,
    public IDispEventImpl<IDC_SHLVW, CMainDlg, &__uuidof(ShBrowserCtlsLibU::_IShListViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>
{
	typedef enum ShellViewModeConstants
	{
		svmFilmstrip,
		svmThumbnails,
		svmTiles,
		svmExtendedTiles,
		svmIcons,
		svmSmallIcons,
		svmList,
		svmDetails
	} ShellViewModeConstants;

	ShellViewModeConstants shellViewMode;

	int COMMANDID_CUSTOMIZE;
	int COMMANDID_VIEW;
	int COMMANDID_VIEW_FIRST;
	int COMMANDID_VIEW_FILMSTRIP;
	int COMMANDID_VIEW_THUMBNAILS;
	int COMMANDID_VIEW_TILES;
	int COMMANDID_VIEW_EXTENDEDTILES;
	int COMMANDID_VIEW_ICONS;
	int COMMANDID_VIEW_SMALLICONS;
	int COMMANDID_VIEW_LIST;
	int COMMANDID_VIEW_DETAILS;
	int COMMANDID_VIEW_LAST;

	LPVOID pToolTipBuffer;

public:
	enum { IDD = IDD_MAINDLG };

	CFont statusFont;
	CSplitterWindow splitter;
	CAnimateCtrl loadingAnimation;
	CContainedWindowT<CTrackBarCtrl> viewModeSlider;
	CStatusBarCtrl statusBar;
	CImageList imageLists[3];

	CContainedWindowT<CAxWindow> extvwWnd;
	CContainedWindowT<CAxWindow> exlvwWnd;

	CMainDlg() :
	    extvwWnd(this, 1),
	    exlvwWnd(this, 2),
	    viewModeSlider(this, 3),
	    shellViewMode(svmTiles)
	{
		pToolTipBuffer = HeapAlloc(GetProcessHeap(), 0, (MAX_PATH + 1) * sizeof(TCHAR));
	}

	~CMainDlg()
	{
		if(pToolTipBuffer) {
			HeapFree(GetProcessHeap(), 0, pToolTipBuffer);
		}
	}

	struct Controls
	{
		CComPtr<ExTVwLibU::IExplorerTreeView> extvw;
		CComPtr<ShBrowserCtlsLibU::IShTreeView> shtvw;
		CComPtr<ExLVwLibU::IExplorerListView> exlvw;
		CComPtr<ShBrowserCtlsLibU::IShListView> shlvw;

		~Controls()
		{
		}
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_UPDATE_UI_MAP(CMainDlg)
		UPDATE_ELEMENT(ID_VIEW_STATUS_BAR, UPDUI_MENUPOPUP)
	END_UPDATE_UI_MAP()

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_CTLCOLORSTATIC, OnCtlColorStatic)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_HSCROLL, OnHScroll)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_SIZE, OnSize)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_VIEW_STATUS_BAR, OnViewStatusBar)
		COMMAND_RANGE_HANDLER(ID_VIEW_FIRST, ID_VIEW_LAST, OnChangeViewMode)
		COMMAND_ID_HANDLER(ID_FAVORITES_ORGANIZE, OnOrganizeFavorites)

		CHAIN_MSG_MAP(CUpdateUI<CMainDlg>)

		REFLECT_NOTIFICATIONS_EX()

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFO, OnToolTipGetDispInfoNotification)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXTVW, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 0, CaretChangedExtvw)
		SINK_ENTRY_EX(IDC_EXTVW, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 13, DestroyedControlWindowExtvw)
		SINK_ENTRY_EX(IDC_EXTVW, __uuidof(ExTVwLibU::_IExplorerTreeViewEvents), 74, RecreatedControlWindowExtvw)

		SINK_ENTRY_EX(IDC_SHTVW, __uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), 1, ChangedContextMenuSelectionShtvw)
		SINK_ENTRY_EX(IDC_SHTVW, __uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), 11, ItemEnumerationCompletedShtvw)
		SINK_ENTRY_EX(IDC_SHTVW, __uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), 13, ItemEnumerationTimedOutShtvw)

		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 0, CaretChangedExlvw)
		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 21, DestroyedControlWindowExlvw)
		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 90, ItemActivateExlvw)
		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 94, ItemGetInfoTipTextExlvw)
		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 124, RecreatedControlWindowExlvw)
		SINK_ENTRY_EX(IDC_EXLVW, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 136, StartingLabelEditingExlvw)

		SINK_ENTRY_EX(IDC_SHLVW, __uuidof(ShBrowserCtlsLibU::_IShListViewEvents), 2, ChangedContextMenuSelectionShlvw)
		SINK_ENTRY_EX(IDC_SHLVW, __uuidof(ShBrowserCtlsLibU::_IShListViewEvents), 12, CreatedShellContextMenuShlvw)
		SINK_ENTRY_EX(IDC_SHLVW, __uuidof(ShBrowserCtlsLibU::_IShListViewEvents), 21, ItemEnumerationCompletedShlvw)
		SINK_ENTRY_EX(IDC_SHLVW, __uuidof(ShBrowserCtlsLibU::_IShListViewEvents), 23, ItemEnumerationTimedOutShlvw)
		SINK_ENTRY_EX(IDC_SHLVW, __uuidof(ShBrowserCtlsLibU::_IShListViewEvents), 28, SelectedShellContextMenuItemShlvw)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	LRESULT OnCtlColorStatic(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	LRESULT OnHScroll(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	LRESULT OnViewStatusBar(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	LRESULT OnChangeViewMode(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	LRESULT OnOrganizeFavorites(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnToolTipGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);

	void CMainDlg::ChangeViewMode(ShellViewModeConstants newView);
	ShellViewModeConstants CMainDlg::ExLvwViewModeToShellViewMode(ExLVwLibU::ViewConstants exlvwViewMode);
	void InsertCustomBackgroundMenuCommands(HMENU hMenu, int minCmdID);
	void SetupListView(void);
	void SetupTreeView(void);
	ExLVwLibU::ViewConstants ShellViewModeToExLvwViewMode(ShellViewModeConstants shViewMode);
	void UpdateLayout(void);
	void CloseDialog(int nVal);

	void __stdcall CaretChangedExtvw(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem, long /*caretChangeReason*/);
	void __stdcall DestroyedControlWindowExtvw(long /*hWnd*/);
	void __stdcall RecreatedControlWindowExtvw(long /*hWnd*/);

	void __stdcall ChangedContextMenuSelectionShtvw(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem);
	void __stdcall ItemEnumerationCompletedShtvw(LPDISPATCH Namespace, VARIANT_BOOL /*aborted*/);
	void __stdcall ItemEnumerationTimedOutShtvw(LPDISPATCH Namespace);

	void __stdcall CaretChangedExlvw(LPDISPATCH /*previousCaretItem*/, LPDISPATCH newCaretItem);
	void __stdcall DestroyedControlWindowExlvw(long /*hWnd*/);
	void __stdcall ItemActivateExlvw(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*shift*/, long /*x*/, long /*y*/);
	void __stdcall ItemGetInfoTipTextExlvw(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/);
	void __stdcall RecreatedControlWindowExlvw(long /*hWnd*/);
	void __stdcall StartingLabelEditingExlvw(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing);

	void __stdcall ChangedContextMenuSelectionShlvw(OLE_HANDLE hContextMenu, VARIANT_BOOL isShellContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem);
	void __stdcall CreatedShellContextMenuShlvw(LPDISPATCH Items, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* /*cancelMenu*/);
	void __stdcall ItemEnumerationCompletedShlvw(LPDISPATCH /*Namespace*/, VARIANT_BOOL /*aborted*/);
	void __stdcall ItemEnumerationTimedOutShlvw(LPDISPATCH /*Namespace*/);
	void __stdcall SelectedShellContextMenuItemShlvw(LPDISPATCH Items, OLE_HANDLE /*hContextMenu*/, LONG commandID, ShBrowserCtlsLibU::WindowModeConstants* /*windowModeToUse*/, ShBrowserCtlsLibU::CommandInvocationFlagsConstants* /*invocationFlagsToUse*/, VARIANT_BOOL* executeCommand);
};
