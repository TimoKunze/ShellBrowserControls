// FavoritesDlg.h : interface of the CFavoritesDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once

class CFavoritesDlg :
    public CAxDialogImpl<CFavoritesDlg>,
    public CMessageFilter,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CFavoritesDlg>,
    public IDispEventImpl<IDC_EXTVW, CFavoritesDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>,
    public IDispEventImpl<IDC_SHTVW, CFavoritesDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>
{
public:
	enum { IDD = IDD_FAVORITES };

	CContainedWindowT<CAxWindow> extvwWnd;

	CFavoritesDlg() :
	    extvwWnd(this, 1)
	{
	}

	struct Controls
	{
		CComPtr<ExTVwLibU::IExplorerTreeView> extvw;
		CComPtr<ShBrowserCtlsLibU::IShTreeView> shtvw;

		~Controls()
		{
		}
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CFavoritesDlg)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_CREATEFOLDER, OnAction)
		COMMAND_ID_HANDLER(ID_RENAME, OnAction)
		COMMAND_ID_HANDLER(ID_MOVETOFOLDER, OnAction)
		COMMAND_ID_HANDLER(ID_DELETE, OnAction)
		COMMAND_ID_HANDLER(IDOK, OnEndDialog)
		COMMAND_ID_HANDLER(IDCANCEL, OnEndDialog)

		REFLECT_NOTIFICATIONS_EX()

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CFavoritesDlg)
	END_SINK_MAP()

	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	LRESULT OnAction(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEndDialog(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};
