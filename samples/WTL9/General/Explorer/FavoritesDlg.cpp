// FavoritesDlg.cpp : implementation of the CFavoritesDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "FavoritesDlg.h"
#include ".\FavoritesDlg.h"


BOOL CFavoritesDlg::PreTranslateMessage(MSG* pMsg)
{
	if((pMsg->message == WM_KEYDOWN) && ((pMsg->wParam == VK_ESCAPE) || (pMsg->wParam == VK_RETURN))) {
		// otherwise label-edit wouldn't work
		return FALSE;
	} else {
		return CWindow::IsDialogMessage(pMsg);
	}
}

LRESULT CFavoritesDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(controls.shtvw) {
		controls.shtvw->Detach();
	}

	if(controls.extvw) {
		IDispEventImpl<IDC_EXTVW, CFavoritesDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>::DispEventUnadvise(controls.extvw);
		controls.extvw.Release();
	}
	if(controls.shtvw) {
		IDispEventImpl<IDC_SHTVW, CFavoritesDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventUnadvise(controls.shtvw);
		controls.shtvw.Release();
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT CFavoritesDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	extvwWnd.SubclassWindow(GetDlgItem(IDC_EXTVW));
	extvwWnd.QueryControl(__uuidof(ExTVwLibU::IExplorerTreeView), reinterpret_cast<LPVOID*>(&controls.extvw));
	if(controls.extvw) {
		IDispEventImpl<IDC_EXTVW, CFavoritesDlg, &__uuidof(ExTVwLibU::_IExplorerTreeViewEvents), &LIBID_ExTVwLibU, 2, 0>::DispEventAdvise(controls.extvw);
	}
	CAxWindow wnd;
	wnd.Attach(GetDlgItem(IDC_SHTVW));
	wnd.QueryControl(__uuidof(ShBrowserCtlsLibU::IShTreeView), reinterpret_cast<LPVOID*>(&controls.shtvw));
	wnd.Detach();
	if(controls.shtvw) {
		IDispEventImpl<IDC_SHTVW, CFavoritesDlg, &__uuidof(ShBrowserCtlsLibU::_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, 1, 5>::DispEventAdvise(controls.shtvw);
	}

	if(controls.shtvw) {
		controls.shtvw->Attach(controls.extvw->GethWnd());
		controls.shtvw->PuthWndShellUIParentWindow(HandleToLong(static_cast<HWND>(*this)));
	}

	LPITEMIDLIST pIDLFavorites = NULL;
	SHGetFolderLocation(*this, CSIDL_FAVORITES, NULL, 0, &pIDLFavorites);
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;
	v.lVal = *reinterpret_cast<LONG*>(&pIDLFavorites);
	controls.shtvw->GetNamespaces()->Add(v, NULL, NULL, NULL, VARIANT_TRUE);
	VariantClear(&v);
	CComQIPtr<ExTVwLibU::ITreeViewItem> pCaret = controls.extvw->GetFirstRootItem();
	if(pCaret) {
		controls.extvw->putref_CaretItem(VARIANT_TRUE, pCaret);
		pCaret->Expand(VARIANT_FALSE);
	}

	return TRUE;
}

LRESULT CFavoritesDlg::OnAction(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	MessageBox(TEXT("TODO"));
	return 0;
}

LRESULT CFavoritesDlg::OnEndDialog(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	EndDialog(wID);
	return 0;
}
