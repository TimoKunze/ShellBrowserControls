// DummyShellBrowser.cpp: Hooks the desktop's IShellFolder implementation

#include "stdafx.h"
#include "DummyShellBrowser.h"


#ifdef ACTIVATE_SECZONE_FEATURES
#if FALSE
	DummyShellBrowser::DummyShellBrowser()
	{
		zone = URLZONE_LOCAL_MACHINE;
	}


	//////////////////////////////////////////////////////////////////////
	// implementation of IOleWindow
	STDMETHODIMP DummyShellBrowser::GetWindow(HWND *phwnd)
	{
		*phwnd = hWnd;
		return S_OK;
	}

	STDMETHODIMP DummyShellBrowser::ContextSensitiveHelp(BOOL fEnterMode)
	{
		return E_NOTIMPL;
	}
	// implementation of IOleWindow
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	// implementation of IShellBrowser
	STDMETHODIMP DummyShellBrowser::InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::RemoveMenusSB(HMENU hmenuShared)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::SetStatusTextSB(LPCWSTR pszStatusText)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::EnableModelessSB(BOOL fEnable)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::TranslateAcceleratorSB(MSG *pmsg, WORD wID)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::BrowseObject(PCUIDLIST_RELATIVE pidl, UINT wFlags)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::GetControlWindow(UINT id, HWND *phwnd)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::QueryActiveShellView(IShellView **ppshv)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::OnViewWindowActive(IShellView *pshv)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::SetToolbarItems(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags)
	{
		return E_NOTIMPL;
	}
	// implementation of IShellBrowser
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	// implementation of IServiceProvider
	STDMETHODIMP DummyShellBrowser::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
	{
		return QueryInterface(riid, ppvObject);
	}
	// implementation of IServiceProvider
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	// implementation of IOleCommandTarget
	STDMETHODIMP DummyShellBrowser::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DummyShellBrowser::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
	{
		if(*pguidCmdGroup == CGID_Explorer) {
			#define SBCMDID_MIXEDZONE 39
			if(nCmdID == SBCMDID_MIXEDZONE) {
				switch(pvaIn->vt) {
					case VT_UI4:
						// seems like the zone is reported in HIWORD and LOWORD contains the zone panel index (WTF?!)
						// however, HIWORD seems to be always 0
						zone = HIWORD(pvaIn->ulVal);
						break;
					case VT_NULL:
						zone = -2;/*ZONE_MIXED?*/;
						break;
					default:
						zone = -3;/*ZONE_UNKNOWN*/;
						break;
				}
				return S_OK;
			}
			return OLECMDERR_E_NOTSUPPORTED;
		}
		return OLECMDERR_E_UNKNOWNGROUP;
	}
	// implementation of IOleCommandTarget
	//////////////////////////////////////////////////////////////////////
#endif
#endif