#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	/// \class DummyShellBrowser
	/// \author Timo "TimoSoft" Kunze
	/// \brief <em>Hooks desktop's \c IShellFolder implementation to make it support everything we need for multi-namespace context menus</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
	//////////////////////////////////////////////////////////////////////
#endif


#pragma once

#include <shobjidl.h>
#include "../helpers.h"
#include "common.h"


#ifdef ACTIVATE_SECZONE_FEATURES
	// {399EE224-4DC7-412d-9676-2F687A219EF4}
	DEFINE_GUID(CLSID_DummyShellBrowser, 0x399ee224, 0x4dc7, 0x412d, 0x96, 0x76, 0x2f, 0x68, 0x7a, 0x21, 0x9e, 0xf4);


	class ATL_NO_VTABLE DummyShellBrowser : 
	    public CComObjectRootEx<CComMultiThreadModel>,
	    public CComCoClass<DummyShellBrowser, &CLSID_DummyShellBrowser>,
	    public IShellBrowser,
	    public IServiceProvider,
	    public IOleCommandTarget
	{
	public:
		/// \brief <em>The constructor of this class</em>
		///
		/// Used for initialization.
		DummyShellBrowser();

		#ifndef DOXYGEN_SHOULD_SKIP_THIS
			BEGIN_COM_MAP(DummyShellBrowser)
				COM_INTERFACE_ENTRY(IOleWindow)
				COM_INTERFACE_ENTRY(IShellBrowser)
				COM_INTERFACE_ENTRY(IServiceProvider)
				COM_INTERFACE_ENTRY(IOleCommandTarget)
			END_COM_MAP()
		#endif

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IOleWindow
		///
		//@{
		/// \brief <em>Wraps \c IOleWindow::GetWindow</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP GetWindow(HWND *phwnd);
		/// \brief <em>Wraps \c IOleWindow::ContextSensitiveHelp</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
		//@}
		//////////////////////////////////////////////////////////////////////

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IShellBrowser
		///
		//@{
		/// \brief <em>Wraps \c IShellBrowser::InsertMenusSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
		/// \brief <em>Wraps \c IShellBrowser::SetMenuSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775066.aspx">IShellFolder::EnumObjects</a>
		virtual STDMETHODIMP SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject);
		/// \brief <em>Wraps \c IShellBrowser::RemoveMenusSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775059.aspx">IShellFolder::BindToObject</a>
		virtual STDMETHODIMP RemoveMenusSB(HMENU hmenuShared);
		/// \brief <em>Wraps \c IShellBrowser::SetStatusTextSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775061.aspx">IShellFolder::BindToStorage</a>
		virtual STDMETHODIMP SetStatusTextSB(LPCWSTR pszStatusText);
		/// \brief <em>Wraps \c IShellBrowser::EnableModelessSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775062.aspx">IShellFolder::CompareIDs</a>
		virtual STDMETHODIMP EnableModelessSB(BOOL fEnable);
		/// \brief <em>Wraps \c IShellBrowser::TranslateAcceleratorSB</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775064.aspx">IShellFolder::CreateViewObject</a>
		virtual STDMETHODIMP TranslateAcceleratorSB(MSG *pmsg, WORD wID);
		/// \brief <em>Wraps \c IShellBrowser::BrowseObject</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775068.aspx">IShellFolder::GetAttributesOf</a>
		virtual STDMETHODIMP BrowseObject(PCUIDLIST_RELATIVE pidl, UINT wFlags);
		/// \brief <em>Wraps \c IShellBrowser::GetViewStateStream</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775073.aspx">IShellFolder::GetUIObjectOf</a>
		virtual STDMETHODIMP GetViewStateStream(DWORD grfMode, IStream **ppStrm);
		/// \brief <em>Wraps \c IShellBrowser::GetControlWindow</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::GetDisplayNameOf</a>
		virtual STDMETHODIMP GetControlWindow(UINT id, HWND *phwnd);
		/// \brief <em>Wraps \c IShellBrowser::SendControlMsg</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::SetNameOf</a>
		virtual STDMETHODIMP SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret);
		/// \brief <em>Wraps \c IShellBrowser::QueryActiveShellView</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::SetNameOf</a>
		virtual STDMETHODIMP QueryActiveShellView(IShellView **ppshv);
		/// \brief <em>Wraps \c IShellBrowser::OnViewWindowActive</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::SetNameOf</a>
		virtual STDMETHODIMP OnViewWindowActive(IShellView *pshv);
		/// \brief <em>Wraps \c IShellBrowser::SetToolbarItems</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775071.aspx">IShellFolder::SetNameOf</a>
		virtual STDMETHODIMP SetToolbarItems(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags);
		//@}
		//////////////////////////////////////////////////////////////////////

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IServiceProvider
		///
		//@{
		/// \brief <em>Wraps \c IServiceProvider::QueryService</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject);
		//@}
		//////////////////////////////////////////////////////////////////////

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IOleCommandTarget
		///
		//@{
		/// \brief <em>Wraps \c IOleCommandTarget::QueryStatus</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
		/// \brief <em>Wraps \c IOleCommandTarget::Exec</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775090.aspx">IShellFolder::ParseDisplayName</a>
		virtual STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut);
		//@}
		//////////////////////////////////////////////////////////////////////

		HWND hWnd;
		LONG zone;

	protected:
	};
#endif