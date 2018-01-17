//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListView, ShBrowserCtlsLibU::_IShListViewEvents
/// \else
///   \sa ShellListView, ShBrowserCtlsLibA::_IShListViewEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c ChangedColumnVisibility event</em>
	///
	/// \param[in] pColumn The column that has been inserted into or removed from the listview control.
	/// \param[in] hasBecomeVisible If \c VARIANT_TRUE, the column has been inserted; otherwise it has been
	///            removed from the listview control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangedColumnVisibility,
	///       ShellListView::Raise_ChangedColumnVisibility, Fire_ChangingColumnVisibility,
	///       Fire_UnloadedColumn
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangedColumnVisibility,
	///       ShellListView::Raise_ChangedColumnVisibility, Fire_ChangingColumnVisibility,
	///       Fire_UnloadedColumn
	/// \endif
	HRESULT Fire_ChangedColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL hasBecomeVisible)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0] = hasBecomeVisible;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGEDCOLUMNVISIBILITY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedContextMenuSelection event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu. Will be \c NULL if the menu is closed.
	/// \param[in] isShellContextMenu If \c VARIANT_TRUE, the menu is a shell context menu; otherwise it is a
	///            column header context menu.
	/// \param[in] commandID The ID of the new selected menu item. May be 0.
	/// \param[in] isCustomMenuItem If \c VARIANT_FALSE, the selected menu item has been inserted by
	///            \c ShellListView; otherwise it is a custom menu item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangedContextMenuSelection,
	///       ShellListView::Raise_ChangedContextMenuSelection, Fire_SelectedShellContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangedContextMenuSelection,
	///       ShellListView::Raise_ChangedContextMenuSelection, Fire_SelectedShellContextMenuItem
	/// \endif
	HRESULT Fire_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, VARIANT_BOOL isShellContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3].lVal = hContextMenu;		p[3].vt = VT_I4;
				p[2] = isShellContextMenu;
				p[1] = commandID;
				p[0] = isCustomMenuItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGEDCONTEXTMENUSELECTION, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedItemPIDL event</em>
	///
	/// \param[in] pListItem The item that was updated.
	/// \param[in] previousPIDL The item's previous fully qualified pIDL.
	/// \param[in] newPIDL The item's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangedItemPIDL, ShellListView::Raise_ChangedItemPIDL,
	///       Fire_ChangedNamespacePIDL
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangedItemPIDL, ShellListView::Raise_ChangedItemPIDL,
	///       Fire_ChangedNamespacePIDL
	/// \endif
	HRESULT Fire_ChangedItemPIDL(IShListViewItem* pListItem, LONG previousPIDL, LONG newPIDL)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pListItem;
				p[1] = previousPIDL;
				p[0] = newPIDL;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGEDITEMPIDL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedNamespacePIDL event</em>
	///
	/// \param[in] pNamespace The namespace that was updated.
	/// \param[in] previousPIDL The namespace's previous fully qualified pIDL.
	/// \param[in] newPIDL The namespace's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangedNamespacePIDL,
	///       ShellListView::Raise_ChangedNamespacePIDL, Fire_ChangedItemPIDL
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangedNamespacePIDL,
	///       ShellListView::Raise_ChangedNamespacePIDL, Fire_ChangedItemPIDL
	/// \endif
	HRESULT Fire_ChangedNamespacePIDL(IShListViewNamespace* pNamespace, LONG previousPIDL, LONG newPIDL)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pNamespace;
				p[1] = previousPIDL;
				p[0] = newPIDL;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGEDNAMESPACEPIDL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedSortColumn event</em>
	///
	/// \param[in] previousSortColumnIndex The zero-based index of the old sort column. This index is the
	///            column's index given by the shell.
	/// \param[in] newSortColumnIndex The zero-based index of the new sort column. This index is the
	///            column's index given by the shell.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangedSortColumn,
	///       ShellListView::Raise_ChangedSortColumn, Fire_ChangingSortColumn
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangedSortColumn,
	///       ShellListView::Raise_ChangedSortColumn, Fire_ChangingSortColumn
	/// \endif
	HRESULT Fire_ChangedSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].lVal = previousSortColumnIndex;
				p[0].lVal = newSortColumnIndex;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGEDSORTCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangingColumnVisibility event</em>
	///
	/// \param[in] pColumn The column that will be inserted into or removed from the listview control.
	/// \param[in] willBecomeVisible If \c VARIANT_TRUE, the column will be inserted; otherwise it will be
	///            removed from the listview control.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion or removal; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangingColumnVisibility,
	///       ShellListView::Raise_ChangingColumnVisibility, Fire_LoadedColumn, Fire_ChangedColumnVisibility,
	///       Fire_UnloadedColumn
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangingColumnVisibility,
	///       ShellListView::Raise_ChangingColumnVisibility, Fire_LoadedColumn, Fire_ChangedColumnVisibility,
	///       Fire_UnloadedColumn
	/// \endif
	HRESULT Fire_ChangingColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL willBecomeVisible, VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pColumn;
				p[1] = willBecomeVisible;
				p[0].pboolVal = pCancel;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGINGCOLUMNVISIBILITY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangingSortColumn event</em>
	///
	/// \param[in] previousSortColumnIndex The zero-based index of the old sort column. This index is the
	///            column's index given by the shell.
	/// \param[in] newSortColumnIndex The zero-based index of the new sort column. This index is the
	///            column's index given by the shell.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort the change; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ChangingSortColumn,
	///       ShellListView::Raise_ChangingSortColumn, Fire_ChangedSortColumn
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ChangingSortColumn,
	///       ShellListView::Raise_ChangingSortColumn, Fire_ChangedSortColumn
	/// \endif
	HRESULT Fire_ChangingSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = previousSortColumnIndex;
				p[1].lVal = newSortColumnIndex;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CHANGINGSORTCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose columns have been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationCompleted,
	///       ShellListView::Raise_ColumnEnumerationCompleted, Fire_ColumnEnumerationStarted,
	///       Fire_ColumnEnumerationTimedOut, Fire_ItemEnumerationCompleted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationCompleted,
	///       ShellListView::Raise_ColumnEnumerationCompleted, Fire_ColumnEnumerationStarted,
	///       Fire_ColumnEnumerationTimedOut, Fire_ItemEnumerationCompleted
	/// \endif
	HRESULT Fire_ColumnEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pNamespace;
				p[0] = aborted;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_COLUMNENUMERATIONCOMPLETED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose columns are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationStarted,
	///       ShellListView::Raise_ColumnEnumerationStarted, Fire_ColumnEnumerationTimedOut,
	///       Fire_ColumnEnumerationCompleted, Fire_ItemEnumerationStarted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationStarted,
	///       ShellListView::Raise_ColumnEnumerationStarted, Fire_ColumnEnumerationTimedOut,
	///       Fire_ColumnEnumerationCompleted, Fire_ItemEnumerationStarted
	/// \endif
	HRESULT Fire_ColumnEnumerationStarted(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_COLUMNENUMERATIONSTARTED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose columns are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationTimedOut,
	///       ShellListView::Raise_ColumnEnumerationTimedOut, Fire_ColumnEnumerationStarted,
	///       Fire_ColumnEnumerationCompleted, Fire_ItemEnumerationTimedOut
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationTimedOut,
	///       ShellListView::Raise_ColumnEnumerationTimedOut, Fire_ColumnEnumerationStarted,
	///       Fire_ColumnEnumerationCompleted, Fire_ItemEnumerationTimedOut
	/// \endif
	HRESULT Fire_ColumnEnumerationTimedOut(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_COLUMNENUMERATIONTIMEDOUT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedHeaderContextMenu event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] minimumCustomCommandID The lowest command ID that may be used for custom menu items.
	///            For custom menu items any command ID larger or equal this value and smaller or equal
	///            65535 may be used.
	/// \param[in,out] pCancelMenu If \c VARIANT_FALSE, the caller should not display the context menu;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::CreatedHeaderContextMenu,
	///       ShellListView::Raise_CreatedHeaderContextMenu, Fire_DestroyingHeaderContextMenu,
	///       Fire_SelectedHeaderContextMenuItem, Fire_CreatedShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::CreatedHeaderContextMenu,
	///       ShellListView::Raise_CreatedHeaderContextMenu, Fire_DestroyingHeaderContextMenu,
	///       Fire_SelectedHeaderContextMenuItem, Fire_CreatedShellContextMenu
	/// \endif
	HRESULT Fire_CreatedHeaderContextMenu(OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = hContextMenu;					p[2].vt = VT_I4;
				p[1] = minimumCustomCommandID;
				p[0].pboolVal = pCancelMenu;			p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CREATEDHEADERCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] minimumCustomCommandID The lowest command ID that may be used for custom menu items.
	///            For custom menu items any command ID larger or equal this value and smaller or equal
	///            65535 may be used.
	/// \param[in,out] pCancelMenu If \c VARIANT_FALSE, the caller should not display the context menu;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::CreatedShellContextMenu,
	///       ShellListView::Raise_CreatedShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingShellContextMenu, Fire_SelectedShellContextMenuItem,
	///       Fire_CreatedHeaderContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::CreatedShellContextMenu,
	///       ShellListView::Raise_CreatedShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingShellContextMenu, Fire_SelectedShellContextMenuItem,
	///       Fire_CreatedHeaderContextMenu
	/// \endif
	HRESULT Fire_CreatedShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pItems;
				p[2].lVal = hContextMenu;					p[2].vt = VT_I4;
				p[1] = minimumCustomCommandID;
				p[0].pboolVal = pCancelMenu;			p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CREATEDSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in,out] pContextMenuStyle A bit field influencing the content of the context menu. Any
	///                combination of the values defined by the \c ShellContextMenuStyleConstants
	///                enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_FALSE, the caller should not create and display the context
	///                menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::CreatingShellContextMenu,
	///       ShellListView::Raise_CreatingShellContextMenu, Fire_CreatedShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::CreatingShellContextMenu,
	///       ShellListView::Raise_CreatingShellContextMenu, Fire_CreatedShellContextMenu
	/// \endif
	HRESULT Fire_CreatingShellContextMenu(LPDISPATCH pItems, ShellContextMenuStyleConstants* pContextMenuStyle, VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pItems;
				p[1].plVal = reinterpret_cast<PLONG>(pContextMenuStyle);		p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pCancel;																		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_CREATINGSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyingHeaderContextMenu event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::DestroyingHeaderContextMenu,
	///       ShellListView::Raise_DestroyingHeaderContextMenu, Fire_CreatedHeaderContextMenu,
	///       Fire_DestroyingShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::DestroyingHeaderContextMenu,
	///       ShellListView::Raise_DestroyingHeaderContextMenu, Fire_CreatedHeaderContextMenu,
	///       Fire_DestroyingShellContextMenu
	/// \endif
	HRESULT Fire_DestroyingHeaderContextMenu(OLE_HANDLE hContextMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].lVal = hContextMenu;		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_DESTROYINGHEADERCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::DestroyingShellContextMenu,
	///       ShellListView::Raise_DestroyingShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingHeaderContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::DestroyingShellContextMenu,
	///       ShellListView::Raise_DestroyingShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingHeaderContextMenu
	/// \endif
	HRESULT Fire_DestroyingShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pItems;
				p[0].lVal = hContextMenu;		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_DESTROYINGSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pListItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::InsertedItem, ShellListView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_InsertedNamespace, Fire_LoadedColumn, Fire_RemovingItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::InsertedItem, ShellListView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_InsertedNamespace, Fire_LoadedColumn, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertedItem(IShListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedNamespace event</em>
	///
	/// \param[in] pNamespace The inserted namespace.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::InsertedNamespace,
	///       ShellListView::Raise_InsertedNamespace, Fire_InsertedItem, Fire_RemovingNamespace
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::InsertedNamespace,
	///       ShellListView::Raise_InsertedNamespace, Fire_InsertedItem, Fire_RemovingNamespace
	/// \endif
	HRESULT Fire_InsertedNamespace(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_INSERTEDNAMESPACE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] fullyQualifiedPIDL The fully qualified pIDL of the item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::InsertingItem, ShellListView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_ChangingColumnVisibility, Fire_RemovingItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::InsertingItem, ShellListView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_ChangingColumnVisibility, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = fullyQualifiedPIDL;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InvokedHeaderContextMenuCommand event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::InvokedHeaderContextMenuCommand,
	///       ShellListView::Raise_InvokedHeaderContextMenuCommand, Fire_SelectedShellContextMenuItem,
	///       Fire_InvokedHeaderContextMenuCommand
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::InvokedHeaderContextMenuCommand,
	///       ShellListView::Raise_InvokedHeaderContextMenuCommand, Fire_SelectedShellContextMenuItem,
	///       Fire_InvokedHeaderContextMenuCommand
	/// \endif
	HRESULT Fire_InvokedHeaderContextMenuCommand(OLE_HANDLE hContextMenu, LONG commandID)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].lVal = hContextMenu;		p[1].vt = VT_I4;
				p[0] = commandID;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_INVOKEDHEADERCONTEXTMENUCOMMAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InvokedShellContextMenuCommand event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the command.
	/// \param[in] usedWindowMode A value specifying how to display the window that may be opened when
	///            executing the command. Any of the values defined by the \c WindowModeConstants enumeration
	///            is valid.
	/// \param[in] usedInvocationFlags A bit field controlling command execution. Any combination of the
	///            values defined by the \c CommandInvocationFlagsConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::InvokedShellContextMenuCommand,
	///       ShellListView::Raise_InvokedShellContextMenuCommand, Fire_SelectedShellContextMenuItem,
	///       Fire_InvokedHeaderContextMenuCommand
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::InvokedShellContextMenuCommand,
	///       ShellListView::Raise_InvokedShellContextMenuCommand, Fire_SelectedShellContextMenuItem,
	///       Fire_InvokedHeaderContextMenuCommand
	/// \endif
	HRESULT Fire_InvokedShellContextMenuCommand(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants usedWindowMode, CommandInvocationFlagsConstants usedInvocationFlags)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pItems;
				p[3].lVal = hContextMenu;															p[3].vt = VT_I4;
				p[2] = commandID;
				p[1].lVal = static_cast<LONG>(usedWindowMode);				p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(usedInvocationFlags);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_INVOKEDSHELLCONTEXTMENUCOMMAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose items have been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationCompleted,
	///       ShellListView::Raise_ItemEnumerationCompleted, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationTimedOut, Fire_ColumnEnumerationCompleted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationCompleted,
	///       ShellListView::Raise_ItemEnumerationCompleted, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationTimedOut, Fire_ColumnEnumerationCompleted
	/// \endif
	HRESULT Fire_ItemEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pNamespace;
				p[0] = aborted;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_ITEMENUMERATIONCOMPLETED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationStarted,
	///       ShellListView::Raise_ItemEnumerationStarted, Fire_ItemEnumerationTimedOut,
	///       Fire_ItemEnumerationCompleted, Fire_ColumnEnumerationStarted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationStarted,
	///       ShellListView::Raise_ItemEnumerationStarted, Fire_ItemEnumerationTimedOut,
	///       Fire_ItemEnumerationCompleted, Fire_ColumnEnumerationStarted
	/// \endif
	HRESULT Fire_ItemEnumerationStarted(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_ITEMENUMERATIONSTARTED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationTimedOut,
	///       ShellListView::Raise_ItemEnumerationTimedOut, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationCompleted, Fire_ColumnEnumerationTimedOut
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationTimedOut,
	///       ShellListView::Raise_ItemEnumerationTimedOut, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationCompleted, Fire_ColumnEnumerationTimedOut
	/// \endif
	HRESULT Fire_ItemEnumerationTimedOut(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_ITEMENUMERATIONTIMEDOUT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c LoadedColumn event</em>
	///
	/// \param[in] pColumn The column that has been loaded.
	/// \param[in,out] pMakeVisible If \c VARIANT_TRUE, the caller should insert the column into the listview
	///                control; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::LoadedColumn, ShellListView::Raise_LoadedColumn,
	///       Fire_ChangingColumnVisibility, Fire_InsertedItem, Fire_InsertedNamespace, Fire_UnloadedColumn
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::LoadedColumn, ShellListView::Raise_LoadedColumn,
	///       Fire_ChangingColumnVisibility, Fire_InsertedItem, Fire_InsertedNamespace, Fire_UnloadedColumn
	/// \endif
	HRESULT Fire_LoadedColumn(IShListViewColumn* pColumn, VARIANT_BOOL* pMakeVisible)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0].pboolVal = pMakeVisible;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_LOADEDCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pListItem The item being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::RemovingItem, ShellListView::Raise_RemovingItem,
	///       Fire_RemovingNamespace, Fire_UnloadedColumn, Fire_InsertedItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::RemovingItem, ShellListView::Raise_RemovingItem,
	///       Fire_RemovingNamespace, Fire_UnloadedColumn, Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovingItem(IShListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingNamespace event</em>
	///
	/// \param[in] pNamespace The namespace being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::RemovingNamespace,
	///       ShellListView::Raise_RemovingNamespace, Fire_RemovingItem, Fire_UnloadedColumn,
	///       Fire_InsertedNamespace
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::RemovingNamespace,
	///       ShellListView::Raise_RemovingNamespace, Fire_RemovingItem, Fire_UnloadedColumn,
	///       Fire_InsertedNamespace
	/// \endif
	HRESULT Fire_RemovingNamespace(IShListViewNamespace* pNamespace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pNamespace;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_REMOVINGNAMESPACE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedHeaderContextMenuItem event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	/// \param[in,out] pExecuteCommand If set to \c VARIANT_TRUE, the caller should execute the menu item;
	///                otherwise not.\n
	///                If the menu item is a custom one, this value should be ignored and always be treated
	///                as being set to \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::SelectedHeaderContextMenuItem,
	///       ShellListView::Raise_SelectedHeaderContextMenuItem, Fire_InvokedHeaderContextMenuCommand
	///       Fire_ChangedContextMenuSelection, Fire_CreatedHeaderContextMenu,
	///       Fire_SelectedShellContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::SelectedHeaderContextMenuItem,
	///       ShellListView::Raise_SelectedHeaderContextMenuItem, Fire_InvokedHeaderContextMenuCommand,
	///       Fire_ChangedContextMenuSelection, Fire_CreatedHeaderContextMenu,
	///       Fire_SelectedShellContextMenuItem
	/// \endif
	HRESULT Fire_SelectedHeaderContextMenuItem(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL* pExecuteCommand)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = hContextMenu;						p[2].vt = VT_I4;
				p[1] = commandID;
				p[0].pboolVal = pExecuteCommand;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_SELECTEDHEADERCONTEXTMENUITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedShellContextMenuItem event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	/// \param[in,out] pWindowModeToUse A value specifying how to display the window that may be opened when
	///                executing the command. Any of the values defined by the \c WindowModeConstants
	///                enumeration is valid.
	/// \param[in,out] pInvocationFlagsToUse A bit field controlling command execution. Any combination of
	///                the values defined by the \c CommandInvocationFlagsConstants enumeration is valid.
	/// \param[in,out] pExecuteCommand If set to \c VARIANT_TRUE, the caller should execute the menu item;
	///                otherwise not.\n
	///                If the menu item is a custom one, this value should be ignored and always be treated
	///                as being set to \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::SelectedShellContextMenuItem,
	///       ShellListView::Raise_SelectedShellContextMenuItem, Fire_InvokedShellContextMenuCommand,
	///       Fire_ChangedContextMenuSelection, Fire_CreatedShellContextMenu,
	///       Fire_SelectedHeaderContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::SelectedShellContextMenuItem,
	///       ShellListView::Raise_SelectedShellContextMenuItem, Fire_InvokedShellContextMenuCommand,
	///       Fire_ChangedContextMenuSelection, Fire_CreatedShellContextMenu,
	///       Fire_SelectedHeaderContextMenuItem
	/// \endif
	HRESULT Fire_SelectedShellContextMenuItem(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants* pWindowModeToUse, CommandInvocationFlagsConstants* pInvocationFlagsToUse, VARIANT_BOOL* pExecuteCommand)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pItems;
				p[4].lVal = hContextMenu;																				p[4].vt = VT_I4;
				p[3] = commandID;
				p[2].plVal = reinterpret_cast<PLONG>(pWindowModeToUse);					p[2].vt = VT_I4 | VT_BYREF;
				p[1].plVal = reinterpret_cast<PLONG>(pInvocationFlagsToUse);		p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pExecuteCommand;																p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_SELECTEDSHELLCONTEXTMENUITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c UnloadedColumn event</em>
	///
	/// \param[in] pColumn The column that has been unloaded. If \c Nothing, all columns have been unloaded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShListViewEvents::UnloadedColumn, ShellListView::Raise_UnloadedColumn,
	///       Fire_LoadedColumn, Fire_RemovingItem, Fire_RemovingNamespace
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShListViewEvents::UnloadedColumn, ShellListView::Raise_UnloadedColumn,
	///       Fire_LoadedColumn, Fire_RemovingItem, Fire_RemovingNamespace
	/// \endif
	HRESULT Fire_UnloadedColumn(IShListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHLVWE_UNLOADEDCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IShListViewEvents