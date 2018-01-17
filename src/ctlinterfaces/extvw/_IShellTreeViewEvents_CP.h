//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShTreeViewEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShTreeViewEvents interface</em>
///
/// \if UNICODE
///   \sa ShellTreeView, ShBrowserCtlsLibU::_IShTreeViewEvents
/// \else
///   \sa ShellTreeView, ShBrowserCtlsLibA::_IShTreeViewEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShTreeViewEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShTreeViewEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c ChangedContextMenuSelection event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu. Will be \c NULL if the menu is closed.
	/// \param[in] commandID The ID of the new selected menu item. May be 0.
	/// \param[in] isCustomMenuItem If \c VARIANT_FALSE, the selected menu item has been inserted by
	///            \c ShellListView; otherwise it is a custom menu item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedContextMenuSelection,
	///       ShellTreeView::Raise_ChangedContextMenuSelection, Fire_SelectedShellContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedContextMenuSelection,
	///       ShellTreeView::Raise_ChangedContextMenuSelection, Fire_SelectedShellContextMenuItem
	/// \endif
	HRESULT Fire_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = hContextMenu;		p[2].vt = VT_I4;
				p[1] = commandID;
				p[0] = isCustomMenuItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHTVWE_CHANGEDCONTEXTMENUSELECTION, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedItemPIDL event</em>
	///
	/// \param[in] pTreeItem The item that was updated.
	/// \param[in] previousPIDL The item's previous fully qualified pIDL.
	/// \param[in] newPIDL The item's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedItemPIDL, ShellTreeView::Raise_ChangedItemPIDL,
	///       Fire_ChangedNamespacePIDL
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedItemPIDL, ShellTreeView::Raise_ChangedItemPIDL,
	///       Fire_ChangedNamespacePIDL
	/// \endif
	HRESULT Fire_ChangedItemPIDL(IShTreeViewItem* pTreeItem, LONG previousPIDL, LONG newPIDL)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pTreeItem;
				p[1] = previousPIDL;
				p[0] = newPIDL;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHTVWE_CHANGEDITEMPIDL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedNamespacePIDL,
	///       ShellTreeView::Raise_ChangedNamespacePIDL, Fire_ChangedItemPIDL
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedNamespacePIDL,
	///       ShellTreeView::Raise_ChangedNamespacePIDL, Fire_ChangedItemPIDL
	/// \endif
	HRESULT Fire_ChangedNamespacePIDL(IShTreeViewNamespace* pNamespace, LONG previousPIDL, LONG newPIDL)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_CHANGEDNAMESPACEPIDL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::CreatedShellContextMenu,
	///       ShellTreeView::Raise_CreatedShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingShellContextMenu, Fire_SelectedShellContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::CreatedShellContextMenu,
	///       ShellTreeView::Raise_CreatedShellContextMenu, Fire_CreatingShellContextMenu,
	///       Fire_DestroyingShellContextMenu, Fire_SelectedShellContextMenuItem
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
				hr = pConnection->Invoke(DISPID_SHTVWE_CREATEDSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
	/// \param[in,out] pContextMenuStyle A bit field influencing the content of the context menu. Any
	///                combination of the values defined by the \c ShellContextMenuStyleConstants
	///                enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_FALSE, the caller should not create and display the context
	///                menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::CreatingShellContextMenu,
	///       ShellTreeView::Raise_CreatingShellContextMenu, Fire_CreatedShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::CreatingShellContextMenu,
	///       ShellTreeView::Raise_CreatingShellContextMenu, Fire_CreatedShellContextMenu
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
				hr = pConnection->Invoke(DISPID_SHTVWE_CREATINGSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::DestroyingShellContextMenu,
	///       ShellTreeView::Raise_DestroyingShellContextMenu, Fire_CreatingShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::DestroyingShellContextMenu,
	///       ShellTreeView::Raise_DestroyingShellContextMenu, Fire_CreatingShellContextMenu
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
				hr = pConnection->Invoke(DISPID_SHTVWE_DESTROYINGSHELLCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pTreeItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::InsertedItem, ShellTreeView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_InsertedNamespace, Fire_RemovingItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::InsertedItem, ShellTreeView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_InsertedNamespace, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertedItem(IShTreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHTVWE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::InsertedNamespace,
	///       ShellTreeView::Raise_InsertedNamespace, Fire_InsertedItem, Fire_RemovingNamespace
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::InsertedNamespace,
	///       ShellTreeView::Raise_InsertedNamespace, Fire_InsertedItem, Fire_RemovingNamespace
	/// \endif
	HRESULT Fire_InsertedNamespace(IShTreeViewNamespace* pNamespace)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_INSERTEDNAMESPACE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] parentItemHandle The handle of the item that will be the immediate parent of the new
	///            item. If \c NULL, the item will become a root item.
	/// \param[in] fullyQualifiedPIDL The fully qualified pIDL of the item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::InsertingItem, ShellTreeView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::InsertingItem, ShellTreeView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(OLE_HANDLE parentItemHandle, LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = parentItemHandle;				p[2].vt = VT_I4;
				p[1] = fullyQualifiedPIDL;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_SHTVWE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InvokedShellContextMenuCommand event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::InvokedShellContextMenuCommand,
	///       ShellTreeView::Raise_InvokedShellContextMenuCommand, Fire_SelectedShellContextMenuItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::InvokedShellContextMenuCommand,
	///       ShellTreeView::Raise_InvokedShellContextMenuCommand, Fire_SelectedShellContextMenuItem
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
				hr = pConnection->Invoke(DISPID_SHTVWE_INVOKEDSHELLCONTEXTMENUCOMMAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose items have been enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items have
	///            been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationCompleted,
	///       ShellTreeView::Raise_ItemEnumerationCompleted, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationTimedOut
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationCompleted,
	///       ShellTreeView::Raise_ItemEnumerationCompleted, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationTimedOut
	/// \endif
	HRESULT Fire_ItemEnumerationCompleted(LPDISPATCH pNamespace, VARIANT_BOOL aborted)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_ITEMENUMERATIONCOMPLETED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items are
	///            being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationStarted,
	///       ShellTreeView::Raise_ItemEnumerationStarted, Fire_ItemEnumerationTimedOut,
	///       Fire_ItemEnumerationCompleted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationStarted,
	///       ShellTreeView::Raise_ItemEnumerationStarted, Fire_ItemEnumerationTimedOut,
	///       Fire_ItemEnumerationCompleted
	/// \endif
	HRESULT Fire_ItemEnumerationStarted(LPDISPATCH pNamespace)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_ITEMENUMERATIONSTARTED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items are
	///            being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationTimedOut,
	///       ShellTreeView::Raise_ItemEnumerationTimedOut, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationCompleted
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationTimedOut,
	///       ShellTreeView::Raise_ItemEnumerationTimedOut, Fire_ItemEnumerationStarted,
	///       Fire_ItemEnumerationCompleted
	/// \endif
	HRESULT Fire_ItemEnumerationTimedOut(LPDISPATCH pNamespace)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_ITEMENUMERATIONTIMEDOUT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pTreeItem The item being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::RemovingItem, ShellTreeView::Raise_RemovingItem,
	///       Fire_RemovingNamespace, Fire_InsertedItem
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::RemovingItem, ShellTreeView::Raise_RemovingItem,
	///       Fire_RemovingNamespace, Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovingItem(IShTreeViewItem* pTreeItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pTreeItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_SHTVWE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::RemovingNamespace,
	///       ShellTreeView::Raise_RemovingNamespace, Fire_RemovingItem, Fire_InsertedNamespace
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::RemovingNamespace,
	///       ShellTreeView::Raise_RemovingNamespace, Fire_RemovingItem, Fire_InsertedNamespace
	/// \endif
	HRESULT Fire_RemovingNamespace(IShTreeViewNamespace* pNamespace)
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
				hr = pConnection->Invoke(DISPID_SHTVWE_REMOVINGNAMESPACE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedShellContextMenuItem event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	///   \sa ShBrowserCtlsLibU::_IShTreeViewEvents::SelectedShellContextMenuItem,
	///       ShellTreeView::Raise_SelectedShellContextMenuItem, Fire_InvokedShellContextMenuCommand,
	///       Fire_ChangedContextMenuSelection, Fire_CreatedShellContextMenu
	/// \else
	///   \sa ShBrowserCtlsLibA::_IShTreeViewEvents::SelectedShellContextMenuItem,
	///       ShellTreeView::Raise_SelectedShellContextMenuItem, Fire_InvokedShellContextMenuCommand,
	///       Fire_ChangedContextMenuSelection, Fire_CreatedShellContextMenu
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
				hr = pConnection->Invoke(DISPID_SHTVWE_SELECTEDSHELLCONTEXTMENUITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IShTreeViewEvents