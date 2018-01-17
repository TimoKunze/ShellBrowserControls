//////////////////////////////////////////////////////////////////////
/// \file definitions.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Common definitions for \c ShellTreeView</em>
//////////////////////////////////////////////////////////////////////


#pragma once

#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif


/// \brief <em>Holds details like the pIDL about a shell item</em>
typedef struct SHELLTREEVIEWITEMDATA
{
	/// \brief <em>The item's fully qualified pIDL</em>
	PCIDLIST_ABSOLUTE pIDL;
	/// \brief <em>The fully qualified pIDL of the namespace this item belongs to</em>
	PCIDLIST_ABSOLUTE pIDLNamespace;
	/// \brief <em>A bit field specifying which of the item's properties are managed by \c ShellTreeView</em>
	///
	/// \if UNICODE
	///   \sa ShellTreeViewItem::get_ManagedProperties,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa ShellTreeViewItem::get_ManagedProperties,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	ShTvwManagedItemPropertiesConstants managedProperties;
	/// \brief <em>If \c TRUE, we've entered \c OnFirstTimeItemExpanding to load this item's sub-items</em>
	///
	/// \sa ShellTreeView::OnFirstTimeItemExpanding
	volatile LONG waitingForLoadingSubItems;
	/// \brief <em>If \c TRUE, we're currently loading this item's sub-items</em>
	UINT isLoadingSubItems : 1;

	SHELLTREEVIEWITEMDATA()
	{
		pIDL = NULL;
		pIDLNamespace = NULL;
		waitingForLoadingSubItems = FALSE;
		isLoadingSubItems = FALSE;
	}

	~SHELLTREEVIEWITEMDATA()
	{
		ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDL));
		// don't free pIDLNamespace
	}
} SHELLTREEVIEWITEMDATA, *LPSHELLTREEVIEWITEMDATA;

/// \brief <em>Holds details reported to \c ShellTreeView by background item enumerator threads together with the enumerated items</em>
///
/// \sa ShTvwBackgroundItemEnumTask, ShTvwInsertSingleItemTask
typedef struct SHTVWBACKGROUNDITEMENUMINFO
{
	/// \brief <em>The unique ID of the background item enumerator task sending the message</em>
	ULONGLONG taskID;
	/// \brief <em>The treeview item that shall become the parent item of the enumerated items</em>
	HTREEITEM hParentItem;
	/// \brief <em>The treeview item after which the enumerated items shall be inserted</em>
	HTREEITEM hInsertAfter;
	/// \brief <em>If \c TRUE, \c ShellTreeView should ensure that the enumerated items do not yet exist within the treeview</em>
	UINT checkForDuplicates : 1;
	/// \brief <em>Specifies the fully qualified pIDL of the item for which label-edit mode should be entered after insertion</em>
	PIDLIST_ABSOLUTE pIDLToLabelEdit;
	/// \brief <em>Specifies the fully qualified pIDL of the namespace to which the enumerated items shall be associated</em>
	PCIDLIST_ABSOLUTE namespacePIDLToSet;
	/// \brief <em>Holds the enumerated items' pIDLs</em>
	HDPA hPIDLBuffer;
} SHTVWBACKGROUNDITEMENUMINFO, *LPSHTVWBACKGROUNDITEMENUMINFO;