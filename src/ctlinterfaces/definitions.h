//////////////////////////////////////////////////////////////////////
/// \file definitions.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Common constants for the controlable controls and \c ShellBrowserControls</em>
///
/// This file defines common constants for \c ShellBrowserControls and the controls it can be attached to.
//////////////////////////////////////////////////////////////////////


#pragma once

#include <windows.h>
#include "IMessageListener.h"
#include "IInternalMessageListener.h"


#define SHBM_COMMON_FIRST WM_APP
#define SHBM_COMMON_LAST (SHBM_COMMON_FIRST + 4)
#define SHBM_LISTVIEW_FIRST (SHBM_COMMON_LAST + 1)
#define SHBM_LISTVIEW_LAST (SHBM_LISTVIEW_FIRST + 4)
#define SHBM_TREEVIEW_FIRST (SHBM_LISTVIEW_LAST + 1)
#define SHBM_TREEVIEW_LAST (SHBM_TREEVIEW_FIRST + 4)

/// \brief <em>Holds data that is transfered with the \c SHBM_HANDSHAKE message</em>
///
/// \sa SHBM_HANDSHAKE
typedef struct SHBHANDSHAKE
{
	/// \brief <em>The struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>A \c BOOL value to determine whether the message was processed at all</em>
	///
	/// If the receiving window processes the message, it will set this member to \c TRUE. Windows that
	/// don't process the message (e. g. because they don't know it), don't touch this member. This way
	/// \c ShellBrowserControls is able to determine whether the window it tries to attach to supports the
	/// required interfaces.
	BOOL processed;
	/// \brief <em>Holds the error code in case the message failed</em>
	HRESULT errorCode;
	/// \brief <em>Holds the build number of the \c ShellBrowserControls control that sent the message</em>
	UINT shellBrowserBuildNumber;
	/// \brief <em>If \c TRUE, the sending \c ShellBrowserControls control is an Unicode build</em>
	BOOL shellBrowserSupportsUnicode;
	/// \brief <em>Holds the build number of the receiving control on return</em>
	PUINT pCtlBuildNumber;
	/// \brief <em>If \c TRUE on return, the receiving control is an Unicode build</em>
	LPBOOL pCtlSupportsUnicode;
	/// \brief <em>The \c IMessageListener implementation of the listener</em>
	///
	/// The implementation of the \c IMessageListener interface of the object that wants to listen to the
	/// receiving window's message streams. If set to \c NULL, the message hook is removed.
	///
	/// \sa IMessageListener
	IMessageListener* pMessageHook;
	/// \brief <em>The \c IInternalMessageListener implementation of the listener for internal messages</em>
	///
	/// The implementation of the \c IInternalMessageListener interface of the object that wants to listen
	/// to the receiving window's internal messages. If set to \c NULL, the message hook is removed.
	///
	/// \sa IInternalMessageListener
	IInternalMessageListener* pShBInternalMessageHook;
	/// \brief <em>Receives the \c IInternalMessageListener implementation of the receiving window</em>
	///
	/// Receives the implementation of the \c IInternalMessageListener interface of the receiving window.
	///
	/// \sa IInternalMessageListener
	IInternalMessageListener** ppCtlInternalMessageHook;
	/// \brief <em>Receives a pointer to the control's \c IUnknown implementation</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms680509.aspx">IUnknown</a>
	LPVOID* pCtlInterface;
} SHBHANDSHAKE, *LPSHBHANDSHAKE;

/// \brief <em>Holds data that is transfered with the \c EXLVM_ADDCOLUMN message</em>
///
/// \sa EXLVM_ADDCOLUMN
typedef struct EXLVMADDCOLUMNDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the new column's zero-based position</em>
	///
	/// Specifies the new column's zero-based position. If set to -1, the column will be inserted as the
	/// last column.
	LONG insertAt;
	/// \brief <em>Specifies the new item's caption text</em>
	///
	/// Specifies the new column's caption text. The maximum number of characters in this text is
	/// \c MAX_LVCOLUMNTEXTLENGTH.
	///
	/// \sa MAX_LVCOLUMNTEXTLENGTH
	LPTSTR pColumnText;
	/// \brief <em>Specifies the new column's format</em>
	///
	/// Specifies the new column's format. Any combination of the \c LVCFMT_* constants is valid.
	int columnFormat;
	/// \brief <em>Specifies the new column's width in pixels</em>
	///
	/// Specifies the new column's width in pixels. If set to a negative value, the column will be
	/// auto-sized after insertion.
	int columnWidth;
	/// \brief <em>Specifies the new column's zero-based index</em>
	int columnSubItemIndex;
	/// \brief <em>Will be set to the inserted column's unique ID</em>
	LONG insertedColumnID;
} EXLVMADDCOLUMNDATA, *LPEXLVMADDCOLUMNDATA;

/// \brief <em>Holds data that is transfered with the \c EXLVM_ADDITEM message</em>
///
/// \sa EXLVM_ADDITEM
typedef struct EXLVMADDITEMDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the new item's zero-based index</em>
	///
	/// Specifies the new item's zero-based index. If set to -1, the item will be inserted as the last
	/// item.
	LONG insertAt;
	/// \brief <em>Specifies the new item's caption text</em>
	///
	/// Specifies the new item's caption text. The maximum number of characters in this text is
	/// \c MAX_LVITEMTEXTLENGTH. If set to \c NULL, the list view control will fire its
	/// \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \sa MAX_LVITEMTEXTLENGTH
	LPTSTR pItemText;
	/// \brief <em>Specifies the new item's icon</em>
	///
	/// Specifies the zero-based index of the item's icon in the listview control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the listview control will fire its
	/// \c ItemGetDisplayInfo event each time this property's value is required. A value of -2 means 'not
	/// specified' and is valid if there's no imagelist associated with the listview control.
	LONG iconIndex;
	/// \brief <em>Specifies the new item's overlay icon</em>
	///
	/// Specifies the one-based index of the item's overlay icon in the listview control's \c ilSmall,
	/// \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to 0, the new item won't
	/// have an overlay.
	LONG overlayIndex;
	/// \brief <em>Specifies the new item's indentation in 'Details' view</em>
	///
	/// Specifies the new item's indentation in 'Details' view in image widths. If set to 1, the item's
	/// indentation will be 1 image width; if set to 2, it will be 2 image widths and so on. If set to -1,
	/// the listview control will fire its \c ItemGetDisplayInfo event each time this property's value is
	/// required.
	LONG itemIndentation;
	/// \brief <em>A \c LONG value that will be associated with the item</em>
	LONG itemData;
	/// \brief <em>Specifies whether the new item is drawn ghosted</em>
	VARIANT_BOOL isGhosted;
	/// \brief <em>Specifies the unique ID of the group that the new item will belong to</em>
	///
	/// Specifies the unique ID of the group that the new item will belong to. If set to \c -2, the item
	/// won't belong to any group.
	LONG groupID;
	/// \brief <em>Specifies the columns that will be used to display additional details in 'Tiles' view</em>
	///
	/// This array of column indexes specifies the columns that will be used to display additional details
	/// below the new item's text in 'Tiles' view. If set to an empty array, no details will be displayed.
	/// If set to \c Empty, the listview control will fire its \c ItemGetDisplayInfo event each time this
	/// property's value is required.
	VARIANT* pTileViewColumns;
	/// \brief <em>Will be set to the inserted item's unique ID</em>
	LONG insertedItemID;
} EXLVMADDITEMDATA, *LPEXLVMADDITEMDATA;

/// \brief <em>Holds data that is transfered with the \c EXLVM_CREATEITEMCONTAINER message</em>
///
/// \sa EXLVM_CREATEITEMCONTAINER
typedef struct EXLVMCREATEITEMCONTAINERDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Holds the number of elements in the array specified by \c pItems</em>
	UINT numberOfItems;
	/// \brief <em>An array containing the unique IDs of the items that the container shall contain</em>
	PLONG pItems;
	/// \brief <em>Receives the container object's \c IDispatch implementation</em>
	LPDISPATCH pContainer;
} EXLVMCREATEITEMCONTAINERDATA, *LPEXLVMCREATEITEMCONTAINERDATA;

/// \brief <em>Holds data that is transfered with the \c EXLVM_GETITEMIDSFROMVARIANT message</em>
///
/// \sa EXLVM_GETITEMIDSFROMVARIANT
typedef struct EXLVMGETITEMIDSDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>The \c VARIANT from which the item IDs shall be retrieved</em>
	VARIANT* pVariant;
	/// \brief <em>Receives the number of item IDs that have been retrieved from the \c VARIANT</em>
	UINT numberOfItems;
	/// \brief <em>Receives the item IDs that have been retrieved from the \c VARIANT</em>
	///
	/// \remarks The array is created by the callee and must be freed by the caller.
	PLONG pItemIDs;
} EXLVMGETITEMIDSDATA, *LPEXLVMGETITEMIDSDATA;

/// \brief <em>Holds data that is transfered with the \c EXLVM_HITTEST message</em>
///
/// \sa EXLVM_HITTEST
typedef struct EXLVMHITTESTDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the coordinates (in pixels relative to the control's upper-left corner) of the point to test</em>
	POINT pointToTest;
	/// \brief <em>Will be set to the hit item's unique ID</em>
	LONG hitItemID;
	/// \brief <em>Will be set to the hit sub-item's index</em>
	LONG hitSubItemIndex;
	/// \brief <em>Receives a bit field of LVHT_* flags, that holds further details about the control's part below the specified point</em>
	UINT hitTestDetails;
} EXLVMHITTESTDATA, *LPEXLVMHITTESTDATA;

/// \brief <em>Holds data that is transfered with the \c EXLVM_OLEDRAG message</em>
///
/// \sa EXLVM_OLEDRAG
typedef struct EXLVMOLEDRAGDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>If \c TRUE, \c SHDoDragDrop is called; otherwise \c IExplorerListView::OLEDrag is called</em>
	BOOL useSHDoDragDrop;
	/// \brief <em>If \c TRUE, the shell's default implementation of \c IDropSource is used; otherwise \c ExplorerListView's implementation is used</em>
	BOOL useShellDropSource;
	/// \brief <em>Specifies the data object to drag</em>
	LPDATAOBJECT pDataObject;
	/// \brief <em>Specifies the drop effects supported by the source</em>
	DWORD supportedEffects;
	/// \brief <em>Specifies the items being dragged</em>
	LPUNKNOWN pDraggedItems;
	/// \brief <em>Receives the effect being performed (as returned by \c DoDragDrop)</em>
	DWORD performedEffects;
} EXLVMOLEDRAGDATA, *LPEXLVMOLEDRAGDATA;

/// \brief <em>Holds data that is transfered with the \c SHLVM_HANDLEDRAGEVENT and \c EXLVM_FIREDRAGDROPEVENT messages</em>
///
/// \sa SHLVM_HANDLEDRAGEVENT, EXLVM_FIREDRAGDROPEVENT
typedef struct SHLVMDRAGDROPEVENTDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies whether the header control rather than the listview control is the drop target</em>
	BOOL headerIsTarget;
	/// \brief <em>Specifies the dragged data object</em>
	LPUNKNOWN pDataObject;
	/// \brief <em>Specifies the current drop effect</em>
	///
	/// Specifies the current drop effect. \c ShellListView may change this value to a more appropriate
	/// effect.
	DWORD effect;
	/// \brief <em>Specifies the current drop target</em>
	///
	/// Specifies the current drop target. \c ShellListView may change this value to a more appropriate
	/// item.
	///
	/// \remarks \c ShellListView may change \c pDropTarget rather than \c dropTarget, because not
	///          \c ShellListView will make the change, but the client app when \c ShellListView calls
	///          into \c ExplorerListView to raise the drag'n'drop event.
	LPUNKNOWN pDropTarget;
	/// \brief <em>Specifies the current drop target's unique ID</em>
	///
	/// Specifies the current drop target's unique ID.
	///
	/// \remarks \c ShellListView may change \c pDropTarget rather than \c dropTarget, because not
	///          \c ShellListView will make the change, but the client app when \c ShellListView calls
	///          into \c ExplorerListView to raise the drag'n'drop event.
	LONG dropTarget;
	/// \brief <em>Specifies the pressed modifier keys and mouse buttons</em>
	DWORD keyState;
	/// \brief <em>Specifies the mouse button that is performing the drag'n'drop operation</em>
	DWORD draggingMouseButton;
	/// \brief <em>Specifies the coordinates (in pixels) of the mouse cursor's position relative to the screen's upper-left corner</em>
	POINTL cursorPosition;
	/// \brief <em>A bit field of \c HitTestConstants flags, specifying the exact part of the control that the position specified by \c cursorPosition lies in</em>
	UINT hitTestDetails;
	/// \brief <em>Specifies the speed multiplier for horizontal auto-scrolling</em>
	///
	/// Specifies the speed multiplier for horizontal auto-scrolling. \c ShellListView may change this
	/// value.
	LONG autoHScrollVelocity;
	/// \brief <em>Specifies the speed multiplier for vertical auto-scrolling</em>
	///
	/// Specifies the speed multiplier for vertical auto-scrolling. \c ShellListView may change this value.
	LONG autoVScrollVelocity;
	/// \brief <em>Specifies the \c IDropTargetHelper instance that \c ExplorerListView is using for drag image support</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	IDropTargetHelper* pDropTargetHelper;
} SHLVMDRAGDROPEVENTDATA, *LPSHLVMDRAGDROPEVENTDATA;

/// \brief <em>Holds data that is transfered with the \c SHLVM_GETINFOTIPEX message</em>
///
/// \sa SHLVM_GETINFOTIPEX
typedef struct SHLVMGETINFOTIPDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Holds any text that \c ShellListView should prepend to the shell info tip</em>
	LPWSTR pPrependedText;
	/// \brief <em>Holds the length of the string specified by \c pPrependedText in characters, not including the terminating \c NULL character</em>
	UINT prependedTextSize;
	/// \brief <em>Holds the shell info tip text (set by \c ShellListView)</em>
	LPWSTR pShellInfoTipText;
	/// \brief <em>Holds the length of the string specified by \c pShellInfoTipText in characters, not including the terminating \c NULL character</em>
	UINT shellInfoTipTextSize;
	/// \brief <em>If \c TRUE, the text specified by \c pPrependedText should be prepended to the text specified by \c pShellInfoTipText (set by \c ShellListView)</em>
	BOOL prependText;
	/// \brief <em>If \c TRUE, the tool tip should not be displayed (set by \c ShellListView)</em>
	BOOL cancelToolTip;
} SHLVMGETINFOTIPDATA, *LPSHLVMGETINFOTIPDATA;

/// \brief <em>Holds data that is transfered with the \c SHLVM_GETSUBITEMCONTROL message</em>
///
/// \sa SHLVM_GETSUBITEMCONTROL
typedef struct SHLVMGETSUBITEMCONTROLDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Holds the unique ID of the list view item for which the message is sent</em>
	LONG itemID;
	/// \brief <em>Holds the unique ID of the list view column for which the message is sent</em>
	LONG columnID;
	/// \brief <em>Specifies the intented use of the sub-item control being requested</em>
	///
	/// Specifies the intented use (in-place editing or visual representation) of the sub-item control
	/// being requested. Any of the values defined by \c ExplorerListView's \c SubItemControlKindConstants
	/// enumeration are valid.
	int controlKind;
	/// \brief <em>Specifies how in-place editing mode is being entered</em>
	///
	/// Specifies how in-place editing mode is being entered, if the sub-item control for in-place editing
	/// the sub-item is being requested. Any of the values defined by \c ExplorerListView's
	/// \c SubItemEditModeConstants enumeration are valid.
	int subItemEditMode;
	/// \brief <em>Receives the sub-item control to use</em>
	///
	/// \c ShellListView sets this member to a value that specifies the sub-item control to use. Any of the
	/// values defined by \c ExplorerListView's \c SubItemControlConstants enumeration are valid.
	int subItemControl;
	/// \brief <em>Identifies the shell property being represented by the sub-item</em>
	///
	/// \c ShellListView sets this member to the unique identifier of the shell property being represented
	/// by the sub-item.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	PROPERTYKEY propertyKey;
	/// \brief <em>Describes the shell property being represented by the sub-item</em>
	///
	/// \c ShellListView sets this member to an implementation of the \c IPropertyDescription interface
	/// that provides information about the shell property being represented by the sub-item.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761561.aspx">IPropertyDescription</a>
	IPropertyDescription* pPropertyDescription;
	/// \brief <em>Holds the current value of the shell property being represented by the sub-item</em>
	///
	/// \c ShellListView sets this member to the current value of the shell property being represented
	/// by the sub-item.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/aa380072.aspx">PROPVARIANT</a>
	PROPVARIANT propertyValue;
	/// \brief <em>If set to \c TRUE, \c ShellListView suggests usage of a sub-item control; otherwise not</em>
	BOOL useSubItemControl;
	/// \brief <em>If set to \c TRUE, \c ShellListView has set the current property value; otherwise \c ExplorerListView should do this</em>
	BOOL hasSetValue;
} SHLVMGETSUBITEMCONTROLDATA, *LPSHLVMGETSUBITEMCONTROLDATA;

/// \brief <em>Holds data that is transfered with the \c EXTVM_ADDITEM message</em>
///
/// \sa EXTVM_ADDITEM
typedef struct EXTVMADDITEMDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the handle of the new item's parent item</em>
	///
	/// Specifies the new item's immediate parent item. If set to \c NULL, the item will be a top-level
	/// item.
	OLE_HANDLE parentItem;
	/// \brief <em>Specifies the handle of the new item's preceding item</em>
	///
	/// Specifies the new item's preceding item. This may be either a treeview item handle or a value
	/// defined by \c ExplorerTreeView's \c InsertAfterConstants enumeration. If set to \c NULL, the item
	/// will be inserted after the last item that has the same immediate parent item.
	OLE_HANDLE insertAfter;
	/// \brief <em>Specifies the new item's caption text</em>
	///
	/// Specifies the new item's caption text. The maximum number of characters in this text is
	/// \c MAX_TVITEMTEXTLENGTH. If set to \c NULL, the treeview control will fire its
	/// \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \sa MAX_TVITEMTEXTLENGTH
	LPTSTR pItemText;
	/// \brief <em>Specifies whether an expando is drawn next to the new item</em>
	///
	/// This value specifies whether to draw a "+" or "-" next to the item indicating the item has
	/// sub-items. Any of the values defined by \c ExplorerTreeView's \c HasExpandoConstants enumeration is
	/// valid. If set to \c heCallback, the treeview control will fire its \c ItemGetDisplayInfo event each
	/// time this property's value is required.
	LONG hasExpando;
	/// \brief <em>Specifies the new item's icon</em>
	///
	/// Specifies the zero-based index of the item's icon in the treeview control's \c ilItems imagelist.
	/// If set to -1, the treeview control will fire its \c ItemGetDisplayInfo event each time this
	/// property's value is required. A value of -2 means 'not specified' and is valid if there's no
	/// imagelist associated with the treeview control.
	LONG iconIndex;
	/// \brief <em>Specifies the new item's selected icon</em>
	///
	/// Specifies the zero-based index of the item's selected icon in the treeview control's \c ilItems
	/// imagelist. This icon will be used instead of the normal icon identified by \c iconIndex if the item
	/// is the caret item. If set to -1, the treeview control will fire its \c ItemGetDisplayInfo event
	/// each time this property's value is required. If set to -2, the normal icon specified by
	/// \c iconIndex will be used.
	LONG selectedIconIndex;
	/// \brief <em>Specifies the new item's expanded icon</em>
	///
	/// Specifies the zero-based index of the item's expanded icon in the treeview control's \c ilItems
	/// imagelist. This icon will be used instead of the normal icon identified by \c iconIndex if the item
	/// is expanded. If set to -1, the treeview control will fire its \c ItemGetDisplayInfo event each time
	/// this property's value is required. If set to -2, the normal icon specified by \c iconIndex will be
	/// used.
	LONG expandedIconIndex;
	/// \brief <em>Specifies the new item's overlay icon</em>
	///
	/// Specifies the one-based index of the item's overlay icon in the treeview control's \c ilItems
	/// imagelist. If set to 0, the new item won't have an overlay.
	LONG overlayIndex;
	/// \brief <em>A \c LONG value that will be associated with the item</em>
	LONG itemData;
	/// \brief <em>Specifies whether the new item is drawn ghosted</em>
	VARIANT_BOOL isGhosted;
	/// \brief <em>Specifies whether the new item is a virtual (i. e. invisible) item</em>
	VARIANT_BOOL isVirtual;
	/// \brief <em>Specifies the new item's height multiplier</em>
	LONG heightIncrement;
	/// \brief <em>Will be set to the inserted item's handle</em>
	HTREEITEM hInsertedItem;
} EXTVMADDITEMDATA, *LPEXTVMADDITEMDATA;

/// \brief <em>Holds data that is transfered with the \c EXTVM_CREATEITEMCONTAINER message</em>
///
/// \sa EXTVM_CREATEITEMCONTAINER
typedef struct EXTVMCREATEITEMCONTAINERDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Holds the number of elements in the array specified by \c pItems</em>
	UINT numberOfItems;
	/// \brief <em>An array containing the handles of the items that the container shall contain</em>
	HTREEITEM* pItems;
	/// \brief <em>Receives the container object's \c IDispatch implementation</em>
	LPDISPATCH pContainer;
} EXTVMCREATEITEMCONTAINERDATA, *LPEXTVMCREATEITEMCONTAINERDATA;

/// \brief <em>Holds data that is transfered with the \c EXTVM_GETITEMHANDLESFROMVARIANT message</em>
///
/// \sa EXTVM_GETITEMHANDLESFROMVARIANT
typedef struct EXTVMGETITEMHANDLESDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>The \c VARIANT from which the item handles shall be retrieved</em>
	VARIANT* pVariant;
	/// \brief <em>Receives the number of item handles that have been retrieved from the \c VARIANT</em>
	UINT numberOfItems;
	/// \brief <em>Receives the item handles that have been retrieved from the \c VARIANT</em>
	///
	/// \remarks The array is created by the callee and must be freed by the caller.
	HTREEITEM* pItemHandles;
} EXTVMGETITEMHANDLESDATA, *LPEXTVMGETITEMHANDLESDATA;

/// \brief <em>Holds data that is transfered with the \c EXTVM_HITTEST message</em>
///
/// \sa EXTVM_HITTEST
typedef struct EXTVMHITTESTDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the coordinates (in pixels relative to the control's upper-left corner) of the point to test</em>
	POINT pointToTest;
	/// \brief <em>Will be set to the hit item's handle</em>
	HTREEITEM hHitItem;
	/// \brief <em>Receives a bit field of TVHT_* flags, that holds further details about the control's part below the specified point</em>
	UINT hitTestDetails;
} EXTVMHITTESTDATA, *LPEXTVMHITTESTDATA;

/// \brief <em>Holds data that is transfered with the \c EXTVM_OLEDRAG message</em>
///
/// \sa EXTVM_OLEDRAG
typedef struct EXTVMOLEDRAGDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>If \c TRUE, \c SHDoDragDrop is called; otherwise \c IExplorerTreeView::OLEDrag is called</em>
	BOOL useSHDoDragDrop;
	/// \brief <em>If \c TRUE, the shell's default implementation of \c IDropSource is used; otherwise \c ExplorerTreeView's implementation is used</em>
	BOOL useShellDropSource;
	/// \brief <em>Specifies the data object to drag</em>
	LPDATAOBJECT pDataObject;
	/// \brief <em>Specifies the drop effects supported by the source</em>
	DWORD supportedEffects;
	/// \brief <em>Specifies the items being dragged</em>
	LPUNKNOWN pDraggedItems;
	/// \brief <em>Receives the effect being performed (as returned by \c DoDragDrop)</em>
	DWORD performedEffects;
} EXTVMOLEDRAGDATA, *LPEXTVMOLEDRAGDATA;

/// \brief <em>Holds data that is transfered with the \c SHTVM_HANDLEDRAGEVENT and \c EXTVM_FIREDRAGDROPEVENT messages</em>
///
/// \sa SHTVM_HANDLEDRAGEVENT, EXTVM_FIREDRAGDROPEVENT
typedef struct SHTVMDRAGDROPEVENTDATA
{
	/// \brief <em>Specifies the struct's size in byte</em>
	DWORD structSize;
	/// \brief <em>Specifies the dragged data object</em>
	LPUNKNOWN pDataObject;
	/// \brief <em>Specifies the current drop effect</em>
	///
	/// Specifies the current drop effect. \c ShellTreeView may change this value to a more appropriate
	/// effect.
	DWORD effect;
	/// \brief <em>Specifies the current drop target</em>
	///
	/// Specifies the current drop target. \c ShellTreeView may change this value to a more appropriate
	/// item.
	///
	/// \remarks \c ShellTreeView may change \c pDropTarget rather than \c hDropTarget, because not
	///          \c ShellTreeView will make the change, but the client app when \c ShellTreeView calls
	///          into \c ExplorerTreeView to raise the drag'n'drop event.
	LPUNKNOWN pDropTarget;
	/// \brief <em>Specifies the current drop target's handle</em>
	///
	/// Specifies the current drop target's item handle.
	///
	/// \remarks \c ShellTreeView may change \c pDropTarget rather than \c hDropTarget, because not
	///          \c ShellTreeView will make the change, but the client app when \c ShellTreeView calls
	///          into \c ExplorerTreeView to raise the drag'n'drop event.
	HTREEITEM hDropTarget;
	/// \brief <em>Specifies the pressed modifier keys and mouse buttons</em>
	DWORD keyState;
	/// \brief <em>Specifies the mouse button that is performing the drag'n'drop operation</em>
	DWORD draggingMouseButton;
	/// \brief <em>Specifies the coordinates (in pixels) of the mouse cursor's position relative to the screen's upper-left corner</em>
	POINTL cursorPosition;
	/// \brief <em>Specifies the y-coordinate (in pixels) of the mouse cursor's position relative to the \c pDropTarget item's upper border</em>
	LONG yToItemTop;
	/// \brief <em>A bit field of \c TVHT_* flags, specifying the exact part of the control that the position specified by \c cursorPosition lies in</em>
	UINT hitTestDetails;
	/// \brief <em>Specifies whether the drop target will be expanded automatically</em>
	///
	/// Specifies whether the drop target will be expanded automatically. \c ShellTreeView may change this
	/// value.
	VARIANT_BOOL autoExpandItem;
	/// \brief <em>Specifies the speed multiplier for horizontal auto-scrolling</em>
	///
	/// Specifies the speed multiplier for horizontal auto-scrolling. \c ShellTreeView may change this
	/// value.
	LONG autoHScrollVelocity;
	/// \brief <em>Specifies the speed multiplier for vertical auto-scrolling</em>
	///
	/// Specifies the speed multiplier for vertical auto-scrolling. \c ShellTreeView may change this value.
	LONG autoVScrollVelocity;
	/// \brief <em>Specifies the \c IDropTargetHelper instance that \c ExplorerTreeView is using for drag image support</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	IDropTargetHelper* pDropTargetHelper;
} SHTVMDRAGDROPEVENTDATA, *LPSHTVMDRAGDROPEVENTDATA;

//////////////////////////////////////////////////////////////////////
/// \name Common ShellBrowserControls messages
///
//@{
/// \brief <em>Sent by \c ShellBrowserControls to the control window as handshake to exchange interface pointers and so on</em>
///
/// \param[in] wParam Must be set to 0.
/// \param[in] lParam A pointer to a \c SHBHANDSHAKE struct containing the details.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa SHBHANDSHAKE
#define SHBM_HANDSHAKE SHBM_COMMON_FIRST
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Internal messages sent from ShellListView to ExplorerListView
///
//@{
/// \brief <em>Queries \c ExplorerListView to insert a column</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXLVMADDCOLUMNDATA struct describing the column to insert. This
///                struct also receives the inserted column's unique ID.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVMADDCOLUMNDATA
#define EXLVM_ADDCOLUMN 1
/// \brief <em>Queries \c ExplorerListView to return the zero-based index of the column with the specified unique ID</em>
///
/// \param[in] wParam The column's unique ID.
/// \param[out] lParam Receives the column's zero-based index.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_COLUMNINDEXTOID
#define EXLVM_COLUMNIDTOINDEX 2
/// \brief <em>Queries \c ExplorerListView to return the unique ID of the column with the specified zero-based index</em>
///
/// \param[in] wParam The column's zero-based index.
/// \param[out] lParam Receives the column's unique ID.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_COLUMNIDTOINDEX
#define EXLVM_COLUMNINDEXTOID 3
/// \brief <em>Queries \c ExplorerListView for a list view column's \c ListViewColumn object</em>
///
/// \param[in] wParam The unique ID of the column to retrieve.
/// \param[out] lParam Receives the \c ListViewColumn object's implementation of the \c IDispatch
///             interface.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETITEMBYID
#define EXLVM_GETCOLUMNBYID 4
/// \brief <em>Sets a column's \c SortArrow property</em>
///
/// \param[in] wParam The zero-based index of the column for which the property shall be set.
/// \param[in] lParam The setting to apply. Any of the values defined by the \c SortArrowConstants
///            enumeration is valid.
///
/// \return An \c HRESULT error code.
#define EXLVM_SETSORTARROW 5
/// \brief <em>Retrieves the current setting of a column's \c BitmapHandle property</em>
///
/// \param[in] wParam The zero-based index of the column for which the property shall be retrieved.
/// \param[out] lParam Receives the property's current setting.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_SETCOLUMNBITMAP
#define EXLVM_GETCOLUMNBITMAP 6
/// \brief <em>Sets a column's \c BitmapHandle property</em>
///
/// \param[in] wParam The zero-based index of the column for which the property shall be set.
/// \param[in] lParam The setting to apply.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETCOLUMNBITMAP
#define EXLVM_SETCOLUMNBITMAP 7
/// \brief <em>Queries \c ExplorerListView to insert an item</em>
///
/// \param[in] wParam If \c TRUE, the item is inserted in fast mode; otherwise not.\n
///            Usually item insertion is done through \c ExplorerListView's COM classes which is more
///            comfortable, but also more time consuming. In fast mode \c ExplorerListView bypasses the
///            the COM classes and inserts the item directly. The trade-off is that some settings are
///            hard-coded.\n
///            In fast mode the members of the \c EXLVMADDITEMDATA struct specified by \c lParam are
///            treated as follows:
///            - \c insertAt: as specified
///            - \c pItemText: will be treated as \c NULL
///            - \c iconIndex: as specified
///            - \c overlayIndex: as specified
///            - \c itemIndentation: as specified
///            - \c itemData: as specified
///            - \c isGhosted: as specified
///            - \c groupID: as specified
///            - \c pTileViewColumns: will be treated as \c Empty
/// \param[in,out] lParam Points to a \c EXLVMADDITEMDATA struct describing the item to insert. This
///                struct also receives the inserted item's unique ID.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVMADDITEMDATA
#define EXLVM_ADDITEM 8
/// \brief <em>Queries \c ExplorerListView to return the zero-based index of the item with the specified unique ID</em>
///
/// \param[in] wParam The item's unique ID.
/// \param[out] lParam Receives the item's zero-based index.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_ITEMINDEXTOID
#define EXLVM_ITEMIDTOINDEX 9
/// \brief <em>Queries \c ExplorerListView to return the unique ID of the item with the specified zero-based index</em>
///
/// \param[in] wParam The item's zero-based index.
/// \param[out] lParam Receives the item's unique ID.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_ITEMIDTOINDEX
#define EXLVM_ITEMINDEXTOID 10
/// \brief <em>Queries \c ExplorerListView for a list view item's \c ListViewItem object</em>
///
/// \param[in] wParam The unique ID of the item to retrieve.
/// \param[out] lParam Receives the \c ListViewItem object's implementation of the \c IDispatch
///             interface.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETCOLUMNBYID
#define EXLVM_GETITEMBYID 11
/// \brief <em>Creates a \c ListViewItemContainer object for the specified items</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXLVMCREATEITEMCONTAINERDATA struct containing the parameters
///                required to create the object. This struct also receives a pointer to the object's
///                \c IDispatch implementation.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVMCREATEITEMCONTAINERDATA
#define EXLVM_CREATEITEMCONTAINER 12
/// \brief <em>Queries \c ExplorerListView to retrieve the unique item IDs from a \c VARIANT that contains either \c IListViewItem objects or a \c IListViewItemContainer object</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXLVMGETITEMIDSDATA struct containing the \c VARIANT from which
///                the IDs shall be retrieved. This struct also receives the extracted item IDs.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVMGETITEMIDSDATA
#define EXLVM_GETITEMIDSFROMVARIANT 13
/// \brief <em>Sets the specified item's icon index</em>
///
/// \param[in] wParam The unique ID of the item for which to set the icon index.
/// \param[in] lParam The icon index to set.
///
/// \return An \c HRESULT error code.
#define EXLVM_SETITEMICONINDEX 14
/// \brief <em>Retrieves the specified item's position within the list view control's client area</em>
///
/// \param[in] wParam The unique ID of the item for which to retrieve the position.
/// \param[out] lParam Points to a \c POINT struct that receives the item's position within the list view
///             control's client area.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_SETITEMPOSITION
#define EXLVM_GETITEMPOSITION 15
/// \brief <em>Sets the specified item's position within the list view control's client area</em>
///
/// \param[in] wParam The unique ID of the item for which to set the position.
/// \param[in] lParam Points to a \c POINT struct that contains the item's position within the list view
///            control's client area.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETITEMPOSITION
#define EXLVM_SETITEMPOSITION 16
/// \brief <em>Moves the positions of the items in the \c IExplorerListView::DraggedItems container by the specified amount of pixels</em>
///
/// \param[in] wParam 0
/// \param[in] lParam Points to a \c SIZE struct that specifies the distance by whicht to move each item.
///
/// \return An \c HRESULT error code.
#define EXLVM_MOVEDRAGGEDITEMS 17
/// \brief <em>Queries \c ExplorerListView to hit-test a point</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXLVMHITTESTDATA struct describing the point to hit-test. This
///                struct also receives the hit item's unique ID, the hit sub-item's index and the hit
///                test flags.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVMHITTESTDATA
#define EXLVM_HITTEST 18
/// \brief <em>Sorts the items</em>
///
/// \param[in] wParam The zero-based index of the column by which to sort. If set to \c -1,
///            \c IExplorerListView::SortItems is called with the \c column parameter being set to
///            \c Empty.
/// \param[in] lParam 0
///
/// \return An \c HRESULT error code.
///
/// \remarks \c IExplorerListView::SortItems will be called with all parameters except \c column being
///          set to the defaults.
#define EXLVM_SORTITEMS 19
/// \brief <em>Retrieves the current setting of the \c SortOrder property</em>
///
/// \param[in] wParam 0
/// \param[out] lParam Receives the current setting of the \c SortOrder property. Any of the values
///             defined by the \c SortOrderConstants enumeration is valid.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_SETSORTORDER
#define EXLVM_GETSORTORDER 20
/// \brief <em>Sets the \c SortOrder property</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The setting to apply. Any of the values defined by the \c SortOrderConstants
///            enumeration is valid.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETSORTORDER
#define EXLVM_SETSORTORDER 21
/// \brief <em>Retrieves the insertion mark position that is closest to the specified point</em>
///
/// \param[in] wParam A pointer to a \c POINT structure defining the position for which to retrieve the
///            closest insertion mark position. The coordinates must be relative to the control's
///            upper-left corner.
/// \param[out] lParam A pointer to an array of two \c int. The first \c int receives the zero-based
///             index of the item, at which the insertion mark should be displayed. The second \c int
///             receives the insertion mark's position relative to this item.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_SETINSERTMARK
#define EXLVM_GETCLOSESTINSERTMARKPOSITION 22
/// \brief <em>Sets the attached control's insertion mark's position</em>
///
/// \param[in] wParam The zero-based index of the item at which the insertion mark will be displayed. If
///            set to \c -1, the insertion mark will be removed.
/// \param[in] lParam The insertion mark's position relative to the specified item. Any of the values
///            defined by the \c InsertMarkPositionConstants enumeration is valid.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETCLOSESTINSERTMARKPOSITION, EXLVM_SETDROPHILITEDITEM
#define EXLVM_SETINSERTMARK 23
/// \brief <em>Sets the attached control's drop-hilited item</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The \c IListViewItem implementation of the item to drop-hilite. If set to \c NULL,
///            the highlighting is removed.
///
/// \return An \c HRESULT error code.
///
/// \remarks We pass an object here rather than an item index, because \c ShellListView gets this object
///          from \c ExplorerListView, so compatibility isn't a problem.
///
/// \sa EXLVM_SETINSERTMARK
#define EXLVM_SETDROPHILITEDITEM 24
/// \brief <em>Asks \c ExplorerListView whether it currently is the source of a drag'n'drop operation</em>
///
/// \param[in] wParam 0
/// \param[in] lParam 0
///
/// \return \c S_OK if the control is currently the source of a drag'n'drop operation, \c S_FALSE if it
///         is not.
#define EXLVM_CONTROLISDRAGSOURCE 25
/// \brief <em>Queries \c ExplorerListView to fire the specified OLE drag'n'drop event</em>
///
/// \param[in] wParam The event to fire. Any of the \c OLEDRAGEVENT_* constants is valid.
/// \param[in,out] lParam Points to a \c SHLVMDRAGDROPEVENTDATA struct containing the event's parameters.
///                Some members of this struct may be customized by the event handler.
///
/// \return An \c HRESULT error code.
///
/// \sa SHLVMDRAGDROPEVENTDATA
#define EXLVM_FIREDRAGDROPEVENT 26
/// \brief <em>Queries \c ExplorerListView to start an OLE Drag'n'Drop operation</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXLVMOLEDRAGDATA struct containing the parameters required to
///                start OLE drag'n'drop. This struct also receives the performed drop effect.
///
/// \return An \c HRESULT error code. If the message fails, the caller should try again with
///         \c EXLVMOLEDRAGDATA::useSHDoDragDrop set to \c FALSE.
///
/// \sa EXLVMOLEDRAGDATA
#define EXLVM_OLEDRAG 27
/// \brief <em>Retrieves an image list</em>
///
/// \param[in] wParam The image list to retrieve. Any of the values defined by the \c ImageListConstants
///            enumeration is valid.
/// \param[out] lParam Receives the handle of the image list to retrieve.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_SETIMAGELIST
#define EXLVM_GETIMAGELIST 28
/// \brief <em>Sets an image list</em>
///
/// \param[in] wParam The image list to set. Any of the values defined by the \c ImageListConstants
///            enumeration is valid.
/// \param[in] lParam The handle of the image list to set.
///
/// \return An \c HRESULT error code.
///
/// \sa EXLVM_GETIMAGELIST
#define EXLVM_SETIMAGELIST 29
/// \brief <em>Retrieves the current setting of the \c View property</em>
///
/// \param[in] wParam 0
/// \param[out] lParam Receives the current setting of the \c View property. Any of the values defined
///             by the \c ViewConstants enumeration is valid.
///
/// \return An \c HRESULT error code.
#define EXLVM_GETVIEWMODE 30
/// \brief <em>Retrieves the current setting of the \c TileViewItemLines property</em>
///
/// \param[in] wParam 0
/// \param[out] lParam Receives the current setting of the \c TileViewItemLines property.
///
/// \return An \c HRESULT error code.
#define EXLVM_GETTILEVIEWITEMLINES 31
/// \brief <em>Retrieves the current setting of the \c ColumnHeaderVisibility property</em>
///
/// \param[in] wParam 0
/// \param[out] lParam Receives the current setting of the \c ColumnHeaderVisibility property. Any of the
///             values defined by the \c ColumnHeaderVisibilityConstants enumeration is valid.
///
/// \return An \c HRESULT error code.
#define EXLVM_GETCOLUMNHEADERVISIBILITY 32
/// \brief <em>Retrieves whether the specified item is currently visible</em>
///
/// \param[in] wParam The unique ID of the item to check.
/// \param[out] lParam Receives \c TRUE if the item is visible; otherwise \c FALSE.
///
/// \return An \c HRESULT error code.
#define EXLVM_ISITEMVISIBLE 33
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Internal messages sent from ExplorerListView to ShellListView
///
//@{
/// \brief <em>Queries \c ShellListView to allocate memory on the process heap</em>
///
/// \param[in] wParam The number of bytes to allocate.
/// \param[in] lParam \c NULL
///
/// \return The address of the allocated memory.
///
/// \remarks The memory must be freed by \c ShellListView.
///
/// \sa SHLVM_FREEMEMORY
#define	SHLVM_ALLOCATEMEMORY 1
/// \brief <em>Queries \c ShellListView to free memory on the process heap</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The address of the memory to free.
///
/// \return The return value is ignored.
///
/// \remarks The memory must have been allocated by \c ShellListView.
///
/// \sa SHLVM_ALLOCATEMEMORY
#define	SHLVM_FREEMEMORY 2
/// \brief <em>Informs \c ShellListView of the unique ID of the listview column that \c ShellListView has inserted into the listview</em>
///
/// \param[in] wParam The inserted column's unique ID.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
#define	SHLVM_INSERTEDCOLUMN 3
/// \brief <em>Informs \c ShellListView that it should free any data it stores for the specified column, because the column has been removed</em>
///
/// \param[in] wParam The removed column's unique ID. If -1, all columns have been removed.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
#define SHLVM_FREECOLUMN 4
/// \brief <em>Queries \c ShellListView for a listview column's \c ShellListViewColumn object</em>
///
/// \param[in] wParam The listview column's unique ID.
/// \param[out] lParam Receives the \c ShellListViewColumn object's implementation of the \c IDispatch
///             interface.
///
/// \return \c S_OK if the message was handled successfully; otherwise \c S_FALSE.
#define SHLVM_GETSHLISTVIEWCOLUMN 5
/// \brief <em>Informs \c ShellListView that a column header has been clicked</em>
///
/// \param[in] wParam The clicked column header's unique ID.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
#define SHLVM_HEADERCLICK 6
/// \brief <em>Queries \c ShellListView to display a column header's drop-down menu</em>
///
/// \param[in] wParam The unique ID of the column header for which to display the drop-down menu.
/// \param[in] lParam The coordinates (in pixels) of the menu's proposed position relative to the
///            control's upper-left corner.
///
/// \return The return value is ignored.
#define SHLVM_HEADERDROPDOWNMENU 7
/// \brief <em>Informs \c ShellListView of the unique ID of the listview item that \c ShellListView has inserted into the listview</em>
///
/// \param[in] wParam The inserted item's unique ID.
/// \param[in] lParam The inserted item's settings.
///
/// \return The return value is ignored.
#define SHLVM_INSERTEDITEM 8
/// \brief <em>Informs \c ShellListView that it should free any data it stores for the specified item, because the item has been removed</em>
///
/// \param[in] wParam The removed item's unique ID. If -1, all items have been removed.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
#define SHLVM_FREEITEM 9
/// \brief <em>Queries \c ShellListView for a listview item's \c ShellListViewItem object</em>
///
/// \param[in] wParam The listview item's unique ID.
/// \param[out] lParam Receives the \c ShellListViewItem object's implementation of the \c IDispatch
///             interface.
///
/// \return \c S_OK if the message was handled successfully; otherwise \c S_FALSE.
#define SHLVM_GETSHLISTVIEWITEM 10
/// \brief <em>Informs \c ShellListView that a listview item has been renamed</em>
///
/// \param[in] wParam The renamed item's unique ID.
/// \param[in] lParam The renamed item's new text as Unicode string.
///
/// \return \c S_OK if \c ShellListView accepts the new name; otherwise \c S_FALSE.
#define SHLVM_RENAMEDITEM 11
/// \brief <em>Asks \c ShellListView to compare the two specified items</em>
///
/// \param[in] wParam A \c LONG array containing the two items' unique IDs.
/// \param[out] lParam Receives \c TRUE if the items have been compared; otherwise \c FALSE.
///
/// \return A negative value if the first item should precede the second, a positive value if the first
///         item should follow the second, or zero if the two items are equivalent.
#define SHLVM_COMPAREITEMS 12
/// \brief <em>Asks \c ShellListView to compare the two specified groups</em>
///
/// \param[in] wParam A \c LONG array containing the two groups' unique IDs.
/// \param[out] lParam Receives \c TRUE if the groups have been compared; otherwise \c FALSE.
///
/// \return A negative value if the first group should precede the second, a positive value if the first
///         group should follow the second, or zero if the two groups are equivalent.
#define SHLVM_COMPAREGROUPS 13
/// \brief <em>Lets \c ShellListView customize handling of the \c LVN_BEGINLABELEDITA notification</em>
///
/// \param[in] wParam \c TRUE if \c ExplorerListView suggests to cancel label-editing; otherwise
///            \c FALSE.
/// \param[in] lParam The \c lParam parameter sent with the \c LVN_BEGINLABELEDITA notification.
///
/// \return \c S_FALSE to cancel label-editing; otherwise \c S_OK.
///
/// \sa SHLVM_BEGINLABELEDITW
#define SHLVM_BEGINLABELEDITA 14
/// \brief <em>Lets \c ShellListView customize handling of the \c LVN_BEGINLABELEDITW notification</em>
///
/// \param[in] wParam \c TRUE if \c ExplorerListView suggests to cancel label-editing; otherwise
///            \c FALSE.
/// \param[in] lParam The \c lParam parameter sent with the \c LVN_BEGINLABELEDITW notification.
///
/// \return \c S_FALSE to cancel label-editing; otherwise \c S_OK.
///
/// \sa SHLVM_BEGINLABELEDITA
#define SHLVM_BEGINLABELEDITW 15
/// \brief <em>Lets \c ShellListView handle the \c LVN_BEGINDRAG notification</em>
///
/// \param[in] wParam \c FALSE if the message is sent to handle the \c LVN_BEGINDRAG notification;
///            \c TRUE if the message is sent to handle the \c LVN_BEGINRDRAG notification.
/// \param[in] lParam The \c lParam parameter sent with the notification.
///
/// \return The return value is ignored.
#define SHLVM_BEGINDRAG 16
/// \brief <em>Queries \c ShellListView to handle the specified OLE drag'n'drop event</em>
///
/// \param[in] wParam The type of the event to handle. Any of the \c OLEDRAGEVENT_* constants is valid.
/// \param[in,out] lParam Points to a \c SHLVMDRAGDROPEVENTDATA struct containing details about the
///                event. Some members of this struct may be customized by \c ShellListView.
///
/// \return \c S_OK if \c ShellListView handled the drag event including \c OLEDrag* events; otherwise
///         \c S_FALSE.
///
/// \sa OLEDRAGEVENT_DRAGENTER, OLEDRAGEVENT_DRAGOVER, OLEDRAGEVENT_DRAGLEAVE, OLEDRAGEVENT_DROP,
///     SHLVMDRAGDROPEVENTDATA
#define SHLVM_HANDLEDRAGEVENT 17
/// \brief <em>Queries \c ShellListView to display the shell context menu</em>
///
/// \param[in] wParam If \c TRUE, the message is sent for the column header control; otherwise for the
///            listview itself.
/// \param[in] lParam The coordinates (in pixels) of the menu's proposed position relative to the
///            control's upper-left corner.
///
/// \return The return value is ignored.
#define SHLVM_CONTEXTMENU 18
/// \brief <em>Queries \c ShellListView for a listview item's info tip text</em>
///
/// \param[in] wParam The listview item's unique ID.
/// \param[out] lParam Receives the address of the Unicode info tip string.
///
/// \return The return value is ignored.
///
/// \remarks \c ExplorerListView is responsible for freeing the string specified by \c lParam using
///          \c CoTaskMemFree.
///
/// \sa SHLVM_GETINFOTIPEX
#define SHLVM_GETINFOTIP 19
/// \brief <em>Lets \c ShellListView do any custom draw stuff</em>
///
/// \param[in] wParam The \c wParam parameter sent with the \c NM_CUSTOMDRAW notification.
/// \param[in] lParam The \c lParam parameter sent with the \c NM_CUSTOMDRAW notification.
///
/// \return The return value is used as the predefined return value of ExplorerListView's
///         \c NM_CUSTOMDRAW notification handler.
#define SHLVM_CUSTOMDRAW 20
/// \brief <em>Informs \c ShellListView that the view mode has been changed</em>
///
/// \param[in] wParam The new view mode as a \c LV_VIEW_* constant.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
#define SHLVM_CHANGEDVIEW 21
/// \brief <em>Queries \c ShellListView for a listview item's info tip text</em>
///
/// \param[in] wParam The listview item's unique ID.
/// \param[in,out] lParam Points to a \c SHLVMGETINFOTIPDATA struct containing the exchanged data.
///
/// \return The return value is ignored.
///
/// \remarks \c ExplorerListView is responsible for freeing the strings specified in the
///          \c SHLVMGETINFOTIPDATA struct using \c CoTaskMemFree.
///
/// \sa SHLVM_GETINFOTIP, SHLVMGETINFOTIPDATA
#define SHLVM_GETINFOTIPEX 22
/// \brief <em>Queries \c ShellListView for a listview sub-item's sub-item control</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c SHLVMGETSUBITEMCONTROLDATA struct containing the exchanged data.
///
/// \return The return value is ignored.
///
/// \sa SHLVMGETSUBITEMCONTROLDATA, SHLVM_ENDSUBITEMEDIT
#define SHLVM_GETSUBITEMCONTROL 23
/// \brief <em>Notifies \c ShellListView about editing of a sub-item</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c SHLVMGETSUBITEMCONTROLDATA struct containing the exchanged data.
///
/// \return The return value is ignored.
///
/// \sa SHLVMGETSUBITEMCONTROLDATA, SHLVM_GETSUBITEMCONTROL
#define SHLVM_ENDSUBITEMEDIT 24
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Internal messages sent from ShellTreeView to ExplorerTreeView
///
//@{
/// \brief <em>Queries \c ExplorerTreeView to insert an item</em>
///
/// \param[in] wParam If \c TRUE, the item is inserted in fast mode; otherwise not.\n
///            Usually item insertion is done through \c ExplorerTreeView's COM classes which is more
///            comfortable, but also more time consuming. In fast mode \c ExplorerTreeView bypasses the
///            the COM classes and inserts the item directly. The trade-off is that some settings are
///            hard-coded.\n
///            In fast mode the members of the \c EXTVMADDITEMDATA struct specified by \c lParam are
///            treated as follows:
///            - \c parentItem: as specified
///            - \c insertAfter: as specified
///            - \c pItemText: will be treated as \c NULL
///            - \c hasExpando: as specified
///            - \c iconIndex: as specified
///            - \c selectedIconIndex: as specified
///            - \c expandedIconIndex: as specified
///            - \c overlayIndex: as specified
///            - \c itemData: as specified
///            - \c isGhosted: as specified
///            - \c isVirtual: as specified
///            - \c heightIncrement: as specified
/// \param[in,out] lParam Points to a \c EXTVMADDITEMDATA struct describing the item to insert. This
///                struct also receives the inserted item's handle.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVMADDITEMDATA
#define EXTVM_ADDITEM 1
/// \brief <em>Queries \c ExplorerTreeView to return the unique ID of the item with the specified handle</em>
///
/// \param[in] wParam The item's handle.
/// \param[out] lParam Receives the item's unique ID.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVM_IDTOITEMHANDLE
#define EXTVM_ITEMHANDLETOID 2
/// \brief <em>Queries \c ExplorerTreeView to return the handle of the item with the specified unique ID</em>
///
/// \param[in] wParam The item's unique ID.
/// \param[out] lParam Receives the item's handle.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVM_ITEMHANDLETOID
#define EXTVM_IDTOITEMHANDLE 3
/// \brief <em>Queries \c ExplorerTreeView for a tree view item's \c TreeViewItem object</em>
///
/// \param[in] wParam The handle of the item to retrieve.
/// \param[out] lParam Receives the \c TreeViewItem object's implementation of the \c IDispatch
///             interface.
///
/// \return An \c HRESULT error code.
#define EXTVM_GETITEMBYHANDLE 4
/// \brief <em>Creates a \c TreeViewItemContainer object for the specified items</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXTVMCREATEITEMCONTAINERDATA struct containing the parameters
///                required to create the object. This struct also receives a pointer to the object's
///                \c IDispatch implementation.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVMCREATEITEMCONTAINERDATA
#define EXTVM_CREATEITEMCONTAINER 5
/// \brief <em>Queries \c ExplorerTreeView to retrieve the item handles from a \c VARIANT that contains either \c ITreeViewItem objects or a \c ITreeViewItemContainer object</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXTVMGETITEMHANDLESDATA struct containing the \c VARIANT from
///                which the handles shall be retrieved. This struct also receives the extracted item
///                handles.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVMGETITEMHANDLESDATA
#define EXTVM_GETITEMHANDLESFROMVARIANT 6
/// \brief <em>Sets the specified item's icon index</em>
///
/// \param[in] wParam The handle of the item for which to set the icon index.
/// \param[in] lParam The icon index to set.
///
/// \return An \c HRESULT error code.
#define EXTVM_SETITEMICONINDEX 7
/// \brief <em>Sets the specified item's selected icon index</em>
///
/// \param[in] wParam The handle of the item for which to set the selected icon index.
/// \param[in] lParam The icon index to set.
///
/// \return An \c HRESULT error code.
#define EXTVM_SETITEMSELECTEDICONINDEX 8
/// \brief <em>Queries \c ExplorerTreeView to hit-test a point</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXTVMHITTESTDATA struct describing the point to hit-test. This
///                struct also receives the hit item's handle and the hit test flags.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVMHITTESTDATA
#define EXTVM_HITTEST 9
/// \brief <em>Sorts the items</em>
///
/// \param[in] wParam If \c TRUE, the items get sorted recursively; otherwise not.
/// \param[in] lParam Specifies the handle of the parent item of the items to sort. If set to \c NULL,
///            the root items get sorted.
///
/// \return An \c HRESULT error code.
///
/// \remarks \c IExplorerTreeView::SortItems will be called with all parameters being set to the
///          defaults.
#define EXTVM_SORTITEMS 10
/// \brief <em>Sets the attached control's drop-hilited item</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The \c ITreeViewItem implementation of the item to drop-hilite. If set to \c NULL,
///            the highlighting is removed.
///
/// \return An \c HRESULT error code.
///
/// \remarks We pass an object here rather than an item index, because \c ShellTreeView gets this object
///          from \c ExplorerTreeView, so compatibility isn't a problem.
#define EXTVM_SETDROPHILITEDITEM 11
/// \brief <em>Queries \c ExplorerTreeView to fire the specified OLE drag'n'drop event</em>
///
/// \param[in] wParam The event to fire. Any of the \c OLEDRAGEVENT_* constants is valid.
/// \param[in,out] lParam Points to a \c SHTVMDRAGDROPEVENTDATA struct containing the event's parameters.
///                Some members of this struct may be customized by the event handler.
///
/// \return An \c HRESULT error code.
///
/// \sa SHTVMDRAGDROPEVENTDATA
#define EXTVM_FIREDRAGDROPEVENT 12
/// \brief <em>Queries \c ExplorerTreeView to start an OLE Drag'n'Drop operation</em>
///
/// \param[in] wParam 0
/// \param[in,out] lParam Points to a \c EXTVMOLEDRAGDATA struct containing the parameters required to
///                start OLE drag'n'drop. This struct also receives the performed drop effect.
///
/// \return An \c HRESULT error code. If the message fails, the caller should try again with
///         \c EXTVMOLEDRAGDATA::useSHDoDragDrop set to \c FALSE.
///
/// \sa EXTVMOLEDRAGDATA
#define EXTVM_OLEDRAG 13
/// \brief <em>Retrieves an image list</em>
///
/// \param[in] wParam The image list to retrieve. Any of the values defined by the \c ImageListConstants
///            enumeration is valid.
/// \param[out] lParam Receives the handle of the image list to retrieve.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVM_SETIMAGELIST
#define EXTVM_GETIMAGELIST 14
/// \brief <em>Sets an image list</em>
///
/// \param[in] wParam The image list to set. Any of the values defined by the \c ImageListConstants
///            enumeration is valid.
/// \param[in] lParam The handle of the image list to set.
///
/// \return An \c HRESULT error code.
///
/// \sa EXTVM_GETIMAGELIST
#define EXTVM_SETIMAGELIST 15
/// \brief <em>Sets the specified item's expanded icon index</em>
///
/// \param[in] wParam The handle of the item for which to set the expanded icon index.
/// \param[in] lParam The icon index to set.
///
/// \return An \c HRESULT error code.
#define EXTVM_SETITEMEXPANDEDICONINDEX 16
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Internal messages sent from ExplorerTreeView to ShellTreeView
///
//@{
/// \brief <em>Queries \c ShellTreeView to allocate memory on the process heap</em>
///
/// \param[in] wParam The number of bytes to allocate.
/// \param[in] lParam \c NULL
///
/// \return The address of the allocated memory.
///
/// \remarks The memory must be freed by \c ShellTreeView.
///
/// \sa SHTVM_FREEMEMORY
#define	SHTVM_ALLOCATEMEMORY 1
/// \brief <em>Queries \c ShellTreeView to free memory on the process heap</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The address of the memory to free.
///
/// \return The return value is ignored.
///
/// \remarks The memory must have been allocated by \c ShellTreeView.
///
/// \sa SHTVM_ALLOCATEMEMORY
#define	SHTVM_FREEMEMORY 2
/// \brief <em>Informs \c ShellTreeView of the ID of the treeview item that \c ShellTreeView is inserting into the treeview</em>
///
/// Informs \c ShellTreeView of the unique ID of the treeview item that \c ShellTreeView is inserting into
/// the treeview. \c ShellTreeView needs this ID to identify the item when handling \c TVN_GETDISPINFO,
/// because this notification is sent before \c ShellTreeView knows the treeview item's handle.
///
/// \param[in] wParam The ID of the item being inserted.
/// \param[in] lParam \c NULL
///
/// \return The return value is ignored.
///
/// \sa SHTVM_INSERTEDITEM
#define SHTVM_INSERTINGITEM 3
/// \brief <em>Informs \c ShellTreeView of the handle of the treeview item that \c ShellTreeView has inserted into the treeview</em>
///
/// \param[in] wParam The inserted item's handle.
/// \param[in] lParam The inserted item's settings.
///
/// \return The return value is ignored.
///
/// \sa SHTVM_INSERTINGITEM
#define SHTVM_INSERTEDITEM 4
/// \brief <em>Informs \c ShellTreeView that it should free any data it stores for the specified item, because the item has been removed</em>
///
/// \param[in] wParam The removed item's handle.
/// \param[in] lParam The removed item's unique ID.
///
/// \return The return value is ignored.
#define SHTVM_FREEITEM 5
/// \brief <em>Informs \c ShellTreeView that a treeview item's handle has changed, i. e. the item has been moved</em>
///
/// \param[in] wParam The item's old handle.
/// \param[in] lParam The item's new handle.
///
/// \return The return value is ignored.
#define SHTVM_UPDATEDITEMHANDLE 6
/// \brief <em>Queries \c ShellTreeView for a treeview item's \c ShellTreeViewItem object</em>
///
/// \param[in] wParam The treeview item's handle.
/// \param[out] lParam Receives the \c ShellTreeViewItem object's implementation of the
///             \c IShTreeViewItem interface.
///
/// \return \c S_OK if the message was handled successfully; otherwise \c S_FALSE.
#define SHTVM_GETSHTREEVIEWITEM 7
/// \brief <em>Informs \c ShellTreeView that a treeview item has been renamed</em>
///
/// \param[in] wParam The renamed item's handle.
/// \param[in] lParam The renamed item's new text as Unicode string.
///
/// \return \c S_OK if \c ShellTreeView accepts the new name; otherwise \c S_FALSE.
#define SHTVM_RENAMEDITEM 8
/// \brief <em>Asks \c ShellTreeView to compare the two specified items</em>
///
/// \param[in] wParam A \c HTREEITEM array containing the two items' handles.
/// \param[out] lParam Receives \c TRUE if the items have been compared; otherwise \c FALSE.
///
/// \return A negative value if the first item should precede the second, a positive value if the first
///         item should follow the second, or zero if the two items are equivalent.
#define SHTVM_COMPAREITEMS 9
/// \brief <em>Lets \c ShellTreeView customize handling of the \c TVN_BEGINLABELEDITA notification</em>
///
/// \param[in] wParam \c TRUE if \c ExplorerTreeView suggests to cancel label-editing; otherwise
///            \c FALSE.
/// \param[in] lParam The \c lParam parameter sent with the \c TVN_BEGINLABELEDITA notification.
///
/// \return \c S_FALSE to cancel label-editing; otherwise \c S_OK.
///
/// \sa SHTVM_BEGINLABELEDITW
#define SHTVM_BEGINLABELEDITA 10
/// \brief <em>Lets \c ShellTreeView customize handling of the \c TVN_BEGINLABELEDITW notification</em>
///
/// \param[in] wParam \c TRUE if \c ExplorerTreeView suggests to cancel label-editing; otherwise
///            \c FALSE.
/// \param[in] lParam The \c lParam parameter sent with the \c TVN_BEGINLABELEDITW notification.
///
/// \return \c S_FALSE to cancel label-editing; otherwise \c S_OK.
///
/// \sa SHTVM_BEGINLABELEDITA
#define SHTVM_BEGINLABELEDITW 11
/// \brief <em>Lets \c ShellTreeView handle the \c TVN_BEGINDRAGA and \c TVN_BEGINRDRAGA notifications</em>
///
/// \param[in] wParam \c FALSE if the message is sent to handle the \c TVN_BEGINDRAGA notification;
///            \c TRUE if the message is sent to handle the \c TVN_BEGINRDRAGA notification.
/// \param[in] lParam The \c lParam parameter sent with the notification.
///
/// \return The return value is ignored.
///
/// \sa SHTVM_BEGINDRAGW
#define SHTVM_BEGINDRAGA 12
/// \brief <em>Lets \c ShellTreeView handle the \c TVN_BEGINDRAGW and \c TVN_BEGINRDRAGW notifications</em>
///
/// \param[in] wParam \c FALSE if the message is sent to handle the \c TVN_BEGINDRAGW notification;
///            \c TRUE if the message is sent to handle the \c TVN_BEGINRDRAGW notification.
/// \param[in] lParam The \c lParam parameter sent with the notification.
///
/// \return The return value is ignored.
///
/// \sa SHTVM_BEGINDRAGA
#define SHTVM_BEGINDRAGW 13
/// \brief <em>Queries \c ShellTreeView to handle the specified OLE drag'n'drop event</em>
///
/// \param[in] wParam The type of the event to handle. Any of the \c OLEDRAGEVENT_* constants is valid.
/// \param[in,out] lParam Points to a \c SHTVMDRAGDROPEVENTDATA struct containing details about the
///                event. Some members of this struct may be customized by \c ShellTreeView.
///
/// \return \c S_OK if \c ShellTreeView handled the drag event including \c OLEDrag* events; otherwise
///         \c S_FALSE.
///
/// \sa OLEDRAGEVENT_DRAGENTER, OLEDRAGEVENT_DRAGOVER, OLEDRAGEVENT_DRAGLEAVE, OLEDRAGEVENT_DROP,
///     SHTVMDRAGDROPEVENTDATA
#define SHTVM_HANDLEDRAGEVENT 14
/// \brief <em>Queries \c ShellTreeView to display the shell context menu</em>
///
/// \param[in] wParam 0
/// \param[in] lParam The coordinates (in pixels) of the menu's proposed position relative to the
///            control's upper-left corner.
///
/// \return The return value is ignored.
#define SHTVM_CONTEXTMENU 15
/// \brief <em>Queries \c ShellTreeView for a treeview item's info tip text</em>
///
/// \param[in] wParam The treeview item's handle.
/// \param[out] lParam Receives the address of the Unicode info tip string.
///
/// \return The return value is ignored.
///
/// \remarks \c ExplorerTreeView is responsible for freeing the string specified by \c lParam using
///          \c CoTaskMemFree.
#define SHTVM_GETINFOTIP 16
/// \brief <em>Lets \c ShellTreeView do any custom draw stuff</em>
///
/// \param[in] wParam The \c wParam parameter sent with the \c NM_CUSTOMDRAW notification.
/// \param[in] lParam The \c lParam parameter sent with the \c NM_CUSTOMDRAW notification.
///
/// \return The return value is used as the predefined return value of ExplorerTreeView's
///         \c NM_CUSTOMDRAW notification handler.
#define SHTVM_CUSTOMDRAW 17
//@}
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
/// \name Additional constants used with \c SHLVM_HANDLEDRAGEVENT and \c SHTVM_HANDLEDRAGEVENT
///
//@{
/// \brief <em>The message was sent from \c IDropTarget::DragEnter</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
#define OLEDRAGEVENT_DRAGENTER 1
/// \brief <em>The message was sent from \c IDropTarget::DragOver</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
#define OLEDRAGEVENT_DRAGOVER 2
/// \brief <em>The message was sent from \c IDropTarget::DragLeave</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
#define OLEDRAGEVENT_DRAGLEAVE 3
/// \brief <em>The message was sent from \c IDropTarget::Drop</em>
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
#define OLEDRAGEVENT_DROP 4
//@}
//////////////////////////////////////////////////////////////////////

/// \brief <em>Listview columns' maximum text length</em>
///
/// This is the maximum length (in characters) of a listview column's text.
///
/// \remarks If we change this value, we also need to update the Doxygen comments in ShBrowserCtls*.idl!
#define MAX_LVCOLUMNTEXTLENGTH MAX_PATH
/// \brief <em>Listview items' maximum text length</em>
///
/// This is the maximum length (in characters) of a listview item's text.
///
/// \remarks If we change this value, we also need to update the Doxygen comments in ShBrowserCtls*.idl!
#define MAX_LVITEMTEXTLENGTH MAX_PATH
/// \brief <em>Treeview items' maximum text length</em>
///
/// This is the maximum length (in characters) of a treeview item's text.
///
/// \remarks If we change this value, we also need to update the Doxygen comments in ShBrowserCtls*.idl!
#define MAX_TVITEMTEXTLENGTH MAX_PATH