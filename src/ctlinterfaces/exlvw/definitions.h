//////////////////////////////////////////////////////////////////////
/// \file definitions.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Common definitions for \c ShellListView</em>
//////////////////////////////////////////////////////////////////////


#pragma once

#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "../definitions.h"


/// \brief <em>Holds details about a shell sub-item</em>
typedef struct SHELLLISTVIEWSUBITEMDATA
{
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		/// \brief <em>Holds the current values of the shell property represented by the list view sub-item</em>
		LPPROPVARIANT pPropertyValue;
	#endif

	SHELLLISTVIEWSUBITEMDATA()
	{
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			pPropertyValue = NULL;
		#endif
	}

	~SHELLLISTVIEWSUBITEMDATA()
	{
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			if(pPropertyValue) {
				PropVariantClear(pPropertyValue);
				delete pPropertyValue;
			}
		#endif
	}
} SHELLLISTVIEWSUBITEMDATA, *LPSHELLLISTVIEWSUBITEMDATA;

/// \brief <em>Holds details like the pIDL about a shell item</em>
typedef struct SHELLLISTVIEWITEMDATA
{
	/// \brief <em>The item's fully qualified pIDL</em>
	PCIDLIST_ABSOLUTE pIDL;
	/// \brief <em>The fully qualified pIDL of the namespace this item belongs to</em>
	PCIDLIST_ABSOLUTE pIDLNamespace;
	/// \brief <em>A bit field specifying which of the item's properties are managed by \c ShellListView</em>
	///
	/// \if UNICODE
	///   \sa ShellListViewItem::get_ManagedProperties,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa ShellListViewItem::get_ManagedProperties,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	ShLvwManagedItemPropertiesConstants managedProperties;
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		#ifdef USE_STL
			/// \brief <em>Holds details about the shell properties represented by the list view item's sub-items</em>
			///
			/// Holds details (e.g. the current value) about the shell properties represented by the list view
			/// item's sub-items. The sub-item's column ID is used as key, the details are stored as value.
			std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA> subItems;
		#else
			/// \brief <em>Holds details about the shell properties represented by the list view item's sub-items</em>
			///
			/// Holds details (e.g. the current value) about the shell properties represented by the list view
			/// item's sub-items. The sub-item's column ID is used as key, the details are stored as value.
			CAtlMap<LONG, LPSHELLLISTVIEWSUBITEMDATA> subItems;
		#endif
	#endif

	SHELLLISTVIEWITEMDATA()
	{
		pIDL = NULL;
		pIDLNamespace = NULL;
	}

	~SHELLLISTVIEWITEMDATA()
	{
		ILFree(const_cast<PIDLIST_ABSOLUTE>(pIDL));
		// don't free pIDLNamespace
		#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
			#ifdef USE_STL
				for(std::unordered_map<LONG, LPSHELLLISTVIEWSUBITEMDATA>::const_iterator it = subItems.cbegin(); it != subItems.cend(); ++it) {
					if(it->second) {
						delete it->second;
					}
				}
				subItems.clear();
			#else
				POSITION p = subItems.GetStartPosition();
				while(p) {
					LPSHELLLISTVIEWSUBITEMDATA pSubItem = subItems.GetValueAt(p);
					if(pSubItem) {
						delete pSubItem;
					}
					subItems.GetNextValue(p);
				}
				subItems.RemoveAll();
			#endif
		#endif
	}
} SHELLLISTVIEWITEMDATA, *LPSHELLLISTVIEWITEMDATA;

/// \brief <em>Holds details like the visibility about a shell column</em>
typedef struct SHELLLISTVIEWCOLUMNDATA
{
	/// \brief <em>Holds the column's caption</em>
	TCHAR pText[MAX_LVCOLUMNTEXTLENGTH];
	/// \brief <em>Holds the column's initial width in pixels</em>
	int width;
	/// \brief <em>Holds the column's format</em>
	DWORD format;
	/// \brief <em>Holds the column's state</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb762538.aspx">SHCOLSTATEF</a>
	SHCOLSTATEF state;
	union
	{
		/// \brief <em>Holds the column's current unique ID with the listview control</em>
		///
		/// If the column is currently visible, i. e. inserted into the listview control, this member holds the
		/// column's unique ID. Otherwise this member is -1.
		///
		/// \remarks If this struct is used for backup purpose, this member will be invalid.
		LONG columnID;
		/// \brief <em>Holds a key that describes the shell property represented by this column</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
		///
		/// \remarks If this struct is not used for backup purpose, this member will be invalid.
		PROPERTYKEY propertyKey;
	};
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		/// \brief <em>Holds the \c PDDT_* flags that control which sub-item control will be used to represent the sub-items of this column</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761535.aspx">PROPDESC_DISPLAYTYPE</a>
		PROPDESC_DISPLAYTYPE displayType;
		/// \brief <em>Holds the \c PDTF_* flags that control which sub-item control will be used to represent the sub-items of this column</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb762527.aspx">PROPDESC_TYPE_FLAGS</a>
		PROPDESC_TYPE_FLAGS typeFlags;
		/// \brief <em>Holds the \c IPropertyDescription object that describes the shell property represented by this column</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb761561.aspx">IPropertyDescription</a>
		IPropertyDescription* pPropertyDescription;
	#endif
	/// \brief <em>Specifies whether the column should be made visible when switching to Details view mode</em>
	///
	/// Columns are used in Details view mode as well as in Tiles view mode. Tiles view mode may require
	/// other columns than Details view mode. Deleting and reloading all columns when switching between both
	/// views seems silly. Instead, when changing the visibility of a column in Details view mode, we store
	/// the new visibility and restore it when switching from Tiles view mode to Details view mode.
	UINT visibleInDetailsView : 1;
	/// \brief <em>Specifies whether the column's visibility in Details view mode has been set explicitly by the client application or the end-user</em>
	///
	/// \c ShellListView can be configured to preserver the column visibility across namespaces. Obviously
	/// this should affect only columns that have been made visible or hidden explicitly by either the
	/// end-user or the client application. The visibility of other columns should still be set by the shell.
	/// To achieve this, we store a flag with each column that we set when the column's visibility is set
	/// explicitly. Later, when preserving the visibility into a new namespace, we use the stored visibility
	/// only if our flag is set.
	UINT visibilityHasBeenChangedExplicitly : 1;
	/// \brief <em>Holds the zero-based index that the column had when it has been visible the last time</em>
	///
	/// If the column is removed from the listview control, its index is stored so that it can be inserted at
	/// its old position when it is made visible again. The initial value of this member is -1.\n
	/// This member is used while in Details view mode.
	///
	/// \sa lastIndex_TileView
	int lastIndex_DetailsView;
	/// \brief <em>Holds the zero-based index that the column had when it has been visible the last time</em>
	///
	/// If the column is removed from the listview control, its index is stored so that it can be inserted at
	/// its old position when it is made visible again. The initial value of this member is -1.\n
	/// This member is used while in Tiles view mode.
	///
	/// \sa lastIndex_DetailsView
	int lastIndex_TilesView;
} SHELLLISTVIEWCOLUMNDATA, *LPSHELLLISTVIEWCOLUMNDATA;

/// \brief <em>Holds details reported to \c ShellListView by background item enumerator threads together with the enumerated items</em>
///
/// \sa ShLvwBackgroundItemEnumTask, ShLvwInsertSingleItemTask
typedef struct SHLVWBACKGROUNDCOLUMNENUMINFO
{
	/// \brief <em>The unique ID of the background item enumerator task sending the message</em>
	ULONGLONG taskID;
	/// \brief <em>The shell index of the first column</em>
	int baseRealColumnIndex;
	#ifdef ACTIVATE_CHUNKED_COLUMNENUMERATIONS
		/// \brief <em>Specifies whether the task will send further \c WM_TRIGGER_COLUMNENUMCOMPLETE messages with more columns from the same enumeration</em>
		UINT moreColumnsToCome : 1;
		/// \brief <em>Specifies the zero-based index of the first new column's details</em>
		///
		/// If the enumeration task reports the enumerated columns in small chunks, it doesn't use a new
		/// dynamic pointer array (DPA) for the column details with each new message. Instead each new column
		/// is added to the same DPA which will contain all columns when the final
		/// \c WM_TRIGGER_COLUMNENUMCOMPLETE message is sent.\n
		/// The members \c firstNewColumn and \c lastNewColumn tell the handler of this message which range of
		/// the DPA contains the new columns that have been enumerated since the previous message.
		///
		/// \sa lastNewColumn
		int firstNewColumn;
		/// \brief <em>Specifies the zero-based index of the last new column's details</em>
		///
		/// If the enumeration task reports the enumerated columns in small chunks, it doesn't use a new
		/// dynamic pointer array (DPA) for the column details with each new message. Instead each new column
		/// is added to the same DPA which will contain all columns when the final
		/// \c WM_TRIGGER_COLUMNENUMCOMPLETE message is sent.\n
		/// The members \c firstNewColumn and \c lastNewColumn tell the handler of this message which range of
		/// the DPA contains the new columns that have been enumerated since the previous message.
		///
		/// \sa firstNewColumn
		int lastNewColumn;
	#endif
	/// \brief <em>Holds the enumerated columns' \c SHELLLISTVIEWCOLUMNDATA structs</em>
	///
	/// \sa SHELLLISTVIEWCOLUMNDATA
	HDPA hColumnBuffer;
} SHLVWBACKGROUNDCOLUMNENUMINFO, *LPSHLVWBACKGROUNDCOLUMNENUMINFO;

/// \brief <em>Holds details reported to \c ShellListView by background item enumerator threads together with the enumerated items</em>
///
/// \sa ShLvwBackgroundItemEnumTask, ShLvwInsertSingleItemTask
typedef struct SHLVWBACKGROUNDITEMENUMINFO
{
	/// \brief <em>The unique ID of the background item enumerator task sending the message</em>
	ULONGLONG taskID;
	/// \brief <em>The zero-based index at which the items shall be inserted into the listview</em>
	LONG insertAt;
	/// \brief <em>If \c TRUE, \c ShellListView should ensure that the enumerated items do not yet exist within the listview</em>
	UINT checkForDuplicates : 1;
	/// \brief <em>If \c TRUE, \c ShellListView should sort the target namespace after item insertion even if the namespace's \c AutoSortItems property is set to \c asiAutoSortOnce</em>
	///
	/// \if UNICODE
	///   \sa ShellListViewNamespace::get_AutoSortItems, ShBrowserCtlsLibU::AutoSortItemsConstants
	/// \else
	///   \sa ShellListViewNamespace::get_AutoSortItems, ShBrowserCtlsLibA::AutoSortItemsConstants
	/// \endif
	UINT forceAutoSort : 1;
	/// \brief <em>Specifies the fully qualified pIDL of the item for which label-edit mode should be entered after insertion</em>
	PIDLIST_ABSOLUTE pIDLToLabelEdit;
	/// \brief <em>Specifies the fully qualified pIDL of the namespace to which the enumerated items shall be associated</em>
	PCIDLIST_ABSOLUTE namespacePIDLToSet;
	/// \brief <em>Holds the enumerated items' pIDLs</em>
	HDPA hPIDLBuffer;
} SHLVWBACKGROUNDITEMENUMINFO, *LPSHLVWBACKGROUNDITEMENUMINFO;

/// \brief <em>Holds details reported to \c ShellListView by background sub-item text extractor threads</em>
///
/// \sa ShLvwSlowColumnTask, ShLvwSubItemControlTask
typedef struct SHLVWBACKGROUNDCOLUMNINFO
{
	/// \brief <em>The unique ID of the item to which the message refers</em>
	LONG itemID;
	/// \brief <em>The unique ID of the column to which the message refers</em>
	LONG columnID;
	/// \brief <em>The text that has been retrieved</em>
	WCHAR pText[MAX_LVITEMTEXTLENGTH];
	#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
		/// \brief <em>The retrieved value of the property represented by the specified column</em>
		LPPROPVARIANT pPropertyValue;
	#endif
} SHLVWBACKGROUNDCOLUMNINFO, *LPSHLVWBACKGROUNDCOLUMNINFO;

/// \brief <em>Holds details reported to \c ShellListView by background info tip text extractor threads</em>
///
/// \sa ShLvwInfoTipTask
typedef struct SHLVWBACKGROUNDINFOTIPINFO
{
	/// \brief <em>The unique ID of the item to which the message refers</em>
	LONG itemID;
	/// \brief <em>The text that has been retrieved</em>
	///
	/// \sa infoTipTextSize
	LPWSTR pInfoTipText;
	/// \brief <em>The length (in characters, without the terminating \c NULL character) of \c pInfoTipText</em>
	///
	/// \sa pInfoTipText
	UINT infoTipTextSize;
} SHLVWBACKGROUNDINFOTIPINFO, *LPSHLVWBACKGROUNDINFOTIPINFO;

/// \brief <em>Holds details reported to \c ShellListView by background thumbnail extractor threads</em>
///
/// \sa ShLvwThumbnailTask, ShLvwLegacyThumbnailTask
typedef struct SHLVWBACKGROUNDTHUMBNAILINFO
{
	/// \brief <em>The unique ID of the item to which the message refers</em>
	LONG itemID;
	/// \brief <em>The \c SIIF_* flags specifying which of the members of this struct contain valid data</em>
	UINT mask;
	/// \brief <em>The maximum size in pixels of the thumbnail that has been retrieved</em>
	SIZE targetThumbnailSize;
	/// \brief <em>The thumbnail or icon of the item to which the message refers</em>
	HBITMAP hThumbnailOrIcon;
	/// \brief <em>The zero-based icon index in the system image list of the executable associated with the item specified by \c itemID</em>
	///
	/// \sa pOverlayImageResource
	int executableIconIndex;
	/// \brief <em>The image resource to display as overlay over the thumbnail of the item specified by \c itemID</em>
	///
	/// \sa executableIconIndex
	LPWSTR pOverlayImageResource;
	/// \brief <em>The \c AII_* flags to set for the item specified by \c itemID</em>
	UINT flags;
	/// \brief <em>If set to \c TRUE, the task has been aborted</em>
	UINT aborted : 1;
} SHLVWBACKGROUNDTHUMBNAILINFO, *LPSHLVWBACKGROUNDTHUMBNAILINFO;

/// \brief <em>Holds details reported to \c ShellListView by background tile view column extractor threads</em>
///
/// \sa ShLvwTileViewTask
typedef struct SHLVWBACKGROUNDTILEVIEWINFO
{
	/// \brief <em>The unique ID of the item to which the message refers</em>
	LONG itemID;
	/// \brief <em>Holds the enumerated columns' indexes</em>
	HDPA hColumnBuffer;
} SHLVWBACKGROUNDTILEVIEWINFO, *LPSHLVWBACKGROUNDTILEVIEWINFO;