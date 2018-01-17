//////////////////////////////////////////////////////////////////////
/// \class ShellListViewItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ShellListViewItem objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ShellListViewItem
/// objects. A \c ShellListViewItems object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewItems, ShellListViewItem, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewItems, ShellListViewItem, ShellListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../res/resource.h"
#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "definitions.h"
#include "_IShellListViewItemsEvents_CP.h"
#include "../../helpers.h"
#include "ShellListView.h"
#include "ShellListViewItem.h"


class ATL_NO_VTABLE ShellListViewItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewItems, &CLSID_ShellListViewItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewItems>,
    public Proxy_IShListViewItemsEvents<ShellListViewItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewItems, &IID_IShListViewItems, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewItems, &IID_IShListViewItems, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;
	friend class ShellListViewNamespace;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWITEMS)

		BEGIN_COM_MAP(ShellListViewItems)
			COM_INTERFACE_ENTRY(IShListViewItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewItemsEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the items</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ShellListViewItem objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x items</em>
	///
	/// Retrieves the next \c numberOfMaxItems items from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of items the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a item's \c IShListViewItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ShellListViewItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// item in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x items</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip items.
	///
	/// \param[in] numberOfItemsToSkip The number of items to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShListViewItems
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaseSensitiveFilters, get_Filter, get_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE get_CaseSensitiveFilters(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CaseSensitiveFilters property</em>
	///
	/// Sets whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaseSensitiveFilters, put_Filter, put_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE put_CaseSensitiveFilters(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ComparisonFunction property</em>
	///
	/// Retrieves an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG, \c BSTR or
	/// \c IShListViewItem). This function must compare its arguments and return a non-zero value if the
	/// arguments are equal and zero otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(ShLvwFilteredPropertyConstants filteredProperty, LONG* pValue);
	/// \brief <em>Sets the \c ComparisonFunction property</em>
	///
	/// Sets an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(LONG itemProperty, LONG pattern);
	/// \endcode
	/// This function must compare its arguments and return a non-zero value if the arguments are equal
	/// and zero otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(ShLvwFilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves an item filter.\n
	/// An \c IShListViewItems collection can be filtered by any of \c IShListViewItem's properties, that
	/// the \c ShLvwFilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction,
	///       ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction,
	///       ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(ShLvwFilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets an item filter.\n
	/// An \c IShListViewItems collection can be filtered by any of \c IShListViewItem's properties, that
	/// the \c ShLvwFilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(ShLvwFilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants,
	///       ShBrowserCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants,
	///       ShBrowserCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(ShLvwFilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShLvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection,
	///       ShBrowserCtlsLibU::ShLvwFilteredPropertyConstants, ShBrowserCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection,
	///       ShBrowserCtlsLibA::ShLvwFilteredPropertyConstants, ShBrowserCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(ShLvwFilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c ShellListViewItem object from the collection</em>
	///
	/// Retrieves a \c ShellListViewItem object from the collection that wraps the listview item identified
	/// by \c itemIdentifier.
	///
	/// \param[in] itemIdentifier A value that identifies the listview item to be retrieved.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShLvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppItem Receives the item's \c IShListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShListViewItem interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellListViewItem, Add, Remove, Contains, ShBrowserCtlsLibU::ShLvwItemIdentifierTypeConstants
	/// \else
	///   \sa ShellListViewItem, Add, Remove, Contains, ShBrowserCtlsLibA::ShLvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType = sliitID, IShListViewItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ShellListViewItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds a new shell listview item to the listview control</em>
	///
	/// Adds a shell listview item with the specified properties at the specified position in the listview
	/// control and returns a \c ShellListViewItem object wrapping the inserted item.
	///
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the new item. The control
	///            takes ownership over the pIDL and will free it if the item is transformed into a normal
	///            treeview item.
	/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
	///            as the last item.
	/// \param[in] managedProperties Specifies which of the listview item's properties are managed by the
	///            \c ShellListView control rather than the listview control/the client application. Any
	///            combination of the values defined by the \c ShLvwManagedItemPropertiesConstants
	///            enumeration is valid. If set to -1, the value of the
	///            \c ShellListView::get_DefaultManagedItemProperties property is used.
	/// \param[in] itemText The new item's caption text. The maximum number of characters in this text
	///            is \c MAX_LVITEMTEXTLENGTH. If set to \c NULL, the listview control will fire its
	///            \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] iconIndex The zero-based index of the item's icon in the listview control's \c ilSmall,
	///            \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the listview
	///            control will fire its \c ItemGetDisplayInfo event each time this property's value is
	///            required. A value of -2 means 'not specified' and is valid if there's no imagelist
	///            associated with the listview control.
	/// \param[in] itemIndentation The new item's indentation in 'Details' view in image widths. If set
	///            to 1, the item's indentation will be 1 image width; if set to 2, it will be 2 image
	///            widths and so on. If set to -1, the listview control will fire its \c ItemGetDisplayInfo
	///            event each time this property's value is required.
	/// \param[in] itemData A \c LONG value that will be associated with the item.
	/// \param[in] groupID The unique ID of the group that the new item will belong to. If set to \c -2,
	///            the item won't belong to any group. Can't be set to -1.
	/// \param[in] tileViewColumns An array of column indexes which specify the columns that will be
	///            used to display additional details below the new item's text in 'Tiles' view. If set
	///            to an empty array, no details will be displayed. If set to \c Empty, the listview
	///            control will fire its \c ItemGetDisplayInfo event each time this property's value is
	///            required.
	/// \param[out] ppAddedItem Receives the added item's \c IShListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c groupID and \c tileViewColumns parameters will be ignored if comctl32.dll is
	///          used in a version older than 6.0.
	///
	/// \if UNICODE
	///   \sa AddExisting, Count, Remove, RemoveAll, ShellListViewItem::FullyQualifiedPIDL,
	///       ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants, MAX_LVITEMTEXTLENGTH
	/// \else
	///   \sa AddExisting, Count, Remove, RemoveAll, ShellListViewItem::FullyQualifiedPIDL,
	///       ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants, MAX_LVITEMTEXTLENGTH
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT pIDLOrParsingName, LONG insertAt = -1, ShLvwManagedItemPropertiesConstants managedProperties = static_cast<ShLvwManagedItemPropertiesConstants>(-1), BSTR itemText = NULL, LONG iconIndex = -2, LONG itemIndentation = 0, LONG itemData = 0, LONG groupID = -2, VARIANT tileViewColumns = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IShListViewItem** ppAddedItem = NULL);
	/// \brief <em>Transfers the specified listview item into a shell listview item</em>
	///
	/// Transfers a listview item into a shell listview item and returns a \c ShellListViewItem object
	/// wrapping the item.
	///
	/// \param[in] itemID The unique ID of the item to transfer. If the item already is a shell item,
	///            its pIDL is freed and replaced with the one specified by \c pIDLOrParsingName.
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the item specified by
	///            \c itemID. The control takes ownership over the pIDL and will free it if the item is
	///            transformed back into a normal listview item.
	/// \param[in] managedProperties Specifies which of the listview item's properties are managed by the
	///            \c ShellListView control rather than the listview control/your application. Any
	///            combination of the values defined by the \c ShLvwManagedItemPropertiesConstants
	///            enumeration is valid. If set to -1, the value of the
	///            \c ShellListView::get_DefaultManagedItemProperties property is used.
	/// \param[out] ppAddedItem Receives the added item's \c IShListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, Remove, RemoveAll, ShellListViewItem::get_ID,
	///       ShellListViewItem::get_FullyQualifiedPIDL, ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa Add, Count, Remove, RemoveAll, ShellListViewItem::get_ID,
	///       ShellListViewItem::get_FullyQualifiedPIDL, ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE AddExisting(LONG itemID, VARIANT pIDLOrParsingName, ShLvwManagedItemPropertiesConstants managedProperties = static_cast<ShLvwManagedItemPropertiesConstants>(-1), IShListViewItem** ppAddedItem = NULL);
	/// \brief <em>Retrieves whether the specified item is part of the item collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the item to be checked.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShLvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the item is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, ShBrowserCtlsLibU::ShLvwItemIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, ShBrowserCtlsLibA::ShLvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType = sliitID, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the items in the collection</em>
	///
	/// Retrieves the number of \c ShellListViewItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Creates the items' shell context menu</em>
	///
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellListView::CreateShellContextMenu,
	///     ShellListViewItem::CreateShellContextMenu, ShellListViewNamespace::CreateShellContextMenu,
	///     ShellListView::DestroyShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateShellContextMenu(OLE_HANDLE* pMenu);
	/// \brief <em>Displays the items' shell context menu</em>
	///
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellListView::DisplayShellContextMenu,
	///     ShellListViewItem::DisplayShellContextMenu, ShellListViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Executes the default command from the items' shell context menu</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellListView::InvokeDefaultShellContextMenuCommand,
	///     ShellListViewItem::InvokeDefaultShellContextMenuCommand
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultShellContextMenuCommand(void);
	/// \brief <em>Removes the specified item from the collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the shell listview item to be removed.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShLvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[in] removeFromListView If \c VARIANT_TRUE, the item is removed from the listview control, too;
	///            otherwise it becomes a normal listview item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibU::ShLvwItemIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibA::ShLvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT itemIdentifier, ShLvwItemIdentifierTypeConstants itemIdentifierType = sliitID, VARIANT_BOOL removeFromListView = VARIANT_TRUE);
	/// \brief <em>Removes all items from the collection</em>
	///
	/// \param[in] removeFromListView If \c VARIANT_TRUE, the items are removed from the listview control,
	///            too; otherwise they become normal listview items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removeFromListView = VARIANT_TRUE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShLvw
	void SetOwner(__in_opt ShellListView* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Filter validation
	///
	//@{
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	/// \param[in] maxValue The maximum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidShellListViewNamespaceFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c IShListViewNamespace type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c IShListViewNamespace.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, put_Filter
	BOOL IsValidShellListViewNamespaceFilter(VARIANT& filter);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsShellListViewNamespaceInSafeArray, get_ComparisonFunction
	BOOL IsIntegerInSafeArray(SAFEARRAY* pSafeArray, int value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified namespace</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] namespacePIDL The fully qualified pIDL of the namespace to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the group; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsIntegerInSafeArray, get_ComparisonFunction
	BOOL IsShellListViewNamespaceInSafeArray(SAFEARRAY* pSafeArray, LONG namespacePIDL, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether an item is part of the collection (applying the filters)</em>
	///
	/// \param[in] itemID The item to check.
	///
	/// \return \c TRUE, if the item is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(LONG itemID);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c ShLvwFilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(ShLvwFilteredPropertyConstants filteredProperty);

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_SHLVW 4
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_SHLVW];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_SHLVW];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_SHLVW];

		/// \brief <em>The \c ShellListView object that owns this collection</em>
		///
		/// \sa SetOwner
		ShellListView* pOwnerShLvw;
		#ifdef USE_STL
			/// \brief <em>Points to the next enumerated item</em>
			std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA>::const_iterator nextEnumeratedItem;
		#else
			/// \brief <em>Points to the next enumerated item</em>
			POSITION nextEnumeratedItem;
		#endif
		/// \brief <em>If \c TRUE, we must filter the items</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties();

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // ShellListViewItems

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewItems), ShellListViewItems)