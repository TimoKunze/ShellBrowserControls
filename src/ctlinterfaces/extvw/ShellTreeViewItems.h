//////////////////////////////////////////////////////////////////////
/// \class ShellTreeViewItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ShellTreeViewItem objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ShellTreeViewItem
/// objects. A \c ShellTreeViewItems object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShTreeViewItems, ShellTreeViewItem, ShellTreeView
/// \else
///   \sa ShBrowserCtlsLibA::IShTreeViewItems, ShellTreeViewItem, ShellTreeView
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
#include "_IShellTreeViewItemsEvents_CP.h"
#include "../../helpers.h"
#include "ShellTreeView.h"
#include "ShellTreeViewItem.h"


class ATL_NO_VTABLE ShellTreeViewItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellTreeViewItems, &CLSID_ShellTreeViewItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellTreeViewItems>,
    public Proxy_IShTreeViewItemsEvents<ShellTreeViewItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IShTreeViewItems, &IID_IShTreeViewItems, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShTreeViewItems, &IID_IShTreeViewItems, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellTreeView;
	friend class ShellTreeViewNamespace;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLTREEVIEWITEMS)

		BEGIN_COM_MAP(ShellTreeViewItems)
			COM_INTERFACE_ENTRY(IShTreeViewItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellTreeViewItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IShTreeViewItemsEvents))
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
	/// the \c ShellTreeViewItem objects managed by this collection object.
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
	///                the pointer to a item's \c IShTreeViewItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ShellTreeViewItem,
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
	/// \name Implementation of IShTreeViewItems
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
	/// \c IShTreeViewItem). This function must compare its arguments and return a non-zero value if the
	/// arguments are equal and zero otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(ShTvwFilteredPropertyConstants filteredProperty, LONG* pValue);
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
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(ShTvwFilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves an item filter.\n
	/// An \c IShTreeViewItems collection can be filtered by any of \c IShTreeViewItem's properties, that
	/// the \c ShTvwFilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction,
	///       ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction,
	///       ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(ShTvwFilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets an item filter.\n
	/// An \c IShTreeViewItems collection can be filtered by any of \c IShTreeViewItem's properties, that
	/// the \c ShTvwFilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(ShTvwFilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants,
	///       ShBrowserCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants,
	///       ShBrowserCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(ShTvwFilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c ShTvwFilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection,
	///       ShBrowserCtlsLibU::ShTvwFilteredPropertyConstants, ShBrowserCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection,
	///       ShBrowserCtlsLibA::ShTvwFilteredPropertyConstants, ShBrowserCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(ShTvwFilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c ShellTreeViewItem object from the collection</em>
	///
	/// Retrieves a \c ShellTreeViewItem object from the collection that wraps the treeview item identified
	/// by \c itemIdentifier.
	///
	/// \param[in] itemIdentifier A value that identifies the treeview item to be retrieved.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShTvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppItem Receives the item's \c IShTreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShTreeViewItem interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellTreeViewItem, Add, Remove, Contains, ShBrowserCtlsLibU::ShTvwItemIdentifierTypeConstants
	/// \else
	///   \sa ShellTreeViewItem, Add, Remove, Contains, ShBrowserCtlsLibA::ShTvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType = stiitHandle, IShTreeViewItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ShellTreeViewItem objects
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

	/// \brief <em>Adds a new shell treeview item to the treeview control</em>
	///
	/// Adds a shell treeview item with the specified properties at the specified position in the treeview
	/// control and returns a \c ShellTreeViewItem object wrapping the inserted item.
	///
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the new item. The control
	///            takes ownership over the pIDL and will free it if the item is transformed into a normal
	///            treeview item.
	/// \param[in] parentItem The new item's immediate parent item. If set to \c NULL, the item will be a
	///            top-level item.
	/// \param[in] insertAfter The new item's preceding item. This may be either a treeview item handle
	///            or a value defined by \c ExplorerTreeView's \c InsertAfterConstants enumeration. If set to
	///            \c NULL, the item will be inserted after the last item that has the same immediate parent
	///            item.
	/// \param[in] managedProperties Specifies which of the treeview item's properties are managed by the
	///            \c ShellTreeView control rather than the treeview control/the client application. Any
	///            combination of the values defined by the \c ShTvwManagedItemPropertiesConstants
	///            enumeration is valid. If set to -1, the value of the
	///            \c ShellTreeView::get_DefaultManagedItemProperties property is used.
	/// \param[in] itemText The new item's caption text. The maximum number of characters in this text
	///            is \c MAX_TVITEMTEXTLENGTH. If set to \c NULL, the treeview control will fire its
	///            \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] hasExpando A value specifying whether to draw a "+" or "-" next to the item
	///            indicating the item has sub-items. Any of the values defined by \c ExplorerTreeView's
	///            \c HasExpandoConstants enumeration is valid. If set to \c heCallback, the treeview control
	///            will fire its \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] iconIndex The zero-based index of the item's icon in the treeview control's imagelist
	///            identified by \c IExplorerTreeView::hImageList. If set to -1, the treeview control will
	///            fire its \c ItemGetDisplayInfo event each time this property's value is required. A
	///            value of -2 means 'not specified' and is valid if there's no imagelist associated with
	///            the treeview control.
	/// \param[in] selectedIconIndex The zero-based index of the item's selected icon in the treeview
	///            control's imagelist identified by \c IExplorerTreeView::hImageList. This icon will be
	///            used instead of the normal icon identified by \c iconIndex if the item is the caret
	///            item. If set to -1, the treeview control will fire its \c ItemGetDisplayInfo event each
	///            time this property's value is required. If set to -2, the normal icon specified by
	///            \c iconIndex will be used.
	/// \param[in] expandedIconIndex The zero-based index of the item's expanded icon in the treeview
	///            control's imagelist identified by \c IExplorerTreeView::hImageList. This icon will be
	///            used instead of the normal icon identified by \c iconIndex if the item is expanded.
	///            If set to -1, the treeview control will fire its \c ItemGetDisplayInfo event each time
	///            this property's value is required. If set to -2, the normal icon specified by
	///            \c iconIndex will be used.
	/// \param[in] itemData A \c LONG value that will be associated with the item.
	/// \param[in] isVirtual If set to \c VARIANT_TRUE, the item will be treated as not existent when
	///            drawing the treeview and therefore won't be drawn. Instead its sub-items will be drawn
	///            at this item's position. If set to \c VARIANT_FALSE, the item and its sub-items will be
	///            drawn normally.
	/// \param[in] heightIncrement The item's height in multiples of the treeview control's basic item
	///            height. E. g. a value of 2 means that the item will be twice as high as an item with
	///            \c heighIncrement set to 1.
	/// \param[out] ppAddedItem Receives the added item's \c IShTreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c expandedIconIndex and \c isVirtual parameters will be ignored if comctl32.dll is
	///          used in a version older than 6.10.
	///
	/// \if UNICODE
	///   \sa AddExisting, Count, Remove, RemoveAll, ShellTreeViewItem::FullyQualifiedPIDL,
	///       ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants, MAX_TVITEMTEXTLENGTH
	/// \else
	///   \sa AddExisting, Count, Remove, RemoveAll, ShellTreeViewItem::FullyQualifiedPIDL,
	///       ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants, MAX_TVITEMTEXTLENGTH
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT pIDLOrParsingName, OLE_HANDLE parentItem = NULL, OLE_HANDLE insertAfter = NULL, ShTvwManagedItemPropertiesConstants managedProperties = static_cast<ShTvwManagedItemPropertiesConstants>(-1), BSTR itemText = NULL, LONG hasExpando = 0/*heNo*/, LONG iconIndex = -2, LONG selectedIconIndex = -2, LONG expandedIconIndex = -2, LONG itemData = 0, VARIANT_BOOL isVirtual = VARIANT_FALSE, LONG heightIncrement = 1, IShTreeViewItem** ppAddedItem = NULL);
	/// \brief <em>Transfers the specified treeview item into a shell treeview item</em>
	///
	/// Transfers a treeview item into a shell treeview item and returns a \c ShellTreeViewItem object
	/// wrapping the item.
	///
	/// \param[in] itemHandle The handle of the item to transfer. If the item already is a shell item,
	///            its pIDL is freed and replaced with the one specified by \c pIDLOrParsingName.
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the item specified by
	///            \c itemHandle. The control takes ownership over the pIDL and will free it if the item is
	///            transformed back into a normal treeview item.
	/// \param[in] managedProperties Specifies which of the treeview item's properties are managed by the
	///            \c ShellTreeView control rather than the treeview control/your application. Any
	///            combination of the values defined by the \c ShTvwManagedItemPropertiesConstants
	///            enumeration is valid. If set to -1, the value of the
	///            \c ShellTreeView::get_DefaultManagedItemProperties property is used.
	/// \param[out] ppAddedItem Receives the added item's \c IShTreeViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, Remove, RemoveAll, ShellTreeViewItem::get_Handle,
	///       ShellTreeViewItem::get_FullyQualifiedPIDL, ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa Add, Count, Remove, RemoveAll, ShellTreeViewItem::get_Handle,
	///       ShellTreeViewItem::get_FullyQualifiedPIDL, ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE AddExisting(OLE_HANDLE itemHandle, VARIANT pIDLOrParsingName, ShTvwManagedItemPropertiesConstants managedProperties = static_cast<ShTvwManagedItemPropertiesConstants>(-1), IShTreeViewItem** ppAddedItem = NULL);
	/// \brief <em>Retrieves whether the specified item is part of the item collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the item to be checked.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShTvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the item is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, ShBrowserCtlsLibU::ShTvwItemIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, ShBrowserCtlsLibA::ShTvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType = stiitHandle, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the items in the collection</em>
	///
	/// Retrieves the number of \c ShellTreeViewItem objects in the collection.
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
	/// \sa DisplayShellContextMenu, ShellTreeView::CreateShellContextMenu,
	///     ShellTreeViewItem::CreateShellContextMenu, ShellTreeViewNamespace::CreateShellContextMenu,
	///     ShellTreeView::DestroyShellContextMenu
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
	/// \sa CreateShellContextMenu, ShellTreeView::DisplayShellContextMenu,
	///     ShellTreeViewItem::DisplayShellContextMenu, ShellTreeViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Executes the default command from the items' shell context menu</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellTreeView::InvokeDefaultShellContextMenuCommand,
	///     ShellTreeViewItem::InvokeDefaultShellContextMenuCommand
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultShellContextMenuCommand(void);
	/// \brief <em>Removes the specified item from the collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the shell treeview item to be removed.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ShTvwItemIdentifierTypeConstants enumeration is valid.
	/// \param[in] removeFromTreeView If \c VARIANT_TRUE, the item is removed from the treeview control, too;
	///            otherwise it becomes a normal treeview item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibU::ShTvwItemIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibA::ShTvwItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT itemIdentifier, ShTvwItemIdentifierTypeConstants itemIdentifierType = stiitHandle, VARIANT_BOOL removeFromTreeView = VARIANT_TRUE);
	/// \brief <em>Removes all items from the collection</em>
	///
	/// \param[in] removeFromTreeView If \c VARIANT_TRUE, the items are removed from the treeview control,
	///            too; otherwise they become normal treeview items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removeFromTreeView = VARIANT_TRUE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShTvw
	void SetOwner(__in_opt ShellTreeView* pOwner);

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
	/// \sa IsValidShellTreeViewNamespaceFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c IShTreeViewNamespace type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c IShTreeViewNamespace.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, put_Filter
	BOOL IsValidShellTreeViewNamespaceFilter(VARIANT& filter);
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
	/// \sa IsPartOfCollection, IsShellTreeViewNamespaceInSafeArray, get_ComparisonFunction
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
	BOOL IsShellTreeViewNamespaceInSafeArray(SAFEARRAY* pSafeArray, LONG namespacePIDL, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether an item is part of the collection (applying the filters)</em>
	///
	/// \param[in] hItem The item to check.
	///
	/// \return \c TRUE, if the item is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(HTREEITEM hItem);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c ShTvwFilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(ShTvwFilteredPropertyConstants filteredProperty);

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_SHTVW 4
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_SHTVW];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_SHTVW];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_SHTVW];

		/// \brief <em>The \c ShellTreeView object that owns this collection</em>
		///
		/// \sa SetOwner
		ShellTreeView* pOwnerShTvw;
		#ifdef USE_STL
			/// \brief <em>Points to the next enumerated item</em>
			std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA>::const_iterator nextEnumeratedItem;
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
};     // ShellTreeViewItems

OBJECT_ENTRY_AUTO(__uuidof(ShellTreeViewItems), ShellTreeViewItems)