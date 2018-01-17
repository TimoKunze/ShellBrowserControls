//////////////////////////////////////////////////////////////////////
/// \class ShellTreeViewNamespaces
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ShellTreeViewNamespace objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ShellTreeViewNamespace
/// objects. A \c ShellTreeViewNamespaces object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShTreeViewNamespaces, ShellTreeViewNamespace, ShellTreeView
/// \else
///   \sa ShBrowserCtlsLibA::IShTreeViewNamespaces, ShellTreeViewNamespace, ShellTreeView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../res/resource.h"
#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "../INamespaceEnumTask.h"
#include "definitions.h"
#include "_IShellTreeViewNamespacesEvents_CP.h"
#include "../../helpers.h"
#include "ShellTreeView.h"
#include "ShellTreeViewNamespace.h"
#include "../common.h"
#include "../NamespaceEnumSettings.h"


class ATL_NO_VTABLE ShellTreeViewNamespaces : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellTreeViewNamespaces, &CLSID_ShellTreeViewNamespaces>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellTreeViewNamespaces>,
    public Proxy_IShTreeViewNamespacesEvents<ShellTreeViewNamespaces>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IShTreeViewNamespaces, &IID_IShTreeViewNamespaces, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShTreeViewNamespaces, &IID_IShTreeViewNamespaces, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellTreeView;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLTREEVIEWNAMESPACES)

		BEGIN_COM_MAP(ShellTreeViewNamespaces)
			COM_INTERFACE_ENTRY(IShTreeViewNamespaces)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellTreeViewNamespaces)
			CONNECTION_POINT_ENTRY(__uuidof(_IShTreeViewNamespacesEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the namespaces</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ShellTreeViewNamespace objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x namespaces</em>
	///
	/// Retrieves the next \c numberOfMaxItems namespaces from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of namespaces the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a namespace's \c IShTreeViewNamespace implementation.
	/// \param[out] pNumberOfItemsReturned The number of namespaces that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ShellTreeViewNamespace,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// namespace in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x namespaces</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip items.
	///
	/// \param[in] numberOfItemsToSkip The number of namespaces to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShTreeViewNamespaces
	///
	//@{
	/// \brief <em>Retrieves a \c ShellTreeViewNamespace object from the collection</em>
	///
	/// Retrieves a \c ShellTreeViewNamespace object from the collection that wraps the namespace identified
	/// by \c namespaceIdentifier.
	///
	/// \param[in] namespaceIdentifier A value that identifies the namespace to be retrieved.
	/// \param[in] namespaceIdentifierType A value specifying the meaning of \c namespaceIdentifier. Any of
	///            the values defined by the \c ShTvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppNamespace Receives the item's \c IShTreeViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShTreeViewNamespace interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellTreeViewNamespace, Add, Remove, Contains,
	///       ShBrowserCtlsLibU::ShTvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa ShellTreeViewNamespace, Add, Remove, Contains,
	///       ShBrowserCtlsLibA::ShTvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType = stnsitExactPIDL, IShTreeViewNamespace** ppNamespace = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ShellTreeViewNamespace objects
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

	/// \brief <em>Adds a new shell namespace to the treeview control</em>
	///
	/// Adds a shell namespace with the specified properties at the specified position in the treeview
	/// control and returns a \c ShellTreeViewNamespace object wrapping the inserted namespace.
	///
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the namespace's parent
	///            shell item. The control takes ownership over the pIDL and will free it if the namespace
	///            is removed from the treeview.\n
	///            The shell item specified by \c pIDLOrParsingName won't be inserted into the treeview,
	///            but its sub-items will.
	/// \param[in] parentItem The new items' immediate parent item. If set to \c NULL, the items will be
	///            top-level items.
	/// \param[in] insertAfter The new items' preceding item. This may be either a treeview item handle
	///            or a value defined by \c ExplorerTreeView's \c InsertAfterConstants enumeration. If set
	///            to \c NULL, the items will be inserted after the last item that has the same immediate
	///            parent item.
	/// \param[in] pEnumerationSettings A \c NamespaceEnumSettings object specifying various item enumeration
	///            options for this namespace. If not specified, the control's default enumeration settings
	///            specified by \c ShellTreeView::get_DefaultNamespaceEnumSettings are used.
	/// \param[in] autoSortItems Specifies whether the sub-tree represented by the new namespace shall be
	///            sorted automatically. If set to \c VARIANT_TRUE, the sub-tree is sorted automatically;
	///            otherwise not.
	/// \param[out] ppAddedNamespace Receives the added item's \c IShTreeViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The items' \c ManagedProperties property is set to the value of the
	///          \c ShellTreeView::get_DefaultManagedItemProperties property.\n
	///          On automatic sorting the items are sorted like in Windows Explorer; shell items are
	///          inserted in front of other items.
	///
	/// \sa Count, Remove, RemoveAll, ShellTreeViewNamespace::get_FullyQualifiedPIDL, NamespaceEnumSettings,
	///     ShellTreeView::get_DefaultNamespaceEnumSettings, ShellTreeView::get_DefaultManagedItemProperties,
	///     ShellTreeViewNamespace::get_AutoSortItems
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT pIDLOrParsingName, OLE_HANDLE parentItem = NULL, OLE_HANDLE insertAfter = NULL, INamespaceEnumSettings* pEnumerationSettings = NULL, VARIANT_BOOL autoSortItems = VARIANT_TRUE, IShTreeViewNamespace** ppAddedNamespace = NULL);
	/// \brief <em>Retrieves whether the specified namespace is part of the namespace collection</em>
	///
	/// \param[in] namespaceIdentifier A value that identifies the namespace to be checked.
	/// \param[in] namespaceIdentifierType A value specifying the meaning of \c namespaceIdentifier. Any of
	///            the values defined by the \c ShTvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the namespace is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, ShBrowserCtlsLibU::ShTvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, ShBrowserCtlsLibA::ShTvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType = stnsitExactPIDL, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the namespaces in the collection</em>
	///
	/// Retrieves the number of \c ShellTreeViewNamespace objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified namespace from the collection</em>
	///
	/// \param[in] namespaceIdentifier A value that identifies the namespace to be removed.
	/// \param[in] namespaceIdentifierType A value specifying the meaning of \c namespaceIdentifier. Any of
	///            the values defined by the \c ShTvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[in] removeFromTreeView If \c VARIANT_TRUE, the items are removed from the treeview control,
	///            too; otherwise they become normal treeview items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibU::ShTvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibA::ShTvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT namespaceIdentifier, ShTvwNamespaceIdentifierTypeConstants namespaceIdentifierType = stnsitExactPIDL, VARIANT_BOOL removeFromTreeView = VARIANT_TRUE);
	/// \brief <em>Removes all namespaces from the collection</em>
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
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellTreeView object that owns this collection</em>
		///
		/// \sa SetOwner
		ShellTreeView* pOwnerShTvw;
		#ifdef USE_STL
			/// \brief <em>Points to the next enumerated namespace</em>
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*>::const_iterator nextEnumeratedNamespace;
		#else
			/// \brief <em>Points to the next enumerated namespace</em>
			POSITION nextEnumeratedNamespace;
		#endif

		Properties();

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // ShellTreeViewNamespaces

OBJECT_ENTRY_AUTO(__uuidof(ShellTreeViewNamespaces), ShellTreeViewNamespaces)