//////////////////////////////////////////////////////////////////////
/// \class ShellListViewNamespaces
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ShellListViewNamespace objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ShellListViewNamespace
/// objects. A \c ShellListViewNamespaces object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewNamespaces, ShellListViewNamespace, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewNamespaces, ShellListViewNamespace, ShellListView
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
#include "LegacyBackgroundItemEnumTask.h"
#include "BackgroundItemEnumTask.h"
#include "_IShellListViewNamespacesEvents_CP.h"
#include "../../helpers.h"
#include "ShellListView.h"
#include "ShellListViewNamespace.h"
#include "../common.h"
#include "../NamespaceEnumSettings.h"


class ATL_NO_VTABLE ShellListViewNamespaces : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewNamespaces, &CLSID_ShellListViewNamespaces>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewNamespaces>,
    public Proxy_IShListViewNamespacesEvents<ShellListViewNamespaces>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewNamespaces, &IID_IShListViewNamespaces, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewNamespaces, &IID_IShListViewNamespaces, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWNAMESPACES)

		BEGIN_COM_MAP(ShellListViewNamespaces)
			COM_INTERFACE_ENTRY(IShListViewNamespaces)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewNamespaces)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewNamespacesEvents))
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
	/// the \c ShellListViewNamespace objects managed by this collection object.
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
	///                the pointer to a namespace's \c IShListViewNamespace implementation.
	/// \param[out] pNumberOfItemsReturned The number of namespaces that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ShellListViewNamespace,
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
	/// \name Implementation of IShListViewNamespaces
	///
	//@{
	/// \brief <em>Retrieves a \c ShellListViewNamespace object from the collection</em>
	///
	/// Retrieves a \c ShellListViewNamespace object from the collection that wraps the namespace identified
	/// by \c namespaceIdentifier.
	///
	/// \param[in] namespaceIdentifier A value that identifies the namespace to be retrieved.
	/// \param[in] namespaceIdentifierType A value specifying the meaning of \c namespaceIdentifier. Any of
	///            the values defined by the \c ShLvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppNamespace Receives the item's \c IShListViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IShListViewNamespace interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellListViewNamespace, Add, Remove, Contains,
	///       ShBrowserCtlsLibU::ShLvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa ShellListViewNamespace, Add, Remove, Contains,
	///       ShBrowserCtlsLibA::ShLvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType = slnsitExactPIDL, IShListViewNamespace** ppNamespace = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ShellListViewNamespace objects
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

	/// \brief <em>Adds a new shell namespace to the listview control</em>
	///
	/// Adds a shell namespace with the specified properties at the specified position in the listview
	/// control and returns a \c ShellListViewNamespace object wrapping the inserted namespace.
	///
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or parsing name of the namespace's parent
	///            shell item. The control takes ownership over the pIDL and will free it if the namespace
	///            is removed from the listview.\n
	///            The shell item specified by \c pIDLOrParsingName won't be inserted into the listview,
	///            but its sub-items will.
	/// \param[in] insertAt The first new item's zero-based index. If set to -1, the items will be inserted
	///            at the end of the list.
	/// \param[in] pEnumerationSettings A \c NamespaceEnumSettings object specifying various item enumeration
	///            options for this namespace. If not specified, the control's default enumeration settings
	///            specified by \c IShListView::DefaultNamespaceEnumSettings are used.
	/// \param[in] autoSortItems Specifies whether the items represented by the new namespace shall be sorted
	///            automatically. Any of the values defined by the \c AutoSortItemsConstants enumeration is
	///            valid.
	/// \param[out] ppAddedNamespace Receives the added item's \c IShListViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The items' \c ManagedProperties property is set to the value of the
	///          \c ShellListView::get_DefaultManagedItemProperties property.\n
	///          On automatic sorting the items are sorted like in Windows Explorer; shell items are
	///          inserted in front of other items.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ShellListViewNamespace::get_FullyQualifiedPIDL,
	///       NamespaceEnumSettings, ShellListView::get_DefaultNamespaceEnumSettings,
	///       ShellListView::get_DefaultManagedItemProperties, ShellListViewNamespace::get_AutoSortItems,
	///       ShBrowserCtlsLibU::AutoSortItemsConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, ShellListViewNamespace::get_FullyQualifiedPIDL,
	///       NamespaceEnumSettings, ShellListView::get_DefaultNamespaceEnumSettings,
	///       ShellListView::get_DefaultManagedItemProperties, ShellListViewNamespace::get_AutoSortItems,
	///       ShBrowserCtlsLibA::AutoSortItemsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT pIDLOrParsingName, LONG insertAt = -1, INamespaceEnumSettings* pEnumerationSettings = NULL, AutoSortItemsConstants autoSortItems = asiAutoSortOnce, IShListViewNamespace** ppAddedNamespace = NULL);
	/// \brief <em>Retrieves whether the specified namespace is part of the namespace collection</em>
	///
	/// \param[in] namespaceIdentifier A value that identifies the namespace to be checked.
	/// \param[in] namespaceIdentifierType A value specifying the meaning of \c namespaceIdentifier. Any of
	///            the values defined by the \c ShLvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the namespace is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, ShBrowserCtlsLibU::ShLvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, ShBrowserCtlsLibA::ShLvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType = slnsitExactPIDL, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the namespaces in the collection</em>
	///
	/// Retrieves the number of \c ShellListViewNamespace objects in the collection.
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
	///            the values defined by the \c ShLvwNamespaceIdentifierTypeConstants enumeration is valid.
	/// \param[in] removeFromListView If \c VARIANT_TRUE, the items are removed from the listview control,
	///            too; otherwise they become normal listview items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibU::ShLvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ShBrowserCtlsLibA::ShLvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT namespaceIdentifier, ShLvwNamespaceIdentifierTypeConstants namespaceIdentifierType = slnsitExactPIDL, VARIANT_BOOL removeFromListView = VARIANT_TRUE);
	/// \brief <em>Removes all namespaces from the collection</em>
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
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellListView object that owns this collection</em>
		///
		/// \sa SetOwner
		ShellListView* pOwnerShLvw;
		#ifdef USE_STL
			/// \brief <em>Points to the next enumerated namespace</em>
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*>::const_iterator nextEnumeratedNamespace;
		#else
			/// \brief <em>Points to the next enumerated namespace</em>
			POSITION nextEnumeratedNamespace;
		#endif

		Properties();

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // ShellListViewNamespaces

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewNamespaces), ShellListViewNamespaces)