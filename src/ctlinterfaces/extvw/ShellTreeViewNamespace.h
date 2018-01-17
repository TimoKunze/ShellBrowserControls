//////////////////////////////////////////////////////////////////////
/// \class ShellTreeViewNamespace
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell namespace</em>
///
/// This class is a wrapper around a shell namespace.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShTreeViewNamespace, ShellTreeViewNamespaces, ShellTreeView
/// \else
///   \sa ShBrowserCtlsLibA::IShTreeViewNamespace, ShellTreeViewNamespaces, ShellTreeView
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
#include "_IShellTreeViewNamespaceEvents_CP.h"
#include "../../helpers.h"
#include "ShellTreeView.h"


class ATL_NO_VTABLE ShellTreeViewNamespace : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellTreeViewNamespace, &CLSID_ShellTreeViewNamespace>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellTreeViewNamespace>,
    public Proxy_IShTreeViewNamespaceEvents<ShellTreeViewNamespace>,
    #ifdef UNICODE
    	public IDispatchImpl<IShTreeViewNamespace, &IID_IShTreeViewNamespace, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShTreeViewNamespace, &IID_IShTreeViewNamespace, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellTreeView;
	friend class ShellTreeViewNamespaces;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLTREEVIEWNAMESPACE)

		BEGIN_COM_MAP(ShellTreeViewNamespace)
			COM_INTERFACE_ENTRY(IShTreeViewNamespace)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellTreeViewNamespace)
			CONNECTION_POINT_ENTRY(__uuidof(_IShTreeViewNamespaceEvents))
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
	/// \name Implementation of IShTreeViewNamespace
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AutoSortItems property</em>
	///
	/// Retrieves whether the sub-tree represented by this namespace shall be sorted automatically. If this
	/// property is set to \c VARIANT_TRUE, the sub-tree is sorted automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On automatic sorting the items are sorted like in Windows Explorer.\n
	///          Shell items are inserted in front of other items.
	///
	/// \sa put_AutoSortItems, get_Items, ShellTreeView::get_ItemTypeSortOrder
	virtual HRESULT STDMETHODCALLTYPE get_AutoSortItems(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoSortItems property</em>
	///
	/// Sets whether the sub-tree represented by this namespace shall be sorted automatically. If this
	/// property is set to \c VARIANT_TRUE, the sub-tree is sorted automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On automatic sorting the items are sorted like in Windows Explorer.\n
	///          Shell items are inserted in front of other items.
	///
	/// \sa get_AutoSortItems, get_Items, ShellTreeView::put_ItemTypeSortOrder
	virtual HRESULT STDMETHODCALLTYPE put_AutoSortItems(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Customizable property</em>
	///
	/// Retrieves whether the namespace can be customized via a Desktop.ini file. If \c VARIANT_TRUE, it can
	/// be customized; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Customize
	virtual HRESULT STDMETHODCALLTYPE get_Customizable(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultDisplayColumnIndex property</em>
	///
	/// Retrieves the zero-based index of the namespace's default display column. This column is the one
	/// that should be displayed if only one column is displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellTreeViewItem::get_DefaultDisplayColumnIndex, get_DefaultSortColumnIndex
	virtual HRESULT STDMETHODCALLTYPE get_DefaultDisplayColumnIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultSortColumnIndex property</em>
	///
	/// Retrieves the zero-based index of the namespace's default sort column. This column is the one by
	/// which the namespace's items should be sorted initially.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellTreeViewItem::get_DefaultSortColumnIndex, get_DefaultDisplayColumnIndex
	virtual HRESULT STDMETHODCALLTYPE get_DefaultSortColumnIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c FullyQualifiedPIDL property</em>
	///
	/// Retrieves the fully qualified pIDL associated with this namespace.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The returned pIDL should NOT be freed!\n
	///          This is the default property of the \c IShTreeViewNamespace interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, get_ParentItemHandle, ShBrowserCtlsLibU::ShTvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa get_Index, get_ParentItemHandle, ShBrowserCtlsLibA::ShTvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FullyQualifiedPIDL(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this namespace.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Adding or removing namespaces changes other namespaces' indexes.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_FullyQualifiedPIDL, ShBrowserCtlsLibU::ShTvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa get_FullyQualifiedPIDL, ShBrowserCtlsLibA::ShTvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Items property</em>
	///
	/// Retrieves a collection object wrapping the shell treeview items belonging to this namespace.
	///
	/// \param[out] ppItems Receives the collection object's \c IShTreeViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellTreeView::get_TreeItems, ShellTreeViewItems
	virtual HRESULT STDMETHODCALLTYPE get_Items(IShTreeViewItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c NamespaceEnumSettings property</em>
	///
	/// Retrieves the namespace enumeration settings used to enumerate items within this namespace.
	///
	/// \param[out] ppEnumSettings Receives the enumeration settings object's \c INamespaceEnumSettings
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa NamespaceEnumSettings, ShellTreeViewItem::get_NamespaceEnumSettings
	virtual HRESULT STDMETHODCALLTYPE get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings);
	/// \brief <em>Retrieves the current setting of the \c ParentItemHandle property</em>
	///
	/// Retrieves a handle identifying the treeview item that is the namespace's parent item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A treeview item's handle will never change.\n
	///          This property is read-only.
	///
	/// \sa get_FullyQualifiedPIDL, get_ParentTreeViewItemObject
	virtual HRESULT STDMETHODCALLTYPE get_ParentItemHandle(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c TreeViewItemObject property</em>
	///
	/// Retrieves the \c TreeViewItem object of the treeview item specified by the \c Handle property from
	/// the attached \c ExplorerTreeView control.
	///
	/// \param[out] ppItem Receives the object's \c IDispatch implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ParentItemHandle
	virtual HRESULT STDMETHODCALLTYPE get_ParentTreeViewItemObject(IDispatch** ppItem);
	#ifdef ACTIVATE_SECZONE_FEATURES
		/// \brief <em>Retrieves the current setting of the \c SecurityZone property</em>
		///
		/// Retrieves the Internet Explorer security zone the namespace belongs to.
		///
		/// \param[out] ppSecurityZone Receives the security zone object's \c ISecurityZone implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		///
		/// \sa SecurityZone
		virtual HRESULT STDMETHODCALLTYPE get_SecurityZone(ISecurityZone** ppSecurityZone);
	#endif
	/// \brief <em>Retrieves the current setting of the \c SupportsNewFolders property</em>
	///
	/// Retrieves whether the namespace supports the creation of new folders as sub-items. If
	/// \c VARIANT_TRUE, it does; otherwise it doesn't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_SupportsNewFolders(VARIANT_BOOL* pValue);

	/// \brief <em>Creates the namespace's background shell context menu</em>
	///
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellTreeView::CreateShellContextMenu,
	///     ShellTreeViewItem::CreateShellContextMenu, ShellTreeViewItems::CreateShellContextMenu,
	///     ShellTreeView::DestroyShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateShellContextMenu(OLE_HANDLE* pMenu);
	/// \brief <em>Opens the folder customization dialog for the namespace</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Customizable
	virtual HRESULT STDMETHODCALLTYPE Customize(void);
	/// \brief <em>Displays the namespace's background shell context menu</em>
	///
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellTreeView::DisplayShellContextMenu,
	///     ShellTreeViewItem::DisplayShellContextMenu, ShellTreeViewItems::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Sorts the attached control's items</em>
	///
	/// Sorts the sub-tree represented by this namespace.\n
	/// \c TreeViewItem::SortSubItems is called with the following parameters:
	/// - \c firstCriterion: \c sobShell
	/// - \c secondCriterion: \c sobText
	/// - \c thirdCriterion: \c sobCustom
	/// - \c fourthCriterion: \c sobNone
	/// - \c fifthCriterion: \c sobNone
	/// - \c recurse: As set using the \c recurse parameter.
	/// - \c caseSensitive: \c VARIANT_FALSE
	///
	/// \param[in] recurse If set to \c VARIANT_FALSE, only the namespace parent item's immediate sub-items
	///            will be sorted; otherwise all sub-items of the namespace will be sorted recursively.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoSortItems, get_Items, ShellTreeView::get_ItemTypeSortOrder
	virtual HRESULT STDMETHODCALLTYPE SortItems(VARIANT_BOOL recurse = VARIANT_FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given shell namespace</em>
	///
	/// Attaches this object to a given shell namespace, so that the namespace's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace to attach to.
	/// \param[in] hRootItem The treeview item that is the root item of the namespace to attach to.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within this namespace.
	///
	/// \sa Detach
	void Attach(PCIDLIST_ABSOLUTE pIDL, HTREEITEM hRootItem, INamespaceEnumSettings* pEnumSettings);
	/// \brief <em>Detaches this object from a shell namespace</em>
	///
	/// Detaches this object from the shell namespace it currently wraps, so that it doesn't wrap any shell
	/// namespace anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications</em>
	///
	/// Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications which are sent if a directory (file) has
	/// been created. If the directory (file) is an immediate child of the namespace, the owner control is
	/// told to insert the shell item into the treeview (applying any filters that might prevent insertion).
	///
	/// \param[in] simplePIDLAddedDir The (simple) pIDL of the created directory (file).
	///
	/// \return Ignored.
	///
	/// \sa ShellTreeView::OnSHChangeNotify_CREATE
	LRESULT OnSHChangeNotify_CREATE(__in PCIDLIST_ABSOLUTE simplePIDLAddedDir);
	/// \brief <em>Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications</em>
	///
	/// Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications which are sent if a directory (file) has
	/// been deleted. If the namespace is equal or a child of the directory (file), the namespace deletes
	/// itself from the treeview.
	///
	/// \param[in] pIDLRemovedDir The pIDL of the deleted directory (file).
	///
	/// \return Ignored.
	///
	/// \sa ShellTreeView::OnSHChangeNotify_DELETE
	LRESULT OnSHChangeNotify_DELETE(__in PCIDLIST_ABSOLUTE pIDLRemovedDir);
	/// \brief <em>Reenumerates the namespace's sub-items and inserts any new ones into the treeview</em>
	///
	/// Starts a background enumeration of the namespace's immediate sub-items. Any new items not already
	/// existing within the treeview will be inserted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeView::OnSHChangeNotify_CREATE
	HRESULT UpdateEnumeration(void);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShTvw
	void SetOwner(__in_opt ShellTreeView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellTreeView object that owns this item</em>
		///
		/// \sa SetOwner
		ShellTreeView* pOwnerShTvw;
		/// \brief <em>Holds the \c AutoSortItems property's setting</em>
		///
		/// \sa get_AutoSortItems, put_AutoSortItems
		UINT autoSortItems : 1;
		/// \brief <em>Holds the \c NamespaceEnumSettings property's setting</em>
		///
		/// \sa get_NamespaceEnumSettings
		INamespaceEnumSettings* pNamespaceEnumSettings;
		/// \brief <em>The fully qualified pIDL of the namespace wrapped by this object</em>
		///
		/// \sa get_FullyQualifiedPIDL
		PCIDLIST_ABSOLUTE pIDL;
		/// \brief <em>The handle of the treeview item being the root item of the namespace wrapped by this object</em>
		///
		/// \sa get_ParentItemHandle, get_ParentTreeViewItemObject
		HTREEITEM hRootItem;

		Properties()
		{
			pOwnerShTvw = NULL;
			autoSortItems = FALSE;
			pNamespaceEnumSettings = NULL;
			pIDL = NULL;
			hRootItem = NULL;
		}

		~Properties();
	} properties;
};     // ShellTreeViewNamespace

OBJECT_ENTRY_AUTO(__uuidof(ShellTreeViewNamespace), ShellTreeViewNamespace)