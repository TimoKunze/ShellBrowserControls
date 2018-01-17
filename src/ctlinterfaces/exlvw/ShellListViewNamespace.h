//////////////////////////////////////////////////////////////////////
/// \class ShellListViewNamespace
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell namespace</em>
///
/// This class is a wrapper around a shell namespace.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewNamespace, ShellListViewNamespaces, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewNamespace, ShellListViewNamespaces, ShellListView
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
#include "_IShellListViewNamespaceEvents_CP.h"
#include "../../helpers.h"
#ifdef ACTIVATE_SECZONE_FEATURES
	#include "../SecurityZones.h"
	#include "../SecurityZone.h"
#endif
#include "ShellListView.h"


class ATL_NO_VTABLE ShellListViewNamespace : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewNamespace, &CLSID_ShellListViewNamespace>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewNamespace>,
    public Proxy_IShListViewNamespaceEvents<ShellListViewNamespace>,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewNamespace, &IID_IShListViewNamespace, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewNamespace, &IID_IShListViewNamespace, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;
	friend class ShellListViewNamespaces;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWNAMESPACE)

		BEGIN_COM_MAP(ShellListViewNamespace)
			COM_INTERFACE_ENTRY(IShListViewNamespace)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewNamespace)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewNamespaceEvents))
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
	/// \name Implementation of IShListViewNamespace
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AutoSortItems property</em>
	///
	/// Retrieves whether the namespace triggers item sorting automatically. Any of the values defined
	/// by the \c AutoSortItemsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On automatic sorting the items are sorted like in Windows Explorer.\n
	///          Automatic sorting sorts by the namespace's default sort column until the column is changed
	///          by manually calling \c IShListView::SortItems.\n
	///          If a namespace triggers a sorting action, the whole attached control is sorted instead
	///          of only the shell items that are children of the namespace.
	///
	/// \if UNICODE
	///   \sa put_AutoSortItems, get_DefaultSortColumnIndex, ShellListView::get_ListItems,
	///       ShellListView::get_ItemTypeSortOrder, ShBrowserCtlsLibU::AutoSortItemsConstants
	/// \else
	///   \sa put_AutoSortItems, get_DefaultSortColumnIndex, ShellListView::get_ListItems,
	///       ShellListView::get_ItemTypeSortOrder, ShBrowserCtlsLibA::AutoSortItemsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoSortItems(AutoSortItemsConstants* pValue);
	/// \brief <em>Sets the \c AutoSortItems property</em>
	///
	/// Sets whether the namespace triggers item sorting automatically. Any of the values defined
	/// by the \c AutoSortItemsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On automatic sorting the items are sorted like in Windows Explorer.\n
	///          Automatic sorting sorts by the namespace's default sort column until the column is changed
	///          by manually calling \c IShListView::SortItems.\n
	///          If a namespace triggers a sorting action, the whole attached control is sorted instead
	///          of only the shell items that are children of the namespace.
	///
	/// \if UNICODE
	///   \sa get_AutoSortItems, get_DefaultSortColumnIndex, ShellListView::get_ListItems,
	///       ShellListView::put_ItemTypeSortOrder, ShBrowserCtlsLibU::AutoSortItemsConstants
	/// \else
	///   \sa get_AutoSortItems, get_DefaultSortColumnIndex, ShellListView::get_ListItems,
	///       ShellListView::put_ItemTypeSortOrder, ShBrowserCtlsLibA::AutoSortItemsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AutoSortItems(AutoSortItemsConstants newValue);
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
	/// \sa ShellListView::get_Columns, ShellListViewItem::get_DefaultDisplayColumnIndex,
	///     get_DefaultSortColumnIndex
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
	/// \sa ShellListView::get_Columns, ShellListViewItem::get_DefaultSortColumnIndex,
	///     get_DefaultDisplayColumnIndex
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
	///          This is the default property of the \c IShListViewNamespace interface.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, ShBrowserCtlsLibU::ShLvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa get_Index, ShBrowserCtlsLibA::ShLvwNamespaceIdentifierTypeConstants
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
	///   \sa get_FullyQualifiedPIDL, ShBrowserCtlsLibU::ShLvwNamespaceIdentifierTypeConstants
	/// \else
	///   \sa get_FullyQualifiedPIDL, ShBrowserCtlsLibA::ShLvwNamespaceIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Items property</em>
	///
	/// Retrieves a collection object wrapping the shell listview items belonging to this namespace.
	///
	/// \param[out] ppItems Receives the collection object's \c IShListViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellListView::get_ListItems, ShellListViewItems
	virtual HRESULT STDMETHODCALLTYPE get_Items(IShListViewItems** ppItems);
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
	/// \sa NamespaceEnumSettings
	virtual HRESULT STDMETHODCALLTYPE get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings);
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
	/// \sa DisplayShellContextMenu, ShellListView::CreateShellContextMenu,
	///     ShellListViewItem::CreateShellContextMenu, ShellListViewItems::CreateShellContextMenu,
	///     ShellListView::DestroyShellContextMenu
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
	/// \sa CreateShellContextMenu, ShellListView::DisplayShellContextMenu,
	///     ShellListViewItem::DisplayShellContextMenu, ShellListViewItems::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given shell namespace</em>
	///
	/// Attaches this object to a given shell namespace, so that the namespace's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace to attach to.
	/// \param[in] pEnumSettings A \c INamespaceEnumSettings implementation specifying the enumeration
	///            settings to use when enumerating items within this namespace.
	///
	/// \sa Detach
	void Attach(PCIDLIST_ABSOLUTE pIDL, INamespaceEnumSettings* pEnumSettings);
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
	/// told to insert the shell item into the listview (applying any filters that might prevent insertion).
	///
	/// \param[in] simplePIDLAddedDir The (simple) pIDL of the created directory (file).
	///
	/// \return Ignored.
	///
	/// \sa ShellListView::OnSHChangeNotify_CREATE
	LRESULT OnSHChangeNotify_CREATE(__in PCIDLIST_ABSOLUTE simplePIDLAddedDir);
	/// \brief <em>Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications</em>
	///
	/// Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications which are sent if a directory (file) has
	/// been deleted. If the namespace is equal or a child of the directory (file), the namespace deletes
	/// itself from the listview.
	///
	/// \param[in] pIDLRemovedDir The pIDL of the deleted directory (file).
	///
	/// \return Ignored.
	///
	/// \sa ShellListView::OnSHChangeNotify_DELETE
	LRESULT OnSHChangeNotify_DELETE(__in PCIDLIST_ABSOLUTE pIDLRemovedDir);
	/// \brief <em>Reenumerates the namespace's sub-items and inserts any new ones into the listview</em>
	///
	/// Starts a background enumeration of the namespace's immediate sub-items. Any new items not already
	/// existing within the listview will be inserted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListView::OnSHChangeNotify_CREATE
	HRESULT UpdateEnumeration(void);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerShLvw
	void SetOwner(__in_opt ShellListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ShellListView object that owns this item</em>
		///
		/// \sa SetOwner
		ShellListView* pOwnerShLvw;
		/// \brief <em>Holds the \c AutoSortItems property's setting</em>
		///
		/// \sa get_AutoSortItems, put_AutoSortItems
		AutoSortItemsConstants autoSortItems;
		/// \brief <em>Holds the \c NamespaceEnumSettings property's setting</em>
		///
		/// \sa get_NamespaceEnumSettings
		INamespaceEnumSettings* pNamespaceEnumSettings;
		/// \brief <em>The fully qualified pIDL of the namespace wrapped by this object</em>
		///
		/// \sa get_FullyQualifiedPIDL
		PCIDLIST_ABSOLUTE pIDL;
		/// \brief <em>The \c IShellImageStore object used to access the namespace's thumbnail disk cache</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa IShellImageStore
		LPSHELLIMAGESTORE pThumbnailDiskCache;
		/// \brief <em>Holds the time stamp at which the thumbnail disk cache specified by \c pThumbnailDiskCache has been accessed the last time</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa pThumbnailDiskCache
		DWORD lastThumbnailDiskCacheAccessTime;
		/// \brief <em>If \c TRUE, the thumbnail disk cache specified by \c pThumbnailDiskCache is currently opened</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa pThumbnailDiskCache
		UINT thumbnailDiskCacheIsOpened : 1;

		Properties()
		{
			pOwnerShLvw = NULL;
			autoSortItems = asiNoAutoSort;
			pNamespaceEnumSettings = NULL;
			pIDL = NULL;
			pThumbnailDiskCache = NULL;
			lastThumbnailDiskCacheAccessTime = 0;
			thumbnailDiskCacheIsOpened = FALSE;
		}

		~Properties();
	} properties;
};     // ShellListViewNamespace

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewNamespace), ShellListViewNamespace)