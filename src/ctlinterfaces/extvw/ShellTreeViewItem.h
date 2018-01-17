//////////////////////////////////////////////////////////////////////
/// \class ShellTreeViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell treeview item</em>
///
/// This class is a wrapper around a shell treeview item.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShTreeViewItem, ShellTreeViewItems, ShellTreeView
/// \else
///   \sa ShBrowserCtlsLibA::IShTreeViewItem, ShellTreeViewItems, ShellTreeView
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
#include "_IShellTreeViewItemEvents_CP.h"
#include "../../helpers.h"
#include "ShellTreeView.h"


class ATL_NO_VTABLE ShellTreeViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellTreeViewItem, &CLSID_ShellTreeViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellTreeViewItem>,
    public Proxy_IShTreeViewItemEvents<ShellTreeViewItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IShTreeViewItem, &IID_IShTreeViewItem, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShTreeViewItem, &IID_IShTreeViewItem, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellTreeView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLTREEVIEWITEM)

		BEGIN_COM_MAP(ShellTreeViewItem)
			COM_INTERFACE_ENTRY(IShTreeViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellTreeViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IShTreeViewItemEvents))
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
	/// \name Implementation of IShTreeViewItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Customizable property</em>
	///
	/// Retrieves whether the item can be customized via a Desktop.ini file. If \c VARIANT_TRUE, it can
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
	/// Retrieves the zero-based index of the item's default display column. This column is the one from
	/// which the shell treeview items should get their text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellTreeViewNamespace::get_DefaultDisplayColumnIndex, get_DefaultSortColumnIndex
	virtual HRESULT STDMETHODCALLTYPE get_DefaultDisplayColumnIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultSortColumnIndex property</em>
	///
	/// Retrieves the zero-based index of the item's default sort column. This column is the one by which
	/// the item's sub-items should be sorted initially.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellTreeViewNamespace::get_DefaultSortColumnIndex, get_DefaultDisplayColumnIndex
	virtual HRESULT STDMETHODCALLTYPE get_DefaultSortColumnIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayName property</em>
	///
	/// Retrieves the item's name as provided by the shell. For items having the \c isaCanBeRenamed shell
	/// attribute set, the \c dntDisplayName display name can be set.
	///
	/// \param[in] displayNameType Specifies the name to retrieve or set. Any of the values defined by the
	///            \c DisplayNameTypeConstants enumeration is valid.
	/// \param[in] relativeToDesktop If \c VARIANT_TRUE, the retrieved name is relative to the Desktop;
	///            otherwise it is relative to the parent shell item.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisplayName, get_ShellAttributes, get_LinkTarget,
	///       ShBrowserCtlsLibU::DisplayNameTypeConstants
	/// \else
	///   \sa put_DisplayName, get_ShellAttributes, get_LinkTarget,
	///       ShBrowserCtlsLibA::DisplayNameTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayName(DisplayNameTypeConstants displayNameType = dntDisplayName, VARIANT_BOOL relativeToDesktop = VARIANT_FALSE, BSTR* pValue = NULL);
	/// \brief <em>Sets the \c DisplayName property</em>
	///
	/// Retrieves the item's name as provided by the shell. For items having the \c isaCanBeRenamed shell
	/// attribute set, the \c dntDisplayName display name can be set.
	///
	/// \param[in] displayNameType Specifies the name to retrieve or set. Any of the values defined by the
	///            \c DisplayNameTypeConstants enumeration is valid.
	/// \param[in] relativeToDesktop If \c VARIANT_TRUE, the retrieved name is relative to the Desktop;
	///            otherwise it is relative to the parent shell item.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisplayName, get_ShellAttributes, ShBrowserCtlsLibU::DisplayNameTypeConstants
	/// \else
	///   \sa get_DisplayName, get_ShellAttributes, ShBrowserCtlsLibA::DisplayNameTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayName(DisplayNameTypeConstants displayNameType = dntDisplayName, VARIANT_BOOL relativeToDesktop = VARIANT_FALSE, BSTR newValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c FileAttributes property</em>
	///
	/// Retrieves the item's file attributes. Any combination of the values defined by the
	/// \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] mask Specifies the file attributes to check. Any combination of the values defined by
	///            the \c ItemFileAttributesConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property will fail for non-filesystem items.
	///
	/// \if UNICODE
	///   \sa put_FileAttributes, get_ShellAttributes, ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa put_FileAttributes, get_ShellAttributes, ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FileAttributes(ItemFileAttributesConstants mask = static_cast<ItemFileAttributesConstants>(0x17FB7), ItemFileAttributesConstants* pValue = NULL);
	/// \brief <em>Sets the \c FileAttributes property</em>
	///
	/// Sets the item's file attributes. Any combination of the values defined by the
	/// \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] mask Specifies the file attributes to set. Any combination of the values defined by
	///            the \c ItemFileAttributesConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property will fail for non-filesystem items.
	///
	/// \if UNICODE
	///   \sa get_FileAttributes, ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa get_FileAttributes, ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FileAttributes(ItemFileAttributesConstants mask = static_cast<ItemFileAttributesConstants>(0x17FB7), ItemFileAttributesConstants newValue = static_cast<ItemFileAttributesConstants>(0));
	/// \brief <em>Retrieves the current setting of the \c FullyQualifiedPIDL property</em>
	///
	/// Retrieves the fully qualified pIDL associated with this treeview item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The returned pIDL should NOT be freed!\n
	///          This is the default property of the \c IShTreeViewItem interface.\n
	///          This property is read-only.
	///
	/// \sa get_Handle
	virtual HRESULT STDMETHODCALLTYPE get_FullyQualifiedPIDL(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Handle property</em>
	///
	/// Retrieves a handle identifying this treeview item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A treeview item's handle will never change.\n
	///          This property is read-only.
	///
	/// \sa get_FullyQualifiedPIDL, get_TreeViewItemObject
	virtual HRESULT STDMETHODCALLTYPE get_Handle(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c InfoTipText property</em>
	///
	/// Retrieves the item's info tip text.
	///
	/// \param[in] flags A bit field influencing the info tip being returned. Any combination of the values
	///            defined by the \c InfoTipFlagsConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ShellTreeView::get_InfoTipFlags, ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa ShellTreeView::get_InfoTipFlags, ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InfoTipText(InfoTipFlagsConstants flags = static_cast<InfoTipFlagsConstants>(itfDefault | itfAllowSlowInfoTipFollowSysSettings), BSTR* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c LinkTarget property</em>
	///
	/// Retrieves the path to which the item is linking.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property raises an error if it is called for an item that isn't a link.\n
	///          This property is read-only.
	///
	/// \sa get_ShellAttributes, get_DisplayName, get_LinkTargetPIDL
	virtual HRESULT STDMETHODCALLTYPE get_LinkTarget(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c LinkTargetPIDL property</em>
	///
	/// Retrieves the pIDL to which the item is linking.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The returned pIDL MUST be freed by the caller!\n
	///          This property raises an error if it is called for an item that isn't a link.\n
	///          This property is read-only.
	///
	/// \sa get_ShellAttributes, get_LinkTarget
	virtual HRESULT STDMETHODCALLTYPE get_LinkTargetPIDL(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ManagedProperties property</em>
	///
	/// Retrieves a bit field specifying which of the treeview item's properties are managed by the
	/// \c ShellTreeView control rather than the treeview control/the client application. Any
	/// combination of the values defined by the \c ShTvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ManagedProperties, ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa put_ManagedProperties, ShellTreeView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ManagedProperties(ShTvwManagedItemPropertiesConstants* pValue);
	/// \brief <em>Sets the \c ManagedProperties property</em>
	///
	/// Sets a bit field specifying which of the treeview item's properties are managed by the
	/// \c ShellTreeView control rather than the treeview control/the client application. Any
	/// combination of the values defined by the \c ShTvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ManagedProperties, ShellTreeView::put_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa get_ManagedProperties, ShellTreeView::put_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ManagedProperties(ShTvwManagedItemPropertiesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Namespace property</em>
	///
	/// Retrieves the shell namespace that the item belongs to. If set to \c NULL, the item doesn't
	/// belong to any namespace.
	///
	/// \param[out] ppNamespace Receives the shell namespace object's \c IShTreeViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_TreeViewItemObject
	virtual HRESULT STDMETHODCALLTYPE get_Namespace(IShTreeViewNamespace** ppNamespace);
	/// \brief <em>Retrieves the current setting of the \c NamespaceEnumSettings property</em>
	///
	/// Retrieves the namespace enumeration settings used to enumerate this item's sub-items. Will be
	/// \c NULL, if the item's sub-items aren't managed by the control.
	///
	/// \param[out] ppEnumSettings Receives the enumeration settings object's \c INamespaceEnumSettings
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa INamespaceEnumSettings, ShellTreeViewNamespace::get_NamespaceEnumSettings
	virtual HRESULT STDMETHODCALLTYPE get_NamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings);
	/// \brief <em>Retrieves the current setting of the \c RequiresElevation property</em>
	///
	/// Retrieves whether opening this item requires elevated rights.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Windows Vista or newer.\n
	///          This property is read-only.
	///
	/// \sa ShellTreeView::get_DisplayElevationShieldOverlays
	virtual HRESULT STDMETHODCALLTYPE get_RequiresElevation(VARIANT_BOOL* pValue);
	#ifdef ACTIVATE_SECZONE_FEATURES
		/// \brief <em>Retrieves the current setting of the \c SecurityZone property</em>
		///
		/// Retrieves the Internet Explorer security zone the item belongs to.
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
	/// \brief <em>Retrieves the current setting of the \c ShellAttributes property</em>
	///
	/// Retrieves the item's shell attributes. Any combination of the values defined by the
	/// \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[in] mask Specifies the shell attributes to check. Any combination of the values defined by the
	///            \c ItemShellAttributesConstants enumeration is valid.
	/// \param[in] validate If \c VARIANT_FALSE, cached data may be used to handle the request; otherwise
	///            not.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_FileAttributes, ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa get_FileAttributes, ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ShellAttributes(ItemShellAttributesConstants mask = static_cast<ItemShellAttributesConstants>(0xFEFFE17F), VARIANT_BOOL validate = VARIANT_FALSE, ItemShellAttributesConstants* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c SupportsNewFolders property</em>
	///
	/// Retrieves whether the item supports the creation of new folders as sub-items. If \c VARIANT_TRUE, it
	/// does; otherwise it doesn't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_SupportsNewFolders(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c TreeViewItemObject property</em>
	///
	/// Retrieves the \c TreeViewItem object of this treeview item from the attached \c ExplorerTreeView
	/// control.
	///
	/// \param[out] ppItem Receives the object's \c IDispatch implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Handle, get_Namespace
	virtual HRESULT STDMETHODCALLTYPE get_TreeViewItemObject(IDispatch** ppItem);

	/// \brief <em>Creates the item's shell context menu</em>
	///
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellTreeView::CreateShellContextMenu,
	///     ShellTreeViewItems::CreateShellContextMenu, ShellTreeViewNamespace::CreateShellContextMenu,
	///     ShellTreeView::DestroyShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateShellContextMenu(OLE_HANDLE* pMenu);
	/// \brief <em>Opens the folder customization dialog for the item</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Customizable
	virtual HRESULT STDMETHODCALLTYPE Customize(void);
	/// \brief <em>Displays the item's shell context menu</em>
	///
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellTreeView::DisplayShellContextMenu,
	///     ShellTreeViewItems::DisplayShellContextMenu, ShellTreeViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Ensures that the item's sub-items have been loaded into the treeview if they are managed by \c ShellTreeView</em>
	///
	/// \c ShellTreeView loads items on demand, i. e. an item's sub-items are not inserted into the control
	/// until the item is expanded for the first time. If this method is called for an item that has not
	/// yet been expanded, it inserts the item's sub-items. Otherwise it does nothing.
	///
	/// \param[in] waitForLastItem If set to \c VARIANT_TRUE, the method doesn't return before all work is
	///            done. If set to \c VARIANT_FALSE, it may move the item enumeration to a background thread
	///            and return before all sub-items have been inserted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ManagedProperties
	virtual HRESULT STDMETHODCALLTYPE EnsureSubItemsAreLoaded(VARIANT_BOOL waitForLastItem = VARIANT_TRUE);
	/// \brief <em>Executes the default command from the item's shell context menu</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellTreeView::InvokeDefaultShellContextMenuCommand,
	///     ShellTreeViewItems::InvokeDefaultShellContextMenuCommand
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultShellContextMenuCommand(void);
	/// \brief <em>Checks whether the shell item wrapped by this object still exists</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if the shell item still exists; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeView::ValidateItem
	virtual HRESULT STDMETHODCALLTYPE Validate(VARIANT_BOOL* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given shell treeview item</em>
	///
	/// Attaches this object to a given shell treeview item, so that the item's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] itemHandle The treeview item to attach to.
	///
	/// \sa Detach
	void Attach(HTREEITEM itemHandle);
	/// \brief <em>Detaches this object from a shell treeview item</em>
	///
	/// Detaches this object from the shell treeview item it currently wraps, so that it doesn't wrap any
	/// shell treeview item anymore.
	///
	/// \sa Attach
	void Detach(void);
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
		/// \brief <em>The handle of the shell treeview item wrapped by this object</em>
		HTREEITEM itemHandle;

		Properties()
		{
			pOwnerShTvw = NULL;
			itemHandle = NULL;
		}

		~Properties();
	} properties;
};     // ShellTreeViewItem

OBJECT_ENTRY_AUTO(__uuidof(ShellTreeViewItem), ShellTreeViewItem)