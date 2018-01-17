//////////////////////////////////////////////////////////////////////
/// \class ShellListViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell listview item</em>
///
/// This class is a wrapper around a shell listview item.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListViewItem, ShellListViewItems, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShListViewItem, ShellListViewItems, ShellListView
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
#include "_IShellListViewItemEvents_CP.h"
#include "../../helpers.h"
#include "ShellListView.h"


class ATL_NO_VTABLE ShellListViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ShellListViewItem, &CLSID_ShellListViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListViewItem>,
    public Proxy_IShListViewItemEvents<ShellListViewItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IShListViewItem, &IID_IShListViewItem, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IShListViewItem, &IID_IShListViewItem, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ShellListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEWITEM)

		BEGIN_COM_MAP(ShellListViewItem)
			COM_INTERFACE_ENTRY(IShListViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewItemEvents))
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
	/// \name Implementation of IShListViewItem
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
	/// Retrieves the zero-based index of the item's default display column. This column is the one that
	/// should be displayed if only one column is displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ShellListView::get_Columns, ShellListViewNamespace::get_DefaultDisplayColumnIndex,
	///     get_DefaultSortColumnIndex
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
	/// \sa ShellListView::get_Columns, ShellListViewNamespace::get_DefaultSortColumnIndex,
	///     get_DefaultDisplayColumnIndex
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
	/// Retrieves the fully qualified pIDL associated with this listview item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The returned pIDL should NOT be freed!\n
	///          This is the default property of the \c IShListViewItem interface.\n
	///          This property is read-only.
	///
	/// \sa get_ID
	virtual HRESULT STDMETHODCALLTYPE get_FullyQualifiedPIDL(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this listview item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A listview item's ID will never change.\n
	///          This property is read-only.
	///
	/// \sa get_FullyQualifiedPIDL, get_ListViewItemObject
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
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
	///   \sa ShellListView::get_InfoTipFlags, ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa ShellListView::get_InfoTipFlags, ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InfoTipText(InfoTipFlagsConstants flags = static_cast<InfoTipFlagsConstants>(itfDefault | itfNoInfoTipFollowSystemSettings | itfAllowSlowInfoTipFollowSysSettings), BSTR* pValue = NULL);
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
	/// \brief <em>Retrieves the current setting of the \c ListViewItemObject property</em>
	///
	/// Retrieves the \c ListViewItem object of this listview item from the attached \c ExplorerListView
	/// control.
	///
	/// \param[out] ppItem Receives the object's \c IDispatch implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ID, get_Namespace
	virtual HRESULT STDMETHODCALLTYPE get_ListViewItemObject(IDispatch** ppItem);
	/// \brief <em>Retrieves the current setting of the \c ManagedProperties property</em>
	///
	/// Retrieves a bit field specifying which of the listview item's properties are managed by the
	/// \c ShellListView control rather than the listview control/the client application. Any
	/// combination of the values defined by the \c ShLvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ManagedProperties, ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa put_ManagedProperties, ShellListView::get_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ManagedProperties(ShLvwManagedItemPropertiesConstants* pValue);
	/// \brief <em>Sets the \c ManagedProperties property</em>
	///
	/// Sets a bit field specifying which of the listview item's properties are managed by the
	/// \c ShellListView control rather than the listview control/the client application. Any
	/// combination of the values defined by the \c ShLvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ManagedProperties, ShellListView::put_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa get_ManagedProperties, ShellListView::put_DefaultManagedItemProperties,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ManagedProperties(ShLvwManagedItemPropertiesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Namespace property</em>
	///
	/// Retrieves the shell namespace that the item belongs to. If set to \c NULL, the item doesn't
	/// belong to any namespace.
	///
	/// \param[out] ppNamespace Receives the shell namespace object's \c IShListViewNamespace implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ListViewItemObject
	virtual HRESULT STDMETHODCALLTYPE get_Namespace(IShListViewNamespace** ppNamespace);
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
	/// \sa ShellListView::get_DisplayElevationShieldOverlays
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

	/// \brief <em>Creates the item's shell context menu</em>
	///
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellListView::CreateShellContextMenu,
	///     ShellListViewItems::CreateShellContextMenu, ShellListViewNamespace::CreateShellContextMenu,
	///     ShellListView::DestroyShellContextMenu
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
	/// \sa CreateShellContextMenu, ShellListView::DisplayShellContextMenu,
	///     ShellListViewItems::DisplayShellContextMenu, ShellListViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Executes the default command from the item's shell context menu</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellListView::InvokeDefaultShellContextMenuCommand,
	///     ShellListViewItems::InvokeDefaultShellContextMenuCommand
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultShellContextMenuCommand(void);
	/// \brief <em>Checks whether the shell item wrapped by this object still exists</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if the shell item still exists; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Validate(VARIANT_BOOL* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given shell listview item</em>
	///
	/// Attaches this object to a given shell listview item, so that the item's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] itemID The listview item to attach to.
	///
	/// \sa Detach
	void Attach(LONG itemID);
	/// \brief <em>Detaches this object from a shell listview item</em>
	///
	/// Detaches this object from the shell listview item it currently wraps, so that it doesn't wrap any
	/// shell listview item anymore.
	///
	/// \sa Attach
	void Detach(void);
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
		/// \brief <em>The unique ID of the shell listview item wrapped by this object</em>
		LONG itemID;

		Properties()
		{
			pOwnerShLvw = NULL;
			itemID = -1;
		}

		~Properties();
	} properties;
};     // ShellListViewItem

OBJECT_ENTRY_AUTO(__uuidof(ShellListViewItem), ShellListViewItem)