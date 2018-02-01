//////////////////////////////////////////////////////////////////////
/// \class ShellListView
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Makes an \c ExplorerListView a shell listview</em>
///
/// This class makes an \c ExplorerListView control a shell listview control.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShListView, ShellTreeView
/// \else
///   \sa ShBrowserCtlsLibA::IShListView, ShellTreeView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../res/resource.h"
#ifdef UNICODE
	#include "../../ShBrowserCtlsU.h"
#else
	#include "../../ShBrowserCtlsA.h"
#endif
#include "../../CWindowEx2.h"
#include "definitions.h"
#include "SetColumnsDlg.h"
#include "BackgroundColumnEnumTask.h"
#include "InsertSingleItemTask.h"
#include "IconTask.h"
#include "InfoTipTask.h"
#include "LegacyThumbnailTask.h"
#include "LegacyTileViewTask.h"
#include "OverlayTask.h"
#include "SlowColumnTask.h"
#include "SubItemControlTask.h"
#include "ThumbnailTask.h"
#include "TileViewTask.h"
#include "../ElevationShieldTask.h"
#include "_IShellListViewEvents_CP.h"
#include "ShLvwDefaultManagedItemPropertiesProperties.h"
#include "../../ICategorizeProperties.h"
#include "../../ICreditsProvider.h"
#include "../../helpers.h"
#include "../../EnumOLEVERB.h"
#include "../../AboutDlg.h"
#include "../../CommonProperties.h"
#include "../AggregateImageList.h"
#include "../UnifiedImageList.h"
#include "../INamespaceEnumTask.h"
#include "../IPersistentChildObject.h"
#include "../NamespaceEnumSettingsProperties.h"
#include "../definitions.h"
#include "../IMessageListener.h"
#include "../IInternalMessageListener.h"
#include "../NamespaceEnumSettings.h"
#ifdef ACTIVATE_SECZONE_FEATURES
	#include "../SecurityZones.h"
#endif
#include "ThumbnailsProperties.h"
#include "ShellListViewItems.h"
#include "ShellListViewNamespaces.h"
#include "ShellListViewColumns.h"


class ATL_NO_VTABLE ShellListView : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IShListView, &IID_IShListView, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IShListView, &IID_IShListView, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ShellListView>,
    public IOleControlImpl<ShellListView>,
    public IOleObjectImpl<ShellListView>,
    public IOleInPlaceActiveObjectImpl<ShellListView>,
    public IViewObjectExImpl<ShellListView>,
    public IOleInPlaceObjectWindowlessImpl<ShellListView>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellListView>,
    public Proxy_IShListViewEvents<ShellListView>,
    public IPersistStorageImpl<ShellListView>,
    public IPersistPropertyBagImpl<ShellListView>,
    public ISpecifyPropertyPages,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ShellListView, &__uuidof(_IShListViewEvents), &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ShellListView, &__uuidof(_IShListViewEvents), &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ShellListView>,
    public CComCoClass<ShellListView, &CLSID_ShellListView>,
    public CComControl<ShellListView>,
    public IPerPropertyBrowsingImpl<ShellListView>,
    public ICategorizeProperties,
    public ICreditsProvider,
    public IMessageListener,
    public IInternalMessageListener,
    public IContextMenuCB
{
	friend class SetColumnsDlg;
	friend class ShellListViewItems;
	friend class ShellListViewItem;
	friend class ShellListViewNamespaces;
	friend class ShellListViewNamespace;
	friend class ShellListViewColumns;
	friend class ShellListViewColumn;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa ~ShellListView, FinalConstruct
	ShellListView();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa ShellListView
	~ShellListView();
	/// \brief <em>This method is called directly after the constructor</em>
	///
	/// This method is called directly after the constructor. It is used for initialization.
	///
	/// \sa ShellListView
	HRESULT FinalConstruct();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_INVISIBLEATRUNTIME | OLEMISC_NOUIACTIVATE | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLLISTVIEW)

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ShellListView)
			COM_INTERFACE_ENTRY(IShListView)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IContextMenuCB)
		END_COM_MAP()

		BEGIN_PROP_MAP(ShellListView)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellListView)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IShListViewEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ShellListView)
			CHAIN_MSG_MAP(CComControl<ShellListView>)
		END_MSG_MAP()
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
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist some read-only properties. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist some read-only properties. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IShListView
	///
	//@{
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c AutoEditNewItems property</em>
	///
	/// If the control detects the creation of a new shell item, it can check how this item has been created.
	/// If the item has been created by selecting a menu item from the <em>New</em> sub menu of a namespace's
	/// background shell context menu, the client may want to enter label-edit mode for the new item so the
	/// user can rename it immediately. This method retrieves whether the control enters label-edit mode
	/// automatically in the described situation.\n
	/// If set to \c VARIANT_TRUE, it enters label-edit mode automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This feature won't work if the \c ProcessShellNotifications property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa put_AutoEditNewItems, get_ProcessShellNotifications
	virtual HRESULT STDMETHODCALLTYPE get_AutoEditNewItems(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoEditNewItems property</em>
	///
	/// If the control detects the creation of a new shell item, it can check how this item has been created.
	/// If the item has been created by selecting a menu item from the <em>New</em> sub menu of a namespace's
	/// background shell context menu, the client may want to enter label-edit mode for the new item so the
	/// user can rename it immediately. This method sets whether the control enters label-edit mode
	/// automatically in the described situation.\n
	/// If set to \c VARIANT_TRUE, it enters label-edit mode automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This feature won't work if the \c ProcessShellNotifications property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa get_AutoEditNewItems, put_ProcessShellNotifications
	virtual HRESULT STDMETHODCALLTYPE put_AutoEditNewItems(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoInsertColumns property</em>
	///
	/// \c ShellListView can setup the columns for details and tile view automatically. If this property is
	/// set to \c VARIANT_TRUE, the columns are inserted automatically; otherwise not. If the columns are
	/// inserted automatically, the columns of the first defined shell namespace are used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutoInsertColumns, get_Namespaces, get_PersistColumnSettingsAcrossNamespaces
	virtual HRESULT STDMETHODCALLTYPE get_AutoInsertColumns(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoInsertColumns property</em>
	///
	/// \c ShellListView can setup the columns for details and tile view automatically. If this property is
	/// set to \c VARIANT_TRUE, the columns are inserted automatically; otherwise not. If the columns are
	/// inserted automatically, the columns of the first defined shell namespace are used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoInsertColumns, get_Namespaces, put_PersistColumnSettingsAcrossNamespaces
	virtual HRESULT STDMETHODCALLTYPE put_AutoInsertColumns(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c ColorCompressedItems property</em>
	///
	/// Retrieves whether compressed items are displayed in another color. If set to \c VARIANT_TRUE,
	/// compressed items are displayed in the system's color for compressed items; otherwise the normal
	/// text color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ColorCompressedItems, get_ColorEncryptedItems, ShellListViewItem::get_ShellAttributes
	virtual HRESULT STDMETHODCALLTYPE get_ColorCompressedItems(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ColorCompressedItems property</em>
	///
	/// Sets whether compressed items are displayed in another color. If set to \c VARIANT_TRUE,
	/// compressed items are displayed in the system's color for compressed items; otherwise the normal
	/// text color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ColorCompressedItems, put_ColorEncryptedItems
	virtual HRESULT STDMETHODCALLTYPE put_ColorCompressedItems(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ColorEncryptedItems property</em>
	///
	/// Retrieves whether encrypted items are displayed in another color. If set to \c VARIANT_TRUE,
	/// encrypted items are displayed in the system's color for encrypted items; otherwise the normal
	/// text color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ColorEncryptedItems, get_ColorCompressedItems, ShellListViewItem::get_ShellAttributes
	virtual HRESULT STDMETHODCALLTYPE get_ColorEncryptedItems(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ColorEncryptedItems property</em>
	///
	/// Sets whether encrypted items are displayed in another color. If set to \c VARIANT_TRUE,
	/// encrypted items are displayed in the system's color for encrypted items; otherwise the normal
	/// text color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ColorEncryptedItems, put_ColorCompressedItems
	virtual HRESULT STDMETHODCALLTYPE put_ColorEncryptedItems(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ColumnEnumerationTimeout property</em>
	///
	/// Retrieves the number of milliseconds the enumeration of a namespace's columns may take before the
	/// \c ColumnEnumerationTimedOut event is raised. The value must be 1000 or greater. If this property is
	/// set to -1, the event isn't raised.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ColumnEnumerationTimeout, get_ItemEnumerationTimeout, get_Columns,
	///     get_PersistColumnSettingsAcrossNamespaces, Raise_ColumnEnumerationTimedOut
	virtual HRESULT STDMETHODCALLTYPE get_ColumnEnumerationTimeout(LONG* pValue);
	/// \brief <em>Sets the \c ColumnEnumerationTimeout property</em>
	///
	/// Sets the number of milliseconds the enumeration of a namespace's columns may take before the
	/// \c ColumnEnumerationTimedOut event is raised. The value must be 1000 or greater. If this property is
	/// set to -1, the event isn't raised.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ColumnEnumerationTimeout, put_ItemEnumerationTimeout, get_Columns,
	///     Raise_ColumnEnumerationTimedOut
	virtual HRESULT STDMETHODCALLTYPE put_ColumnEnumerationTimeout(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ShellListColumns property</em>
	///
	/// Retrieves a collection object wrapping the listview control's shell columns.
	///
	/// \param[out] ppColumns Receives the collection object's \c IShListViewColumns implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ColumnEnumerationTimeout, EnsureShellColumnsAreLoaded, get_ListItems, get_Namespaces,
	///     ShellListViewColumns
	virtual HRESULT STDMETHODCALLTYPE get_Columns(IShListViewColumns** ppColumns);
	/// \brief <em>Retrieves the current setting of the \c DefaultManagedItemProperties property</em>
	///
	/// Retrieves a bit field specifying which of the listview items' properties by default are managed by
	/// the \c ShellListView control rather than the listview control/the client application. Any combination
	/// of the values defined by the \c ShLvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The default settings are used for items that are inserted as part of a namespace or in
	///          response to shell notifications. Changing this property doesn't affect items that already
	///          have been inserted into the listview.
	///
	/// \if UNICODE
	///   \sa put_DefaultManagedItemProperties, get_ProcessShellNotifications,
	///       ShellListViewItem::get_ManagedProperties, ShellListViewNamespaces::Add,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa put_DefaultManagedItemProperties, get_ProcessShellNotifications,
	///       ShellListViewItem::get_ManagedProperties, ShellListViewNamespaces::Add,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DefaultManagedItemProperties(ShLvwManagedItemPropertiesConstants* pValue);
	/// \brief <em>Sets the \c DefaultManagedItemProperties property</em>
	///
	/// Sets a bit field specifying which of the listview items' properties by default are managed by
	/// the \c ShellListView control rather than the listview control/the client application. Any combination
	/// of the values defined by the \c ShLvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The default settings are used for items that are inserted as part of a namespace or in
	///          response to shell notifications. Changing this property doesn't affect items that already
	///          have been inserted into the listview.
	///
	/// \if UNICODE
	///   \sa get_DefaultManagedItemProperties, put_ProcessShellNotifications,
	///       ShellListViewItem::put_ManagedProperties, ShellListViewNamespaces::Add,
	///       ShBrowserCtlsLibU::ShLvwManagedItemPropertiesConstants
	/// \else
	///   \sa get_DefaultManagedItemProperties, put_ProcessShellNotifications,
	///       ShellListViewItem::put_ManagedProperties, ShellListViewNamespaces::Add,
	///       ShBrowserCtlsLibA::ShLvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DefaultManagedItemProperties(ShLvwManagedItemPropertiesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultNamespaceEnumSettings property</em>
	///
	/// Retrieves the control's default namespace enumeration settings.
	///
	/// \param[out] ppEnumSettings Receives the enumeration settings object's \c INamespaceEnumSettings
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_DefaultNamespaceEnumSettings, NamespaceEnumSettings,
	///     ShellListViewItem::get_NamespaceEnumSettings, ShellListViewNamespace::get_NamespaceEnumSettings
	virtual HRESULT STDMETHODCALLTYPE get_DefaultNamespaceEnumSettings(INamespaceEnumSettings** ppEnumSettings);
	/// \brief <em>Sets the \c DefaultNamespaceEnumSettings property</em>
	///
	/// Sets the control's default namespace enumeration settings.
	///
	/// \param[in] pEnumSettings The default enumeration settings to use.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DefaultNamespaceEnumSettings, NamespaceEnumSettings,
	///     ShellListViewItem::get_NamespaceEnumSettings, ShellListViewNamespace::get_NamespaceEnumSettings
	virtual HRESULT STDMETHODCALLTYPE putref_DefaultNamespaceEnumSettings(INamespaceEnumSettings* pEnumSettings);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, ShBrowserCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, ShBrowserCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, ShBrowserCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, ShBrowserCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayElevationShieldOverlays property</em>
	///
	/// Retrieves whether the elevation shield is displayed as overlay image, if the item requires elevation.
	/// If set to \c VARIANT_TRUE, the elevation shield is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Windows Vista or newer.\n
	///          If the item has a thumbnail image and requires elevation, the executable's icon is
	///          displayed, ignoring the setting of the \c DisplayFileTypeOverlays property.
	///
	/// \sa put_DisplayElevationShieldOverlays, get_DisplayThumbnails, get_DisplayThumbnailAdornments,
	///     ShellListViewItem::get_RequiresElevation, get_DisplayFileTypeOverlays
	virtual HRESULT STDMETHODCALLTYPE get_DisplayElevationShieldOverlays(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayElevationShieldOverlays property</em>
	///
	/// Sets whether the elevation shield is displayed as overlay image, if the item requires elevation.
	/// If set to \c VARIANT_TRUE, the elevation shield is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Windows Vista or newer.\n
	///          If the item has a thumbnail image and requires elevation, the executable's icon is
	///          displayed, ignoring the setting of the \c DisplayFileTypeOverlays property.
	///
	/// \sa get_DisplayElevationShieldOverlays, put_DisplayThumbnails, put_DisplayThumbnailAdornments,
	///     put_DisplayFileTypeOverlays
	virtual HRESULT STDMETHODCALLTYPE put_DisplayElevationShieldOverlays(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayFileTypeOverlays property</em>
	///
	/// Retrieves which kind of image is drawn over the thumbnail's bottom-right corner in thumbnail mode.
	/// Any of the values defined by the \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Those overlays are displayed only for items that have a thumbnail image.
	///
	/// \if UNICODE
	///   \sa put_DisplayFileTypeOverlays, get_DisplayThumbnails, get_DisplayThumbnailAdornments,
	///       get_DisplayElevationShieldOverlays, ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
	/// \else
	///   \sa put_DisplayFileTypeOverlays, get_DisplayThumbnails, get_DisplayThumbnailAdornments,
	///       get_DisplayElevationShieldOverlays, ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayFileTypeOverlays(ShLvwDisplayFileTypeOverlaysConstants* pValue);
	/// \brief <em>Sets the \c DisplayFileTypeOverlays property</em>
	///
	/// Sets which kind of image is drawn over the thumbnail's bottom-right corner in thumbnail mode.
	/// Any of the values defined by the \c ShLvwDisplayFileTypeOverlaysConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Those overlays are displayed only for items that have a thumbnail image.
	///
	/// \if UNICODE
	///   \sa get_DisplayFileTypeOverlays, put_DisplayThumbnails, put_DisplayThumbnailAdornments,
	///       put_DisplayElevationShieldOverlays, ShBrowserCtlsLibU::ShLvwDisplayFileTypeOverlaysConstants
	/// \else
	///   \sa get_DisplayFileTypeOverlays, put_DisplayThumbnails, put_DisplayThumbnailAdornments,
	///       put_DisplayElevationShieldOverlays, ShBrowserCtlsLibA::ShLvwDisplayFileTypeOverlaysConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayFileTypeOverlays(ShLvwDisplayFileTypeOverlaysConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayThumbnailAdornments property</em>
	///
	/// Retrieves which adornments are drawn to improve the visual appearance of thumbnail images. Any
	/// combination of the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
	/// valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisplayThumbnailAdornments, get_DisplayThumbnails, get_DisplayFileTypeOverlays,
	///       get_DisplayElevationShieldOverlays, ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants
	/// \else
	///   \sa put_DisplayThumbnailAdornments, get_DisplayThumbnails, get_DisplayFileTypeOverlays,
	///       get_DisplayElevationShieldOverlays, ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayThumbnailAdornments(ShLvwDisplayThumbnailAdornmentsConstants* pValue);
	/// \brief <em>Sets the \c DisplayThumbnailAdornments property</em>
	///
	/// Sets which adornments are drawn to improve the visual appearance of thumbnail images. Any
	/// combination of the values defined by the \c ShLvwDisplayThumbnailAdornmentsConstants enumeration is
	/// valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisplayThumbnailAdornments, put_DisplayThumbnails, put_DisplayFileTypeOverlays,
	///       put_DisplayElevationShieldOverlays, ShBrowserCtlsLibU::ShLvwDisplayThumbnailAdornmentsConstants
	/// \else
	///   \sa get_DisplayThumbnailAdornments, put_DisplayThumbnails, put_DisplayFileTypeOverlays,
	///       put_DisplayElevationShieldOverlays, ShBrowserCtlsLibA::ShLvwDisplayThumbnailAdornmentsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayThumbnailAdornments(ShLvwDisplayThumbnailAdornmentsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayThumbnails property</em>
	///
	/// Retrieves whether thumbnail images instead of normal icons are displayed. If set to \c VARIANT_TRUE,
	/// thumbnail images, otherwise normal icons are displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisplayThumbnails, get_ThumbnailSize, get_UseThumbnailDiskCache,
	///     get_DisplayThumbnailAdornments, get_DisplayFileTypeOverlays, get_DisplayElevationShieldOverlays
	virtual HRESULT STDMETHODCALLTYPE get_DisplayThumbnails(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayThumbnails property</em>
	///
	/// Sets whether thumbnail images instead of normal icons are displayed. If set to \c VARIANT_TRUE,
	/// thumbnail images, otherwise normal icons are displayed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DisplayThumbnails, put_ThumbnailSize, put_UseThumbnailDiskCache,
	///     put_DisplayThumbnailAdornments, put_DisplayFileTypeOverlays, put_DisplayElevationShieldOverlays
	virtual HRESULT STDMETHODCALLTYPE put_DisplayThumbnails(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HandleOLEDragDrop property</em>
	///
	/// Retrieves which parts of OLE drag'n'drop are handled automatically. Any combination of the values
	/// defined by the \c HandleOLEDragDropConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_HandleOLEDragDrop, ShBrowserCtlsLibU::HandleOLEDragDropConstants
	/// \else
	///   \sa put_HandleOLEDragDrop, ShBrowserCtlsLibA::HandleOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HandleOLEDragDrop(HandleOLEDragDropConstants* pValue);
	/// \brief <em>Sets the \c HandleOLEDragDrop property</em>
	///
	/// Sets which parts of OLE drag'n'drop are handled automatically. Any combination of the values
	/// defined by the \c HandleOLEDragDropConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_HandleOLEDragDrop, ShBrowserCtlsLibU::HandleOLEDragDropConstants
	/// \else
	///   \sa get_HandleOLEDragDrop, ShBrowserCtlsLibA::HandleOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HandleOLEDragDrop(HandleOLEDragDropConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HiddenItemsStyle property</em>
	///
	/// Retrieves the display style of hidden shell items. Any of the values defined by the
	/// \c HiddenItemsStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Setting this property to \c hisGhostedOnDemand increases performance, but the
	///          \c IListViewItem::Ghosted property always returns the <em>current</em> state which isn't
	///          necessarily correct until the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \if UNICODE
	///   \sa put_HiddenItemsStyle, ShBrowserCtlsLibU::HiddenItemsStyleConstants
	/// \else
	///   \sa put_HiddenItemsStyle, ShBrowserCtlsLibA::HiddenItemsStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HiddenItemsStyle(HiddenItemsStyleConstants* pValue);
	/// \brief <em>Sets the \c HiddenItemsStyle property</em>
	///
	/// Sets the display style of hidden shell items. Any of the values defined by the
	/// \c HiddenItemsStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Setting this property to \c hisGhostedOnDemand increases performance, but the
	///          \c IListViewItem::Ghosted property always returns the <em>current</em> state which isn't
	///          necessarily correct until the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \if UNICODE
	///   \sa get_HiddenItemsStyle, ShBrowserCtlsLibU::HiddenItemsStyleConstants
	/// \else
	///   \sa get_HiddenItemsStyle, ShBrowserCtlsLibA::HiddenItemsStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HiddenItemsStyle(HiddenItemsStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle of the specified image list.
	///
	/// \param[in] imageList The image list to retrieve. Any of the values defined by the
	///            \c ImageListConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa get_UseSystemImageList, ShBrowserCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_UseSystemImageList, ShBrowserCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle of the specified image list.
	///
	/// \param[in] imageList The image list to set. Any of the values defined by the \c ImageListConstants
	///            enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa put_UseSystemImageList, ShBrowserCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_UseSystemImageList, ShBrowserCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the handle of the \c SysListView32 window that the object is currently attached to.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWndShellUIParentWindow, Attach, Detach
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndShellUIParentWindow property</em>
	///
	/// Retrieves the handle of the window that is used as parent window for any UI that the shell may
	/// display.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c NULL, many UI isn't displayed at all.
	///
	/// \sa put_hWndShellUIParentWindow, get_hWnd, GethWndShellUIParentWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWndShellUIParentWindow(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndShellUIParentWindow property</em>
	///
	/// Sets the handle of the window that is used as parent window for any UI that the shell may display.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c NULL, many UI isn't displayed at all.
	///
	/// \sa get_hWndShellUIParentWindow, get_hWnd, GethWndShellUIParentWindow
	virtual HRESULT STDMETHODCALLTYPE put_hWndShellUIParentWindow(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c InfoTipFlags property</em>
	///
	/// Retrieves a bit field influencing the listview items info tips if they are managed by the
	/// \c ShellListView control. Any combination of the values defined by the \c InfoTipFlagsConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_InfoTipFlags, ShellListViewItem::get_ManagedProperties, ShellListViewItem::get_InfoTipText,
	///       ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa put_InfoTipFlags, ShellListViewItem::get_ManagedProperties, ShellListViewItem::get_InfoTipText,
	///       ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InfoTipFlags(InfoTipFlagsConstants* pValue);
	/// \brief <em>Sets the \c InfoTipFlags property</em>
	///
	/// Sets a bit field influencing the listview items info tips if they are managed by the
	/// \c ShellListView control. Any combination of the values defined by the \c InfoTipFlagsConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_InfoTipFlags, ShellListViewItem::put_ManagedProperties,
	///       ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa get_InfoTipFlags, ShellListViewItem::put_ManagedProperties,
	///       ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_InfoTipFlags(InfoTipFlagsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c InitialSortOrder property</em>
	///
	/// Retrieves the sort order initially used when loading a new shell namespace. Any of the values defined
	/// by the \c SortOrderConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_InitialSortOrder, SortItems, ShBrowserCtlsLibU::SortOrderConstants
	/// \else
	///   \sa put_InitialSortOrder, SortItems, ShBrowserCtlsLibA::SortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InitialSortOrder(SortOrderConstants* pValue);
	/// \brief <em>Sets the \c InitialSortOrder property</em>
	///
	/// Sets the sort order initially used when loading a new shell namespace. Any of the values defined
	/// by the \c SortOrderConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_InitialSortOrder, SortItems, ShBrowserCtlsLibU::SortOrderConstants
	/// \else
	///   \sa get_InitialSortOrder, SortItems, ShBrowserCtlsLibA::SortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_InitialSortOrder(SortOrderConstants newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemEnumerationTimeout property</em>
	///
	/// Retrieves the number of milliseconds the enumeration of a namespace's items may take before the
	/// \c ItemEnumerationTimedOut event is raised. The value must be 1000 or greater. If this property is
	/// set to -1, the event isn't raised.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ItemEnumerationTimeout, get_ColumnEnumerationTimeout, Raise_ItemEnumerationTimedOut
	virtual HRESULT STDMETHODCALLTYPE get_ItemEnumerationTimeout(LONG* pValue);
	/// \brief <em>Sets the \c ItemEnumerationTimeout property</em>
	///
	/// Sets the number of milliseconds the enumeration of a namespace's items may take before the
	/// \c ItemEnumerationTimedOut event is raised. The value must be 1000 or greater. If this property is
	/// set to -1, the event isn't raised.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemEnumerationTimeout, put_ColumnEnumerationTimeout, Raise_ItemEnumerationTimedOut
	virtual HRESULT STDMETHODCALLTYPE put_ItemEnumerationTimeout(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemTypeSortOrder property</em>
	///
	/// Retrieves the order of the different kinds of items (shell items, normal items) within the attached
	/// listview control. This order is applied when sorting items. Any of the values defined by the
	/// \c ItemTypeSortOrderConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ItemTypeSortOrder, ShellListViewNamespace::get_AutoSortItems, SortItems,
	///       ShBrowserCtlsLibU::ItemTypeSortOrderConstants
	/// \else
	///   \sa put_ItemTypeSortOrder, ShellListViewNamespace::get_AutoSortItems, SortItems,
	///       ShBrowserCtlsLibA::ItemTypeSortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ItemTypeSortOrder(ItemTypeSortOrderConstants* pValue);
	/// \brief <em>Sets the \c ItemTypeSortOrder property</em>
	///
	/// Sets the order of the different kinds of items (shell items, normal items) within the attached
	/// listview control. This order is applied when sorting items. Any of the values defined by the
	/// \c ItemTypeSortOrderConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ItemTypeSortOrder, ShellListViewNamespace::put_AutoSortItems, SortItems,
	///       ShBrowserCtlsLibU::ItemTypeSortOrderConstants
	/// \else
	///   \sa get_ItemTypeSortOrder, ShellListViewNamespace::put_AutoSortItems, SortItems,
	///       ShBrowserCtlsLibA::ItemTypeSortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ItemTypeSortOrder(ItemTypeSortOrderConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c LimitLabelEditInput property</em>
	///
	/// Retrieves whether characters, that are not allowed in the affected shell item's name, are recognized
	/// and declined at input time. If set to \c VARIANT_TRUE, invalid characters are declined immediately
	/// and an info tip is displayed explaining that the character was invalid. If set to \c VARIANT_FALSE,
	/// no characters are declined, but later the actual renaming will fail and an error message will be
	/// displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LimitLabelEditInput, ShellListViewItem::get_ManagedProperties,
	///     get_PreselectBasenameOnLabelEdit
	virtual HRESULT STDMETHODCALLTYPE get_LimitLabelEditInput(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c LimitLabelEditInput property</em>
	///
	/// Sets whether characters, that are not allowed in the affected shell item's name, are recognized
	/// and declined at input time. If set to \c VARIANT_TRUE, invalid characters are declined immediately
	/// and an info tip is displayed explaining that the character was invalid. If set to \c VARIANT_FALSE,
	/// no characters are declined, but later the actual renaming will fail and an error message will be
	/// displayed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LimitLabelEditInput, ShellListViewItem::put_ManagedProperties,
	///     put_PreselectBasenameOnLabelEdit
	virtual HRESULT STDMETHODCALLTYPE put_LimitLabelEditInput(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShellListItems property</em>
	///
	/// Retrieves a collection object wrapping the listview control's shell items.
	///
	/// \param[out] ppItems Receives the collection object's \c IShListViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Namespaces, get_Columns, ShellListViewItems
	virtual HRESULT STDMETHODCALLTYPE get_ListItems(IShListViewItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c LoadOverlaysOnDemand property</em>
	///
	/// Retrieves whether an item's overlay icon is loaded on demand or when adding the item. If this
	/// property is set to \c VARIANT_TRUE, the overlay icon is loaded when it is needed; otherwise it is
	/// loaded when the item is added.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Loading overlay icons on demand is faster, but the \c IListViewItem::OverlayIndex property
	///          always returns the <em>current</em> overlay icon index which isn't necessarily correct until
	///          the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \sa put_LoadOverlaysOnDemand, ShellListViewItem::get_ManagedProperties
	virtual HRESULT STDMETHODCALLTYPE get_LoadOverlaysOnDemand(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c LoadOverlaysOnDemand property</em>
	///
	/// Sets whether an item's overlay icon is loaded on demand or when adding the item. If this
	/// property is set to \c VARIANT_TRUE, the overlay icon is loaded when it is needed; otherwise it is
	/// loaded when the item is added.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Loading overlay icons on demand is faster, but the \c IListViewItem::OverlayIndex property
	///          always returns the <em>current</em> overlay icon index which isn't necessarily correct until
	///          the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \sa get_LoadOverlaysOnDemand, ShellListViewItem::put_ManagedProperties
	virtual HRESULT STDMETHODCALLTYPE put_LoadOverlaysOnDemand(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShellListNamespaces property</em>
	///
	/// Retrieves a collection object wrapping the shell namespaces managed by this control.
	///
	/// \param[out] ppNamespaces Receives the collection object's \c IShListViewNamespaces implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ListItems, get_Columns, ShellListViewNamespaces
	virtual HRESULT STDMETHODCALLTYPE get_Namespaces(IShListViewNamespaces** ppNamespaces);
	/// \brief <em>Retrieves the current setting of the \c PersistColumnSettingsAcrossNamespaces property</em>
	///
	/// Retrieves whether the control keeps the current column settings when loading the columns of another
	/// namespace. Any of the values defined by the \c ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// enumeration is valid. The following column settings are preserved:
	/// - \c Alignment
	/// - \c Caption
	/// - \c Visible
	/// - \c Width
	/// - Column order
	/// - Sorting
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_PersistColumnSettingsAcrossNamespaces, get_Columns, get_AutoInsertColumns,
	///       ShellListViewColumn::get_Alignment, ShellListViewColumn::get_Caption,
	///       ShellListViewColumn::get_Visible, ShellListViewColumn::get_Width,
	///       ShBrowserCtlsLibU::ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// \else
	///   \sa put_PersistColumnSettingsAcrossNamespaces, get_Columns, get_AutoInsertColumns,
	///       ShellListViewColumn::get_Alignment, ShellListViewColumn::get_Caption,
	///       ShellListViewColumn::get_Visible, ShellListViewColumn::get_Width,
	///       ShBrowserCtlsLibA::ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_PersistColumnSettingsAcrossNamespaces(ShLvwPersistColumnSettingsAcrossNamespacesConstants* pValue);
	/// \brief <em>Sets the \c PersistColumnSettingsAcrossNamespaces property</em>
	///
	/// Sets whether the control keeps the current column settings when loading the columns of another
	/// namespace. Any of the values defined by the \c ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// enumeration is valid. The following column settings are preserved:
	/// - \c Alignment
	/// - \c Caption
	/// - \c Visible
	/// - \c Width
	/// - Column order
	/// - Sorting
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_PersistColumnSettingsAcrossNamespaces, get_Columns, put_AutoInsertColumns,
	///       ShellListViewColumn::put_Alignment, ShellListViewColumn::put_Caption,
	///       ShellListViewColumn::put_Visible, ShellListViewColumn::put_Width,
	///       ShBrowserCtlsLibU::ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// \else
	///   \sa get_PersistColumnSettingsAcrossNamespaces, get_Columns, put_AutoInsertColumns,
	///       ShellListViewColumn::put_Alignment, ShellListViewColumn::put_Caption,
	///       ShellListViewColumn::put_Visible, ShellListViewColumn::put_Width,
	///       ShBrowserCtlsLibA::ShLvwPersistColumnSettingsAcrossNamespacesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_PersistColumnSettingsAcrossNamespaces(ShLvwPersistColumnSettingsAcrossNamespacesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c PreselectBasenameOnLabelEdit property</em>
	///
	/// Retrieves whether the control selects the file name only when label-editing files. If set to
	/// \c VARIANT_TRUE, the label-edit control contains the whole file name including the file extension,
	/// but only the base name without the extension is selected. If set to \c VARIANT_FALSE, the whole file
	/// name including the file extension is selected.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_PreselectBasenameOnLabelEdit, ShellListViewItem::get_ManagedProperties,
	///     get_LimitLabelEditInput
	virtual HRESULT STDMETHODCALLTYPE get_PreselectBasenameOnLabelEdit(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c PreselectBasenameOnLabelEdit property</em>
	///
	/// Sets whether the control selects the file name only when label-editing files. If set to
	/// \c VARIANT_TRUE, the label-edit control contains the whole file name including the file extension,
	/// but only the base name without the extension is selected. If set to \c VARIANT_FALSE, the whole file
	/// name including the file extension is selected.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_PreselectBasenameOnLabelEdit, ShellListViewItem::put_ManagedProperties,
	///     put_LimitLabelEditInput
	virtual HRESULT STDMETHODCALLTYPE put_PreselectBasenameOnLabelEdit(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessShellNotifications property</em>
	///
	/// Retrieves whether the control checks for shell notifications and updates the shell items
	/// automatically on file deletions and similar actions. If set to \c VARIANT_TRUE, shell notifications
	/// are processed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessShellNotifications, Raise_ChangedItemPIDL, Raise_ChangedNamespacePIDL
	virtual HRESULT STDMETHODCALLTYPE get_ProcessShellNotifications(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessShellNotifications property</em>
	///
	/// Sets whether the control checks for shell notifications and updates the shell items automatically
	/// on file deletions and similar actions. If set to \c VARIANT_TRUE, shell notifications are processed;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessShellNotifications, Raise_ChangedItemPIDL, Raise_ChangedNamespacePIDL
	virtual HRESULT STDMETHODCALLTYPE put_ProcessShellNotifications(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
  #ifdef ACTIVATE_SECZONE_FEATURES
		/// \brief <em>Retrieves the current setting of the \c SecurityZones property</em>
		///
		/// Retrieves a collection object wrapping the Internet Explorer security zones.
		///
		/// \param[out] ppZones Receives the collection object's \c ISecurityZones implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		///
		/// \sa SecurityZones
		virtual HRESULT STDMETHODCALLTYPE get_SecurityZones(ISecurityZones** ppZones);
	#endif
	/// \brief <em>Retrieves the current setting of the \c SelectSortColumn property</em>
	///
	/// Retrieves whether the control selects the column, by which the items are sorted, automatically.
	/// If set to \c VARIANT_TRUE, the column by which the control is sorted, is set as the attached
	/// control's selected column automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To make this feature work, the \c SortItems method must be used for sorting.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_SelectSortColumn, get_SetSortArrows, SortItems, get_SortOnHeaderClick
	virtual HRESULT STDMETHODCALLTYPE get_SelectSortColumn(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SelectSortColumn property</em>
	///
	/// Sets whether the control selects the column, by which the items are sorted, automatically.
	/// If set to \c VARIANT_TRUE, the column by which the control is sorted, is set as the attached
	/// control's selected column automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To make this feature work, the \c SortItems method must be used for sorting.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_SelectSortColumn, put_SetSortArrows, SortItems, put_SortOnHeaderClick
	virtual HRESULT STDMETHODCALLTYPE put_SelectSortColumn(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SetSortArrows property</em>
	///
	/// Retrieves whether the control sets the sort arrows, that indicate by which column and in which
	/// direction the items are sorted, automatically. If set to \c VARIANT_TRUE, sort arrows are set
	/// automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To make this feature work, the \c SortItems method must be used for sorting.
	///
	/// \sa put_SetSortArrows, get_SelectSortColumn, SortItems, get_SortOnHeaderClick
	virtual HRESULT STDMETHODCALLTYPE get_SetSortArrows(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SetSortArrows property</em>
	///
	/// Sets whether the control sets the sort arrows, that indicate by which column and in which
	/// direction the items are sorted, automatically. If set to \c VARIANT_TRUE, sort arrows are set
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks To make this feature work, the \c SortItems method must be used for sorting.
	///
	/// \sa get_SetSortArrows, put_SelectSortColumn, SortItems, put_SortOnHeaderClick
	virtual HRESULT STDMETHODCALLTYPE put_SetSortArrows(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SortColumnIndex property</em>
	///
	/// Retrieves the zero-based index of the shell column by which the attached listview control is
	/// currently sorted. If the control is sorted by a custom column instead of a shell column, this
	/// property is -1.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa SortItems, ShellListViewNamespace::get_AutoSortItems,
	///     ShellListViewNamespace::get_DefaultSortColumnIndex, Raise_ChangingSortColumn,
	///     Raise_ChangedSortColumn
	virtual HRESULT STDMETHODCALLTYPE get_SortColumnIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c SortOnHeaderClick property</em>
	///
	/// Retrieves whether the control updates the sort direction and column automatically if a column header
	/// is clicked. If set to \c VARIANT_TRUE, the settings for item sorting are updated automatically if a
	/// column header is clicked; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SortOnHeaderClick, SortItems, get_SelectSortColumn, get_SetSortArrows
	virtual HRESULT STDMETHODCALLTYPE get_SortOnHeaderClick(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SortOnHeaderClick property</em>
	///
	/// Sets whether the control updates the sort direction and column automatically if a column header
	/// is clicked. If set to \c VARIANT_TRUE, the settings for item sorting are updated automatically if a
	/// column header is clicked; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SortOnHeaderClick, SortItems, put_SelectSortColumn, put_SetSortArrows
	virtual HRESULT STDMETHODCALLTYPE put_SortOnHeaderClick(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c ThumbnailSize property</em>
	///
	/// Retrieves the width and height in pixels of thumbnail images. If set to a negative value, the
	/// system's default setting is used according to the following list:
	/// - \c -1 The system's default thumbnail size is used.
	/// - \c -2 The size of images in the small system imagelist is used.
	/// - \c -3 The size of images in the large system imagelist is used.
	/// - \c -4 The size of images in the extra-large system imagelist is used.
	/// - \c -5 The size of images in the jumbo-size system imagelist is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ThumbnailSize, get_DisplayThumbnails
	virtual HRESULT STDMETHODCALLTYPE get_ThumbnailSize(LONG* pValue);
	/// \brief <em>Sets the \c ThumbnailSize property</em>
	///
	/// Sets the width and height in pixels of thumbnail images. If set to a negative value, the
	/// system's default setting is used according to the following list:
	/// - \c -1 The system's default thumbnail size is used.
	/// - \c -2 The size of images in the small system imagelist is used.
	/// - \c -3 The size of images in the large system imagelist is used.
	/// - \c -4 The size of images in the extra-large system imagelist is used.
	/// - \c -5 The size of images in the jumbo-size system imagelist is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ThumbnailSize, put_DisplayThumbnails
	virtual HRESULT STDMETHODCALLTYPE put_ThumbnailSize(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c UseColumnResizability property</em>
	///
	/// Retrieves whether the control sets the column resizability according to the column's \c Resizable
	/// property when inserting a shell column into the attached \c ExplorerListView control. If set
	/// to \c VARIANT_TRUE, the column resizability is set according to the columns' \c Resizable property;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseColumnResizability, ShellListViewColumn::get_Resizable
	virtual HRESULT STDMETHODCALLTYPE get_UseColumnResizability(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseColumnResizability property</em>
	///
	/// Sets whether the control sets the column resizability according to the column's \c Resizable
	/// property when inserting a shell column into the attached \c ExplorerListView control. If set
	/// to \c VARIANT_TRUE, the column resizability is set according to the columns' \c Resizable property;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseColumnResizability, ShellListViewColumn::get_Resizable
	virtual HRESULT STDMETHODCALLTYPE put_UseColumnResizability(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseFastButImpreciseItemHandling property</em>
	///
	/// Retrieves whether the control skips time-consuming extra-checks when e. g. removing items. The
	/// control can do extra-checks to ensure, that it doesn't remove or modify items that shouldn't be
	/// removed or modified. These extra-checks are time-consuming and may lead to delays of several
	/// seconds when browsing from one namespace with many items to another one.\n
	/// If set to \c VARIANT_TRUE, these extra-checks are skipped; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Setting this property to \c True may result in unwanted behavior. E. g. removing a
	///          namespace will result in <strong>any</strong> item being removed, not just those that
	///          belong to the removed namespace.
	///
	/// \sa put_UseFastButImpreciseItemHandling, get_Namespaces
	virtual HRESULT STDMETHODCALLTYPE get_UseFastButImpreciseItemHandling(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseFastButImpreciseItemHandling property</em>
	///
	/// Sets whether the control skips time-consuming extra-checks when e. g. removing items. The
	/// control can do extra-checks to ensure, that it doesn't remove or modify items that shouldn't be
	/// removed or modified. These extra-checks are time-consuming and may lead to delays of several
	/// seconds when browsing from one namespace with many items to another one.\n
	/// If set to \c VARIANT_TRUE, these extra-checks are skipped; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Setting this property to \c True may result in unwanted behavior. E. g. removing a
	///          namespace will result in <strong>any</strong> item being removed, not just those that
	///          belong to the removed namespace.
	///
	/// \sa get_UseFastButImpreciseItemHandling, get_Namespaces
	virtual HRESULT STDMETHODCALLTYPE put_UseFastButImpreciseItemHandling(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseGenericIcons property</em>
	///
	/// Retrieves when the control displays generic icons and when it displays item-specific icons. Any of
	/// the values defined by the \c UseGenericIconsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_UseGenericIcons, get_UseSystemImageList, ShBrowserCtlsLibU::UseGenericIconsConstants
	/// \else
	///   \sa put_UseGenericIcons, get_UseSystemImageList, ShBrowserCtlsLibA::UseGenericIconsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UseGenericIcons(UseGenericIconsConstants* pValue);
	/// \brief <em>Sets the \c UseGenericIcons property</em>
	///
	/// Sets when the control displays generic icons and when it displays item-specific icons. Any of
	/// the values defined by the \c UseGenericIconsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_UseGenericIcons, put_UseSystemImageList, ShBrowserCtlsLibU::UseGenericIconsConstants
	/// \else
	///   \sa get_UseGenericIcons, put_UseSystemImageList, ShBrowserCtlsLibA::UseGenericIconsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UseGenericIcons(UseGenericIconsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemImageList property</em>
	///
	/// Retrieves a bit field indicating which of the attached listview's imagelists are set to the system
	/// imagelist. Any combination of the values defined by the \c UseSystemImageListConstants enumeration
	/// is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_UseSystemImageList, get_DisplayThumbnails, get_hImageList, get_UseGenericIcons,
	///       ShBrowserCtlsLibU::UseSystemImageListConstants
	/// \else
	///   \sa put_UseSystemImageList, get_DisplayThumbnails, get_hImageList, get_UseGenericIcons,
	///       ShBrowserCtlsLibA::UseSystemImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemImageList(UseSystemImageListConstants* pValue);
	/// \brief <em>Sets the \c UseSystemImageList property</em>
	///
	/// Sets a bit field indicating which of the attached listview's imagelists are set to the system
	/// imagelist. Any combination of the values defined by the \c UseSystemImageListConstants enumeration
	/// is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_UseSystemImageList, put_DisplayThumbnails, put_hImageList, put_UseGenericIcons,
	///       ShBrowserCtlsLibU::UseSystemImageListConstants
	/// \else
	///   \sa get_UseSystemImageList, put_DisplayThumbnails, put_hImageList, put_UseGenericIcons,
	///       ShBrowserCtlsLibA::UseSystemImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemImageList(UseSystemImageListConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseThreadingForSlowColumns property</em>
	///
	/// Retrieves whether the control extracts the text of shell items' sub-items in a background thread if
	/// access to the shell column's content is flagged (by the shell) as being slow. If set to
	/// \c VARIANT_TRUE, a background thread is used; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseThreadingForSlowColumns, get_Columns, ShellListViewColumn::get_Slow,
	///     OnGetDispInfoNotification
	virtual HRESULT STDMETHODCALLTYPE get_UseThreadingForSlowColumns(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseThreadingForSlowColumns property</em>
	///
	/// Sets whether the control extracts the text of shell items' sub-items in a background thread if
	/// access to the shell column's content is flagged (by the shell) as being slow. If set to
	/// \c VARIANT_TRUE, a background thread is used; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseThreadingForSlowColumns, get_Columns, ShellListViewColumn::get_Slow,
	///     OnGetDispInfoNotification
	virtual HRESULT STDMETHODCALLTYPE put_UseThreadingForSlowColumns(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseThumbnailDiskCache property</em>
	///
	/// Retrieves whether the thumbnail disk cache (Thumbs.db files) is used to improve the performance of
	/// thumbnail mode. If set to \c VARIANT_TRUE, the disk cache is used; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The disk cache can be used only if the setting of the \c ThumbnailSize property matches
	///          the system setting for the size of thumbnails.\n
	///          Starting with Windows Vista, the disk cache is always used (for any thumbnail size).
	///
	/// \sa put_UseThumbnailDiskCache, get_DisplayThumbnails, get_ThumbnailSize
	virtual HRESULT STDMETHODCALLTYPE get_UseThumbnailDiskCache(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseThumbnailDiskCache property</em>
	///
	/// Sets whether the thumbnail disk cache (Thumbs.db files) is used to improve the performance of
	/// thumbnail mode. If set to \c VARIANT_TRUE, the disk cache is used; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The disk cache can be used only if the setting of the \c ThumbnailSize property matches
	///          the system setting for the size of thumbnails.\n
	///          Starting with Windows Vista, the disk cache is always used (for any thumbnail size).
	///
	/// \sa get_UseThumbnailDiskCache, put_DisplayThumbnails, put_ThumbnailSize
	virtual HRESULT STDMETHODCALLTYPE put_UseThumbnailDiskCache(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Attaches the object to the specified \c SysListView32 window</em>
	///
	/// \param[in] hWnd The \c SysListView32 window to attach to.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Detach, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE Attach(OLE_HANDLE hWnd);
	/// \brief <em>Compares two shell items</em>
	///
	/// \param[in] pFirstItem The first shell item to compare.
	/// \param[in] pSecondItem The second shell item to compare.
	/// \param[in] sortColumnIndex The zero-based index of the shell column by which to compare. If set to
	///            -1, the default sort column of the first defined shell namespace is used.
	/// \param[out] pResult Receives a negative value if the first item should preceed the second item; a
	///             positive value if the second item should preceed the first item; otherwise 0.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewNamespace::get_AutoSortItems, ShellListViewNamespace::SortItems
	virtual HRESULT STDMETHODCALLTYPE CompareItems(IShListViewItem* pFirstItem, IShListViewItem* pSecondItem, LONG sortColumnIndex = -1, LONG* pResult = NULL);
	/// \brief <em>Creates the default context menu for the header control</em>
	///
	/// Creates the default context menu for the header control. The control uses the current column set.
	/// \c EnsureShellColumnsAreLoaded should be called before this method.
	///
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayHeaderContextMenu, DestroyHeaderContextMenu, CreateShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateHeaderContextMenu(OLE_HANDLE* pMenu);
	/// \brief <em>Creates a shell context menu</em>
	///
	/// \param[in] items An object specifying the listview items for which to create the shell context
	///            menu. The following values may be used to identify the items:
	///            - A single item ID.
	///            - An array of item IDs.
	///            - A \c ListViewItem object.
	///            - A \c ListViewItems object.
	///            - A \c ListViewItemContainer object.
	///            - A \c ShellListViewItem object.
	///            - A \c ShellListViewItems object.
	///            To create the background context menu of a shell namespace, pass the
	///            \c ShellListViewNamespace object of the shell namespace.
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellListViewItem::CreateShellContextMenu,
	///     ShellListViewItems::CreateShellContextMenu, ShellListViewNamespace::CreateShellContextMenu,
	///     DestroyShellContextMenu, CreateHeaderContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateShellContextMenu(VARIANT items, OLE_HANDLE* pMenu);
	/// \brief <em>Destroys the control's current header context menu</em>
	///
	/// Destroys the header context menu that was created by the last call to a \c CreateHeaderContextMenu
	/// method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateHeaderContextMenu, DestroyShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DestroyHeaderContextMenu(void);
	/// \brief <em>Destroys the control's current shell context menu</em>
	///
	/// Destroys the shell context menu that was created by the last call to a \c CreateShellContextMenu
	/// method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellListViewItem::CreateShellContextMenu,
	///     ShellListViewItems::CreateShellContextMenu, ShellListViewNamespace::CreateShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DestroyShellContextMenu(void);
	/// \brief <em>Detaches the object from the specified \c SysListView32 window</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Attach, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE Detach(void);
	/// \brief <em>Displays the default context menu for the header control</em>
	///
	/// Displays the default context menu for the header control. The control uses the current column set.
	/// \c EnsureShellColumnsAreLoaded should be called before this method.
	///
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateHeaderContextMenu, DestroyHeaderContextMenu, DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayHeaderContextMenu(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Displays a shell context menu</em>
	///
	/// \param[in] items An object specifying the listview items for which to display the shell context
	///            menu. The following values may be used to identify the items:
	///            - A single item ID.
	///            - An array of item IDs.
	///            - A \c ListViewItem object.
	///            - A \c ListViewItems object.
	///            - A \c ListViewItemContainer object.
	///            - A \c ShellListViewItem object.
	///            - A \c ShellListViewItems object.
	///            To display the background context menu of a shell namespace, pass the
	///            \c ShellListViewNamespace object of the shell namespace.
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DestroyShellContextMenu, DisplayHeaderContextMenu,
	///     ShellListViewItem::DisplayShellContextMenu, ShellListViewItems::DisplayShellContextMenu,
	///     ShellListViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(VARIANT items, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Ensures that the shell columns have been loaded</em>
	///
	/// \c ShellListView loads the shell columns as late as possible. As long as the active view mode doesn't
	/// require the shell columns to be loaded ('Tiles' and 'Details' view mode are the only view modes for
	/// which columns are required), the control doesn't load them.\n
	/// This method loads the shell columns if they have not yet been loaded; otherwise it does nothing. It
	/// should be called before accessing the \c Columns property if the control is in e. q. 'Icons' view
	/// mode.
	///
	/// \param[in] waitForLastColumn If set to \c VARIANT_TRUE, the method doesn't return before all work is
	///            done. If set to \c VARIANT_FALSE, it may move the shell column enumeration to a background
	///            thread and return before all columns have been loaded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Columns, get_AutoInsertColumns
	virtual HRESULT STDMETHODCALLTYPE EnsureShellColumnsAreLoaded(VARIANT_BOOL waitForLastColumn = VARIANT_TRUE);
	/// \brief <em>Reloads all icons managed by the control</em>
	///
	/// \param[in] includeOverlays If set to \c VARIANT_TRUE, not only the icons managed by the control are
	///            flushed, but also the overlay images; otherwise only the icons are flushed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessShellNotifications, FlushIcons
	virtual HRESULT STDMETHODCALLTYPE FlushManagedIcons(VARIANT_BOOL includeOverlays = VARIANT_TRUE);
	/// \brief <em>Retrieves the specified header menu item's caption</em>
	///
	/// Retrieves the caption of the specified header context menu item.
	///
	/// \param[in] commandID The unique ID of the context menu item.
	/// \param[in,out] pItemCaption Receives the caption of the item specified by \c commandID.
	/// \param[out] pCommandIDWasValid If \c VARIANT_TRUE, the menu item specified by \c commandID is a valid
	///             header context menu item; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetShellContextMenuItemString, Raise_ChangedContextMenuSelection
	virtual HRESULT STDMETHODCALLTYPE GetHeaderContextMenuItemString(LONG commandID, VARIANT* pItemCaption = NULL, VARIANT_BOOL* pCommandIDWasValid = NULL);
	/// \brief <em>Retrieves the specified shell menu item's description or verb</em>
	///
	/// Retrieves the help text and/or the language-independent command name of the specified shell context
	/// menu item. The help text may be displayed in a status bar if the menu item is selected; the command
	/// name ('verb') may be used to identify the command (the command ID may be different on different
	/// systems).
	///
	/// \param[in] commandID The unique ID of the context menu item.
	/// \param[in,out] pItemDescription Receives the help text of the item specified by \c commandID.
	/// \param[in,out] pItemVerb Receives the language-independent command name of the item specified by
	///                \c commandID.
	/// \param[out] pCommandIDWasValid If \c VARIANT_TRUE, the menu item specified by \c commandID is a valid
	///             shell context menu item; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Many shell extensions (some being part of Windows) are not implemented correctly. Therefore
	///          \c pCommandIDWasValid may report an invalid menu item although the item is a valid shell
	///          context menu item and valid strings are returned.
	///
	/// \sa GetHeaderContextMenuItemString, Raise_ChangedContextMenuSelection
	virtual HRESULT STDMETHODCALLTYPE GetShellContextMenuItemString(LONG commandID, VARIANT* pItemDescription = NULL, VARIANT* pItemVerb = NULL, VARIANT_BOOL* pCommandIDWasValid = NULL);
	/// \brief <em>Executes the default command from the specified items' shell context menu</em>
	///
	/// \param[in] items An object specifying the listview items for which to invoke the default command.
	///            The following values may be used to identify the items:
	///            - A single item ID.
	///            - An array of item IDs.
	///            - A \c ListViewItem object.
	///            - A \c ListViewItems object.
	///            - A \c ListViewItemContainer object.
	///            - A \c ShellListViewItem object.
	///            - A \c ShellListViewItems object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellListViewItem::InvokeDefaultShellContextMenuCommand,
	///     ShellListViewItems::InvokeDefaultShellContextMenuCommand
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultShellContextMenuCommand(VARIANT items);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Sorts the attached control's items</em>
	///
	/// Sorts the attached control's items.\n
	/// \c ExplorerListView::SortItems is called with the following parameters:
	/// - \c firstCriterion: \c sobShell
	/// - \c secondCriterion: \c sobText
	/// - \c thirdCriterion: \c sobCustom
	/// - \c fourthCriterion: \c sobNone
	/// - \c fifthCriterion: \c sobNone
	/// - \c column: The specified shell column's index within the attached control, if the column is
	///   currently activated; otherwise this parameter is not specified.
	/// - \c caseSensitive: \c VARIANT_FALSE
	///
	/// \param[in] shellColumnIndex The zero-based index of the shell column by which to sort. If set to -1,
	///            the default sort column of the first defined shell namespace is used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemTypeSortOrder, ShellListViewNamespace::get_AutoSortItems, get_SelectSortColumn,
	///     get_SetSortArrows, get_SortOnHeaderClick, get_ListItems, get_Columns
	virtual HRESULT STDMETHODCALLTYPE SortItems(LONG shellColumnIndex = -1);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Overrides ATL's \c CComControlBase::OnDraw method</em>
	///
	/// This method is called if the control needs to be drawn.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa imglstIDEIcon,
	///     <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Informs an object of how much display space its container has assigned it</em>
	///
	/// \param[in] drawAspect A value describing which form, or "aspect", of an object is to be displayed.
	///            Any of the values defined by the \c DVASPECT enumeration is valid.
	/// \param[in] The size limit for the object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694330.aspx">IOleObject::SetExtent</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetExtent(DWORD /*drawAspect*/, SIZEL* /*pSize*/);

	/// \brief <em>Retrieves the current setting of the \c hWndShellUIParentWindow property</em>
	///
	/// \return The current setting of the \c hWndShellUIParentWindow property.
	///
	/// \sa get_hWndShellUIParentWindow, put_hWndShellUIParentWindow
	HWND GethWndShellUIParentWindow(void);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IContextMenuCB
	///
	//@{
	/// \brief <em>Callback function required for cross namespace context menus</em>
	///
	/// \param[in] pShellFolder The \c IShellFolder object the message applies to.
	/// \param[in] hWnd The window containing the view that receives the message.
	/// \param[in] pDataObject The \c IDataObject object representing the items that the context menu refers
	///            to.
	/// \param[in] message A \c DFM_* notification.
	/// \param[in] wParam Additional information specific for the notification.
	/// \param[in] lParam Additional information specific for the notification.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetCrossNamespaceContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776081.aspx">IContextMenuCB::CallBack</a>
	virtual HRESULT STDMETHODCALLTYPE CallBack(LPSHELLFOLDER pShellFolder, HWND hWnd, LPDATAOBJECT pDataObject, UINT message, WPARAM wParam, LPARAM lParam);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	// \param[in] languageID The locale identifier identifying the language in which name should be
	//            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties, ShLvwDefaultManagedItemPropertiesProperties, NamespaceEnumSettingsProperties,
	///     ThumbnailsProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IMessageListener
	///
	//@{
	/// \brief <em>Allows \c ShellListView to post-process a message</em>
	///
	/// This method is the very last one that processes a received message. It allows \c ShellListView to
	/// process the message after the listview control and after \c DefWindowProc.
	///
	/// \param[in] hWnd The window that received the message.
	/// \param[in] message The received message.
	/// \param[in] wParam The received message's \c wParam parameter.
	/// \param[in] lParam The received message's \c lParam parameter.
	/// \param[out] pResult Receives the result of message processing.
	/// \param[in] cookie A value specified by the \c PreMessageFilter method.
	/// \param[in] eatenMessage If \c TRUE, the listview control has eaten the message, i. e. it has not
	///            forwarded it to the default window procedure.
	///
	/// \return \c S_OK if the message was processed; \c S_FALSE if the message was not processed;
	///         \c E_NOTIMPL if we do not filter any messages; \c E_POINTER if \c pResult is an illegal
	///         pointer.
	///
	/// \sa PreMessageFilter, IMessageListener::PostMessageFilter
	virtual HRESULT PostMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, DWORD cookie, BOOL eatenMessage);
	/// \brief <em>Allows \c ShellListView to pre-process a message</em>
	///
	/// This method is the very first one that processes a received message. It allows \c ShellListView to
	/// process the message before the listview control.
	///
	/// \param[in] hWnd The window that received the message.
	/// \param[in] message The received message.
	/// \param[in] wParam The received message's \c wParam parameter.
	/// \param[in] lParam The received message's \c lParam parameter.
	/// \param[out] pResult Receives the result of message processing.
	/// \param[out] pCookie Receives a \c DWORD value that is passed to the \c PostMessageFilter method.
	///
	/// \return \c S_OK if the message was processed; \c S_FALSE if the message was not processed;
	///         \c E_NOTIMPL if we do not filter any messages; \c E_POINTER if \c pResult or \c pCookie is
	///         an illegal pointer.
	///
	/// \sa PostMessageFilter, IMessageListener::PreMessageFilter
	virtual HRESULT PreMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, LPDWORD pCookie);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IInternalMessageListener
	///
	//@{
	/// \brief <em>Processes the messages sent by the attached listview control</em>
	///
	/// \param[in] message The message to process.
	/// \param[in] wParam Used to transfer data with the message.
	/// \param[in] lParam Used to transfer data with the message.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filtered message handlers
	///
	//@{
	/// \brief <em>\c WM_DRAWITEM, \c WM_INITMENUPOPUP, \c WM_MEASUREITEM, \c WM_MENUCHAR, \c WM_MENUSELECT and \c WM_NEXTMENU handler</em>
	///
	/// Will be called if the attached control receives a menu message.
	/// We use this handler for shell context menu support.
	///
	/// \sa OnInternalContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646352.aspx">WM_MENUSELECT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647612.aspx">WM_NEXTMENU</a>
	LRESULT OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c LVM_GETITEM handler</em>
	///
	/// Will be called if some or all of a listview item's attributes are requested from the attached
	/// control.
	/// We use this handler to intercept and handle requests of non-shell item's icon indexes.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms670761.aspx">LVM_GETITEM</a>
	LRESULT OnGetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c LVM_SETITEM handler</em>
	///
	/// Will be called if the attached control is requested to update some or all of a listview item's
	/// attributes.
	/// We use this handler to intercept and handle changes of non-shell item's icon indexes.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms670845.aspx">LVM_SETITEM</a>
	LRESULT OnSetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the attached control's parent window receives a notification from the attached
	/// control.
	/// We use this handler to do some shell stuff on some notifications.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& cookie, BOOL& /*eatenMessage*/, BOOL preProcessing);
	/// \brief <em>\c WM_REPORT_THUMBNAILDISKCACHEACCESS handler</em>
	///
	/// Will be called if the attached control is queried by a background task to schedule a new background
	/// task.
	/// We use this handler to allow the creation of new background tasks in a background task.
	///
	/// \remarks Beginning with Windows Vista, this member isn't required anymore.
	///
	/// \sa WM_REPORT_THUMBNAILDISKCACHEACCESS, ShLvwLegacyTestThumbnailCacheTask,
	///     ShLvwLegacyExtractThumbnailTask, ShLvwLegacyExtractThumbnailFromDiskCacheTask,
	///     LazyCloseThumbnailDiskCaches
	LRESULT OnReportThumbnailDiskCacheAccess(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_SHCHANGENOTIFY handler</em>
	///
	/// Will be called if the attached control receives a notification that the state of the shell has
	/// changed.
	/// We use this handler to update the shell items.
	///
	/// \sa WM_SHCHANGENOTIFY
	LRESULT OnSHChangeNotify(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called if the attached control receives a notification that a timer has expired.
	/// We use this handler to raise the \c EnumerationTimedOut event.
	///
	/// \sa Raise_EnumerationTimedOut,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& eatenMessage, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_ENQUEUETASK handler</em>
	///
	/// Will be called if the attached control is queried by a background task to schedule a new background
	/// task.
	/// We use this handler to allow the creation of new background tasks in a background task.
	///
	/// \sa WM_TRIGGER_ENQUEUETASK, SHCTLSBACKGROUNDTASKINFO
	LRESULT OnTriggerEnqueueTask(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_ITEMENUMCOMPLETE handler</em>
	///
	/// Will be called if the attached control receives a namespace's sub-items retrieved by a
	/// \c ShLvwBackgroundItemEnumTask task.
	/// We use this handler to dynamically load items.
	///
	/// \sa WM_TRIGGER_ITEMENUMCOMPLETE, ShLvwBackgroundItemEnumTask, OnTriggerColumnEnumComplete
	LRESULT OnTriggerItemEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_COLUMNENUMCOMPLETE handler</em>
	///
	/// Will be called if the attached control receives a namespace's columns retrieved by a
	/// \c ShLvwBackgroundColumnEnumTask task.
	/// We use this handler to dynamically load columns.
	///
	/// \sa WM_TRIGGER_COLUMNENUMCOMPLETE, ShLvwBackgroundColumnEnumTask, OnTriggerItemEnumComplete
	LRESULT OnTriggerColumnEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_SETELEVATIONSHIELD handler</em>
	///
	/// Will be called if the attached control receives an item's elevation requirements retrieved by a
	/// \c ElevationShieldTask task.
	/// We use this handler to dynamically load UAC elevation shields.
	///
	/// \sa OnTriggerUpdateIcon, OnTriggerUpdateThumbnail, WM_TRIGGER_SETELEVATIONSHIELD, ElevationShieldTask
	LRESULT OnTriggerSetElevationShield(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_SETINFOTIP handler</em>
	///
	/// Will be called if the attached control receives an item's info tip text retrieved by a
	/// \c ShLvwInfoTipTask task.
	/// We use this handler to dynamically load item info tips.
	///
	/// \sa WM_TRIGGER_SETINFOTIP, ShLvwInfoTipTask, OnInternalGetInfoTip
	LRESULT OnTriggerSetInfoTip(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATEICON handler</em>
	///
	/// Will be called if the attached control receives an item's icon index retrieved by a \c ShLvwIconTask
	/// task.
	/// We use this handler to dynamically load item icons.
	///
	/// \sa OnTriggerUpdateThumbnail, WM_TRIGGER_UPDATEICON, ShLvwIconTask, OnTriggerUpdateOverlay
	LRESULT OnTriggerUpdateIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATEOVERLAY handler</em>
	///
	/// Will be called if the attached control receives an item's overlay icon index retrieved by a
	/// \c ShLvwOverlayTask task.
	/// We use this handler to dynamically load item overlays.
	///
	/// \sa OnTriggerUpdateIcon, OnTriggerUpdateThumbnail, WM_TRIGGER_UPDATEOVERLAY, ShLvwOverlayTask
	LRESULT OnTriggerUpdateOverlay(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATESUBITEMCONTROL handler</em>
	///
	/// Will be called if the attached control receives a sub-item control's current value retrieved by a
	/// \c ShLvwSubItemControlTask task.
	/// We use this handler to dynamically load sub-item control values.
	///
	/// \sa OnTriggerUpdateText, WM_TRIGGER_UPDATESUBITEMCONTROL, ShLvwSubItemControlTask
	LRESULT OnTriggerUpdateSubItemControl(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATETEXT handler</em>
	///
	/// Will be called if the attached control receives a sub-item's text retrieved by a
	/// \c ShLvwSlowColumnTask task.
	/// We use this handler to dynamically load sub-items for columns that are flagged with
	/// \c SHCOLSTATE_SLOW.
	///
	/// \sa OnTriggerUpdateSubItemControl, WM_TRIGGER_UPDATETEXT, ShLvwSlowColumnTask
	LRESULT OnTriggerUpdateText(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATETHUMBNAIL handler</em>
	///
	/// Will be called if the attached control receives an item's thumbnail retrieved by a
	/// \c ShLvwThumbnailTask task.
	/// We use this handler to dynamically load item thumbnails.
	///
	/// \sa OnTriggerUpdateIcon, WM_TRIGGER_UPDATETHUMBNAIL, ShLvwThumbnailTask, OnTriggerUpdateOverlay
	LRESULT OnTriggerUpdateThumbnail(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATETILEVIEWCOLUMNS handler</em>
	///
	/// Will be called if the attached control receives an item's tile view column indexes retrieved by a
	/// \c ShLvwTileViewTask or \c ShLvwLegacyTileViewTask task.
	/// We use this handler to dynamically load sub-items in tile view.
	///
	/// \sa WM_TRIGGER_UPDATETILEVIEWCOLUMNS, ShLvwTileViewTask, ShLvwLegacyTileViewTask,
	///     OnTriggerColumnEnumComplete
	LRESULT OnTriggerUpdateTileViewColumns(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Internal message handlers
	///
	//@{
	/// \brief <em>Handles the \c SHLVM_ALLOCATEMEMORY message</em>
	///
	/// \sa SHLVM_ALLOCATEMEMORY
	HRESULT OnInternalAllocateMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_FREEMEMORY message</em>
	///
	/// \sa SHLVM_FREEMEMORY
	HRESULT OnInternalFreeMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_INSERTEDCOLUMN message</em>
	///
	/// \sa SHLVM_INSERTEDCOLUMN
	HRESULT OnInternalInsertedColumn(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_FREECOLUMN message</em>
	///
	/// \sa SHLVM_FREECOLUMN
	HRESULT OnInternalFreeColumn(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_GETSHLISTVIEWCOLUMN message</em>
	///
	/// \sa SHLVM_GETSHLISTVIEWCOLUMN
	HRESULT OnInternalGetShListViewColumn(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_HEADERCLICK message</em>
	///
	/// \sa SHLVM_HEADERCLICK
	HRESULT OnInternalHeaderClick(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_INSERTEDITEM message</em>
	///
	/// \sa SHLVM_INSERTEDITEM
	HRESULT OnInternalInsertedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_FREEITEM message</em>
	///
	/// \sa SHLVM_FREEITEM
	HRESULT OnInternalFreeItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_GETSHLISTVIEWITEM message</em>
	///
	/// \sa SHLVM_GETSHLISTVIEWITEM
	HRESULT OnInternalGetShListViewItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_RENAMEDITEM message</em>
	///
	/// \sa SHLVM_RENAMEDITEM
	HRESULT OnInternalRenamedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_COMPAREITEMS message</em>
	///
	/// \sa SHLVM_COMPAREITEMS
	HRESULT OnInternalCompareItems(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_BEGINLABELEDITA and \c SHLVM_BEGINLABELEDITW messages</em>
	///
	/// \sa SHLVM_BEGINLABELEDITA, SHLVM_BEGINLABELEDITW
	HRESULT OnInternalBeginLabelEdit(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_BEGINDRAG message</em>
	///
	/// \sa SHLVM_BEGINDRAG
	HRESULT OnInternalBeginDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_HANDLEDRAGEVENT message</em>
	///
	/// \sa SHLVM_HANDLEDRAGEVENT
	HRESULT OnInternalHandleDragEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_CONTEXTMENU message</em>
	///
	/// \sa SHLVM_CONTEXTMENU
	HRESULT OnInternalContextMenu(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_GETINFOTIP message</em>
	///
	/// \sa SHLVM_GETINFOTIP, OnInternalGetInfoTipEx
	HRESULT OnInternalGetInfoTip(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_CUSTOMDRAW message</em>
	///
	/// \sa SHLVM_CUSTOMDRAW
	HRESULT OnInternalCustomDraw(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_CHANGEDVIEW message</em>
	///
	/// \sa SHLVM_CHANGEDVIEW
	HRESULT OnInternalViewChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHLVM_GETINFOTIPEX message</em>
	///
	/// \sa SHLVM_GETINFOTIPEX, OnInternalGetInfoTip
	HRESULT OnInternalGetInfoTipEx(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHLVM_GETSUBITEMCONTROL message</em>
	///
	/// \sa SHLVM_GETSUBITEMCONTROL
	HRESULT OnInternalGetSubItemControl(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	// \brief <em>Handles the \c SHLVM_ENDSUBITEMEDIT message</em>
	//
	// \sa SHLVM_ENDSUBITEMEDIT
	//HRESULT OnInternalEndSubItemEdit(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filtered notification handlers
	///
	//@{
	/// \brief <em>\c LVN_GETDISPINFO handler</em>
	///
	/// Will be called if the attached control requests display information about a listview item from its
	/// parent window.
	/// We use this handler to dynamically set the item's managed properties.
	///
	/// \sa OnGetDispInfoNotification, ShellListViewItem::get_ManagedProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774818.aspx">LVN_GETDISPINFO</a>
	LRESULT OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c ChangedColumnVisibility event</em>
	///
	/// \param[in] pColumn The column that has been inserted into or removed from the listview control.
	/// \param[in] hasBecomeVisible If \c VARIANT_TRUE, the column has been inserted; otherwise it has been
	///            removed from the listview control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangedColumnVisibility,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangedColumnVisibility, Raise_ChangingColumnVisibility,
	///       Raise_UnloadedColumn
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangedColumnVisibility,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangedColumnVisibility, Raise_ChangingColumnVisibility,
	///       Raise_UnloadedColumn
	/// \endif
	inline HRESULT Raise_ChangedColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL hasBecomeVisible);
	/// \brief <em>Raises the \c ChangedContextMenuSelection event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu. Will be \c NULL if the menu is closed.
	/// \param[in] isShellContextMenu If \c VARIANT_TRUE, the menu is a shell context menu; otherwise it is a
	///            column header context menu.
	/// \param[in] commandID The ID of the new selected menu item. May be 0.
	/// \param[in] isCustomMenuItem If \c VARIANT_FALSE, the selected menu item has been inserted by
	///            \c ShellListView; otherwise it is a custom menu item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangedContextMenuSelection,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangedContextMenuSelection,
	///       Raise_SelectedShellContextMenuItem, GetShellContextMenuItemString,
	///       GetHeaderContextMenuItemString
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangedContextMenuSelection,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangedContextMenuSelection,
	///       Raise_SelectedShellContextMenuItem, GetShellContextMenuItemString,
	///       GetHeaderContextMenuItemString
	/// \endif
	inline HRESULT Raise_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, VARIANT_BOOL isShellContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem);
	/// \brief <em>Raises the \c ChangedItemPIDL event</em>
	///
	/// \param[in] itemID The unique ID of the item that was updated.
	/// \param[in] previousPIDL The item's previous fully qualified pIDL.
	/// \param[in] newPIDL The item's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangedItemPIDL,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangedItemPIDL, Raise_ChangedNamespacePIDL
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangedItemPIDL,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangedItemPIDL, Raise_ChangedNamespacePIDL
	/// \endif
	inline HRESULT Raise_ChangedItemPIDL(LONG itemID, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL);
	/// \brief <em>Raises the \c ChangedNamespacePIDL event</em>
	///
	/// \param[in] pNamespace The namespace that was updated.
	/// \param[in] previousPIDL The namespace's previous fully qualified pIDL.
	/// \param[in] newPIDL The namespace's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangedNamespacePIDL,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangedNamespacePIDL, Raise_ChangedItemPIDL
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangedNamespacePIDL,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangedNamespacePIDL, Raise_ChangedItemPIDL
	/// \endif
	inline HRESULT Raise_ChangedNamespacePIDL(IShListViewNamespace* pNamespace, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL);
	/// \brief <em>Raises the \c ChangedSortColumn event</em>
	///
	/// \param[in] previousSortColumnIndex The zero-based index of the old sort column. This index is the
	///            column's index given by the shell.
	/// \param[in] newSortColumnIndex The zero-based index of the new sort column. This index is the
	///            column's index given by the shell.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangedSortColumn,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangedSortColumn, Raise_ChangingSortColumn
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangedSortColumn,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangedSortColumn, Raise_ChangingSortColumn
	/// \endif
	inline HRESULT Raise_ChangedSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex);
	/// \brief <em>Raises the \c ChangingColumnVisibility event</em>
	///
	/// \param[in] pColumn The column that will be inserted into or removed from the listview control.
	/// \param[in] willBecomeVisible If \c VARIANT_TRUE, the column will be inserted; otherwise it will be
	///            removed from the listview control.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion or removal; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangingColumnVisibility,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangingColumnVisibility, Raise_LoadedColumn,
	///       Raise_ChangedColumnVisibility, Raise_UnloadedColumn
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangingColumnVisibility,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangingColumnVisibility, Raise_LoadedColumn,
	///       Raise_ChangedColumnVisibility, Raise_UnloadedColumn
	/// \endif
	inline HRESULT Raise_ChangingColumnVisibility(IShListViewColumn* pColumn, VARIANT_BOOL willBecomeVisible, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ChangingSortColumn event</em>
	///
	/// \param[in] previousSortColumnIndex The zero-based index of the old sort column. This index is the
	///            column's index given by the shell.
	/// \param[in] newSortColumnIndex The zero-based index of the new sort column. This index is the
	///            column's index given by the shell.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the change; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ChangingSortColumn,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ChangingSortColumn, Raise_ChangedSortColumn
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ChangingSortColumn,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ChangingSortColumn, Raise_ChangedSortColumn
	/// \endif
	inline HRESULT Raise_ChangingSortColumn(LONG previousSortColumnIndex, LONG newSortColumnIndex, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ColumnEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose columns have been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationCompleted,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationCompleted,
	///       Raise_ColumnEnumerationStarted, Raise_ColumnEnumerationTimedOut, Raise_ItemEnumerationCompleted
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationCompleted,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationCompleted,
	///       Raise_ColumnEnumerationStarted, Raise_ColumnEnumerationTimedOut, Raise_ItemEnumerationCompleted
	/// \endif
	inline HRESULT Raise_ColumnEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted);
	/// \brief <em>Raises the \c ColumnEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose columns are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationStarted,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationStarted,
	///       Raise_ColumnEnumerationTimedOut, Raise_ColumnEnumerationCompleted, Raise_ItemEnumerationStarted
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationStarted,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationStarted,
	///       Raise_ColumnEnumerationTimedOut, Raise_ColumnEnumerationCompleted, Raise_ItemEnumerationStarted
	/// \endif
	HRESULT Raise_ColumnEnumerationStarted(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c ColumnEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose columns are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationTimedOut,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ColumnEnumerationTimedOut,
	///       Raise_ColumnEnumerationStarted, Raise_ColumnEnumerationCompleted, Raise_ItemEnumerationTimedOut
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ColumnEnumerationTimedOut,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ColumnEnumerationTimedOut,
	///       Raise_ColumnEnumerationStarted, Raise_ColumnEnumerationCompleted, Raise_ItemEnumerationTimedOut
	/// \endif
	inline HRESULT Raise_ColumnEnumerationTimedOut(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c CreatedHeaderContextMenu event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] minimumCustomCommandID The lowest command ID that may be used for custom menu items.
	///            For custom menu items any command ID larger or equal this value and smaller or equal
	///            65535 may be used.
	/// \param[in,out] pCancelMenu If \c VARIANT_FALSE, the caller should not display the context menu;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_CreatedHeaderContextMenu,
	///       ShBrowserCtlsLibU::_IShListViewEvents::CreatedHeaderContextMenu,
	///       Raise_DestroyingHeaderContextMenu, Raise_SelectedHeaderContextMenuItem,
	///       Raise_CreatedShellContextMenu
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_CreatedHeaderContextMenu,
	///       ShBrowserCtlsLibA::_IShListViewEvents::CreatedHeaderContextMenu,
	///       Raise_DestroyingHeaderContextMenu, Raise_SelectedHeaderContextMenuItem,
	///       Raise_CreatedShellContextMenu
	/// \endif
	inline HRESULT Raise_CreatedHeaderContextMenu(OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu);
	/// \brief <em>Raises the \c CreatedShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] minimumCustomCommandID The lowest command ID that may be used for custom menu items.
	///            For custom menu items any command ID larger or equal this value and smaller or equal
	///            65535 may be used.
	/// \param[in,out] pCancelMenu If \c VARIANT_FALSE, the caller should not display the context menu;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Background context menus require that the control contains exactly 1 namespace.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_CreatedShellContextMenu,
	///       ShBrowserCtlsLibU::_IShListViewEvents::CreatedShellContextMenu, Raise_CreatingShellContextMenu,
	///       Raise_SelectedShellContextMenuItem, Raise_CreatedHeaderContextMenu, get_Namespaces
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_CreatedShellContextMenu,
	///       ShBrowserCtlsLibA::_IShListViewEvents::CreatedShellContextMenu, Raise_CreatingShellContextMenu,
	///       Raise_SelectedShellContextMenuItem, Raise_CreatedHeaderContextMenu, get_Namespaces
	/// \endif
	inline HRESULT Raise_CreatedShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu);
	/// \brief <em>Raises the \c CreatingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in,out] pContextMenuStyle A bit field influencing the content of the context menu. Any
	///                combination of the values defined by the \c ShellContextMenuStyleConstants
	///                enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_FALSE, the caller should not create and display the context
	///                menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Background context menus require that the control contains exactly 1 namespace.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_CreatingShellContextMenu,
	///       ShBrowserCtlsLibU::_IShListViewEvents::CreatingShellContextMenu, Raise_CreatedShellContextMenu,
	///       ShBrowserCtlsLibU::ShellContextMenuStyleConstants, get_Namespaces
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_CreatingShellContextMenu,
	///       ShBrowserCtlsLibA::_IShListViewEvents::CreatingShellContextMenu, Raise_CreatedShellContextMenu,
	///       ShBrowserCtlsLibA::ShellContextMenuStyleConstants, get_Namespaces
	/// \endif
	inline HRESULT Raise_CreatingShellContextMenu(LPDISPATCH pItems, ShellContextMenuStyleConstants* pContextMenuStyle, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c DestroyingHeaderContextMenu event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_DestroyingHeaderContextMenu,
	///       ShBrowserCtlsLibU::_IShListViewEvents::DestroyingHeaderContextMenu,
	///       Raise_CreatedHeaderContextMenu, Raise_DestroyingShellContextMenu
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_DestroyingHeaderContextMenu,
	///       ShBrowserCtlsLibA::_IShListViewEvents::DestroyingHeaderContextMenu,
	///       Raise_CreatedHeaderContextMenu, Raise_DestroyingShellContextMenu
	/// \endif
	inline HRESULT Raise_DestroyingHeaderContextMenu(OLE_HANDLE hContextMenu);
	/// \brief <em>Raises the \c DestroyingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Background context menus require that the control contains exactly 1 namespace.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_DestroyingShellContextMenu,
	///       ShBrowserCtlsLibU::_IShListViewEvents::DestroyingShellContextMenu,
	///       Raise_CreatingShellContextMenu, Raise_DestroyingHeaderContextMenu, get_Namespaces
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_DestroyingShellContextMenu,
	///       ShBrowserCtlsLibA::_IShListViewEvents::DestroyingShellContextMenu,
	///       Raise_CreatingShellContextMenu, Raise_DestroyingHeaderContextMenu, get_Namespaces
	/// \endif
	inline HRESULT Raise_DestroyingShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu);
	/// \brief <em>Raises the \c InsertedItem event</em>
	///
	/// \param[in] pListItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_InsertedItem,
	///       ShBrowserCtlsLibU::_IShListViewEvents::InsertedItem, Raise_InsertingItem,
	///       Raise_InsertedNamespace, Raise_LoadedColumn, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_InsertedItem,
	///       ShBrowserCtlsLibA::_IShListViewEvents::InsertedItem, Raise_InsertingItem,
	///       Raise_InsertedNamespace, Raise_LoadedColumn, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertedItem(IShListViewItem* pListItem);
	/// \brief <em>Raises the \c InsertedNamespace event</em>
	///
	/// \param[in] pNamespace The inserted namespace.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_InsertedNamespace,
	///       ShBrowserCtlsLibU::_IShListViewEvents::InsertedNamespace, Raise_InsertedItem,
	///       Raise_LoadedColumn, Raise_RemovingNamespace
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_InsertedNamespace,
	///       ShBrowserCtlsLibA::_IShListViewEvents::InsertedNamespace, Raise_InsertedItem,
	///       Raise_LoadedColumn, Raise_RemovingNamespace
	/// \endif
	inline HRESULT Raise_InsertedNamespace(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c InsertingItem event</em>
	///
	/// \param[in] fullyQualifiedPIDL The fully qualified pIDL of the item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_InsertingItem,
	///       ShBrowserCtlsLibU::_IShListViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_InsertingItem,
	///       ShBrowserCtlsLibA::_IShListViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertingItem(LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion);
	/// \brief <em>Raises the \c InvokedHeaderContextMenuCommand event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_InvokedHeaderContextMenuCommand,
	///       ShBrowserCtlsLibU::_IShListViewEvents::InvokedHeaderContextMenuCommand,
	///       Raise_SelectedHeaderContextMenuItem, Raise_InvokedShellContextMenuCommand
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_InvokedHeaderContextMenuCommand,
	///       ShBrowserCtlsLibA::_IShListViewEvents::InvokedHeaderContextMenuCommand,
	///       Raise_SelectedHeaderContextMenuItem, Raise_InvokedShellContextMenuCommand
	/// \endif
	inline HRESULT Raise_InvokedHeaderContextMenuCommand(OLE_HANDLE hContextMenu, LONG commandID);
	/// \brief <em>Raises the \c InvokedShellContextMenuCommand event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	/// \param[in] usedWindowMode A value specifying how to display the window that may be opened when
	///            executing the command. Any of the values defined by the \c WindowModeConstants enumeration
	///            is valid.
	/// \param[in] usedInvocationFlags A bit field controlling command execution. Any combination of the
	///            values defined by the \c CommandInvocationFlagsConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Background context menus require that the control contains exactly 1 namespace.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_InvokedShellContextMenuCommand,
	///       ShBrowserCtlsLibU::_IShListViewEvents::InvokedShellContextMenuCommand,
	///       Raise_SelectedShellContextMenuItem, Raise_InvokedHeaderContextMenuCommand,
	///       ShBrowserCtlsLibU::WindowModeConstants, ShBrowserCtlsLibU::CommandInvocationFlagsConstants,
	///       get_Namespaces
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_InvokedShellContextMenuCommand,
	///       ShBrowserCtlsLibA::_IShListViewEvents::InvokedShellContextMenuCommand,
	///       Raise_SelectedShellContextMenuItem, Raise_InvokedHeaderContextMenuCommand,
	///       ShBrowserCtlsLibA::WindowModeConstants, ShBrowserCtlsLibA::CommandInvocationFlagsConstants,
	///       get_Namespaces
	/// \endif
	inline HRESULT Raise_InvokedShellContextMenuCommand(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants usedWindowMode, CommandInvocationFlagsConstants usedInvocationFlags);
	/// \brief <em>Raises the \c ItemEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose items have been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationCompleted,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationCompleted, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationTimedOut, Raise_ColumnEnumerationCompleted
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationCompleted,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationCompleted, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationTimedOut, Raise_ColumnEnumerationCompleted
	/// \endif
	inline HRESULT Raise_ItemEnumerationCompleted(IShListViewNamespace* pNamespace, VARIANT_BOOL aborted);
	/// \brief <em>Raises the \c ItemEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationStarted,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationStarted, Raise_ItemEnumerationTimedOut,
	///       Raise_ItemEnumerationCompleted, Raise_ColumnEnumerationStarted
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationStarted,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationStarted, Raise_ItemEnumerationTimedOut,
	///       Raise_ItemEnumerationCompleted, Raise_ColumnEnumerationStarted
	/// \endif
	HRESULT Raise_ItemEnumerationStarted(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c ItemEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationTimedOut,
	///       ShBrowserCtlsLibU::_IShListViewEvents::ItemEnumerationTimedOut, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationCompleted, Raise_ColumnEnumerationTimedOut
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_ItemEnumerationTimedOut,
	///       ShBrowserCtlsLibA::_IShListViewEvents::ItemEnumerationTimedOut, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationCompleted, Raise_ColumnEnumerationTimedOut
	/// \endif
	inline HRESULT Raise_ItemEnumerationTimedOut(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c LoadedColumn event</em>
	///
	/// \param[in] pColumn The column that has been loaded.
	/// \param[in,out] pMakeVisible If \c VARIANT_TRUE, the caller should insert the column into the listview
	///                control; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_LoadedColumn,
	///       ShBrowserCtlsLibU::_IShListViewEvents::LoadedColumn, Raise_UnloadedColumn,
	///       Raise_ChangingColumnVisibility, Raise_InsertedItem, Raise_InsertedNamespace
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_LoadedColumn,
	///       ShBrowserCtlsLibA::_IShListViewEvents::LoadedColumn, Raise_UnloadedColumn,
	///       Raise_ChangingColumnVisibility, Raise_InsertedItem, Raise_InsertedNamespace
	/// \endif
	inline HRESULT Raise_LoadedColumn(IShListViewColumn* pColumn, VARIANT_BOOL* pMakeVisible);
	/// \brief <em>Raises the \c RemovingItem event</em>
	///
	/// \param[in] pListItem The item being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_RemovingItem,
	///       ShBrowserCtlsLibU::_IShListViewEvents::RemovingItem, Raise_RemovingNamespace,
	///       Raise_UnloadedColumn, Raise_InsertedItem
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_RemovingItem,
	///       ShBrowserCtlsLibA::_IShListViewEvents::RemovingItem, Raise_RemovingNamespace,
	///       Raise_UnloadedColumn, Raise_InsertedItem
	/// \endif
	inline HRESULT Raise_RemovingItem(IShListViewItem* pListItem);
	/// \brief <em>Raises the \c RemovingNamespace event</em>
	///
	/// \param[in] pNamespace The namespace being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_RemovingNamespace,
	///       ShBrowserCtlsLibU::_IShListViewEvents::RemovingNamespace, Raise_RemovingItem,
	///       Raise_UnloadedColumn, Raise_InsertedNamespace
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_RemovingNamespace,
	///       ShBrowserCtlsLibA::_IShListViewEvents::RemovingNamespace, Raise_RemovingItem,
	///       Raise_UnloadedColumn, Raise_InsertedNamespace
	/// \endif
	inline HRESULT Raise_RemovingNamespace(IShListViewNamespace* pNamespace);
	/// \brief <em>Raises the \c SelectedHeaderContextMenuItem event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	/// \param[in,out] pExecuteCommand If set to \c VARIANT_TRUE, the caller should execute the menu item;
	///                otherwise not.\n
	///                If the menu item is a custom one, this value should be ignored and always be treated
	///                as being set to \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_SelectedHeaderContextMenuItem,
	///       ShBrowserCtlsLibU::_IShListViewEvents::SelectedHeaderContextMenuItem,
	///       Raise_InvokedHeaderContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedHeaderContextMenu, Raise_SelectedShellContextMenuItem
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_SelectedHeaderContextMenuItem,
	///       ShBrowserCtlsLibA::_IShListViewEvents::SelectedHeaderContextMenuItem,
	///       Raise_InvokedHeaderContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedHeaderContextMenu, Raise_SelectedShellContextMenuItem
	/// \endif
	inline HRESULT Raise_SelectedHeaderContextMenuItem(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL* pExecuteCommand);
	/// \brief <em>Raises the \c SelectedShellContextMenuItem event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a IListViewItemContainer or
	///            a \c IShListViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	/// \param[in] commandID The command ID identifying the selected command.
	/// \param[in,out] pWindowModeToUse A value specifying how to display the window that may be opened when
	///                executing the command. Any of the values defined by the \c WindowModeConstants
	///                enumeration is valid.
	/// \param[in,out] pInvocationFlagsToUse A bit field controlling command execution. Any combination of
	///                the values defined by the \c CommandInvocationFlagsConstants enumeration is valid.
	/// \param[in,out] pExecuteCommand If set to \c VARIANT_TRUE, the caller should execute the menu item;
	///                otherwise not.\n
	///                If the menu item is a custom one, this value should be ignored and always be treated
	///                as being set to \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Background context menus require that the control contains exactly 1 namespace.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_SelectedShellContextMenuItem,
	///       ShBrowserCtlsLibU::_IShListViewEvents::SelectedShellContextMenuItem,
	///       Raise_InvokedShellContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedShellContextMenu, Raise_SelectedHeaderContextMenuItem,
	///       ShBrowserCtlsLibU::WindowModeConstants, ShBrowserCtlsLibU::CommandInvocationFlagsConstants,
	///       get_Namespaces
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_SelectedShellContextMenuItem,
	///       ShBrowserCtlsLibA::_IShListViewEvents::SelectedShellContextMenuItem,
	///       Raise_InvokedShellContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedShellContextMenu, Raise_SelectedHeaderContextMenuItem,
	///       ShBrowserCtlsLibA::WindowModeConstants, ShBrowserCtlsLibA::CommandInvocationFlagsConstants,
	///       get_Namespaces
	/// \endif
	inline HRESULT Raise_SelectedShellContextMenuItem(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants* pWindowModeToUse, CommandInvocationFlagsConstants* pInvocationFlagsToUse, VARIANT_BOOL* pExecuteCommand);
	/// \brief <em>Raises the \c UnloadedColumn event</em>
	///
	/// \param[in] pColumn The column that has been unloaded. If \c Nothing, all columns have been unloaded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShListViewEvents::Fire_UnloadedColumn,
	///       ShBrowserCtlsLibU::_IShListViewEvents::UnloadedColumn, Raise_LoadedColumn, Raise_RemovingItem,
	///       Raise_RemovingNamespace
	/// \else
	///   \sa Proxy_IShListViewEvents::Fire_UnloadedColumn,
	///       ShBrowserCtlsLibA::_IShListViewEvents::UnloadedColumn, Raise_LoadedColumn, Raise_RemovingItem,
	///       Raise_RemovingNamespace
	/// \endif
	inline HRESULT Raise_UnloadedColumn(IShListViewColumn* pColumn);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);

	/// \brief <em>Configures the attached control</em>
	///
	/// \sa Attach
	void ConfigureControl(void);
	/// \brief <em>Retrieves the size in pixels in which thumbnails should be displayed</em>
	///
	/// \param[in] userSetSize The thumbnail size in pixels as set by the user.
	///
	/// \return The thumbnail size in pixels.
	///
	/// \sa SetSystemImageLists, get_DisplayThumbnails
	SIZE GetThumbnailsSize(long userSetSize);
	/// \brief <em>Makes the attached control use the system imagelists as specified by the \c UseSystemImageList property</em>
	///
	/// \sa GetThumbnailsSize, ConfigureControl, get_UseSystemImageList
	void SetSystemImageLists(void);
	/// \brief <em>Configures the unified image list for the current view mode</em>
	///
	/// \sa Properties::pUnifiedImageList
	void SetupUnifiedImageListForCurrentView(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Item management
	///
	//@{
	/// \brief <em>Adds the specified listview item to the collection of listview items</em>
	///
	/// \param[in] itemID The unique ID of the listview item to add.
	/// \param[in] pItemData The data of the listview item to add.
	///
	/// \sa AddNamespace, ChangedColumnVisibility, Properties::items, OnInternalInsertedItem
	void AddItem(LONG itemID, LPSHELLLISTVIEWITEMDATA pItemData);
	/// \brief <em>Applies a shell listview item's \c ManagedProperties property to the attached listview</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to apply the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsShellItem, ShellListViewItem::get_ManagedProperties
	HRESULT ApplyManagedProperties(LONG itemID);
	/// \brief <em>Buffers an item's \c SHELLLISTVIEWITEMDATA struct until it is required by \c OnInternalInsertedItem</em>
	///
	/// \param[in] pItemData The data to buffer.
	///
	/// \sa Properties::shellItemDataOfInsertedItems, OnInternalInsertedItem, SHELLLISTVIEWITEMDATA,
	///     ShellListViewItems::Add
	void BufferShellItemData(LPSHELLLISTVIEWITEMDATA pItemData);
	/// \brief <em>Inserts a new shell item faster than \c ShellListViewItems::Add</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item to insert.
	/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted as the
	///            last item.
	///
	/// \return The new item's unique ID.
	///
	/// \sa OnTriggerItemEnumComplete
	LONG FastInsertShellItem(PIDLIST_ABSOLUTE pIDL, LONG insertAt);
	/// \brief <em>Retrieves the specified shell listview item's details</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to retrieve the details.
	///
	/// \return The item's details or \c NULL if the item isn't a shell item.
	///
	/// \sa ItemIDToPIDL, SHELLLISTVIEWITEMDATA, IsShellItem
	LPSHELLLISTVIEWITEMDATA GetItemDetails(LONG itemID);
	/// \brief <em>Retrieves whether the specified listview item is a shell listview item</em>
	///
	/// \param[in] itemID The unique ID of the listview item to check.
	///
	/// \return \c TRUE if the item is a shell listview item; otherwise \c FALSE.
	///
	/// \sa GetItemDetails, ItemIDToPIDL, ApplyManagedProperties
	BOOL IsShellItem(LONG itemID);
	/// \brief <em>Retrieves the specified shell listview item's zero-based index</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to retrieve the index.
	///
	/// \return The item's zero-based index or -1 if an error occured.
	///
	/// \sa ItemIndexToID, EXLVM_ITEMIDTOINDEX
	int ItemIDToIndex(LONG itemID);
	/// \brief <em>Retrieves the specified shell listview item's unique ID</em>
	///
	/// \param[in] itemIndex The zero-based index of the listview item for which to retrieve the unique ID.
	///
	/// \return The item's unique ID or -1 if an error occured.
	///
	/// \sa ItemIDToIndex, EXLVM_ITEMINDEXTOID
	LONG ItemIndexToID(int itemIndex);
	/// \brief <em>Retrieves the specified shell listview item's fully qualified pIDL</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to retrieve the pIDL.
	///
	/// \return The item's fully qualified pIDL or \c NULL if the item isn't a shell item.
	///
	/// \sa ItemIDsToPIDLs, ItemIDToNamespace, PIDLToItemID, GetItemDetails, IsShellItem
	PCIDLIST_ABSOLUTE ItemIDToPIDL(LONG itemID);
	/// \brief <em>Retrieves the specified shell listview items' fully qualified pIDLs</em>
	///
	/// \param[in] pItemIDs The unique IDs of the listview items for which to retrieve the pIDLs.
	/// \param[in] itemCount The number of elements in the array \c pItemIDs.
	/// \param[out] ppIDLs Receives the pIDLs of the listview items.
	/// \param[in] keepNonShellItems If \c TRUE, the pIDLs of listview items that are no shell items, are set
	///            to \c NULL. Otherwise \c ppIDLs does not contain these items.
	///
	/// \return The number of elements in the array returned in \c ppIDLs.
	///
	/// \sa ItemIDToPIDL
	UINT ItemIDsToPIDLs(PLONG pItemIDs, UINT itemCount, PIDLIST_ABSOLUTE*& ppIDLs, BOOL keepNonShellItems);
	/// \brief <em>Retrieves the shell namespace that the specified shell listview item belongs to</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to retrieve the namespace.
	///
	/// \return A \c ShellListViewNamespace object wrapping the shell namespace the item belongs to or
	///         \c NULL if the item isn't a shell item.
	///
	/// \sa ItemIDToNamespacePIDL, ItemIDToPIDL, GetItemDetails
	CComObject<ShellListViewNamespace>* ItemIDToNamespace(LONG itemID);
	/// \brief <em>Retrieves the fully qualified pIDL of the shell namespace that the specified shell listview item belongs to</em>
	///
	/// \param[in] itemID The unique ID of the listview item for which to retrieve the namespace pIDL.
	///
	/// \return The fully qualified pIDL of the shell namespace the item belongs to or \c NULL if the item
	///         isn't a shell item.
	///
	/// \sa ItemIDToNamespace, ItemIDToPIDL, GetItemDetails
	PCIDLIST_ABSOLUTE ItemIDToNamespacePIDL(LONG itemID);
	/// \brief <em>Retrieves the item ID of the specified fully qualified pIDL</em>
	///
	/// \param[in] pIDL The fully qualified pIDL for which to retrieve the shell item's ID.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The item ID or -1 if no matching item was found.
	///
	/// \sa ParsingNameToItemID, VariantToItemIDs, ItemIDToPIDL
	LONG PIDLToItemID(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch);
	/// \brief <em>Retrieves the item ID of the specified fully qualified parsing name</em>
	///
	/// \param[in] pParsingName The fully qualified parsing name of the shell item for which to retrieve the
	///            shell item's ID.
	///
	/// \return The item ID or -1 if no matching item was found.
	///
	/// \sa PIDLToItemID, VariantToItemIDs, ItemIDToPIDL
	LONG ParsingNameToItemID(LPWSTR pParsingName);
	/// \brief <em>Retrieves the item IDs of the specified listview items</em>
	///
	/// \param[in] items A \c VARIANT specifying the items for which to retrieve the IDs. The \c VARIANT may
	///            contain the following data:
	///            - A single item ID.
	///            - An array of item IDs.
	///            - A \c ListViewItem object.
	///            - A \c ListViewItems object.
	///            - A \c ListViewItemContainer object.
	///            - A \c ShellListViewItem object.
	///            - A \c ShellListViewItems object.
	/// \param[out] pItemIDs Receives the item IDs.
	/// \param[out] itemCount Receives the number of elements in the array returned in \c pItemIDs.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PIDLToItemID
	HRESULT VariantToItemIDs(VARIANT& items, PLONG& pItemIDs, UINT& itemCount);
	/// \brief <em>Calls \c SHBindToParent for the specified shell listview item</em>
	///
	/// \param[in] itemID The shell listview item for which to call \c SHBindToParent.
	/// \param[in] requiredInterface The \c riid parameter of \c SHBindToParent.
	/// \param[out] ppInterfaceImpl The \c ppv parameter of \c SHBindToParent.
	/// \param[out] pPIDLLast The \c ppidlLast parameter of \c SHBindToParent.
	/// \param[out] pPIDL The item's fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms647661.aspx">SHBindToParent</a>
	HRESULT SHBindToParent(LONG itemID, REFIID requiredInterface, LPVOID* ppInterfaceImpl, PCUITEMID_CHILD* pPIDLLast, PCIDLIST_ABSOLUTE* pPIDL = NULL);
	/// \brief <em>Checks whether the specified shell item still exists</em>
	///
	/// \param[in] itemID The unique ID of the listview item to validate.
	/// \param[in] pItemDetails Optional. The listview item's details. If set to \c NULL (the default), the
	///            details are retrieved by the method.
	///
	/// \return \c TRUE if the item still exists; otherwise \c FALSE.
	///
	/// \sa ShellListViewItem::Validate, ValidatePIDL
	BOOL ValidateItem(LONG itemID, LPSHELLLISTVIEWITEMDATA pItemDetails = NULL);
	/// \brief <em>Updates the specified item's pIDL</em>
	///
	/// \param[in] itemID The unique ID of the listview item to update.
	/// \param[in] pItemDetails The listview item's details.
	/// \param[in] pIDLNew The fully qualified new pIDL. If the new pIDL consists of as many item ids as the
	///            item's current pIDL, the pIDL is simply exchanged. If it contains less item ids, the head
	///            of the item's current pIDL is replaced with the new pIDL.
	///
	/// \remarks This method does not call \c ApplyManagedProperties.
	///
	/// \sa Raise_ChangedItemPIDL, ApplyManagedProperties
	void UpdateItemPIDL(LONG itemID, __in LPSHELLLISTVIEWITEMDATA pItemDetails, __in PIDLIST_ABSOLUTE pIDLNew);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Namespace management
	///
	//@{
	/// \brief <em>Adds the specified namespace to the collection of namespaces</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the namespace to add.
	/// \param[in] pNamespace The namespace to add.
	///
	/// \sa AddItem, ChangedColumnVisibility, Properties::namespaces, ShellListViewNamespaces::Add
	void AddNamespace(PCIDLIST_ABSOLUTE pIDL, CComObject<ShellListViewNamespace>* pNamespace);
	/// \brief <em>Buffers an item's associated namespace pIDL until it is required by \c OnInternalInsertedItem</em>
	///
	/// \param[in] pIDLItem The fully qualified pIDL of the item for which to buffer the namespace pIDL.
	/// \param[in] pIDLNamespace The fully qualified pIDL of the item's namespace.
	///
	/// \sa Properties::namespacesOfInsertedItems, OnInternalInsertedItem, ShellListViewNamespaces::Add,
	///     RemoveBufferedShellItemNamespace, BufferShellItemData
	void BufferShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem, PCIDLIST_ABSOLUTE pIDLNamespace);
	/// \brief <em>Deletes an item's buffered namespace pIDL from the buffer</em>
	///
	/// \param[in] pIDLItem The fully qualified pIDL of the item.
	///
	/// \sa Properties::namespacesOfInsertedItems, OnInternalInsertedItem, ShellListViewNamespaces::Add,
	///     BufferShellItemNamespace, BufferShellItemData
	void RemoveBufferedShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem);
	/// \brief <em>Retrieves the n-th \c ShellListViewNamespace object</em>
	///
	/// \param[in] index The zero-based index for which to retrieve the shell namespace object.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespacePIDL, NamespacePIDLToIndex, NamespacePIDLToNamespace,
	///     NamespaceParsingNameToNamespace
	CComObject<ShellListViewNamespace>* IndexToNamespace(UINT index);
	/// \brief <em>Retrieves the fully qualified pIDL of the n-th \c ShellListViewNamespace object</em>
	///
	/// \param[in] index The zero-based index for which to retrieve the shell namespace object.
	///
	/// \return The namespace object's pIDL or NULL if no matching namespace was found.
	///
	/// \sa NamespacePIDLToIndex, IndexToNamespace, NamespacePIDLToNamespace
	PCIDLIST_ABSOLUTE IndexToNamespacePIDL(UINT index);
	/// \brief <em>Retrieves the zero-based index of the namespace object identified by the specified fully qualified pIDL</em>
	///
	/// \param[in] pIDL The fully qualified pIDL for which to retrieve the index.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The namespace's zero-based index or -1 if no matching namespace was found.
	///
	/// \sa IndexToNamespacePIDL
	LONG NamespacePIDLToIndex(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch);
	/// \brief <em>Retrieves the \c ShellListViewNamespace object of the specified fully qualified pIDL</em>
	///
	/// \param[in] pIDL The fully qualified pIDL for which to retrieve the shell namespace object.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespace, NamespaceParsingNameToNamespace
	CComObject<ShellListViewNamespace>* NamespacePIDLToNamespace(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch);
	/// \brief <em>Retrieves the \c ShellListViewNamespace object of the specified fully qualified parsing name</em>
	///
	/// \param[in] pParsingName The fully qualified parsing name of the shell item for which to retrieve the
	///            shell namespace object.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespace, NamespacePIDLToNamespace
	CComObject<ShellListViewNamespace>* NamespaceParsingNameToNamespace(LPWSTR pParsingName);
	/// \brief <em>Removes a shell namespace</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the shell namespace to remove.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	/// \param[in] removeItemsFromListView If \c TRUE, the namespace's items are removed from the attached
	///            listview; otherwise they are converted to normal items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RemoveAllNamespaces, get_Namespaces, ShellListViewNamespaces::RemoveAll
	HRESULT RemoveNamespace(PCIDLIST_ABSOLUTE pIDLNamespace, BOOL exactMatch, BOOL removeItemsFromListView);
	/// \brief <em>Removes all shell namespaces</em>
	///
	/// \param[in] removeItemsFromListView If \c TRUE, the namespaces' items are removed from the attached
	///            listview; otherwise they are converted to normal items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RemoveNamespace, get_Namespaces, ShellListViewNamespaces::RemoveAll
	HRESULT RemoveAllNamespaces(BOOL removeItemsFromListView);
	/// \brief <em>Closes any open thumbnail disk caches that have not recently been accessed</em>
	///
	/// \param[in] forceClose If \c TRUE, the caches are closed even if they have been accessed recently.
	///
	/// \remarks Beginning with Windows Vista, this member isn't required anymore.
	///
	/// \sa OnReportThumbnailDiskCacheAccess, OnTimer
	void LazyCloseThumbnailDiskCaches(BOOL forceClose = FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Column management
	///
	//@{
	/// \brief <em>Retrieves whether the attached listview control is currently set up to display columns</em>
	///
	/// \return \c TRUE if the attached listview control is set up to display columns; otherwise \c FALSE.
	///
	/// \sa EnsureShellColumnsAreLoaded, IsInDetailsView
	BOOL CurrentViewNeedsColumns(void);
	/// \brief <em>Retrieves whether the attached listview control is currently in Details view mode</em>
	///
	/// \return \c TRUE if the attached listview control is in Details view mode; otherwise \c FALSE.
	///
	/// \sa IsInTilesView, IsInSmallIconsView, CurrentViewNeedsColumns
	BOOL IsInDetailsView(void);
	/// \brief <em>Retrieves whether the attached listview control is currently in Small Icons view mode</em>
	///
	/// \return \c TRUE if the attached listview control is in Small Icons view mode; otherwise \c FALSE.
	///
	/// \sa IsInDetailsView, IsInTilesView
	BOOL IsInSmallIconsView(void);
	/// \brief <em>Retrieves whether the attached listview control is currently in Tiles view mode</em>
	///
	/// \return \c TRUE if the attached listview control is in Tiles view mode; otherwise \c FALSE.
	///
	/// \sa IsInDetailsView, IsInSmallIconsView, CurrentViewNeedsColumns
	BOOL IsInTilesView(void);
	/// \brief <em>Sets the namespace that acts as the main namespace</em>
	///
	/// Sets the namespace that acts as the main namespace, i.e. its the namespace whose columns are loaded
	/// and which receives special treatment for auto-update.
	///
	/// \param[in] pIDLNamespace The fully qualified pIDL of the namespace whose columns shall be
	///            displayed. May be \c NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnsureShellColumnsAreLoaded, UnloadShellColumns
	HRESULT SetMainNamespace(PCIDLIST_ABSOLUTE pIDLNamespace);
	/// \brief <em>Ensures that the columns of the main namespace are loaded</em>
	///
	/// \param[in] threadingMode Controls whether the columns are enumerated synchronously or asynchronously.
	///            Any of the values defined by the \c ThreadingMode enumeration except
	///            \c tmImmediateThreading is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CurrentViewNeedsColumns, SetMainNamespace, UnloadShellColumns, ThreadingMode
	HRESULT EnsureShellColumnsAreLoaded(ThreadingMode threadingMode);
	/// \brief <em>Unloads the currently loaded shell commands</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnsureShellColumnsAreLoaded, SetMainNamespace
	HRESULT UnloadShellColumns(void);
	/// \brief <em>Calculates the average width of a character in the attached listview's column header</em>
	///
	/// \return The average character width.
	///
	/// \sa EnsureShellColumnsAreLoaded
	int GetAverageColumnHeaderCharWidth(void);
	/// \brief <em>Buffers a column's index until it is required by \c OnInternalInsertedColumn</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	///
	/// \sa Properties::shellColumnIndexesOfInsertedColumns, ShellColumnsStatus::pAllColumns,
	///     OnInternalInsertedColumn
	void BufferShellColumnRealIndex(int realColumnIndex);
	/// \brief <em>Makes a shell column visible or invisible, i. e. inserts it into or removes it from the attached listview control</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	/// \param[in] makeVisible If \c TRUE, the column shall be inserted; otherwise it shall be removed.
	/// \param[in] changeColumnVisibilityflags A bit-field of \c CCVF_* flags that influence the method's
	///            behavior.
	///
	/// \return The new column's unique ID.
	///
	/// \sa OnTriggerColumnEnumComplete, ChangedColumnVisibility, ShellColumnsStatus::pAllColumns,
	///     UnloadShellColumns, CCVF_ISEXPLICITCHANGE, CCVF_ISEXPLICITCHANGEIFDIFFERENT,
	///     CCVF_FORUNLOADSHELLCOLUMNS
	LONG ChangeColumnVisibility(int realColumnIndex, BOOL makeVisible, DWORD changeColumnVisibilityflags);
	/// \brief <em>Adds the specified listview column to the collection of visible listview columns or removes it from there</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	/// \param[in] newColumnID The unique ID of the listview column that was made visible or invisible.
	/// \param[in] madeVisible If \c TRUE, the column has been made visible, otherwise invisible.
	///
	/// \sa AddItem, AddNamespace, Properties::visibleColumns, ShellColumnsStatus::pAllColumns,
	///     OnInternalInsertedColumn
	void ChangedColumnVisibility(int realColumnIndex, LONG newColumnID, BOOL madeVisible);
	/// \brief <em>Retrieves the specified shell listview column's details</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	///
	/// \return The column's details or \c NULL if an error occurs.
	///
	/// \sa SHELLLISTVIEWCOLUMNDATA
	LPSHELLLISTVIEWCOLUMNDATA GetColumnDetails(int realColumnIndex);
	/// \brief <em>Retrieves the specified shell column's property key</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	/// \param[out] pPropertyKey Receives the column's property key.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewColumn::get_FormatIdentifier, ShellListViewColumn::get_PropertyIdentifier,
	///     GetColumnPropertyDescription, PropertyKeyToRealIndex,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	HRESULT GetColumnPropertyKey(int realColumnIndex, __in PROPERTYKEY* pPropertyKey);
	/// \brief <em>Retrieves an object describing the property displayed in the specified shell column</em>
	///
	/// \param[in] realColumnIndex The column's zero-based index within the
	///            \c ShellColumnsStatus::pAllColumns array.
	/// \param[in] requiredInterface The interface that the returned object must implement.
	/// \param[out] ppInterfaceImpl Receives a pointer to the object describing the property. This object can
	///             be accessed through the queried interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetColumnPropertyKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761561.aspx">IPropertyDescription</a>
	HRESULT GetColumnPropertyDescription(int realColumnIndex, __in REFIID requiredInterface, __in LPVOID* ppInterfaceImpl);
	/// \brief <em>Retrieves whether the specified listview column is a shell listview column</em>
	///
	/// \param[in] columnID The unique ID of the listview column to check.
	///
	/// \return \c TRUE if the column is a shell listview column; otherwise \c FALSE.
	///
	/// \sa GetColumnDetails
	BOOL IsShellColumn(LONG columnID);
	/// \brief <em>Retrieves the specified shell listview column's zero-based index</em>
	///
	/// \param[in] columnID The unique ID of the listview column for which to retrieve the index.
	///
	/// \return The column's zero-based index or -1 if an error occured.
	///
	/// \sa ColumnIndexToID, ColumnIDToRealIndex, RealColumnIndexToIndex, EXLVM_COLUMNIDTOINDEX
	int ColumnIDToIndex(LONG columnID);
	/// \brief <em>Retrieves the specified shell listview column's zero-based index within the \c ShellColumnsStatus::pAllColumns array</em>
	///
	/// \param[in] columnID The unique ID of the listview column for which to retrieve the index.
	///
	/// \return The column's zero-based index within the \c ShellColumnsStatus::pAllColumns array.
	///
	/// \sa RealColumnIndexToID, ColumnIDToIndex, ColumnIndexToRealIndex, ShellColumnsStatus::pAllColumns
	int ColumnIDToRealIndex(LONG columnID);
	/// \brief <em>Retrieves the specified shell listview column's unique ID</em>
	///
	/// \param[in] columnIndex The zero-based index of the listview column for which to retrieve the unique
	///            ID.
	///
	/// \return The column's unique ID or -1 if an error occured.
	///
	/// \sa ColumnIDToIndex, ColumnIndexToRealIndex, RealColumnIndexToID, EXLVM_COLUMNINDEXTOID
	LONG ColumnIndexToID(int columnIndex);
	/// \brief <em>Retrieves the specified shell listview column's zero-based index within the \c ShellColumnsStatus::pAllColumns array</em>
	///
	/// \param[in] columnIndex The zero-based index of the listview column for which to retrieve the index.
	///
	/// \return The column's zero-based index within the \c ShellColumnsStatus::pAllColumns array.
	///
	/// \sa RealColumnIndexToIndex, ColumnIndexToID, ColumnIDToRealIndex, ShellColumnsStatus::pAllColumns
	int ColumnIndexToRealIndex(int columnIndex);
	/// \brief <em>Retrieves the specified shell column's zero-based index within the \c ShellColumnsStatus::pAllColumns array</em>
	///
	/// \param[in] propertyKey The property key specifying the shell column for which to retrieve the index.
	///
	/// \return The column's zero-based index within the \c ShellColumnsStatus::pAllColumns array.
	///
	/// \sa ColumnIndexToRealIndex, GetColumnPropertyKey, ShellColumnsStatus::pAllColumns
	int PropertyKeyToRealIndex(const PROPERTYKEY& propertyKey);
	/// \brief <em>Retrieves the specified shell listview column's unique ID</em>
	///
	/// \param[in] realColumnIndex The zero-based index (within the \c ShellColumnsStatus::pAllColumns array)
	///            of the listview column for which to retrieve the unique ID.
	///
	/// \return The column's unique ID or -1 if an error occured.
	///
	/// \sa ColumnIDToRealIndex, RealColumnIndexToIndex, ColumnIndexToID, ShellColumnsStatus::pAllColumns
	LONG RealColumnIndexToID(int realColumnIndex);
	/// \brief <em>Retrieves the specified shell listview column's zero-based index</em>
	///
	/// \param[in] realColumnIndex The zero-based index (within the \c ShellColumnsStatus::pAllColumns array)
	///            of the listview column for which to retrieve the index.
	///
	/// \return The column's zero-based index or -1 if an error occured.
	///
	/// \sa ColumnIndexToRealIndex, RealColumnIndexToID, ColumnIDToIndex, ShellColumnsStatus::pAllColumns
	int RealColumnIndexToIndex(int realColumnIndex);
	/// \brief <em>Sets the specified column's alignment</em>
	///
	/// \param[in] realColumnIndex The zero-based index (within the \c ShellColumnsStatus::pAllColumns array)
	///            of the listview column for which to set the alignment.
	/// \param[in] alignment The alignment to set.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa ShellListViewColumn::put_Alignment
	BOOL SetColumnAlignment(int realColumnIndex, AlignmentConstants alignment);
	/// \brief <em>Sets the specified column's caption</em>
	///
	/// \param[in] realColumnIndex The zero-based index (within the \c ShellColumnsStatus::pAllColumns array)
	///            of the listview column for which to set the caption.
	/// \param[in] pText The text to set.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa ShellListViewColumn::put_Caption
	BOOL SetColumnCaption(int realColumnIndex, LPTSTR pText);
	/// \brief <em>Sets the specified column's default width</em>
	///
	/// \param[in] realColumnIndex The zero-based index (within the \c ShellColumnsStatus::pAllColumns array)
	///            of the listview column for which to set the width.
	/// \param[in] width The width to set.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa ShellListViewColumn::put_Width
	BOOL SetColumnDefaultWidth(int realColumnIndex, OLE_XSIZE_PIXELS width);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Context menu support
	///
	//@{
	/// \brief <em>Creates the shell context menu for the specified listview items</em>
	///
	/// \param[in] pItems The listview items for which to create the context menu.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	/// \param[in] predefinedFlags Specifies the \c CMF_ flags that shall always be set.
	/// \param[out] hMenu Receives the handle of the created shell context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewItem::CreateShellContextMenu, ShellListViewItems::CreateShellContextMenu,
	///     ShellListViewNamespace::CreateShellContextMenu, DestroyShellContextMenu, DisplayShellContextMenu
	HRESULT CreateShellContextMenu(PLONG pItems, UINT itemCount, UINT predefinedFlags, HMENU& hMenu);
	/// \brief <em>Creates the shell context menu for the specified shell namespace</em>
	///
	/// \param[in] pIDLNamespace The fully qualified pIDL of the shell namespace for which to create the
	///            context menu.
	/// \param[out] hMenu Receives the handle of the created shell context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellListViewNamespace::CreateShellContextMenu, ShellListViewItem::CreateShellContextMenu,
	///     ShellListViewItems::CreateShellContextMenu, DestroyShellContextMenu, DisplayShellContextMenu
	HRESULT CreateShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, HMENU& hMenu);
	/// \brief <em>Creates the context menu for the header control based on the current column settings</em>
	///
	/// \param[in] defaultDisplayColumn The index of the default display column which can't be removed.
	/// \param[out] hMenu Receives the handle of the created shell context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayHeaderContextMenu, DestroyHeaderContextMenu
	HRESULT CreateHeaderContextMenu(ULONG defaultDisplayColumn, HMENU& hMenu);
	/// \brief <em>Displays the shell context menu for the specified listview items at the specified position</em>
	///
	/// \param[in] pItems The listview items for which to show the context menu.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	/// \param[in] position The coordinates (in pixels) of the position relative to the screen at which the
	///            context menu shall be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnInternalContextMenu, ShellListViewItem::DisplayShellContextMenu,
	///     ShellListViewItems::DisplayShellContextMenu, ShellListViewNamespace::DisplayShellContextMenu,
	///     CreateShellContextMenu
	HRESULT DisplayShellContextMenu(PLONG pItems, UINT itemCount, POINT& position);
	/// \brief <em>Displays the shell context menu for the specified shell namespace at the specified position</em>
	///
	/// \param[in] pIDLNamespace The fully qualified pIDL of the shell namespace for which to show the
	///            context menu.
	/// \param[in] position The coordinates (in pixels) of the position relative to the screen at which the
	///            context menu shall be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnInternalContextMenu, ShellListViewNamespace::DisplayShellContextMenu,
	///     ShellListViewItem::DisplayShellContextMenu, ShellListViewItems::DisplayShellContextMenu,
	///     CreateShellContextMenu
	HRESULT DisplayShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, POINT& position);
	/// \brief <em>Displays the context menu for the column header control at the specified position</em>
	///
	/// \param[in] position The coordinates (in pixels) of the position relative to the screen at which the
	///            context menu shall be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnInternalContextMenu, CreateHeaderContextMenu
	HRESULT DisplayHeaderContextMenu(POINT& position);
	/// \brief <em>Invokes the default shell context menu command for the specified listview items</em>
	///
	/// \param[in] pItems The listview items for which to invoke the command.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellListViewItem::InvokeDefaultShellContextMenuCommand,
	///     ShellListViewItems::InvokeDefaultShellContextMenuCommand
	HRESULT InvokeDefaultShellContextMenuCommand(PLONG pItems, UINT itemCount);

	/// \brief <em>Stores any data required for context menus</em>
	typedef struct ContextMenuData
	{
		/// \brief <em>The currently used \c IContextMenu implementation</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb776095.aspx">IContextMenu</a>
		LPCONTEXTMENU pIContextMenu;
		/// \brief <em>The control's current shell context menu</em>
		///
		/// \sa CreateShellContextMenu, DestroyShellContextMenu
		CMenu currentShellContextMenu;
		/// \brief <em>The control's current column header context menu</em>
		///
		/// \sa CreateHeaderContextMenu, DestroyHeaderContextMenu
		CMenu currentHeaderContextMenu;
		/// \brief <em>The items or namespace that the control's current shell context menu refers to</em>
		///
		/// \sa CreateShellContextMenu, DestroyShellContextMenu
		LPDISPATCH pContextMenuItems;
		/// \brief <em>If \c TRUE, the CTRL key is hold down</em>
		///
		/// \sa CreateShellContextMenu, DestroyShellContextMenu
		UINT bufferedCtrlDown : 1;
		/// \brief <em>If \c TRUE, the SHIFT key is hold down</em>
		///
		/// \sa CreateShellContextMenu, DestroyShellContextMenu
		UINT bufferedShiftDown : 1;
		/// \brief <em>The \c IShellFolder object required for cross namespace context menus</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
		CComPtr<IShellFolder> pMultiNamespaceParentISF;
		/// \brief <em>The \c IDataObject object required for cross namespace context menus</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
		CComPtr<IDataObject> pMultiNamespaceDataObject;
		/// \brief <em>The popup menu containing the context menu item that has been clicked</em>
		///
		/// \sa OnMenuMessage, DisplayShellContextMenu
		CMenuHandle clickedMenu;

		ContextMenuData()
		{
			pIContextMenu = NULL;
			pContextMenuItems = NULL;
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			if(pIContextMenu) {
				pIContextMenu->Release();
				pIContextMenu = NULL;
			}
			pContextMenuItems = NULL;
			pMultiNamespaceParentISF = NULL;
			pMultiNamespaceDataObject = NULL;
			if(currentShellContextMenu.IsMenu()) {
				currentShellContextMenu.DestroyMenu();
			}
			if(currentHeaderContextMenu.IsMenu()) {
				currentHeaderContextMenu.DestroyMenu();
			}
			if(clickedMenu.IsMenu()) {
				clickedMenu.DestroyMenu();
			}
		}

		~ContextMenuData()
		{
			if(pIContextMenu) {
				pIContextMenu->Release();
				pIContextMenu = NULL;
			}
		}
	} ContextMenuData;
	/// \brief <em>Holds any data required for context menus</em>
	ContextMenuData contextMenuData;
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Shell notifications
	///
	//@{
	/// \brief <em>Deregisters the attached listview control for shell notifications</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RegisterForShellNotifications, put_ProcessShellNotifications
	HRESULT DeregisterForShellNotifications(void);
	/// \brief <em>Registers the attached listview control for shell notifications</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DeregisterForShellNotifications, put_ProcessShellNotifications
	HRESULT RegisterForShellNotifications(void);

	/// \brief <em>Searches the items for a shell item with the specified pIDL</em>
	///
	/// \param[in] pIDL The absolute pIDL to search.
	/// \param[in] simplePIDL If \c TRUE, \c pIDL is a simple pIDL; otherwise not.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The handle of the found immediate sub-item or \c NULL if no matching item was found.
	///
	/// \sa OnTriggerItemEnumComplete, PIDLToItemID
	LONG FindItem(__in PCIDLIST_ABSOLUTE pIDL, BOOL simplePIDL, BOOL exactMatch = FALSE);
	/// \brief <em>Reloads any icons managed by the control</em>
	///
	/// \param[in] iconIndex The zero-based index of the icon in the system imagelist to update. Any icon
	///            managed by the control that is currently set to this icon, will be updated.\n
	///            If set to \c -1, all icons managed by the control will be updated. Additionally the system
	///            image list will be rebuilt and re-assigned if it is managed by the control.
	/// \param[in] flushOverlays Specifies whether overlay icons managed by the control will be updated.
	///            If set to \c TRUE, overlays are updated; otherwise not.
	/// \param[in] dontRebuildSysImageList Specifies whether the system image list will be rebuilt if
	///            \c iconIndex is set to \c -1. If set to \c TRUE, the system image list is not rebuilt if
	///            \c iconIndex is set to \c -1; otherwise it is rebuilt.
	///
	/// \sa OnSHChangeNotify_UPDATEIMAGE, FlushSystemImageList
	void FlushIcons(int iconIndex, BOOL flushOverlays, BOOL dontRebuildSysImageList);

	/// \brief <em>Handles the \c SHCNE_ASSOCCHANGED notification</em>
	///
	/// Handles the \c SHCNE_ASSOCCHANGED notification which is sent if:
	/// -# the icon of a file type or virtual item has changed.
	/// -# a file type link has changed.
	/// -# the sort order of "My Computer" and "My Documents" has changed.
	/// -# the visibility of a virtual item has changed.
	///
	/// \param[in] ppIDLs An array containing the following pIDL as only element:
	///            -# The desktop's pIDL.
	///            -# The file type's pIDL (display name: "*.xxx", path: "C:\*.xxx").
	///            -# Unknown.
	///            -# Unknown.
	///
	/// \return Ignored.
	///
	/// \remarks Case 4 (changed visibility of virtual items) currently is not implemented, because maybe
	///          it happens on Windows 9x/ME only. In tests on Windows 2000 and XP, changing the
	///          visibility of virtual items did not trigger SHCNE_ASSOCCHANGED.
	///
	/// \sa OnSHChangeNotify
	LRESULT OnSHChangeNotify_ASSOCCHANGED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_NETSHARE and \c SHCNE_NETUNSHARE notifications</em>
	///
	/// Handles the \c SHCNE_NETSHARE and \c SHCNE_NETUNSHARE notifications which are sent if a network share
	/// has been created or removed.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the changed network share as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify
	LRESULT OnSHChangeNotify_CHANGEDSHARE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications</em>
	///
	/// Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications which are sent if a directory (file) has
	/// been created.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the created directory as only element. This may be
	///            a simple pIDL.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, ShellListViewNamespace::OnSHChangeNotify_CREATE
	LRESULT OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications for a namespace</em>
	///
	/// Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications for the specified affected namespace.
	/// This notifications are sent if a directory (file) has been created.
	///
	/// \param[in] simplePIDLAddedDir The (simple) pIDL of the created directory (file).
	/// \param[in] pIDLNamespace The pIDL of the affected namespace.
	/// \param[in] pEnumSettings The value of the affected namespace's \c NamespaceEnumSettings property.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, ShellListViewNamespace::OnSHChangeNotify_CREATE,
	///     ShellListViewNamespace::get_NamespaceEnumSettings, OnSHChangeNotify_DELETE
	LRESULT OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir, PCIDLIST_ABSOLUTE pIDLNamespace, INamespaceEnumSettings* pEnumSettings);
	/// \brief <em>Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications</em>
	///
	/// Handles the \c SHCNE_RMDIR and \c SHCNE_DELETE notifications which are sent if a directory (file) has
	/// been deleted.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the deleted directory (file) as 1st element. If
	///            the item has been moved to the recycle bin, the array's 2nd element is the new pIDL.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, ShellListViewNamespace::OnSHChangeNotify_DELETE, OnSHChangeNotify_CREATE
	LRESULT OnSHChangeNotify_DELETE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_DRIVEADD and SHCNE_DRIVEADDGUI notifications</em>
	///
	/// Handles the \c SHCNE_DRIVEADD and \c SHCNE_DRIVEADDGUI notifications which are sent if a drive has
	/// been added.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the added drive as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_DRIVEREMOVED
	LRESULT OnSHChangeNotify_DRIVEADD(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_DRIVEREMOVED notification</em>
	///
	/// Handles the \c SHCNE_DRIVEREMOVED notification which is sent if a drive has been removed.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the removed drive as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_DRIVEADD
	LRESULT OnSHChangeNotify_DRIVEREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_FREESPACE notification</em>
	///
	/// Handles the \c SHCNE_FREESPACE notification which is sent if a drive's free space has changed.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the changed drive as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify
	LRESULT OnSHChangeNotify_FREESPACE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_MEDIAINSERTED notification</em>
	///
	/// Handles the \c SHCNE_MEDIAINSERTED notification which is sent if a medium was inserted into a drive.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the affected drive as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_MEDIAREMOVED
	LRESULT OnSHChangeNotify_MEDIAINSERTED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_MEDIAREMOVED notification</em>
	///
	/// Handles the \c SHCNE_MEDIAREMOVED notification which is sent if a medium was removed from a drive.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the affected drive as only element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_MEDIAINSERTED
	LRESULT OnSHChangeNotify_MEDIAREMOVED(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_RENAMEFOLDER and \c SHCNE_RENAMEITEM notifications</em>
	///
	/// Handles the \c SHCNE_RENAMEFOLDER and \c SHCNE_RENAMEITEM notifications which are sent if a directory
	/// (file) has been renamed or moved.
	///
	/// \param[in] ppIDLs An array containing the old pIDL of the renamed/moved directory (file) as 1st
	///            and its new one as 2nd element. The 2nd element may be a simple pIDL.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, ShellListViewNamespace::OnSHChangeNotify_RENAME
	LRESULT OnSHChangeNotify_RENAME(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_UPDATEDIR notification</em>
	///
	/// Handles the \c SHCNE_UPDATEDIR notification which is sent if a directory has changed.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the changed directory as only element. This may be
	///            a simple pIDL.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_UPDATEITEM
	LRESULT OnSHChangeNotify_UPDATEDIR(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_UPDATEIMAGE notification</em>
	///
	/// Handles the \c SHCNE_UPDATEIMAGE notification which is sent if a shell item's icon or overlay icon
	/// has been changed.
	///
	/// \param[in] ppIDLs An array containing a \c SHChangeDWORDAsIDList style pIDL as 1st and optionally
	///            a \c SHChangeUpdateImageIDList style pIDL as 2nd element.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, FlushIcons
	LRESULT OnSHChangeNotify_UPDATEIMAGE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_UPDATEITEM notification</em>
	///
	/// Handles the \c SHCNE_UPDATEITEM notification which is sent if a shell item has changed.
	///
	/// \param[in] ppIDLs An array containing the pIDL of the changed item as only element. This may be
	///            a simple pIDL.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, OnSHChangeNotify_UPDATEDIR
	LRESULT OnSHChangeNotify_UPDATEITEM(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);

	/// \brief <em>Stores any data related to shell notifications</em>
	typedef struct ShellNotificationData
	{
		/// \brief <em>The registration ID returned by \c SHChangeNotifyRegister</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb762120.aspx">SHChangeNotifyRegister</a>
		ULONG registrationID;

		ShellNotificationData()
		{
			registrationID = 0;
		}

		/// \brief <em>Checks whether we're currently registered for shell notifications</em>
		///
		/// Checks whether the attached control is registered for shell notifications.
		///
		/// \return \c TRUE if shell notifications are received; otherwise \c FALSE.
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb762120.aspx">SHChangeNotifyRegister</a>
		BOOL IsRegistered(void)
		{
			return (registrationID > 0);
		}
	} ShellNotificationData;
	/// \brief <em>Holds any data related to shell notifications</em>
	ShellNotificationData shellNotificationData;

	//////////////////////////////////////////////////////////////////////
	/// \name Directory watching
	///
	/// In some situations, shell notifications won't be sent. For instance creating a new file using Notepad
	/// causes a notification only if Windows Explorer is opened for the directory that the file is stored
	/// in. To detect these events, we watch the current main namespace on a background thread using
	/// \c ReadDirectoryChanges.
	///
	//@{
	/// \brief <em>Used to hold the current state of the background thread that watches the current main namespace using \c ReadDirectoryChanges</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/aa365465.aspx">ReadDirectoryChanges</a>
	typedef struct DirectoryWatchState
	{
		/// \brief <em>The handle of the background thread</em>
		HANDLE hThread;
		#ifndef SYNCHRONOUS_READDIRECTORYCHANGES
			/// \brief <em>The handle of the event that notifies of changes in the watched namespace</em>
			HANDLE hChangeEvent;
		#endif
		/// \brief <em>The fully qualified pIDL of the namespace to watch</em>
		LPVOID pIDLWatchedNamespace;
		/// \brief <em>Informs the background thread that \c pIDLWatchedNamespace has been changed</em>
		///
		/// \sa pIDLWatchedNamespace
		LONG watchNewDirectory;
		/// \brief <em>Instructs the background thread to suspend itself</em>
		LONG suspend;
		/// \brief <em>If \c TRUE, the background thread is currently suspended</em>
		LONG isSuspended;
		/// \brief <em>Instructs the background thread to terminate itself</em>
		LONG terminate;

		DirectoryWatchState()
		{
			hThread = NULL;
			#ifndef SYNCHRONOUS_READDIRECTORYCHANGES
				hChangeEvent = INVALID_HANDLE_VALUE;
			#endif
			pIDLWatchedNamespace = NULL;
			watchNewDirectory = FALSE;
			suspend = FALSE;
			isSuspended = FALSE;
			terminate = FALSE;
		}

		~DirectoryWatchState()
		{
			#ifndef SYNCHRONOUS_READDIRECTORYCHANGES
				if(hChangeEvent != INVALID_HANDLE_VALUE) {
					CloseHandle(hChangeEvent);
					hChangeEvent = INVALID_HANDLE_VALUE;
				}
			#endif
		}
	} DirectoryWatchState;
	/// \brief <em>Holds the current state of the background thread that watches the current main namespace using \c ReadDirectoryChanges</em>
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/aa365465.aspx">ReadDirectoryChanges</a>
	DirectoryWatchState directoryWatchState;

	/// \brief <em>The entry point of the background thread</em>
	///
	/// \param[in] lpParameter Address of a \c DirectoryWatchState struct that stores the thread's state.
	///
	/// \return The exit code of the thread.
	///
	/// \sa DirectoryWatchState
	static DWORD WINAPI DirWatchThreadProc(__in LPVOID lpParameter);
	//@}
	//////////////////////////////////////////////////////////////////////
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Threading
	///
	//@{
	/// \brief <em>Enqueues a task in the task scheduler</em>
	///
	/// \param[in] pTask The task to schedule.
	/// \param[in] taskGroupID A \c GUID used to group tasks for counting and removal.
	/// \param[in] lParam Points to a custom \c DWORD value associated with the task. May be \c NULL.
	/// \param[in] priority Specifies the task's priority.
	/// \param[in] operationStart Specifies the time at which the operation has been started. For certain
	///            tasks the method will check for timeouts and notify the client that the operation is
	///            taking very long.\n
	///            If this parameter is set to 0, the method won't check for timeouts.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconAsync, GetOverlayAsync,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632523.aspx">IRunnableTask::AddTask</a>
	HRESULT EnqueueTask(__in IRunnableTask* pTask, REFTASKOWNERID taskGroupID, DWORD_PTR lParam, DWORD priority, DWORD operationStart = 0);
	/// \brief <em>Creates a runnable task which retrieves the specified item's icon in a separate thread</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon index.
	/// \param[in] itemID The item for which to retrieve the icon index.
	/// \param[in] pParentISI The \c IShellIcon object to be used to retrieve the icon index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetOverlayAsync, GetThumbnailAsync, EnqueueTask, OnGetDispInfoNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761277.aspx">IShellIcon</a>
	HRESULT GetIconAsync(__in PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIcon* pParentISI);
	/// \brief <em>Creates a runnable task which retrieves the specified item's overlay icon in a separate thread</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the overlay icon.
	/// \param[in] itemID The item for which to retrieve the overlay icon.
	/// \param[in] pParentISIO The \c IShellIconOverlay object to be used to retrieve the overlay icon.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconAsync, EnqueueTask, OnGetDispInfoNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761273.aspx">IShellIconOverlay</a>
	HRESULT GetOverlayAsync(__in PCIDLIST_ABSOLUTE pIDL, LONG itemID, IShellIconOverlay* pParentISIO);
	/// \brief <em>Creates a runnable task which retrieves the specified item's thumbnail image in a separate thread</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the thumbnail image.
	/// \param[in] itemID The item for which to retrieve the thumbnail image.
	/// \param[in] thumbnailSize The size in pixels in which the thumbnail image shall be retrieved.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconAsync, EnqueueTask, OnGetDispInfoNotification, OnTriggerUpdateIcon
	HRESULT GetThumbnailAsync(__in PCIDLIST_ABSOLUTE pIDL, LONG itemID, SIZE& thumbnailSize);
	/// \brief <em>Creates a runnable task which retrieves the specified item's elevation requirements</em>
	///
	/// \param[in] itemID The item for which to retrieve the elevation requirements.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the elevation
	///            requirements.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconAsync, GetThumbnailAsync, EnqueueTask, OnTriggerUpdateIcon
	HRESULT GetElevationShieldAsync(LONG itemID, __in PCIDLIST_ABSOLUTE pIDL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Wraps the attached listview window</em>
	///
	/// \sa get_hWnd, Attach, Detach
	CWindow attachedControl;

	/// \brief <em>Holds details about the shell columns</em>
	typedef struct ShellColumnsStatus
	{
		typedef DWORD SHELLCOLUMNSLOADSTATUS;
		/// \brief <em>Possible value for the \c loaded member, meaning that the shell columns are not loaded</em>
		#define SCLS_NOTLOADED		0x1
		/// \brief <em>Possible value for the \c loaded member, meaning that the shell columns are currently being loaded</em>
		#define SCLS_LOADING			0x2
		/// \brief <em>Possible value for the \c loaded member, meaning that the shell columns have been loaded</em>
		#define SCLS_LOADED				0x3

		/// \brief <em>Specifies whether the columns have already been loaded</em>
		SHELLCOLUMNSLOADSTATUS loaded : 2;
		/// \brief <em>Holds the average width of a character in the column header caption</em>
		int averageCharWidth;
		/// \brief <em>Buffers the current value of the attached \c ExplorerListView control's \c TileViewItemLines property</em>
		LONG maxTileViewColumnCount;
		/// \brief <em>The fully qualified pIDL of the namespace whose columns are displayed</em>
		PCIDLIST_ABSOLUTE pIDLNamespace;
		/// \brief <em>The \c IShellFolder implementation of the namespace whose columns are displayed</em>
		CComPtr<IShellFolder> pNamespaceSHF;
		/// \brief <em>The \c IShellFolder2 implementation of the namespace whose columns are displayed</em>
		CComPtr<IShellFolder2> pNamespaceSHF2;
		/// \brief <em>The \c IShellDetails implementation of the namespace whose columns are displayed</em>
		CComPtr<IShellDetails> pNamespaceSHD;
		/// \brief <em>Holds the number of elements in \c pAllColumns</em>
		///
		/// \sa pAllColumns, numberOfAllColumnsOfPreviousNamespace
		UINT numberOfAllColumns;
		/// \brief <em>Holds details about all columns provided by the current namespace</em>
		///
		/// \sa numberOfAllColumns, SHELLLISTVIEWCOLUMNDATA, pAllColumnsOfPreviousNamespace
		LPSHELLLISTVIEWCOLUMNDATA pAllColumns;
		/// \brief <em>Holds the number of elements in \c pAllColumnsOfPreviousNamespace</em>
		///
		/// \sa pAllColumnsOfPreviousNamespace, numberOfAllColumns
		UINT numberOfAllColumnsOfPreviousNamespace;
		/// \brief <em>Holds details about all columns provided by the previous namespace</em>
		///
		/// \sa numberOfAllColumnsOfPreviousNamespace, SHELLLISTVIEWCOLUMNDATA, pAllColumns
		LPSHELLLISTVIEWCOLUMNDATA pAllColumnsOfPreviousNamespace;
		/// \brief <em>Holds the number of elements in \c pVisibleColumnOrderOfPreviousNamespace</em>
		///
		/// \sa pVisibleColumnOrderOfPreviousNamespace, numberOfAllColumns
		UINT numberOfVisibleColumnsOfPreviousNamespace;
		/// \brief <em>Holds all visible columns provided by the previous namespace</em>
		///
		/// \sa numberOfVisibleColumnsOfPreviousNamespace, pAllColumns,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
		PROPERTYKEY* pVisibleColumnOrderOfPreviousNamespace;
		/// \brief <em>Holds the current size of \c pAllColumns in bytes</em>
		SIZE_T currentArraySize;
		/// \brief <em>Holds the zero-based index of the current column namespace's default sort column</em>
		int defaultSortColumn;
		/// \brief <em>Holds the key of the last sort column of the previous namespace</em>
		///
		/// \sa previousSortOrder,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
		PROPERTYKEY previousSortColumn;
		/// \brief <em>Holds the last sort order of the previous namespace</em>
		///
		/// \sa previousSortColumn
		int previousSortOrder;
		/// \brief <em>The event used to signal that the control has finished loading the shell columns</em>
		HANDLE hColumnsReadyEvent;

		ShellColumnsStatus()
		{
			loaded = SCLS_NOTLOADED;
			maxTileViewColumnCount = 0;
			averageCharWidth = 6;
			pIDLNamespace = NULL;
			pNamespaceSHF = NULL;
			pNamespaceSHF2 = NULL;
			pNamespaceSHD = NULL;
			numberOfAllColumns = 0;
			pAllColumns = NULL;
			numberOfAllColumnsOfPreviousNamespace = 0;
			pAllColumnsOfPreviousNamespace = NULL;
			currentArraySize = 0;
			defaultSortColumn = 0;
			ZeroMemory(&previousSortColumn, sizeof(previousSortColumn));
			previousSortOrder = 0;
			numberOfVisibleColumnsOfPreviousNamespace = 0;
			pVisibleColumnOrderOfPreviousNamespace = NULL;
			hColumnsReadyEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
			ATLASSERT(hColumnsReadyEvent);
		}

		~ShellColumnsStatus()
		{
			if(pAllColumns) {
				HeapFree(GetProcessHeap(), 0, pAllColumns);
			}
			if(pAllColumnsOfPreviousNamespace) {
				HeapFree(GetProcessHeap(), 0, pAllColumnsOfPreviousNamespace);
			}
			if(pVisibleColumnOrderOfPreviousNamespace) {
				HeapFree(GetProcessHeap(), 0, pVisibleColumnOrderOfPreviousNamespace);
			}
			if(hColumnsReadyEvent) {
				CloseHandle(hColumnsReadyEvent);
			}
		}

		/// \brief <em>Resets all properties to their defaults</em>
		///
		/// \param[in] forUnloadShellColumns If \c TRUE, only those members are reset, that have to be reset
		///            when loading the columns of another namespace.
		///
		/// \sa UnloadShellColumns
		void ResetToDefaults(BOOL forUnloadShellColumns = FALSE)
		{
			currentArraySize = 0;
			#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
				UINT oldNumberOfAllColumns = numberOfAllColumns;
				UINT oldNumberOfAllColumnsOfPreviousNamespace = numberOfAllColumnsOfPreviousNamespace;
			#endif
			numberOfAllColumns = 0;
			if(pAllColumns) {
				#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
					for(UINT i = 0; i < oldNumberOfAllColumns; ++i) {
						if(pAllColumns[i].pPropertyDescription) {
							pAllColumns[i].pPropertyDescription->Release();
							pAllColumns[i].pPropertyDescription = NULL;
						}
					}
				#endif
				HeapFree(GetProcessHeap(), 0, pAllColumns);
				pAllColumns = NULL;
			}
			if(!forUnloadShellColumns) {
				numberOfAllColumnsOfPreviousNamespace = 0;
				if(pAllColumnsOfPreviousNamespace) {
					#ifdef ACTIVATE_SUBITEMCONTROL_SUPPORT
						for(UINT i = 0; i < oldNumberOfAllColumnsOfPreviousNamespace; ++i) {
							if(pAllColumnsOfPreviousNamespace[i].pPropertyDescription) {
								pAllColumnsOfPreviousNamespace[i].pPropertyDescription->Release();
								pAllColumnsOfPreviousNamespace[i].pPropertyDescription = NULL;
							}
						}
					#endif
					HeapFree(GetProcessHeap(), 0, pAllColumnsOfPreviousNamespace);
					pAllColumnsOfPreviousNamespace = NULL;
				}
				numberOfVisibleColumnsOfPreviousNamespace = 0;
				if(pVisibleColumnOrderOfPreviousNamespace) {
					HeapFree(GetProcessHeap(), 0, pVisibleColumnOrderOfPreviousNamespace);
					pVisibleColumnOrderOfPreviousNamespace = NULL;
				}
			}
			loaded = SCLS_NOTLOADED;
			defaultSortColumn = 0;
			if(!forUnloadShellColumns) {
				ZeroMemory(&previousSortColumn, sizeof(previousSortColumn));
				previousSortOrder = 0;
				maxTileViewColumnCount = 0;
				averageCharWidth = 6;
				pIDLNamespace = NULL;
				pNamespaceSHF = NULL;
				pNamespaceSHF2 = NULL;
				pNamespaceSHD = NULL;
				if(hColumnsReadyEvent) {
					CloseHandle(hColumnsReadyEvent);
				}
				hColumnsReadyEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
			}
		}
	} ShellColumnsStatus;

	/// \brief <em>Holds details about the thumbnails</em>
	typedef struct ThumbnailsStatus
	{
		/// \brief <em>Holds the effective thumbnail size</em>
		SIZE thumbnailSize;
		/// \brief <em>Holds the handle of the aggregate image list that contains the thumbnail images</em>
		///
		/// \sa AggregateImageList_CreateInstance
		CComPtr<IImageList> pThumbnailImageList;

		ThumbnailsStatus()
		{
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			pThumbnailImageList = NULL;
		}
	} ThumbnailsStatus;

	/// \brief <em>Holds the object's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>An empty \c VARIANT shared between time-critical methods</em>
		VARIANT emptyVariant;
		/// \brief <em>A \c EXLVMADDCOLUMNDATA struct prepared to be used by \c ChangeColumnVisibility</em>
		///
		/// \sa ChangeColumnVisibility
		EXLVMADDCOLUMNDATA columnDataForFastInsertion;
		/// \brief <em>A \c EXLVMADDITEMDATA struct prepared to be used by \c FastInsertShellItem</em>
		///
		/// \sa FastInsertShellItem
		EXLVMADDITEMDATA itemDataForFastInsertion;

		/// \brief <em>Holds the \c AutoEditNewItems property's setting</em>
		///
		/// \sa get_AutoEditNewItems, put_AutoEditNewItems
		UINT autoEditNewItems : 1;
		/// \brief <em>Holds the \c AutoInsertColumns property's setting</em>
		///
		/// \sa get_AutoInsertColumns, put_AutoInsertColumns
		UINT autoInsertColumns : 1;
		/// \brief <em>Holds the \c ColorCompressedItems property's setting</em>
		///
		/// \sa get_ColorCompressedItems, put_ColorCompressedItems
		UINT colorCompressedItems : 1;
		/// \brief <em>Holds the \c ColorEncryptedItems property's setting</em>
		///
		/// \sa get_ColorEncryptedItems, put_ColorEncryptedItems
		UINT colorEncryptedItems : 1;
		/// \brief <em>Holds the \c ColumnEnumerationTimeout property's setting</em>
		///
		/// \sa get_ColumnEnumerationTimeout, put_ColumnEnumerationTimeout
		long columnEnumerationTimeout;
		/// \brief <em>Holds the \c Columns property's setting</em>
		///
		/// \sa get_Columns
		CComObject<ShellListViewColumns>* pColumns;
		/// \brief <em>Holds the \c DefaultManagedItemProperties property's setting</em>
		///
		/// \sa get_DefaultManagedItemProperties, put_DefaultManagedItemProperties
		ShLvwManagedItemPropertiesConstants defaultManagedItemProperties;
		/// \brief <em>Holds the \c DefaultNamespaceEnumSettings property's setting</em>
		///
		/// \sa get_DefaultNamespaceEnumSettings, putref_DefaultNamespaceEnumSettings
		INamespaceEnumSettings* pDefaultNamespaceEnumSettings;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DisplayElevationShieldOverlays property's setting</em>
		///
		/// \sa get_DisplayElevationShieldOverlays, put_DisplayElevationShieldOverlays
		UINT displayElevationShieldOverlays : 1;
		/// \brief <em>Holds the \c DisplayFileTypeOverlays property's setting</em>
		///
		/// \sa get_DisplayFileTypeOverlays, put_DisplayFileTypeOverlays
		ShLvwDisplayFileTypeOverlaysConstants displayFileTypeOverlays;
		/// \brief <em>Holds the \c DisplayThumbnailAdornments property's setting</em>
		///
		/// \sa get_DisplayThumbnailAdornments, put_DisplayThumbnailAdornments
		ShLvwDisplayThumbnailAdornmentsConstants displayThumbnailAdornments;
		/// \brief <em>Holds the \c DisplayThumbnails property's setting</em>
		///
		/// \sa get_DisplayThumbnails, put_DisplayThumbnails
		UINT displayThumbnails : 1;
		/// \brief <em>Holds the \c HandleOLEDragDrop property's setting</em>
		///
		/// \sa get_HandleOLEDragDrop, put_HandleOLEDragDrop
		HandleOLEDragDropConstants handleOLEDragDrop;
		/// \brief <em>Holds the \c HiddenItemsStyle property's setting</em>
		///
		/// \sa get_HiddenItemsStyle, put_HiddenItemsStyle
		HiddenItemsStyleConstants hiddenItemsStyle;
		/// \brief <em>Holds the \c hImageList property's setting</em>
		///
		/// \sa get_hImageList, put_hImageList
		HIMAGELIST hImageList[1];
		/// \brief <em>Holds the \c hWndShellUIParentWindow property's setting</em>
		///
		/// \sa get_hWndShellUIParentWindow, put_hWndShellUIParentWindow, GethWndShellUIParentWindow
		HWND hWndShellUIParentWindow;
		/// \brief <em>Holds the \c InfoTipFlags property's setting</em>
		///
		/// get_InfoTipFlags
		InfoTipFlagsConstants infoTipFlags;
		/// \brief <em>Holds the \c InitialSortOrder property's setting</em>
		///
		/// \sa get_InitialSortOrder, put_InitialSortOrder
		SortOrderConstants initialSortOrder;
		/// \brief <em>Holds the \c ItemEnumerationTimeout property's setting</em>
		///
		/// \sa get_ItemEnumerationTimeout, put_ItemEnumerationTimeout
		long itemEnumerationTimeout;
		/// \brief <em>Holds the \c ItemTypeSortOrder property's setting</em>
		///
		/// \sa get_ItemTypeSortOrder, put_ItemTypeSortOrder
		ItemTypeSortOrderConstants itemTypeSortOrder;
		/// \brief <em>Holds the \c LimitLabelEditInput property's setting</em>
		///
		/// \sa get_LimitLabelEditInput, put_LimitLabelEditInput
		UINT limitLabelEditInput : 1;
		/// \brief <em>Holds the \c ListItems property's setting</em>
		///
		/// \sa get_ListItems
		CComObject<ShellListViewItems>* pListItems;
		/// \brief <em>Holds the \c LoadOverlaysOnDemand property's setting</em>
		///
		/// \sa get_LoadOverlaysOnDemand, put_LoadOverlaysOnDemand
		UINT loadOverlaysOnDemand : 1;
		/// \brief <em>Holds the \c Namespaces property's setting</em>
		///
		/// \sa get_Namespaces
		CComObject<ShellListViewNamespaces>* pNamespaces;
		/// \brief <em>Holds the \c PersistColumnSettingsAcrossNamespaces property's setting</em>
		///
		/// \sa get_PersistColumnSettingsAcrossNamespaces, put_PersistColumnSettingsAcrossNamespaces
		ShLvwPersistColumnSettingsAcrossNamespacesConstants persistColumnSettingsAcrossNamespaces;
		/// \brief <em>Holds the \c PreselectBasenameOnLabelEdit property's setting</em>
		///
		/// \sa get_PreselectBasenameOnLabelEdit, put_PreselectBasenameOnLabelEdit
		UINT preselectBasenameOnLabelEdit : 1;
		/// \brief <em>Holds the \c ProcessShellNotifications property's setting</em>
		///
		/// \sa get_ProcessShellNotifications, put_ProcessShellNotifications
		UINT processShellNotifications : 1;
		/// \brief <em>Holds the \c SelectSortColumn property's setting</em>
		///
		/// \sa get_SelectSortColumn, put_SelectSortColumn
		UINT selectSortColumn : 1;
		/// \brief <em>Holds the \c SetSortArrows property's setting</em>
		///
		/// \sa get_SetSortArrows, put_SetSortArrows
		UINT setSortArrows : 1;
		/// \brief <em>Holds the \c SortOnHeaderClick property's setting</em>
		///
		/// \sa get_SortOnHeaderClick, put_SortOnHeaderClick
		UINT sortOnHeaderClick : 1;
		/// \brief <em>Holds the \c ThumbnailSize property's setting</em>
		///
		/// \sa get_ThumbnailSize, put_ThumbnailSize
		long thumbnailSize;
		/// \brief <em>Holds the \c UseColumnResizability property's setting</em>
		///
		/// \sa get_UseColumnResizability, put_UseColumnResizability
		UINT useColumnResizability : 1;
		/// \brief <em>Holds the \c UseFastButImpreciseItemHandling property's setting</em>
		///
		/// \sa get_UseFastButImpreciseItemHandling, put_UseFastButImpreciseItemHandling
		UINT useFastButImpreciseItemHandling : 1;
		/// \brief <em>Holds the \c UseGenericIcons property's setting</em>
		///
		/// \sa get_UseGenericIcons, put_UseGenericIcons
		UseGenericIconsConstants useGenericIcons;
		/// \brief <em>Holds the \c UseSystemImageList property's setting</em>
		///
		/// \sa get_UseSystemImageList, put_UseSystemImageList
		UseSystemImageListConstants useSystemImageList;
		/// \brief <em>Holds the \c UseThreadingForSlowColumns property's setting</em>
		///
		/// \sa get_UseThreadingForSlowColumns, put_UseThreadingForSlowColumns
		UINT useThreadingForSlowColumns : 1;
		/// \brief <em>Holds the \c UseThumbnailDiskCache property's setting</em>
		///
		/// \sa get_UseThumbnailDiskCache, put_UseThumbnailDiskCache
		UINT useThumbnailDiskCache : 1;

		/// \brief <em>A pointer to the attached control's \c IUnknown implementation</em>
		///
		/// \remarks This member is used only to hold a reference on the attached control.
		LPUNKNOWN pAttachedCOMControl;
		/// \brief <em>Holds the \c IInternalMessageListener implementation of the attached control</em>
		IInternalMessageListener* pAttachedInternalMessageListener;

		#ifdef USE_STL
			/// \brief <em>A map of all listview items managed by this object</em>
			///
			/// Holds the details of all shell-items in the attached listview that are managed by this object.
			/// The item's ID is stored as key, the item's details are stored as value.
			std::unordered_map<LONG, LPSHELLLISTVIEWITEMDATA> items;
			/// \brief <em>A map of all shell namespaces managed by this object</em>
			///
			/// Holds all shell namespaces in the attached listview that are managed by this object. The
			/// namespace's fully qualified pIDL is stored as key, the namespace object is stored as value.
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*> namespaces;
			/// \brief <em>A map of all currently visible listview columns managed by this object</em>
			///
			/// Maps the IDs of all shell-columns in the attached listview that are managed by this object to
			/// the columns' indexes within the \c ShellColumnsStatus::pAllColumns array.
			/// The column's ID is stored as key, the column's index is stored as value.
			///
			/// \sa ShellColumnsStatus::pAllColumns
			std::unordered_map<LONG, int> visibleColumns;
		#else
			/// \brief <em>A map of all listview items managed by this object</em>
			///
			/// Holds the details of all shell-items in the attached listview that are managed by this object.
			/// The item's ID is stored as key, the item's details are stored as value.
			CAtlMap<LONG, LPSHELLLISTVIEWITEMDATA> items;
			/// \brief <em>A map of all shell namespaces managed by this object</em>
			///
			/// Holds all shell namespaces in the attached listview that are managed by this object. The
			/// namespace's fully qualified pIDL is stored as key, the namespace object is stored as value.
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellListViewNamespace>*> namespaces;
			/// \brief <em>A map of all currently visible listview columns managed by this object</em>
			///
			/// Maps the IDs of all shell-columns in the attached listview that are managed by this object to
			/// the columns' indexes within the \c ShellColumnsStatus::pAllColumns array.
			/// The column's ID is stored as key, the column's index is stored as value.
			///
			/// \sa ShellColumnsStatus::pAllColumns
			CAtlMap<LONG, int> visibleColumns;
		#endif
		/// \brief <em>Holds details about the columns managed by \c ShellListView</em>
		///
		/// \sa ShellColumnsStatus
		ShellColumnsStatus columnsStatus;
		/// \brief <em>Holds details about the thumbnail icons managed by \c ShellListView</em>
		///
		/// \sa ThumbnailsStatus
		ThumbnailsStatus thumbnailsStatus;
		/// \brief <em>The Desktop's \c IShellFolder implementation</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
		LPSHELLFOLDER pDesktopISF;
		/// \brief <em>The \c IShellTaskScheduler implementation of the task scheduler we use for multithreading</em>
		///
		/// This is the task scheduler that is used for item and column enumeration tasks.
		///
		/// \sa ShLvwBackgroundItemEnumTask, ShLvwBackgroundColumnEnumTask,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774864.aspx">IShellTaskScheduler</a>
		CComPtr<IShellTaskScheduler> pEnumTaskScheduler;
		#ifdef USE_STL
			/// \brief <em>Buffers the items enumerated by the background threads until they are inserted into the list view</em>
			///
			/// \sa ShLvwBackgroundItemEnumTask, ShLvwInsertSingleItemTask, SHLVWBACKGROUNDITEMENUMINFO
			std::queue<LPSHLVWBACKGROUNDITEMENUMINFO> backgroundItemEnumResults;
			/// \brief <em>Buffers the columns enumerated by the background thread until they are inserted into the list view</em>
			///
			/// \sa ShLvwBackgroundColumnEnumTask, SHLVWBACKGROUNDCOLUMNENUMINFO
			std::queue<LPSHLVWBACKGROUNDCOLUMNENUMINFO> backgroundColumnEnumResults;
			/// \brief <em>Buffers the sub-item text for a slow column retrieved by the background thread until it is inserted into the list view</em>
			///
			/// \sa ShLvwSlowColumnTask, SHLVWBACKGROUNDCOLUMNINFO
			std::queue<LPSHLVWBACKGROUNDCOLUMNINFO> slowColumnResults;
			/// \brief <em>Buffers the sub-item control property value retrieved by the background thread until it is inserted into the list view</em>
			///
			/// \sa ShLvwSubItemControlTask, SHLVWBACKGROUNDCOLUMNINFO
			std::queue<LPSHLVWBACKGROUNDCOLUMNINFO> subItemControlResults;
			/// \brief <em>Buffers the info tip text retrieved by the background thread until it is used by the list view</em>
			///
			/// \sa ShLvwInfoTipTask, SHLVWBACKGROUNDINFOTIPINFO
			std::queue<LPSHLVWBACKGROUNDINFOTIPINFO> infoTipResults;
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa ShLvwThumbnailTask, SHLVWBACKGROUNDTHUMBNAILINFO
			std::queue<LPSHLVWBACKGROUNDTHUMBNAILINFO> backgroundThumbnailResults;
			/// \brief <em>Buffers the tile view information retrieved by the background thread until the list view is updated</em>
			///
			/// \sa ShLvwTileViewTask, ShLvwLegacyTileViewTask, SHLVWBACKGROUNDTILEVIEWINFO
			std::queue<LPSHLVWBACKGROUNDTILEVIEWINFO> backgroundTileViewResults;
			/// \brief <em>Buffers tasks created by a background thread until it is sent to the scheduler</em>
			///
			/// \sa ShLvwLegacyThumbnailTask, ShLvwLegacyTestThumbnailCacheTask, SHCTLSBACKGROUNDTASKINFO
			std::queue<LPSHCTLSBACKGROUNDTASKINFO> tasksToEnqueue;
		#else
			/// \brief <em>Buffers the items enumerated by the background threads until they are inserted into the list view</em>
			///
			/// \sa ShLvwBackgroundItemEnumTask, ShLvwInsertSingleItemTask, SHLVWBACKGROUNDITEMENUMINFO
			CAtlList<LPSHLVWBACKGROUNDITEMENUMINFO> backgroundItemEnumResults;
			/// \brief <em>Buffers the columns enumerated by the background thread until they are inserted into the list view</em>
			///
			/// \sa ShLvwBackgroundColumnEnumTask, SHLVWBACKGROUNDCOLUMNENUMINFO
			CAtlList<LPSHLVWBACKGROUNDCOLUMNENUMINFO> backgroundColumnEnumResults;
			/// \brief <em>Buffers the sub-item text for a slow column retrieved by the background thread until it is inserted into the list view</em>
			///
			/// \sa ShLvwSlowColumnTask, SHLVWBACKGROUNDCOLUMNINFO
			CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO> slowColumnResults;
			/// \brief <em>Buffers the sub-item control property value retrieved by the background thread until it is inserted into the list view</em>
			///
			/// \sa ShLvwSubItemControlTask, SHLVWBACKGROUNDCOLUMNINFO
			CAtlList<LPSHLVWBACKGROUNDCOLUMNINFO> subItemControlResults;
			/// \brief <em>Buffers the info tip text retrieved by the background thread until it is used by the list view</em>
			///
			/// \sa ShLvwInfoTipTask, SHLVWBACKGROUNDINFOTIPINFO
			CAtlList<LPSHLVWBACKGROUNDINFOTIPINFO> infoTipResults;
			/// \brief <em>Buffers the thumbnail information retrieved by the background thread until the image list is updated</em>
			///
			/// \sa ShLvwThumbnailTask, SHLVWBACKGROUNDTHUMBNAILINFO
			CAtlList<LPSHLVWBACKGROUNDTHUMBNAILINFO> backgroundThumbnailResults;
			/// \brief <em>Buffers the tile view information retrieved by the background thread until the list view is updated</em>
			///
			/// \sa ShLvwTileViewTask, ShLvwLegacyTileViewTask, SHLVWBACKGROUNDTILEVIEWINFO
			CAtlList<LPSHLVWBACKGROUNDTILEVIEWINFO> backgroundTileViewResults;
			/// \brief <em>Buffers tasks created by a background thread until it is sent to the scheduler</em>
			///
			/// \sa ShLvwLegacyThumbnailTask, ShLvwLegacyTestThumbnailCacheTask, SHCTLSBACKGROUNDTASKINFO
			CAtlList<LPSHCTLSBACKGROUNDTASKINFO> tasksToEnqueue;
		#endif
		/// \brief <em>The critical section used to synchronize access to \c backgroundItemEnumResults</em>
		///
		/// \sa backgroundItemEnumResults
		CRITICAL_SECTION backgroundItemEnumResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c backgroundColumnEnumResults</em>
		///
		/// \sa backgroundColumnEnumResults
		CRITICAL_SECTION backgroundColumnEnumResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c slowColumnResults</em>
		///
		/// \sa slowColumnResults
		CRITICAL_SECTION slowColumnResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c subItemControlResults</em>
		///
		/// \sa subItemControlResults
		CRITICAL_SECTION subItemControlResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c infoTipResults</em>
		///
		/// \sa infoTipResults
		CRITICAL_SECTION infoTipResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c backgroundThumbnailResults</em>
		///
		/// \sa backgroundThumbnailResults
		CRITICAL_SECTION backgroundThumbnailResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c backgroundTileViewResults</em>
		///
		/// \sa backgroundTileViewResults
		CRITICAL_SECTION backgroundTileViewResultsCritSection;
		/// \brief <em>The critical section used to synchronize access to \c tasksToEnqueue</em>
		///
		/// \sa tasksToEnqueue
		CRITICAL_SECTION tasksToEnqueueCritSection;
		/// \brief <em>The \c IShellTaskScheduler implementation of the task scheduler we use for multithreading</em>
		///
		/// This is the task scheduler that is used for the following tasks:
		/// - icon retrieval
		/// - overlay retrieval
		/// - sub-item text retrieval
		/// - tile view columns retrieval
		/// - insertion of single items in response to a shell notification
		///
		/// \sa ShLvwIconTask, ShLvwOverlayTask, ShLvwSlowColumnTask, ShLvwLegacyTileViewTask,
		///     ShLvwTileViewTask, ShLvwInsertSingleItemTask,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774864.aspx">IShellTaskScheduler</a>
		CComPtr<IShellTaskScheduler> pNormalTaskScheduler;
		/// \brief <em>The \c IShellTaskScheduler implementation of the task scheduler we use for multithreading</em>
		///
		/// This is the task scheduler that is used for thumbnail tasks.
		///
		/// \sa ShLvwThumbnailTask,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774864.aspx">IShellTaskScheduler</a>
		CComPtr<IShellTaskScheduler> pThumbnailTaskScheduler;
		/// \brief <em>The default icon index for files in the system imagelist</em>
		int DEFAULTICON_FILE;
		/// \brief <em>The default icon index for folders in the system imagelist</em>
		int DEFAULTICON_FOLDER;
		/// \brief <em>Holds the handle of the unified image list that contains the items' images</em>
		///
		/// \sa UnifiedImageList_CreateInstance
		CComPtr<IImageList> pUnifiedImageList;

		#ifdef USE_STL
			/// \brief <em>Holds the \c LPSHELLLISTVIEWITEMDATA pointers that shall be stored in \c OnInternalInsertedItem</em>
			///
			/// \sa BufferShellItemData, OnInternalInsertedItem, SHELLLISTVIEWITEMDATA, ShellListViewItems::Add
			std::stack<LPSHELLLISTVIEWITEMDATA> shellItemDataOfInsertedItems;
			/// \brief <em>Holds the column index that shall be stored in \c OnInternalInsertedColumn</em>
			///
			/// \sa BufferShellColumnRealIndex, OnInternalInsertedColumn
			std::stack<int> shellColumnIndexesOfInsertedColumns;
			/// \brief <em>Holds the namespace pIDLs of items that are going to be inserted</em>
			///
			/// \sa BufferShellItemNamespace, OnInternalInsertedItem, ShellListViewNamespaces::Add,
			///     shellItemDataOfInsertedItems
			std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacesOfInsertedItems;
			/// \brief <em>Holds details about currently running background enumerations</em>
			///
			/// Holds details about currently running background enumerations. These details include the
			/// time the task has been started and whether to check for timeouts.\n
			/// The enumeration task's unique ID is stored as key, the details are stored as value.
			///
			/// \sa EnqueueTask, BACKGROUNDENUMERATION
			std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION> backgroundEnums;
		#else
			/// \brief <em>Holds the \c LPSHELLLISTVIEWITEMDATA pointers that shall be stored in \c OnInternalInsertedItem</em>
			///
			/// \sa BufferShellItemData, OnInternalInsertedItem, SHELLLISTVIEWITEMDATA, ShellListViewItems::Add
			CAtlList<LPSHELLLISTVIEWITEMDATA> shellItemDataOfInsertedItems;
			/// \brief <em>Holds the column index that shall be stored in \c OnInternalInsertedColumn</em>
			///
			/// \sa BufferShellColumnRealIndex, OnInternalInsertedColumn
			CAtlList<int> shellColumnIndexesOfInsertedColumns;
			/// \brief <em>Holds the namespace pIDLs of items that are going to be inserted</em>
			///
			/// \sa BufferShellItemNamespace, OnInternalInsertedItem, ShellListViewNamespaces::Add,
			///     shellItemDataOfInsertedItems
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacesOfInsertedItems;
			/// \brief <em>Holds details about currently running background enumerations</em>
			///
			/// Holds details about currently running background enumerations. These details include the
			/// time the task has been started and whether to check for timeouts.\n
			/// The enumeration task's unique ID is stored as key, the details are stored as value.
			///
			/// \sa EnqueueTask, BACKGROUNDENUMERATION
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION> backgroundEnums;
		#endif

		/// \brief <em>Holds the latest tick count at which a command from the "New" shell context menu was invoked</em>
		///
		/// When invoking a shell context menu command from the "New" menu (e. g. "New Folder"), it might be
		/// desired to enter label-edit mode for the item created by the command after inserting it.
		/// Unfortunately this is very tricky for the following reasons:
		/// - There's no reliable way to determine whether the invoked context menu command was from the "New"
		///   menu. We use some heuristics for this task.
		/// - The pIDL of the created item is unknown. However, we receive a \c SHCNE_CREATE or \c SHCNE_MKDIR
		///   notification for it. To determine whether the item for which the notification was sent is the
		///   item created by the invoked command, we store the tick count when the command was invoked. If
		///   the \c SHCNE_* notification was received within 250 ms after the command invocation, it is
		///   assumed that the notification was sent for the item we're looking for.
		DWORD timeOfLastItemCreatorInvocation;
		/// \brief <em>Holds the coordinates (in pixels) of the context menu's position relative to the attached listview control's upper-left corner</em>
		///
		/// When invoking a shell context menu command from the "New" menu (e. g. "New Folder"), it might be
		/// desired to position the item created by the command at the point that has been clicked to create
		/// it. To do this, we store the context menu's position if it has been invoked using the mouse. Then,
		/// after the created item has been inserted, we check whether the stored position is valid and move
		/// the item to it, if it is.
		POINT contextMenuPosition;

		Properties();
		~Properties();

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			autoEditNewItems = TRUE;
			autoInsertColumns = TRUE;
			colorCompressedItems = TRUE;
			colorEncryptedItems = TRUE;
			columnEnumerationTimeout = 3000;
			defaultManagedItemProperties = slmipAll;
			disabledEvents = static_cast<DisabledEventsConstants>(deNamespacePIDLChangeEvents | deNamespaceInsertionEvents | deNamespaceDeletionEvents | deItemPIDLChangeEvents | deItemInsertionEvents | deItemDeletionEvents | deColumnLoadingEvents | deColumnVisibilityEvents);
			displayElevationShieldOverlays = TRUE;
			displayFileTypeOverlays = static_cast<ShLvwDisplayFileTypeOverlaysConstants>(sldftoFollowSystemSettings);
			displayThumbnailAdornments = sldtaAny;
			displayThumbnails = FALSE;
			handleOLEDragDrop = static_cast<HandleOLEDragDropConstants>(hoddSourcePart | hoddTargetPart | hoddTargetPartWithDropHilite);
			hiddenItemsStyle = hisGhostedOnDemand;
			hImageList[0] = NULL;
			infoTipFlags = static_cast<InfoTipFlagsConstants>(itfDefault | itfNoInfoTipFollowSystemSettings | itfAllowSlowInfoTipFollowSysSettings);
			initialSortOrder = soAscending;
			itemEnumerationTimeout = 3000;
			itemTypeSortOrder = itsoShellItemsLast;
			limitLabelEditInput = TRUE;
			loadOverlaysOnDemand = TRUE;
			persistColumnSettingsAcrossNamespaces = slpcsanDontPersist;
			preselectBasenameOnLabelEdit = TRUE;
			processShellNotifications = TRUE;
			selectSortColumn = TRUE;
			setSortArrows = TRUE;
			sortOnHeaderClick = TRUE;
			thumbnailSize = -1;
			useColumnResizability = TRUE;
			useFastButImpreciseItemHandling = TRUE;
			useGenericIcons = ugiOnlyForSlowItems;
			useSystemImageList = static_cast<UseSystemImageListConstants>(usilSmallImageList | usilLargeImageList | usilExtraLargeImageList);
			useThreadingForSlowColumns = TRUE;
			useThumbnailDiskCache = TRUE;

			pUnifiedImageList = NULL;
			timeOfLastItemCreatorInvocation = 0;
			contextMenuPosition.x = -1;
			contextMenuPosition.y = -1;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds some frequently used settings</em>
	///
	/// Holds some settings that otherwise would be requested from the control window very often.
	/// The cached settings are updated whenever the corresponding control window's settings change.
	struct CachedSettings
	{
		/// \brief <em>Holds the attached listview's current view mode</em>
		///
		/// \sa OnInternalViewChanged, IsInDetailsView, IsInTilesView, CurrentViewNeedsColumns
		int viewMode;

		CachedSettings()
		{
			viewMode = -1;
		}
	} cachedSettings;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, \c EnqueueTask accepts new tasks; otherwise it doesn't</em>
		///
		/// \sa EnqueueTask, Detach
		UINT acceptNewTasks : 1;
		/// \brief <em>If \c TRUE, we're detaching from the target control</em>
		///
		/// \sa Attach, Detach
		UINT detaching : 1;
		/// \brief <em>If \c TRUE, \c SortItems is allowed to toggle the sort order</em>
		///
		/// \sa OnInternalHeaderClick, SortItems
		UINT toggleSortOrder : 1;
		/// \brief <em>If \c TRUE, the timer that does lazy-closing of thumbnail disk caches, is currently active</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa OnReportThumbnailDiskCacheAccess
		UINT thumbnailDiskCacheTimerIsActive : 1;

		Flags()
		{
			detaching = FALSE;
			acceptNewTasks = TRUE;
			toggleSortOrder = FALSE;
			thumbnailDiskCacheTimerIsActive = FALSE;
		}
	} flags;

	/// \brief <em>Holds the settings of the current sorting process</em>
	///
	/// \sa SortItems
	struct SortingSettings
	{
		/// \brief <em>Specifies whether the next call of \c SortItems is treated like the initial call of this method</em>
		UINT firstSortItemsCall : 1;
		/// \brief <em>The zero-based index of the shell column by which to sort</em>
		int columnIndex;
		/// \brief <em>The bitmap handle of the current sort column's header</em>
		///
		/// If we're using an old (5.x) version of comctl32.dll, we use bitmaps as sort arrows. These bitmaps
		/// must be deleted by ourselves, so we have to store them.
		///
		/// \sa FreeBitmapHandle
		OLE_HANDLE headerBitmapHandle;

		SortingSettings()
		{
			columnIndex = -1;
			firstSortItemsCall = TRUE;
			headerBitmapHandle = NULL;
		}

		~SortingSettings()
		{
			FreeBitmapHandle();
		}

		/// \brief <em>Deletes the bitmap identified by \c headerBitmapHandle</em>
		void FreeBitmapHandle(void)
		{
			if(headerBitmapHandle) {
				DeleteObject(static_cast<HGDIOBJ>(LongToHandle(headerBitmapHandle)));
				headerBitmapHandle = NULL;
			}
		}
	} sortingSettings;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	typedef struct DroppedPIDLOffsets
	{
		LPIDA pDroppedPIDLs;
		LPPOINT pOffsets;
		DWORD performedEffect;
		POINT cursorPosition;
		PIDLIST_ABSOLUTE pTargetNamespace;

		DroppedPIDLOffsets(LPDATAOBJECT pDataObject)
		{
			pDroppedPIDLs = NULL;
			pOffsets = NULL;

			if(pDataObject) {
				// copy CFSTR_SHELLIDLIST data
				FORMATETC format = {0};
				format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_SHELLIDLIST));
				format.lindex = -1;
				format.dwAspect = DVASPECT_CONTENT;
				format.tymed = TYMED_HGLOBAL;
				STGMEDIUM data = {0};
				if(SUCCEEDED(pDataObject->GetData(&format, &data))) {
					LPIDA pDroppedPIDLsSource = static_cast<LPIDA>(GlobalLock(data.hGlobal));
					ATLASSERT(pDroppedPIDLsSource);
					if(pDroppedPIDLsSource) {
						SIZE_T size = GlobalSize(data.hGlobal);
						ATLASSERT(size > 0);
						if(size > 0) {
							pDroppedPIDLs = static_cast<LPIDA>(HeapAlloc(GetProcessHeap(), 0, size));
							ATLASSERT(pDroppedPIDLs);
							if(pDroppedPIDLs) {
								CopyMemory(pDroppedPIDLs, pDroppedPIDLsSource, size);
							}
						}
					}
					GlobalUnlock(data.hGlobal);
					ReleaseStgMedium(&data);
				}

				// copy CFSTR_SHELLIDLISTOFFSET data
				format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_SHELLIDLISTOFFSET));
				ZeroMemory(&data, sizeof(data));
				if(SUCCEEDED(pDataObject->GetData(&format, &data))) {
					LPPOINT pOffsetsSource = static_cast<LPPOINT>(GlobalLock(data.hGlobal));
					ATLASSERT(pOffsetsSource);
					if(pOffsetsSource) {
						SIZE_T size = GlobalSize(data.hGlobal);
						ATLASSERT(size > 0);
						if(size > 0) {
							pOffsets = static_cast<LPPOINT>(HeapAlloc(GetProcessHeap(), 0, size));
							ATLASSERT(pOffsets);
							if(pOffsets) {
								CopyMemory(pOffsets, pOffsetsSource, size);
							}
						}
					}
					GlobalUnlock(data.hGlobal);
					ReleaseStgMedium(&data);
				}
			}
			performedEffect = DROPEFFECT_NONE;
			pTargetNamespace = NULL;
		}

		~DroppedPIDLOffsets()
		{
			if(pDroppedPIDLs) {
				HeapFree(GetProcessHeap(), 0, pDroppedPIDLs);
			}
			if(pOffsets) {
				HeapFree(GetProcessHeap(), 0, pOffsets);
			}
			if(pTargetNamespace) {
				ILFree(pTargetNamespace);
			}
		}
	} DroppedPIDLOffsets;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>Holds the coordinates (in pixels) of the position (relative to the attached listview control's origin) at which the drag operation has been started</em>
		///
		/// If the attached listview control is the source of the drag'n'drop operation, this member holds the
		/// coordinates of the position at which the operation has been started. This information is used to
		/// calculate the dragged items' new positions in case of the listview also being the drop target.
		///
		/// \sa OnInternalBeginDrag, OnInternalHandleDragEvent
		POINT positionOfDragStart;
		/// \brief <em>Holds the unique ID of the listview item currently being the drop target</em>
		///
		/// \sa OnInternalHandleDragEvent, pCurrentDropTarget
		LONG idCurrentDropTarget;
		/// \brief <em>Holds the \c IDropTarget implementation of the shell listview item currently being the drop target</em>
		///
		/// \sa OnInternalHandleDragEvent, idCurrentDropTarget,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms679679.aspx">IDropTarget</a>
		CComPtr<IDropTarget> pCurrentDropTarget;
		/// \brief <em>Holds the key state reported with the latest \c OLEDRAGEVENT_DRAG* event</em>
		///
		/// While handling \c OLEDRAGEVENT_DROP, we also call other methods of \c IDropTarget than just
		/// \c Drop. The \c keyState parameter of \c IDropTarget::Drop doesn't include the dragging mouse
		/// button, which is used by the shell to distinguish between left-button drag and right-button drag.
		/// So for proper right-button drag we save the \c keyState parameter of any \c OLEDRAGEVENT_* event
		/// except \c OLEDRAGEVENT_DROP and use this value in \c OLEDRAGEVENT_DROP when calling any other
		/// method of \c IDropTarget than \c IDropTarget::Drop.
		///
		/// \sa OnInternalHandleDragEvent, OLEDRAGEVENT_DROP,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms679679.aspx">IDropTarget</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
		DWORD lastKeyState;
		#ifdef USE_STL
			/// \brief <em>Holds details like pIDLs and relative positions about data that has been dropped into the control</em>
			///
			/// \sa DroppedPIDLOffsets
			std::vector<DroppedPIDLOffsets*> bufferedPIDLOffsets;
		#else
			/// \brief <em>Holds details like pIDLs and relative positions about data that has been dropped into the control</em>
			///
			/// \sa DroppedPIDLOffsets
			CAtlArray<DroppedPIDLOffsets*> bufferedPIDLOffsets;
		#endif

		DragDropStatus()
		{
			positionOfDragStart.x = -1;
			positionOfDragStart.y = -1;
			idCurrentDropTarget = -2;
			pCurrentDropTarget = NULL;
			lastKeyState = 0;
		}

		~DragDropStatus()
		{
			#ifdef USE_STL
				for(std::vector<DroppedPIDLOffsets*>::const_iterator it = bufferedPIDLOffsets.cbegin(); it != bufferedPIDLOffsets.cend(); ++it) {
					ATLASSERT(*it);
					delete *it;
				}
				bufferedPIDLOffsets.clear();
			#else
				for(size_t i = 0; i < bufferedPIDLOffsets.GetCount(); ++i) {
					ATLASSERT(bufferedPIDLOffsets[i]);
					delete bufferedPIDLOffsets[i];
				}
				bufferedPIDLOffsets.RemoveAll();
			#endif
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			#ifdef USE_STL
				for(std::vector<DroppedPIDLOffsets*>::const_iterator it = bufferedPIDLOffsets.cbegin(); it != bufferedPIDLOffsets.cend(); ++it) {
					ATLASSERT(*it);
					delete *it;
				}
				bufferedPIDLOffsets.clear();
			#else
				for(size_t i = 0; i < bufferedPIDLOffsets.GetCount(); ++i) {
					ATLASSERT(bufferedPIDLOffsets[i]);
					delete bufferedPIDLOffsets[i];
				}
				bufferedPIDLOffsets.RemoveAll();
			#endif
			positionOfDragStart.x = -1;
			positionOfDragStart.y = -1;
			idCurrentDropTarget = -2;
			pCurrentDropTarget = NULL;
			lastKeyState = 0;
		}
	} dragDropStatus;
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to raise the \c ColumnEnumerationTimedOut and \c ItemEnumerationTimedOut events</em>
		///
		/// \sa get_ColumnEnumerationTimeout, get_ItemEnumerationTimeout, Raise_ColumnEnumerationTimedOut,
		///     Raise_ItemEnumerationTimedOut
		static const UINT_PTR ID_BACKGROUNDENUMTIMEOUT = 20;
		/// \brief <em>The ID of the timer that is used to do lazy-closing of thumbnail disk caches</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa OnReportThumbnailDiskCacheAccess, OnTimer, LazyCloseThumbnailDiskCaches
		static const UINT_PTR ID_LAZYCLOSETHUMBNAILDISKCACHE = 21;

		/// \brief <em>The interval of the timer that is used to raise the \c ColumnEnumerationTimedOut and \c ItemEnumerationTimedOut events</em>
		///
		/// \sa get_ColumnEnumerationTimeout, get_ItemEnumerationTimeout, Raise_ColumnEnumerationTimedOut,
		///     Raise_ItemEnumerationTimedOut
		static const UINT INT_BACKGROUNDENUMTIMEOUT = 100;
		/// \brief <em>The interval of the timer that is used to do lazy-closing of thumbnail disk caches</em>
		///
		/// \remarks Beginning with Windows Vista, this member isn't required anymore.
		///
		/// \sa OnReportThumbnailDiskCacheAccess, OnTimer, LazyCloseThumbnailDiskCaches
		static const UINT INT_LAZYCLOSETHUMBNAILDISKCACHE = 2000;
	} timers;

	/// \brief <em>Holds the imagelist that contains the icon that is drawn as the control's logo during design time</em>
	///
	/// \sa OnDraw, ShellListView
	CImageList imglstIDEIcon;

	/// \brief <em>Holds the reference count for the control's main thread</em>
	///
	/// To support dragging e. g. a picture from Firefox to the control, we need to call \c SHCreateThreadRef
	/// and \c SHSetThreadRef. \c SHCreateThread requires a \c LONG variable in the calling thread to store
	/// the reference count.
	///
	/// \sa ConfigureControl, pThreadObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759872.aspx">SHCreateThreadRef</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759886.aspx">SHSetThreadRef</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776787.aspx">Managing Thread References</a>
	LONG threadReferenceCounter;
	/// \brief <em>Holds the control's main thread object used to keep the thread alive until all child threads have died</em>
	///
	/// To support dragging e. g. a picture from Firefox to the control, we need to call \c SHCreateThreadRef
	/// and \c SHSetThreadRef. \c SHCreateThread returns an object that implements \c IUnknown. This object
	/// keeps the parent thread alive until any child thread that references this object has died.
	///
	/// \sa ConfigureControl, threadReferenceCounter,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759872.aspx">SHCreateThreadRef</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759886.aspx">SHSetThreadRef</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb776787.aspx">Managing Thread References</a>
	LPUNKNOWN pThreadObject;
};     // ShellListView

OBJECT_ENTRY_AUTO(__uuidof(ShellListView), ShellListView)