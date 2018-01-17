//////////////////////////////////////////////////////////////////////
/// \class ShellTreeView
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Makes an \c ExplorerTreeView a shell treeview</em>
///
/// This class makes an \c ExplorerTreeView control a shell treeview control.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::IShTreeView, ShellListView
/// \else
///   \sa ShBrowserCtlsLibA::IShTreeView, ShellListView
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
#include "LegacyBackgroundEnumTask.h"
#include "BackgroundEnumTask.h"
#include "InsertSingleItemTask.h"
#include "IconTask.h"
#include "OverlayTask.h"
#include "../ElevationShieldTask.h"
#include "_IShellTreeViewEvents_CP.h"
#include "ShTvwDefaultManagedItemPropertiesProperties.h"
#include "../../ICategorizeProperties.h"
#include "../../ICreditsProvider.h"
#include "../../helpers.h"
#include "../../EnumOLEVERB.h"
#include "../../AboutDlg.h"
#include "../../CommonProperties.h"
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
#include "ShellTreeViewItems.h"
#include "ShellTreeViewNamespaces.h"


class ATL_NO_VTABLE ShellTreeView : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IShTreeView, &IID_IShTreeView, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IShTreeView, &IID_IShTreeView, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ShellTreeView>,
    public IOleControlImpl<ShellTreeView>,
    public IOleObjectImpl<ShellTreeView>,
    public IOleInPlaceActiveObjectImpl<ShellTreeView>,
    public IViewObjectExImpl<ShellTreeView>,
    public IOleInPlaceObjectWindowlessImpl<ShellTreeView>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ShellTreeView>,
    public Proxy_IShTreeViewEvents<ShellTreeView>,
    public IPersistStorageImpl<ShellTreeView>,
    public IPersistPropertyBagImpl<ShellTreeView>,
    public ISpecifyPropertyPages,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ShellTreeView, &__uuidof(_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ShellTreeView, &__uuidof(_IShTreeViewEvents), &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ShellTreeView>,
    public CComCoClass<ShellTreeView, &CLSID_ShellTreeView>,
    public CComControl<ShellTreeView>,
    public IPerPropertyBrowsingImpl<ShellTreeView>,
    public ICategorizeProperties,
    public ICreditsProvider,
    public IMessageListener,
    public IInternalMessageListener,
    public IContextMenuCB
{
	friend class ShellTreeViewItems;
	friend class ShellTreeViewItem;
	friend class ShellTreeViewNamespaces;
	friend class ShellTreeViewNamespace;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa ~ShellTreeView, FinalConstruct
	ShellTreeView();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa ShellTreeView
	~ShellTreeView();
	/// \brief <em>This method is called directly after the constructor</em>
	///
	/// This method is called directly after the constructor. It is used for initialization.
	///
	/// \sa ShellTreeView
	HRESULT FinalConstruct();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_INVISIBLEATRUNTIME | OLEMISC_NOUIACTIVATE | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_SHELLTREEVIEW)

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ShellTreeView)
			COM_INTERFACE_ENTRY(IShTreeView)
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

		BEGIN_PROP_MAP(ShellTreeView)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ShellTreeView)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IShTreeViewEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ShellTreeView)
			CHAIN_MSG_MAP(CComControl<ShellTreeView>)
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
	/// \name Implementation of IShTreeView
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
	/// \sa put_ColorCompressedItems, get_ColorEncryptedItems, ShellTreeViewItem::get_ShellAttributes
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
	/// \sa put_ColorEncryptedItems, get_ColorCompressedItems, ShellTreeViewItem::get_ShellAttributes
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
	/// \brief <em>Retrieves the current setting of the \c DefaultManagedItemProperties property</em>
	///
	/// Retrieves a bit field specifying which of the treeview items' properties by default are managed by
	/// the \c ShellTreeView control rather than the treeview control/the client application. Any combination
	/// of the values defined by the \c ShTvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The default settings are used for items that are inserted as part of a namespace or in
	///          response to shell notifications. Changing this property doesn't affect items that already
	///          have been inserted into the treeview.
	///
	/// \if UNICODE
	///   \sa put_DefaultManagedItemProperties, get_ProcessShellNotifications,
	///       ShellTreeViewItem::get_ManagedProperties, ShellTreeViewNamespaces::Add,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa put_DefaultManagedItemProperties, get_ProcessShellNotifications,
	///       ShellTreeViewItem::get_ManagedProperties, ShellTreeViewNamespaces::Add,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DefaultManagedItemProperties(ShTvwManagedItemPropertiesConstants* pValue);
	/// \brief <em>Sets the \c DefaultManagedItemProperties property</em>
	///
	/// Sets a bit field specifying which of the treeview items' properties by default are managed by
	/// the \c ShellTreeView control rather than the treeview control/the client application. Any combination
	/// of the values defined by the \c ShTvwManagedItemPropertiesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The default settings are used for items that are inserted as part of a namespace or in
	///          response to shell notifications. Changing this property doesn't affect items that already
	///          have been inserted into the treeview.
	///
	/// \if UNICODE
	///   \sa get_DefaultManagedItemProperties, put_ProcessShellNotifications,
	///       ShellTreeViewItem::put_ManagedProperties, ShellTreeViewNamespaces::Add,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants
	/// \else
	///   \sa get_DefaultManagedItemProperties, put_ProcessShellNotifications,
	///       ShellTreeViewItem::put_ManagedProperties, ShellTreeViewNamespaces::Add,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DefaultManagedItemProperties(ShTvwManagedItemPropertiesConstants newValue);
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
	///     ShellTreeViewItem::get_NamespaceEnumSettings, ShellTreeViewNamespace::get_NamespaceEnumSettings
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
	///     ShellTreeViewItem::get_NamespaceEnumSettings, ShellTreeViewNamespace::get_NamespaceEnumSettings
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
	/// \remarks Requires Windows Vista or newer.
	///
	/// \sa put_DisplayElevationShieldOverlays
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
	/// \remarks Requires Windows Vista or newer.
	///
	/// \sa get_DisplayElevationShieldOverlays
	virtual HRESULT STDMETHODCALLTYPE put_DisplayElevationShieldOverlays(VARIANT_BOOL newValue);
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
	///          \c ITreeViewItem::Ghosted property always returns the <em>current</em> state which isn't
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
	///          \c ITreeViewItem::Ghosted property always returns the <em>current</em> state which isn't
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
	/// Retrieves the handle of the \c SysTreeView32 window that the object is currently attached to.
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
	/// Retrieves a bit field influencing the treeview items info tips if they are managed by the
	/// \c ShellListView control. Any combination of the values defined by the \c InfoTipFlagsConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_InfoTipFlags, ShellTreeViewItem::get_ManagedProperties, ShellTreeViewItem::get_InfoTipText,
	///       ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa put_InfoTipFlags, ShellTreeViewItem::get_ManagedProperties, ShellTreeViewItem::get_InfoTipText,
	///       ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InfoTipFlags(InfoTipFlagsConstants* pValue);
	/// \brief <em>Sets the \c InfoTipFlags property</em>
	///
	/// Sets a bit field influencing the treeview items info tips if they are managed by the
	/// \c ShellListView control. Any combination of the values defined by the \c InfoTipFlagsConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_InfoTipFlags, ShellTreeViewItem::put_ManagedProperties,
	///       ShBrowserCtlsLibU::InfoTipFlagsConstants
	/// \else
	///   \sa get_InfoTipFlags, ShellTreeViewItem::put_ManagedProperties,
	///       ShBrowserCtlsLibA::InfoTipFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_InfoTipFlags(InfoTipFlagsConstants newValue);
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
	/// \sa put_ItemEnumerationTimeout, Raise_ItemEnumerationTimedOut
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
	/// \sa get_ItemEnumerationTimeout, Raise_ItemEnumerationTimedOut
	virtual HRESULT STDMETHODCALLTYPE put_ItemEnumerationTimeout(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ItemTypeSortOrder property</em>
	///
	/// Retrieves the order of the different kinds of items (shell items, normal items) within the attached
	/// treeview control. This order is applied when sorting items. Any of the values defined by the
	/// \c ItemTypeSortOrderConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ItemTypeSortOrder, ShellTreeViewNamespace::get_AutoSortItems,
	///       ShellTreeViewNamespace::SortItems, ShBrowserCtlsLibU::ItemTypeSortOrderConstants
	/// \else
	///   \sa put_ItemTypeSortOrder, ShellTreeViewNamespace::get_AutoSortItems,
	///       ShellTreeViewNamespace::SortItems, ShBrowserCtlsLibA::ItemTypeSortOrderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ItemTypeSortOrder(ItemTypeSortOrderConstants* pValue);
	/// \brief <em>Sets the \c ItemTypeSortOrder property</em>
	///
	/// Sets the order of the different kinds of items (shell items, normal items) within the attached
	/// treeview control. This order is applied when sorting items. Any of the values defined by the
	/// \c ItemTypeSortOrderConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ItemTypeSortOrder, ShellTreeViewNamespace::put_AutoSortItems,
	///       ShellTreeViewNamespace::SortItems, ShBrowserCtlsLibU::ItemTypeSortOrderConstants
	/// \else
	///   \sa get_ItemTypeSortOrder, ShellTreeViewNamespace::put_AutoSortItems,
	///       ShellTreeViewNamespace::SortItems, ShBrowserCtlsLibA::ItemTypeSortOrderConstants
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
	/// \sa put_LimitLabelEditInput, ShellTreeViewItem::get_ManagedProperties,
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
	/// \sa get_LimitLabelEditInput, ShellTreeViewItem::put_ManagedProperties,
	///     put_PreselectBasenameOnLabelEdit
	virtual HRESULT STDMETHODCALLTYPE put_LimitLabelEditInput(VARIANT_BOOL newValue);
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
	/// \remarks Loading overlay icons on demand is faster, but the \c ITreeViewItem::OverlayIndex property
	///          always returns the <em>current</em> overlay icon index which isn't necessarily correct until
	///          the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \sa put_LoadOverlaysOnDemand, ShellTreeViewItem::get_ManagedProperties
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
	/// \remarks Loading overlay icons on demand is faster, but the \c ITreeViewItem::OverlayIndex property
	///          always returns the <em>current</em> overlay icon index which isn't necessarily correct until
	///          the item is initially drawn.\n
	///          Changing this property won't update existing shell items.
	///
	/// \sa get_LoadOverlaysOnDemand, ShellTreeViewItem::put_ManagedProperties
	virtual HRESULT STDMETHODCALLTYPE put_LoadOverlaysOnDemand(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShellTreeNamespaces property</em>
	///
	/// Retrieves a collection object wrapping the shell namespaces managed by this control.
	///
	/// \param[out] ppNamespaces Receives the collection object's \c IShTreeViewNamespaces implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_TreeItems, ShellTreeViewNamespaces
	virtual HRESULT STDMETHODCALLTYPE get_Namespaces(IShTreeViewNamespaces** ppNamespaces);
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
	/// \sa put_PreselectBasenameOnLabelEdit, ShellTreeViewItem::get_ManagedProperties,
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
	/// \sa get_PreselectBasenameOnLabelEdit, ShellTreeViewItem::put_ManagedProperties,
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
	/// \brief <em>Retrieves the current setting of the \c TreeItems property</em>
	///
	/// Retrieves a collection object wrapping the treeview control's shell items.
	///
	/// \param[out] ppItems Receives the collection object's \c IShTreeViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Namespaces, ShellTreeViewItems
	virtual HRESULT STDMETHODCALLTYPE get_TreeItems(IShTreeViewItems** ppItems);
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
	/// Retrieves a bit field indicating which of the attached treeview's imagelists are set to the system
	/// imagelist. Only \c usilSmallImageList (defined by the \c UseSystemImageListConstants enumeration)
	/// is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_UseSystemImageList, get_hImageList, get_UseGenericIcons,
	///       ShBrowserCtlsLibU::UseSystemImageListConstants
	/// \else
	///   \sa put_UseSystemImageList, get_hImageList, get_UseGenericIcons,
	///       ShBrowserCtlsLibA::UseSystemImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemImageList(UseSystemImageListConstants* pValue);
	/// \brief <em>Sets the \c UseSystemImageList property</em>
	///
	/// Sets a bit field indicating which of the attached treeview's imagelists are set to the system
	/// imagelist. Only \c usilSmallImageList (defined by the \c UseSystemImageListConstants enumeration)
	/// is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_UseSystemImageList, put_hImageList, put_UseGenericIcons,
	///       ShBrowserCtlsLibU::UseSystemImageListConstants
	/// \else
	///   \sa get_UseSystemImageList, put_hImageList, put_UseGenericIcons,
	///       ShBrowserCtlsLibA::UseSystemImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemImageList(UseSystemImageListConstants newValue);
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
	/// \brief <em>Attaches the object to the specified \c SysTreeView32 window</em>
	///
	/// \param[in] hWnd The \c SysTreeView32 window to attach to.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Detach, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE Attach(OLE_HANDLE hWnd);
	/// \brief <em>Compares two shell items</em>
	///
	/// \param[in] pFirstItem The first shell item to compare.
	/// \param[in] pSecondItem The second shell item to compare.
	/// \param[in] sortColumnIndex The zero-based index of the shell column by which to compare.
	/// \param[out] pResult Receives a negative value if the first item should preceed the second item; a
	///             positive value if the second item should preceed the first item; otherwise 0.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ShellTreeViewNamespace::SortItems,
	///       ShBrowserCtlsLibU::ShTvwManagedItemPropertiesConstants::stmipSubItemsSorting
	/// \else
	///   \sa ShellTreeViewNamespace::SortItems,
	///       ShBrowserCtlsLibA::ShTvwManagedItemPropertiesConstants::stmipSubItemsSorting
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE CompareItems(IShTreeViewItem* pFirstItem, IShTreeViewItem* pSecondItem, LONG sortColumnIndex = 0, LONG* pResult = NULL);
	/// \brief <em>Creates a shell context menu</em>
	///
	/// \param[in] items An object specifying the treeview items for which to create the shell context
	///            menu. The following values may be used to identify the items:
	///            - A single item handle.
	///            - An array of item handles.
	///            - A \c TreeViewItem object.
	///            - A \c TreeViewItems object.
	///            - A \c TreeViewItemContainer object.
	///            - A \c ShellTreeViewItem object.
	///            - A \c ShellTreeViewItems object.
	///            To create the background context menu of a shell namespace, pass the
	///            \c ShellTreeViewNamespace object of the shell namespace.
	/// \param[out] pMenu The shell context menu's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DisplayShellContextMenu, ShellTreeViewItem::CreateShellContextMenu,
	///     ShellTreeViewItems::CreateShellContextMenu, ShellTreeViewNamespace::CreateShellContextMenu,
	///     DestroyShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE CreateShellContextMenu(VARIANT items, OLE_HANDLE* pMenu);
	/// \brief <em>Destroys the control's current shell context menu</em>
	///
	/// Destroys the shell context menu that was created by the last call to a \c CreateShellContextMenu
	/// method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellTreeViewItem::CreateShellContextMenu,
	///     ShellTreeViewItems::CreateShellContextMenu, ShellTreeViewNamespace::CreateShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DestroyShellContextMenu(void);
	/// \brief <em>Detaches the object from the specified \c SysTreeView32 window</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Attach, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE Detach(void);
	/// \brief <em>Displays a shell context menu</em>
	///
	/// \param[in] items An object specifying the treeview items for which to display the shell context
	///            menu. The following values may be used to identify the items:
	///            - A single item handle.
	///            - An array of item handles.
	///            - A \c TreeViewItem object.
	///            - A \c TreeViewItems object.
	///            - A \c TreeViewItemContainer object.
	///            - A \c ShellTreeViewItem object.
	///            - A \c ShellTreeViewItems object.
	///            To display the background context menu of a shell namespace, pass the
	///            \c ShellTreeViewNamespace object of the shell namespace.
	/// \param[in] x The x-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's position relative to the attached control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, ShellTreeViewItem::DisplayShellContextMenu,
	///     ShellTreeViewItems::DisplayShellContextMenu, ShellTreeViewNamespace::DisplayShellContextMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayShellContextMenu(VARIANT items, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Ensures that a specific shell item has been loaded into the treeview</em>
	///
	/// Walks the specified fully qualified pIDL or parsing name and inserts any items part of this path if
	/// they have not yet been inserted.
	///
	/// \param[in] pIDLOrParsingName The fully qualified pIDL or the fully qualified parsing name of the
	///            item to insert.
	/// \param[out] ppLastValidItem Receives the item identified by \c pIDLOrParsingName. If this item is
	///             invalid, its last valid parent item is returned.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeViewItems::Add, get_DefaultNamespaceEnumSettings, ShellTreeViewItem
	virtual HRESULT STDMETHODCALLTYPE EnsureItemIsLoaded(VARIANT pIDLOrParsingName, IShTreeViewItem** ppLastValidItem);
	/// \brief <em>Reloads all icons managed by the control</em>
	///
	/// \param[in] includeOverlays If set to \c VARIANT_TRUE, not only the icons managed by the control are
	///            flushed, but also the overlay images; otherwise only the icons are flushed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessShellNotifications, FlushIcons
	virtual HRESULT STDMETHODCALLTYPE FlushManagedIcons(VARIANT_BOOL includeOverlays = VARIANT_TRUE);
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
	/// \sa Raise_ChangedContextMenuSelection
	virtual HRESULT STDMETHODCALLTYPE GetShellContextMenuItemString(LONG commandID, VARIANT* pItemDescription = NULL, VARIANT* pItemVerb = NULL, VARIANT_BOOL* pCommandIDWasValid = NULL);
	/// \brief <em>Executes the default command from the specified items' shell context menu</em>
	///
	/// \param[in] items An object specifying the treeview items for which to invoke the default command.
	///            The following values may be used to identify the items:
	///            - A single item handle.
	///            - An array of item handles.
	///            - A \c TreeViewItem object.
	///            - A \c TreeViewItems object.
	///            - A \c TreeViewItemContainer object.
	///            - A \c ShellTreeViewItem object.
	///            - A \c ShellTreeViewItems object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellTreeViewItem::InvokeDefaultShellContextMenuCommand,
	///     ShellTreeViewItems::InvokeDefaultShellContextMenuCommand
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
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT /*category*/, LCID /*languageID*/, BSTR* /*pName*/);
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
	/// \sa CommonProperties, ShTvwDefaultManagedItemPropertiesProperties, NamespaceEnumSettingsProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IMessageListener
	///
	//@{
	/// \brief <em>Allows \c ShellTreeView to post-process a message</em>
	///
	/// This method is the very last one that processes a received message. It allows \c ShellTreeView to
	/// process the message after the treeview control and after \c DefWindowProc.
	///
	/// \param[in] hWnd The window that received the message.
	/// \param[in] message The received message.
	/// \param[in] wParam The received message's \c wParam parameter.
	/// \param[in] lParam The received message's \c lParam parameter.
	/// \param[out] pResult Receives the result of message processing.
	/// \param[in] cookie A value specified by the \c PreMessageFilter method.
	/// \param[in] eatenMessage If \c TRUE, the treeview control has eaten the message, i. e. it has not
	///            forwarded it to the default window procedure.
	///
	/// \return \c S_OK if the message was processed; \c S_FALSE if the message was not processed;
	///         \c E_NOTIMPL if we do not filter any messages; \c E_POINTER if \c pResult is an illegal
	///         pointer.
	///
	/// \sa PreMessageFilter, IMessageListener::PostMessageFilter
	virtual HRESULT PostMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, DWORD cookie, BOOL eatenMessage);
	/// \brief <em>Allows \c ShellTreeView to pre-process a message</em>
	///
	/// This method is the very first one that processes a received message. It allows \c ShellTreeView to
	/// process the message before the treeview control.
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
	/// \brief <em>\c TVM_GETITEM handler</em>
	///
	/// Will be called if some or all of a tree view item's attributes are requested from the attached
	/// control.
	/// We use this handler to intercept and handle requests of non-shell item's icon indexes.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms650168.aspx">TVM_GETITEM</a>
	LRESULT OnGetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c TVM_SETITEM handler</em>
	///
	/// Will be called if the attached control is requested to update some or all of a tree view item's
	/// attributes.
	/// We use this handler to intercept and handle changes of non-shell item's icon indexes.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms650189.aspx">TVM_SETITEM</a>
	LRESULT OnSetItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the attached control's parent window receives a notification from the attached
	/// control.
	/// We use this handler to do some shell stuff on some notifications.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled, DWORD& cookie, BOOL& eatenMessage, BOOL preProcessing);
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
	LRESULT OnTriggerEnqueueTask(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_ITEMENUMCOMPLETE handler</em>
	///
	/// Will be called if the attached control receives an item's sub-items retrieved by a
	/// \c ShTvwBackgroundItemEnumTask task.
	/// We use this handler to dynamically load sub-items.
	///
	/// \sa WM_TRIGGER_ITEMENUMCOMPLETE, ShTvwBackgroundItemEnumTask
	LRESULT OnTriggerItemEnumComplete(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_SETELEVATIONSHIELD handler</em>
	///
	/// Will be called if the attached control receives an item's elevation requirements retrieved by a
	/// \c ElevationShieldTask task.
	/// We use this handler to dynamically load UAC elevation shields.
	///
	/// \sa OnTriggerUpdateIcon, OnTriggerUpdateThumbnail, WM_TRIGGER_SETELEVATIONSHIELD, ElevationShieldTask
	LRESULT OnTriggerSetElevationShield(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATEICON handler</em>
	///
	/// Will be called if the attached control receives an item's icon index retrieved by a \c ShTvwIconTask
	/// task.
	/// We use this handler to dynamically load item icons.
	///
	/// \sa OnTriggerUpdateSelectedIcon, WM_TRIGGER_UPDATEICON, ShTvwIconTask, OnTriggerUpdateOverlay
	LRESULT OnTriggerUpdateIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATESELECTEDICON handler</em>
	///
	/// Will be called if the attached control receives an item's selected icon index retrieved by a
	/// \c ShTvwIconTask task.
	/// We use this handler to dynamically load item icons.
	///
	/// \sa OnTriggerUpdateIcon, WM_TRIGGER_UPDATESELECTEDICON, ShTvwIconTask
	LRESULT OnTriggerUpdateSelectedIcon(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	/// \brief <em>\c WM_TRIGGER_UPDATEOVERLAY handler</em>
	///
	/// Will be called if the attached control receives an item's overlay icon index retrieved by a
	/// \c ShTvwOverlayTask task.
	/// We use this handler to dynamically load item overlays.
	///
	/// \sa OnTriggerUpdateIcon, WM_TRIGGER_UPDATEOVERLAY, ShTvwOverlayTask
	LRESULT OnTriggerUpdateOverlay(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL& /*eatenMessage*/, BOOL /*preProcessing*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Internal message handlers
	///
	//@{
	/// \brief <em>Handles the \c SHTVM_ALLOCATEMEMORY message</em>
	///
	/// \sa SHTVM_ALLOCATEMEMORY
	HRESULT OnInternalAllocateMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHTVM_FREEMEMORY message</em>
	///
	/// \sa SHTVM_FREEMEMORY
	HRESULT OnInternalFreeMemory(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHTVM_INSERTINGITEM message</em>
	///
	/// \sa SHTVM_INSERTINGITEM
	HRESULT OnInternalInsertingItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHTVM_INSERTEDITEM message</em>
	///
	/// \sa SHTVM_INSERTEDITEM
	HRESULT OnInternalInsertedItem(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/);
	/// \brief <em>Handles the \c SHTVM_FREEITEM message</em>
	///
	/// \sa SHTVM_FREEITEM
	HRESULT OnInternalFreeItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_UPDATEDITEMHANDLE message</em>
	///
	/// \sa SHTVM_UPDATEDITEMHANDLE
	HRESULT OnInternalUpdatedItemHandle(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_GETSHTREEVIEWITEM message</em>
	///
	/// \sa SHTVM_GETSHTREEVIEWITEM
	HRESULT OnInternalGetShTreeViewItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_RENAMEDITEM message</em>
	///
	/// \sa SHTVM_RENAMEDITEM
	HRESULT OnInternalRenamedItem(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_COMPAREITEMS message</em>
	///
	/// \sa SHTVM_COMPAREITEMS
	HRESULT OnInternalCompareItems(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_BEGINLABELEDITA and \c SHTVM_BEGINLABELEDITW messages</em>
	///
	/// \sa SHTVM_BEGINLABELEDITA, SHTVM_BEGINLABELEDITW
	HRESULT OnInternalBeginLabelEdit(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_BEGINDRAGA and \c SHTVM_BEGINDRAGW messages</em>
	///
	/// \sa SHTVM_BEGINDRAGA, SHTVM_BEGINDRAGW
	HRESULT OnInternalBeginDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_HANDLEDRAGEVENT message</em>
	///
	/// \sa SHTVM_HANDLEDRAGEVENT
	HRESULT OnInternalHandleDragEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_CONTEXTMENU message</em>
	///
	/// \sa SHTVM_CONTEXTMENU
	HRESULT OnInternalContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_GETINFOTIP message</em>
	///
	/// \sa SHTVM_GETINFOTIP
	HRESULT OnInternalGetInfoTip(UINT /*message*/, WPARAM wParam, LPARAM lParam);
	/// \brief <em>Handles the \c SHTVM_CUSTOMDRAW message</em>
	///
	/// \sa SHTVM_CUSTOMDRAW
	HRESULT OnInternalCustomDraw(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filtered notification handlers
	///
	//@{
	/// \brief <em>\c TVN_GETDISPINFO handler</em>
	///
	/// Will be called if the attached control requests display information about a treeview item from its
	/// parent window.
	/// We use this handler to dynamically set the item's managed properties.
	///
	/// \sa ShellTreeViewItem::get_ManagedProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650146.aspx">TVN_GETDISPINFO</a>
	LRESULT OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/);
	/// \brief <em>\c TVN_ITEMEXPANDING handler</em>
	///
	/// Will be called if the attached control's parent window is notified, that a treeview item's sub-items
	/// are about to be collapsed or expanded.
	/// We use this handler to dynamically load the item's sub-items.
	///
	/// \sa OnFirstTimeItemExpanding, OnItemExpandedNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650149.aspx">TVN_ITEMEXPANDING</a>
	LRESULT OnItemExpandingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/);
	/// \brief <em>\c TVN_ITEMEXPANDED handler</em>
	///
	/// Will be called if the attached control's parent window is notified, that a treeview item's sub-items
	/// have been collapsed or expanded.
	/// We use this handler to clear or set the \c AII_USEEXPANDEDICON flag.
	///
	/// \sa OnItemExpandingNotification, AII_USEEXPANDEDICON,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650148.aspx">TVN_ITEMEXPANDED</a>
	LRESULT OnItemExpandedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled, DWORD& /*cookie*/, BOOL /*preProcessing*/);
	/// \brief <em>\c TVN_SELCHANGED handler</em>
	///
	/// Will be called if the attached control's parent window is notified, that the caret item has changed.
	/// We use this handler to clear or set the \c AII_USESELECTEDICON flag.
	///
	/// \sa AII_USESELECTEDICON,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650151.aspx">TVN_SELCHANGED</a>
	LRESULT OnCaretChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/, DWORD& /*cookie*/, BOOL /*preProcessing*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c ChangedContextMenuSelection event</em>
	///
	/// \param[in] hContextMenu The handle of the context menu. Will be \c NULL if the menu is closed.
	/// \param[in] commandID The ID of the new selected menu item. May be 0.
	/// \param[in] isCustomMenuItem If \c VARIANT_FALSE, the selected menu item has been inserted by
	///            \c ShellListView; otherwise it is a custom menu item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedContextMenuSelection,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedContextMenuSelection,
	///       Raise_SelectedShellContextMenuItem, GetShellContextMenuItemString
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedContextMenuSelection,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedContextMenuSelection,
	///       Raise_SelectedShellContextMenuItem, GetShellContextMenuItemString
	/// \endif
	inline HRESULT Raise_ChangedContextMenuSelection(OLE_HANDLE hContextMenu, LONG commandID, VARIANT_BOOL isCustomMenuItem);
	/// \brief <em>Raises the \c ChangedItemPIDL event</em>
	///
	/// \param[in] itemHandle The handle of the item that was updated.
	/// \param[in] previousPIDL The item's previous fully qualified pIDL.
	/// \param[in] newPIDL The item's new fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedItemPIDL,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedItemPIDL, Raise_ChangedNamespacePIDL
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedItemPIDL,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedItemPIDL, Raise_ChangedNamespacePIDL
	/// \endif
	inline HRESULT Raise_ChangedItemPIDL(HTREEITEM itemHandle, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL);
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
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedNamespacePIDL,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ChangedNamespacePIDL, Raise_ChangedItemPIDL
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ChangedNamespacePIDL,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ChangedNamespacePIDL, Raise_ChangedItemPIDL
	/// \endif
	inline HRESULT Raise_ChangedNamespacePIDL(IShTreeViewNamespace* pNamespace, PCIDLIST_ABSOLUTE previousPIDL, PCIDLIST_ABSOLUTE newPIDL);
	/// \brief <em>Raises the \c CreatedShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	///   \sa Proxy_IShTreeViewEvents::Fire_CreatedShellContextMenu,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::CreatedShellContextMenu, Raise_CreatingShellContextMenu,
	///       Raise_SelectedShellContextMenuItem
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_CreatedShellContextMenu,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::CreatedShellContextMenu, Raise_CreatingShellContextMenu,
	///       Raise_SelectedShellContextMenuItem
	/// \endif
	inline HRESULT Raise_CreatedShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG minimumCustomCommandID, VARIANT_BOOL* pCancelMenu);
	/// \brief <em>Raises the \c CreatingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
	/// \param[in,out] pContextMenuStyle A bit field influencing the content of the context menu. Any
	///                combination of the values defined by the \c ShellContextMenuStyleConstants
	///                enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_FALSE, the caller should not create and display the context
	///                menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_CreatingShellContextMenu,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::CreatingShellContextMenu, Raise_CreatedShellContextMenu,
	///       ShBrowserCtlsLibU::ShellContextMenuStyleConstants
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_CreatingShellContextMenu,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::CreatingShellContextMenu, Raise_CreatedShellContextMenu,
	///       ShBrowserCtlsLibA::ShellContextMenuStyleConstants
	/// \endif
	inline HRESULT Raise_CreatingShellContextMenu(LPDISPATCH pItems, ShellContextMenuStyleConstants* pContextMenuStyle, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c DestroyingShellContextMenu event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
	/// \param[in] hContextMenu The handle of the context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_DestroyingShellContextMenu,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::DestroyingShellContextMenu,
	///       Raise_CreatingShellContextMenu
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_DestroyingShellContextMenu,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::DestroyingShellContextMenu,
	///       Raise_CreatingShellContextMenu
	/// \endif
	inline HRESULT Raise_DestroyingShellContextMenu(LPDISPATCH pItems, OLE_HANDLE hContextMenu);
	/// \brief <em>Raises the \c InsertedItem event</em>
	///
	/// \param[in] pTreeItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertedItem,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::InsertedItem, Raise_InsertingItem,
	///       Raise_InsertedNamespace, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertedItem,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::InsertedItem, Raise_InsertingItem,
	///       Raise_InsertedNamespace, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertedItem(IShTreeViewItem* pTreeItem);
	/// \brief <em>Raises the \c InsertedNamespace event</em>
	///
	/// \param[in] pNamespace The inserted namespace.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertedNamespace,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::InsertedNamespace, Raise_InsertedItem,
	///       Raise_RemovingNamespace
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertedNamespace,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::InsertedNamespace, Raise_InsertedItem,
	///       Raise_RemovingNamespace
	/// \endif
	inline HRESULT Raise_InsertedNamespace(IShTreeViewNamespace* pNamespace);
	/// \brief <em>Raises the \c InsertingItem event</em>
	///
	/// \param[in] parentItemHandle The handle of the item that will be the immediate parent of the new
	///            item. If \c NULL, the item will become a root item.
	/// \param[in] fullyQualifiedPIDL The fully qualified pIDL of the item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertingItem,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_InsertingItem,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::InsertingItem, Raise_InsertedItem, Raise_RemovingItem
	/// \endif
	inline HRESULT Raise_InsertingItem(OLE_HANDLE parentItemHandle, LONG fullyQualifiedPIDL, VARIANT_BOOL* pCancelInsertion);
	/// \brief <em>Raises the \c InvokedShellContextMenuCommand event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_InvokedShellContextMenuCommand,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::InvokedShellContextMenuCommand,
	///       Raise_SelectedShellContextMenuItem, ShBrowserCtlsLibU::WindowModeConstants,
	///       ShBrowserCtlsLibU::CommandInvocationFlagsConstants
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_InvokedShellContextMenuCommand,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::InvokedShellContextMenuCommand,
	///       Raise_SelectedShellContextMenuItem, ShBrowserCtlsLibA::WindowModeConstants,
	///       ShBrowserCtlsLibA::CommandInvocationFlagsConstants
	/// \endif
	inline HRESULT Raise_InvokedShellContextMenuCommand(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants usedWindowMode, CommandInvocationFlagsConstants usedInvocationFlags);
	/// \brief <em>Raises the \c ItemEnumerationCompleted event</em>
	///
	/// \param[in] pNamespace The namespace whose items have been enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items have
	///            been enumerated.
	/// \param[in] aborted Specifies whether the enumeration has been aborted, e. g. because the client has
	///            removed the namespace before it completed. If \c VARIANT_TRUE, the enumeration has been
	///            aborted; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationCompleted,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationCompleted, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationTimedOut
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationCompleted,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationCompleted, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationTimedOut
	/// \endif
	inline HRESULT Raise_ItemEnumerationCompleted(LPDISPATCH pNamespace, VARIANT_BOOL aborted);
	/// \brief <em>Raises the \c ItemEnumerationStarted event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items are
	///            being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationStarted,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationStarted, Raise_ItemEnumerationTimedOut,
	///       Raise_ItemEnumerationCompleted
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationStarted,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationStarted, Raise_ItemEnumerationTimedOut,
	///       Raise_ItemEnumerationCompleted
	/// \endif
	HRESULT Raise_ItemEnumerationStarted(LPDISPATCH pNamespace);
	/// \brief <em>Raises the \c ItemEnumerationTimedOut event</em>
	///
	/// \param[in] pNamespace The namespace whose items are being enumerated. If the event refers to a
	///            namespace, this is a \c ShellTreeViewNamespace object. If it refers to an item's
	///            sub-items, this is a \c ShellTreeViewItem object wrapping the item whose sub-items are
	///            being enumerated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationTimedOut,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::ItemEnumerationTimedOut, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationCompleted
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_ItemEnumerationTimedOut,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::ItemEnumerationTimedOut, Raise_ItemEnumerationStarted,
	///       Raise_ItemEnumerationCompleted
	/// \endif
	inline HRESULT Raise_ItemEnumerationTimedOut(LPDISPATCH pNamespace);
	/// \brief <em>Raises the \c RemovingItem event</em>
	///
	/// \param[in] pTreeItem The item being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_RemovingItem,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::RemovingItem, Raise_RemovingNamespace,
	///       Raise_InsertedItem
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_RemovingItem,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::RemovingItem, Raise_RemovingNamespace,
	///       Raise_InsertedItem
	/// \endif
	inline HRESULT Raise_RemovingItem(IShTreeViewItem* pTreeItem);
	/// \brief <em>Raises the \c RemovingNamespace event</em>
	///
	/// \param[in] pNamespace The namespace being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_RemovingNamespace,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::RemovingNamespace, Raise_RemovingItem,
	///       Raise_InsertedNamespace
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_RemovingNamespace,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::RemovingNamespace, Raise_RemovingItem,
	///       Raise_InsertedNamespace
	/// \endif
	inline HRESULT Raise_RemovingNamespace(IShTreeViewNamespace* pNamespace);
	/// \brief <em>Raises the \c SelectedShellContextMenuItem event</em>
	///
	/// \param[in] pItems The items the context menu refers to. \c This is a ITreeViewItemContainer or
	///            a \c IShTreeViewNamespace object.
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
	/// \if UNICODE
	///   \sa Proxy_IShTreeViewEvents::Fire_SelectedShellContextMenuItem,
	///       ShBrowserCtlsLibU::_IShTreeViewEvents::SelectedShellContextMenuItem,
	///       Raise_InvokedShellContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedShellContextMenu, ShBrowserCtlsLibU::WindowModeConstants,
	///       ShBrowserCtlsLibU::CommandInvocationFlagsConstants
	/// \else
	///   \sa Proxy_IShTreeViewEvents::Fire_SelectedShellContextMenuItem,
	///       ShBrowserCtlsLibA::_IShTreeViewEvents::SelectedShellContextMenuItem,
	///       Raise_InvokedShellContextMenuCommand, Raise_ChangedContextMenuSelection,
	///       Raise_CreatedShellContextMenu, ShBrowserCtlsLibA::WindowModeConstants,
	///       ShBrowserCtlsLibA::CommandInvocationFlagsConstants
	/// \endif
	inline HRESULT Raise_SelectedShellContextMenuItem(LPDISPATCH pItems, OLE_HANDLE hContextMenu, LONG commandID, WindowModeConstants* pWindowModeToUse, CommandInvocationFlagsConstants* pInvocationFlagsToUse, VARIANT_BOOL* pExecuteCommand);
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
	/// \brief <em>Handles the specified item's first expansion</em>
	///
	/// Handles the specified item's first expansion. If it's a shell item, its sub-items are inserted;
	/// otherwise nothing is done.
	///
	/// \param[in] itemHandle The handle of the treeview item being expanded.
	/// \param[in] threadingMode Controls whether the items are enumerated synchronously or asynchronously.
	///            Any of the values defined by the \c ThreadingMode enumeration is valid.
	/// \param[in] checkForDuplicates If set to \c TRUE, the method takes care, that no item is inserted
	///            twice. This is used by the \c EnsureItemIsLoaded method, which inserts an item and then
	///            calls \c EnsureItemIsLoaded to insert the remaining items.
	/// \param[out] isShellItem If set to \c TRUE, the item is a shell item; otherwise not.
	///
	/// \return \c TRUE to prevent the item expansion; otherwise \c FALSE.
	///
	/// \sa OnItemExpandingNotification, ShellTreeViewItem::EnsureSubItemsAreLoaded, ThreadingMode
	BOOL OnFirstTimeItemExpanding(HTREEITEM itemHandle, ThreadingMode threadingMode, BOOL checkForDuplicates, BOOL& isShellItem);
	/// \brief <em>Makes the attached control use the system imagelists as specified by the \c UseSystemImageList property</em>
	///
	/// \sa ConfigureControl, get_UseSystemImageList
	void SetSystemImageLists(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Item management
	///
	//@{
	/// \brief <em>Adds the specified treeview item to the collection of treeview items</em>
	///
	/// \param[in] itemHandle The handle of the treeview item to add.
	/// \param[in] pItemData The data of the treeview item to add.
	///
	/// \sa AddNamespace, Properties::items, OnInternalInsertedItem
	void AddItem(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemData);
	/// \brief <em>Applies a shell treeview item's \c ManagedProperties property to the attached treeview</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to apply the property.
	/// \param[in] dontTouchSubItems If \c TRUE, only the item's text, icons and similar are updated;
	///            otherwise also its sub-items are re-loaded if the item has been expanded at least once.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeViewItem::get_ManagedProperties, IsShellItem
	HRESULT ApplyManagedProperties(HTREEITEM itemHandle, BOOL dontTouchSubItems = FALSE);
	/// \brief <em>Buffers an item's \c SHELLTREEVIEWITEMDATA struct until it is required by \c OnInternalInsertedItem</em>
	///
	/// \param[in] pItemData The data to buffer.
	///
	/// \sa Properties::shellItemDataOfInsertedItems, OnInternalInsertedItem, SHELLTREEVIEWITEMDATA,
	///     ShellTreeViewItems::Add, BufferShellItemNamespace
	void BufferShellItemData(LPSHELLTREEVIEWITEMDATA pItemData);
	/// \brief <em>Inserts a new shell item faster than \c ShellTreeViewItems::Add</em>
	///
	/// \param[in] hParentItem The parent item of the item to insert.
	/// \param[in] pIDL The fully qualified pIDL of the item to insert.
	/// \param[in] hInsertAfter The new item's preceding item.
	/// \param[in] hasExpando A value specifying whether to draw a "+" or "-" next to the item
	///            indicating the item has sub-items. Any of the values defined by \c ExplorerTreeView's
	///            \c HasExpandoConstants enumeration is valid. If set to \c heCallback, the treeview control
	///            will fire its \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \return The new item's handle.
	///
	/// \sa OnFirstTimeItemExpanding, OnTriggerItemEnumComplete
	HTREEITEM FastInsertShellItem(OLE_HANDLE hParentItem, PIDLIST_ABSOLUTE pIDL, HTREEITEM hInsertAfter, LONG hasExpando);
	/// \brief <em>Retrieves the specified shell treeview item's details</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the details.
	///
	/// \return The item's details or \c NULL if the item isn't a shell item.
	///
	/// \sa ItemHandleToPIDL, SHELLTREEVIEWITEMDATA, IsShellItem
	LPSHELLTREEVIEWITEMDATA GetItemDetails(HTREEITEM itemHandle);
	/// \brief <em>Retrieves whether the specified treeview item is a shell treeview item</em>
	///
	/// \param[in] itemHandle The handle of the treeview item to check.
	///
	/// \return \c TRUE if the item is a shell treeview item; otherwise \c FALSE.
	///
	/// \sa GetItemDetails, ItemHandleToPIDL, ApplyManagedProperties
	BOOL IsShellItem(HTREEITEM itemHandle);
	/// \brief <em>Retrieves the specified shell treeview item's unique ID</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the unique ID.
	///
	/// \return The item's unique ID or -1 if an error occured.
	///
	/// \sa EXTVM_ITEMHANDLETOID, ItemIDToHandle
	LONG ItemHandleToID(HTREEITEM itemHandle);
	/// \brief <em>Retrieves the specified shell treeview item's handle</em>
	///
	/// \param[in] itemID The unique ID of the treeview item for which to retrieve the handle.
	///
	/// \return The item's handle or \c NULL if an error occured.
	///
	/// \sa EXTVM_ITEMIDTOHANDLE, ItemHandleToID
	HTREEITEM ItemIDToHandle(LONG itemID);
	/// \brief <em>Retrieves the specified shell treeview item's fully qualified pIDL</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the pIDL.
	///
	/// \return The item's fully qualified pIDL or \c NULL if the item isn't a shell item.
	///
	/// \sa ItemHandlesToPIDLs, ItemHandleToNamespace, PIDLToItemHandle, GetItemDetails, IsShellItem
	PCIDLIST_ABSOLUTE ItemHandleToPIDL(HTREEITEM itemHandle);
	/// \brief <em>Retrieves the specified shell treeview items' fully qualified pIDLs</em>
	///
	/// \param[in] pItemHandles The handles of the treeview items for which to retrieve the pIDLs.
	/// \param[in] itemCount The number of elements in the array \c pItemHandles.
	/// \param[out] ppIDLs Receives the pIDLs of the treeview items.
	/// \param[in] keepNonShellItems If \c TRUE, the pIDLs of treeview items that are no shell items, are set
	///            to \c NULL. Otherwise \c ppIDLs does not contain these items.
	///
	/// \return The number of elements in the array returned in \c ppIDLs.
	///
	/// \sa ItemHandleToPIDL
	UINT ItemHandlesToPIDLs(HTREEITEM* pItemHandles, UINT itemCount, PIDLIST_ABSOLUTE*& ppIDLs, BOOL keepNonShellItems);
	/// \brief <em>Retrieves the shell namespace that the specified shell treeview item belongs to</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the namespace.
	///
	/// \return A \c ShellTreeViewNamespace object wrapping the shell namespace the item belongs to or
	///         \c NULL if the item isn't a shell item.
	///
	/// \sa ItemHandleToNamespacePIDL, ItemHandleToNameSpaceEnumSettings, ItemHandleToPIDL, GetItemDetails
	CComObject<ShellTreeViewNamespace>* ItemHandleToNamespace(HTREEITEM itemHandle);
	/// \brief <em>Retrieves the enumeration settings that shall be used for the specified shell treeview item</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the enumeration settings.
	/// \param[in] pItemDetails Optional. The treeview item's details. If set to \c NULL (the default), the
	///            details are retrieved by the method.
	///
	/// \return A \c INamespaceEnumSettings object wrapping the enumeration settings to use. \c NULL if the
	///         item's sub-items aren't managed by the control or if an error occured.
	///
	/// \sa INamespaceEnumSettings, ItemHandleToNamespace, GetItemDetails
	INamespaceEnumSettings* ItemHandleToNameSpaceEnumSettings(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemDetails = NULL);
	/// \brief <em>Retrieves the fully qualified pIDL of the shell namespace that the specified shell treeview item belongs to</em>
	///
	/// \param[in] itemHandle The handle of the treeview item for which to retrieve the namespace pIDL.
	///
	/// \return The fully qualified pIDL of the shell namespace the item belongs to or \c NULL if the item
	///         isn't a shell item.
	///
	/// \sa ItemHandleToNamespace, ItemHandleToPIDL, GetItemDetails
	PCIDLIST_ABSOLUTE ItemHandleToNamespacePIDL(HTREEITEM itemHandle);
	/// \brief <em>Retrieves the item handle of the specified fully qualified pIDL</em>
	///
	/// \param[in] pIDL The fully qualified pIDL for which to retrieve the shell item's handle.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The item handle or NULL if no matching item was found.
	///
	/// \sa ParsingNameToItemHandle, VariantToItemHandles, ItemHandleToPIDL
	HTREEITEM PIDLToItemHandle(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch);
	/// \brief <em>Retrieves the item handle of the specified fully qualified parsing name</em>
	///
	/// \param[in] pParsingName The fully qualified parsing name of the shell item for which to retrieve the
	///            shell item's handle.
	///
	/// \return The item handle or NULL if no matching item was found.
	///
	/// \sa PIDLToItemHandle, VariantToItemHandles, ItemHandleToPIDL
	HTREEITEM ParsingNameToItemHandle(LPCWSTR pParsingName);
	/// \brief <em>Retrieves the item handles of the specified treeview items</em>
	///
	/// \param[in] items A \c VARIANT specifying the items for which to retrieve the handles. The \c VARIANT
	///            may contain the following data:
	///            - A single item handle.
	///            - An array of item handles.
	///            - A \c TreeViewItem object.
	///            - A \c TreeViewItems object.
	///            - A \c TreeViewItemContainer object.
	///            - A \c ShellTreeViewItem object.
	///            - A \c ShellTreeViewItems object.
	/// \param[out] pItemHandles Receives the item handles.
	/// \param[out] itemCount Receives the number of elements in the array returned in \c pItemHandles.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PIDLToItemHandle
	HRESULT VariantToItemHandles(VARIANT& items, HTREEITEM*& pItemHandles, UINT& itemCount);
	/// \brief <em>Calls \c SHBindToParent for the specified shell treeview item</em>
	///
	/// \param[in] itemHandle The shell treeview item for which to call \c SHBindToParent.
	/// \param[in] requiredInterface The \c riid parameter of \c SHBindToParent.
	/// \param[out] ppInterfaceImpl The \c ppv parameter of \c SHBindToParent.
	/// \param[out] pPIDLLast The \c ppidlLast parameter of \c SHBindToParent.
	/// \param[out] pPIDL The item's fully qualified pIDL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms647661.aspx">SHBindToParent</a>
	HRESULT SHBindToParent(HTREEITEM itemHandle, REFIID requiredInterface, LPVOID* ppInterfaceImpl, PCUITEMID_CHILD* pPIDLLast, PCIDLIST_ABSOLUTE* pPIDL = NULL);
	/// \brief <em>Checks whether the specified shell item still exists</em>
	///
	/// \param[in] itemHandle The handle of the treeview item to validate.
	/// \param[in] pItemDetails Optional. The treeview item's details. If set to \c NULL (the default), the
	///            details are retrieved by the method.
	///
	/// \return \c TRUE if the item still exists; otherwise \c FALSE.
	///
	/// \sa ShellTreeViewItem::Validate, ValidatePIDL
	BOOL ValidateItem(HTREEITEM itemHandle, LPSHELLTREEVIEWITEMDATA pItemDetails = NULL);
	/// \brief <em>Updates the specified item's pIDL</em>
	///
	/// \param[in] itemHandle The handle of the treeview item to update.
	/// \param[in] pItemDetails The treeview item's details.
	/// \param[in] pIDLNew The fully qualified new pIDL. If the new pIDL consists of as many item ids as the
	///            item's current pIDL, the pIDL is simply exchanged. If it contains less item ids, the head
	///            of the item's current pIDL is replaced with the new pIDL.
	///
	/// \remarks This method does not call \c ApplyManagedProperties.
	///
	/// \sa Raise_ChangedItemPIDL, ApplyManagedProperties
	void UpdateItemPIDL(HTREEITEM itemHandle, __in LPSHELLTREEVIEWITEMDATA pItemDetails, __in PIDLIST_ABSOLUTE pIDLNew);
	/// \brief <em>Removes sub-items of collapsed items to improve performance</em>
	///
	/// Scans a sub tree for items with sub-items that are collapsed and collapses these items again with
	/// the reset flag being set so that the sub-items get removed. This may reduce the number of items
	/// significantly, improving auto-update performance.
	///
	/// \param[in] hParentItem The handle of the treeview item to check.
	///
	/// \sa OnSHChangeNotify_UPDATEDIR
	void RemoveCollapsedSubTree(HTREEITEM hParentItem);
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
	/// \sa AddItem, Properties::namespaces, ShellTreeViewNamespaces::Add
	void AddNamespace(PCIDLIST_ABSOLUTE pIDL, CComObject<ShellTreeViewNamespace>* pNamespace);
	/// \brief <em>Buffers an item's associated namespace pIDL until it is required by \c OnInternalInsertedItem</em>
	///
	/// \param[in] pIDLItem The fully qualified pIDL of the item for which to buffer the namespace pIDL.
	/// \param[in] pIDLNamespace The fully qualified pIDL of the item's namespace.
	///
	/// \sa Properties::namespacesOfInsertedItems, OnInternalInsertedItem, ShellTreeViewNamespaces::Add,
	///     RemoveBufferedShellItemNamespace, BufferShellItemData
	void BufferShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem, PCIDLIST_ABSOLUTE pIDLNamespace);
	/// \brief <em>Deletes an item's buffered namespace pIDL from the buffer</em>
	///
	/// \param[in] pIDLItem The fully qualified pIDL of the item.
	///
	/// \sa Properties::namespacesOfInsertedItems, OnInternalInsertedItem, ShellTreeViewNamespaces::Add,
	///     BufferShellItemNamespace, BufferShellItemData
	void RemoveBufferedShellItemNamespace(PCIDLIST_ABSOLUTE pIDLItem);
	/// \brief <em>Retrieves the n-th \c ShellTreeViewNamespace object</em>
	///
	/// \param[in] index The zero-based index for which to retrieve the shell namespace object.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespacePIDL, NamespacePIDLToIndex, NamespacePIDLToNamespace,
	///     NamespaceParsingNameToNamespace
	CComObject<ShellTreeViewNamespace>* IndexToNamespace(UINT index);
	/// \brief <em>Retrieves the fully qualified pIDL of the n-th \c ShellTreeViewNamespace object</em>
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
	/// \brief <em>Retrieves the \c ShellTreeViewNamespace object of the specified fully qualified pIDL</em>
	///
	/// \param[in] pIDL The fully qualified pIDL for which to retrieve the shell namespace object.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespace, NamespaceParsingNameToNamespace
	CComObject<ShellTreeViewNamespace>* NamespacePIDLToNamespace(PCIDLIST_ABSOLUTE pIDL, BOOL exactMatch);
	/// \brief <em>Retrieves the \c ShellListViewNamespace object of the specified fully qualified parsing name</em>
	///
	/// \param[in] pParsingName The fully qualified parsing name of the shell item for which to retrieve the
	///            shell namespace object.
	///
	/// \return The namespace object or NULL if no matching namespace was found.
	///
	/// \sa IndexToNamespace, NamespacePIDLToNamespace
	CComObject<ShellTreeViewNamespace>* NamespaceParsingNameToNamespace(LPWSTR pParsingName);
	/// \brief <em>Removes a shell namespace</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the shell namespace to remove.
	/// \param[in] exactMatch If \c TRUE, the pIDLs are compared directly (using the \c == operator);
	///            otherwise \c ILIsEqual is used.
	/// \param[in] removeItemsFromTreeView If \c TRUE, the namespace's items are removed from the attached
	///            treeview; otherwise they are converted to normal items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RemoveAllNamespaces, get_Namespaces, ShellTreeViewNamespaces::RemoveAll
	HRESULT RemoveNamespace(PCIDLIST_ABSOLUTE pIDLNamespace, BOOL exactMatch, BOOL removeItemsFromTreeView);
	/// \brief <em>Removes all shell namespaces</em>
	///
	/// \param[in] removeItemsFromTreeView If \c TRUE, the namespaces' items are removed from the attached
	///            treeview; otherwise they are converted to normal items.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RemoveNamespace, get_Namespaces, ShellTreeViewNamespaces::RemoveAll
	HRESULT RemoveAllNamespaces(BOOL removeItemsFromTreeView);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Context menu support
	///
	//@{
	/// \brief <em>Creates the shell context menu for the specified treeview items</em>
	///
	/// \param[in] pItems The treeview items for which to create the context menu.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	/// \param[in] predefinedFlags Specifies the \c CMF_ flags that shall always be set.
	/// \param[out] hMenu Receives the handle of the created shell context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeViewItem::CreateShellContextMenu, ShellTreeViewItems::CreateShellContextMenu,
	///     ShellTreeViewNamespace::CreateShellContextMenu, DestroyShellContextMenu, DisplayShellContextMenu
	HRESULT CreateShellContextMenu(HTREEITEM* pItems, UINT itemCount, UINT predefinedFlags, HMENU& hMenu);
	/// \brief <em>Creates the shell context menu for the specified shell namespace</em>
	///
	/// \param[in] pIDLNamespace The fully qualified pIDL of the shell namespace for which to create the
	///            context menu.
	/// \param[out] hMenu Receives the handle of the created shell context menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ShellTreeViewNamespace::CreateShellContextMenu, ShellTreeViewItem::CreateShellContextMenu,
	///     ShellTreeViewItems::CreateShellContextMenu, DestroyShellContextMenu, DisplayShellContextMenu
	HRESULT CreateShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, HMENU& hMenu);
	/// \brief <em>Displays the shell context menu for the specified treeview items at the specified position</em>
	///
	/// \param[in] pItems The treeview items for which to show the context menu.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	/// \param[in] position The coordinates (in pixels) of the position relative to the screen at which the
	///            context menu shall be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnInternalContextMenu, ShellTreeViewItem::DisplayShellContextMenu,
	///     ShellTreeViewItems::DisplayShellContextMenu, ShellTreeViewNamespace::DisplayShellContextMenu,
	///     CreateShellContextMenu
	HRESULT DisplayShellContextMenu(HTREEITEM* pItems, UINT itemCount, POINT& position);
	/// \brief <em>Displays the shell context menu for the specified shell namespace at the specified position</em>
	///
	/// \param[in] pIDLNamespace The fully qualified pIDL of the shell namespace for which to show the
	///            context menu.
	/// \param[in] position The coordinates (in pixels) of the position relative to the screen at which the
	///            context menu shall be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnInternalContextMenu, ShellTreeViewNamespace::DisplayShellContextMenu,
	///     ShellTreeViewItem::DisplayShellContextMenu, ShellTreeViewItems::DisplayShellContextMenu,
	///     CreateShellContextMenu
	HRESULT DisplayShellContextMenu(PCIDLIST_ABSOLUTE pIDLNamespace, POINT& position);
	/// \brief <em>Invokes the default shell context menu command for the specified treeview items</em>
	///
	/// \param[in] pItems The treeview items for which to invoke the command.
	/// \param[in] itemCount The number of elements in the array \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CreateShellContextMenu, DisplayShellContextMenu,
	///     ShellTreeViewItem::InvokeDefaultShellContextMenuCommand,
	///     ShellTreeViewItems::InvokeDefaultShellContextMenuCommand
	HRESULT InvokeDefaultShellContextMenuCommand(HTREEITEM* pItems, UINT itemCount);

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
	/// \brief <em>Deregisters the attached treeview control for shell notifications</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RegisterForShellNotifications, put_ProcessShellNotifications
	HRESULT DeregisterForShellNotifications(void);
	/// \brief <em>Registers the attached treeview control for shell notifications</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DeregisterForShellNotifications, put_ProcessShellNotifications
	HRESULT RegisterForShellNotifications(void);

	/// \brief <em>Reloads any icons managed by the control</em>
	///
	/// \param[in] iconIndex The zero-based index of the icon in the system imagelist to update. Any icon
	///            managed by the control that is currently set to this icon, will be updated.\n
	///            If set to \c -1, all icons managed by the control will be updated. Additionally the system
	///            imagelist will be rebuilt and re-assigned if it is managed by the control.
	/// \param[in] flushOverlays Specifies whether overlay icons managed by the control will be updated.
	///            If set to \c TRUE, overlays are updated; otherwise not.
	///
	/// \sa OnSHChangeNotify_UPDATEIMAGE, FlushSystemImageList
	void FlushIcons(int iconIndex, BOOL flushOverlays);
	/// \brief <em>Searches the immediate sub-items of the specified item for a shell item with the specified pIDL</em>
	///
	/// \param[in] hParentItem The treeview item whose immediate sub-items shall be checked.
	/// \param[in] pIDL The absolute pIDL to search.
	/// \param[in] simplePIDL If \c TRUE, \c pIDL is a simple pIDL; otherwise not.
	///
	/// \return The handle of the found immediate sub-item or \c NULL if no matching item was found.
	///
	/// \sa OnTriggerItemEnumComplete
	HTREEITEM ItemFindImmediateSubItem(HTREEITEM hParentItem, __in PIDLIST_ABSOLUTE pIDL, BOOL simplePIDL);
	/// \brief <em>Checks whether the specified treeview item has been expanded at least once</em>
	///
	/// \param[in] itemHandle The treeview item to check.
	///
	/// \return \c TRUE if the item has been expanded at least once; otherwise \c FALSE.
	BOOL ItemHasBeenExpandedOnce(HTREEITEM itemHandle);
	#ifdef USE_STL
		/// \brief <em>Removes the specified items from the control</em>
		///
		/// \param[in] itemsToRemove The items to remove.
		///
		/// \sa SelectNewCaret
		void RemoveItems(std::vector<HTREEITEM>& itemsToRemove);
		/// \brief <em>Selects a new caret item if the current caret item is being removed</em>
		///
		/// If the caret item is removed, SysTreeView32 changes the caret automatically. The item it selects
		/// isn't always the best choice, e. g. it might be an item which will be removed, too, so not only do
		/// we get multiple caret changes (which are expensive for most use cases of this control), but also
		/// chances are high, that the new caret item refers to a shell item that has become invalid and we'll
		/// run into trouble.\n
		/// To avoid this, this method can be called. It takes the list of items that will be removed and
		/// searches for a new caret item not on this list. Additionally a shell item may be specified, which
		/// won't become the new caret item except it already is the caret item.
		///
		/// \param[in] itemsBeingRemoved The items being removed. The new caret item will be neither in this
		///            list nor a sub item of an item in this list.
		/// \param[in] pIDLToAvoid Optional. If set to an absolute pIDL, the new caret item won't be an item
		///            whose pIDL equals the pIDL specified by this parameter.
		///
		/// \sa RemoveItems
		void SelectNewCaret(std::vector<HTREEITEM>& itemsBeingRemoved, PCIDLIST_ABSOLUTE pIDLToAvoid);
	#else
		/// \brief <em>Removes the specified items from the control</em>
		///
		/// \param[in] itemsToRemove The items to remove.
		///
		/// \sa SelectNewCaret
		void RemoveItems(CAtlList<HTREEITEM>& itemsToRemove);
		/// \brief <em>Selects a new caret item if the current caret item is being removed</em>
		///
		/// If the caret item is removed, SysTreeView32 changes the caret automatically. The item it selects
		/// isn't always the best choice, e. g. it might be an item which will be removed, too, so not only do
		/// we get multiple caret changes (which are expensive for most use cases of this control), but also
		/// chances are high, that the new caret item refers to a shell item that has become invalid and we'll
		/// run into trouble.\n
		/// To avoid this, this method can be called. It takes the list of items that will be removed and
		/// searches for a new caret item not on this list. Additionally a shell item may be specified, which
		/// won't become the new caret item except it already is the caret item.
		///
		/// \param[in] itemsBeingRemoved The items being removed. The new caret item will be neither in this
		///            list nor a sub item of an item in this list.
		/// \param[in] pIDLToAvoid Optional. If set to an absolute pIDL, the new caret item won't be an item
		///            whose pIDL equals the pIDL specified by this parameter.
		///
		/// \sa RemoveItems
		void SelectNewCaret(CAtlList<HTREEITEM>& itemsBeingRemoved, PCIDLIST_ABSOLUTE pIDLToAvoid);
	#endif

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
	/// \sa OnSHChangeNotify, ShellTreeViewNamespace::OnSHChangeNotify_CREATE
	LRESULT OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE_ARRAY ppIDLs);
	/// \brief <em>Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications for a namespace</em>
	///
	/// Handles the \c SHCNE_MKDIR and \c SHCNE_CREATE notifications for the specified affected namespace.
	/// This notifications are sent if a directory (file) has been created.
	///
	/// \param[in] simplePIDLAddedDir The (simple) pIDL of the created directory (file).
	/// \param[in] pIDLNamespace The pIDL of the affected namespace.
	/// \param[in] hParentItem The affected namespace's root item.
	/// \param[in] pEnumSettings The value of the affected namespace's \c NamespaceEnumSettings property.
	///
	/// \return Ignored.
	///
	/// \sa OnSHChangeNotify, ShellTreeViewNamespace::OnSHChangeNotify_CREATE,
	///     ShellTreeViewNamespace::get_ParentItemHandle, ShellTreeViewNamespace::get_NamespaceEnumSettings,
	///     OnSHChangeNotify_DELETE
	LRESULT OnSHChangeNotify_CREATE(PCIDLIST_ABSOLUTE simplePIDLAddedDir, PCIDLIST_ABSOLUTE pIDLNamespace, HTREEITEM hParentItem, INamespaceEnumSettings* pEnumSettings);
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
	/// \sa OnSHChangeNotify, ShellTreeViewNamespace::OnSHChangeNotify_DELETE, OnSHChangeNotify_CREATE
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
	/// \sa OnSHChangeNotify, ShellTreeViewNamespace::OnSHChangeNotify_RENAME
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
	/// \sa GetIconsAsync, GetOverlayAsync,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775201.aspx">IRunnableTask</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632523.aspx">IRunnableTask::AddTask</a>
	HRESULT EnqueueTask(__in IRunnableTask* pTask, REFTASKOWNERID taskGroupID, DWORD_PTR lParam, DWORD priority, DWORD operationStart = 0);
	/// \brief <em>Creates a runnable task which retrieves the specified item's icon in a separate thread</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the icon indexes.
	/// \param[in] itemHandle The item for which to retrieve the icon indexes.
	/// \param[in] retrieveNormalImage If \c TRUE, the item's normal icon is retrieved.
	/// \param[in] retrieveSelectedImage If \c TRUE, the item's selected icon is retrieved.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetOverlayAsync, EnqueueTask, OnGetDispInfoNotification
	HRESULT GetIconsAsync(__in PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle, BOOL retrieveNormalImage, BOOL retrieveSelectedImage);
	/// \brief <em>Creates a runnable task which retrieves the specified item's overlay icon in a separate thread</em>
	///
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the overlay icon.
	/// \param[in] itemHandle The item for which to retrieve the overlay icon.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconsAsync, EnqueueTask, OnGetDispInfoNotification
	HRESULT GetOverlayAsync(__in PCIDLIST_ABSOLUTE pIDL, HTREEITEM itemHandle);
	/// \brief <em>Creates a runnable task which retrieves the specified item's elevation requirements</em>
	///
	/// \param[in] itemID The item for which to retrieve the elevation requirements.
	/// \param[in] pIDL The fully qualified pIDL of the item for which to retrieve the elevation
	///            requirements.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetIconAsync, EnqueueTask, OnTriggerUpdateIcon
	HRESULT GetElevationShieldAsync(LONG itemID, __in PCIDLIST_ABSOLUTE pIDL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Wraps the attached treeview window</em>
	///
	/// \sa get_hWnd, Attach, Detach
	CWindow attachedControl;

	/// \brief <em>Holds the object's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>An empty \c VARIANT shared between time-critical methods</em>
		VARIANT emptyVariant;
		/// \brief <em>A \c EXTVMADDITEMDATA struct prepared to be used by \c FastInsertShellItem</em>
		///
		/// \sa FastInsertShellItem
		EXTVMADDITEMDATA itemDataForFastInsertion;

		/// \brief <em>Holds the \c AutoEditNewItems property's setting</em>
		///
		/// \sa get_AutoEditNewItems, put_AutoEditNewItems
		UINT autoEditNewItems : 1;
		/// \brief <em>Holds the \c ColorCompressedItems property's setting</em>
		///
		/// \sa get_ColorCompressedItems, put_ColorCompressedItems
		UINT colorCompressedItems : 1;
		/// \brief <em>Holds the \c ColorEncryptedItems property's setting</em>
		///
		/// \sa get_ColorEncryptedItems, put_ColorEncryptedItems
		UINT colorEncryptedItems : 1;
		/// \brief <em>Holds the \c DefaultManagedItemProperties property's setting</em>
		///
		/// \sa get_DefaultManagedItemProperties, put_DefaultManagedItemProperties
		ShTvwManagedItemPropertiesConstants defaultManagedItemProperties;
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
		/// \brief <em>Holds the \c LoadOverlaysOnDemand property's setting</em>
		///
		/// \sa get_LoadOverlaysOnDemand, put_LoadOverlaysOnDemand
		UINT loadOverlaysOnDemand : 1;
		/// \brief <em>Holds the \c Namespaces property's setting</em>
		///
		/// \sa get_Namespaces
		CComObject<ShellTreeViewNamespaces>* pNamespaces;
		/// \brief <em>Holds the \c PreselectBasenameOnLabelEdit property's setting</em>
		///
		/// \sa get_PreselectBasenameOnLabelEdit, put_PreselectBasenameOnLabelEdit
		UINT preselectBasenameOnLabelEdit : 1;
		/// \brief <em>Holds the \c ProcessShellNotifications property's setting</em>
		///
		/// \sa get_ProcessShellNotifications, put_ProcessShellNotifications
		UINT processShellNotifications : 1;
		/// \brief <em>Holds the \c TreeItems property's setting</em>
		///
		/// \sa get_TreeItems
		CComObject<ShellTreeViewItems>* pTreeItems;
		/// \brief <em>Holds the \c UseGenericIcons property's setting</em>
		///
		/// \sa get_UseGenericIcons, put_UseGenericIcons
		UseGenericIconsConstants useGenericIcons;
		/// \brief <em>Holds the \c UseSystemImageList property's setting</em>
		///
		/// \sa get_UseSystemImageList, put_UseSystemImageList
		UseSystemImageListConstants useSystemImageList;

		/// \brief <em>A pointer to the attached control's \c IUnknown implementation</em>
		///
		/// \remarks This member is used only to hold a reference on the attached control.
		LPUNKNOWN pAttachedCOMControl;
		/// \brief <em>Holds the \c IInternalMessageListener implementation of the attached control</em>
		IInternalMessageListener* pAttachedInternalMessageListener;
		/// \brief <em>Holds the build number of the attached control</em>
		UINT attachedControlBuildNumber;

		#ifdef USE_STL
			/// \brief <em>A map of all treeview items managed by this object</em>
			///
			/// Holds the details of all shell-items in the attached treeview that are managed by this object.
			/// The item's handle is stored as key, the item's details are stored as value.
			std::unordered_map<HTREEITEM, LPSHELLTREEVIEWITEMDATA> items;
			/// \brief <em>A map of all shell namespaces managed by this object</em>
			///
			/// Holds all shell namespaces in the attached treeview that are managed by this object. The
			/// namespace's fully qualified pIDL is stored as key, the namespace object is stored as value.
			std::unordered_map<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*> namespaces;
		#else
			/// \brief <em>A map of all treeview items managed by this object</em>
			///
			/// Holds the details of all shell-items in the attached treeview that are managed by this object.
			/// The item's handle is stored as key, the item's details are stored as value.
			CAtlMap<HTREEITEM, LPSHELLTREEVIEWITEMDATA> items;
			/// \brief <em>A map of all shell namespaces managed by this object</em>
			///
			/// Holds all shell namespaces in the attached treeview that are managed by this object. The
			/// namespace's fully qualified pIDL is stored as key, the namespace object is stored as value.
			CAtlMap<PCIDLIST_ABSOLUTE, CComObject<ShellTreeViewNamespace>*> namespaces;
		#endif
		/// \brief <em>The Desktop's \c IShellFolder implementation</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb775075.aspx">IShellFolder</a>
		LPSHELLFOLDER pDesktopISF;
		/// \brief <em>The \c IShellTaskScheduler implementation of the task scheduler we use for multithreading</em>
		///
		/// This is the task scheduler that is used for item enumeration tasks.
		///
		/// \sa ShTvwBackgroundItemEnumTask,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774864.aspx">IShellTaskScheduler</a>
		CComPtr<IShellTaskScheduler> pEnumTaskScheduler;
		/// \brief <em>The \c IShellTaskScheduler implementation of the task scheduler we use for multithreading</em>
		///
		/// This is the task scheduler that is used for the following tasks:
		/// - icon retrieval
		/// - overlay retrieval
		/// - insertion of single items in response to a shell notification
		///
		/// \sa ShTvwIconTask, ShTvwOverlayTask, ShTvwInsertSingleItemTask,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774864.aspx">IShellTaskScheduler</a>
		CComPtr<IShellTaskScheduler> pNormalTaskScheduler;
		#ifdef USE_STL
			/// \brief <em>Buffers the items enumerated by the background threads until they are inserted into the tree view</em>
			///
			/// \sa ShTvwBackgroundItemEnumTask, ShTvwInsertSingleItemTask, SHTVWBACKGROUNDITEMENUMINFO
			std::queue<LPSHTVWBACKGROUNDITEMENUMINFO> backgroundEnumResults;
		#else
			/// \brief <em>Buffers the items enumerated by the background threads until they are inserted into the tree view</em>
			///
			/// \sa ShTvwBackgroundItemEnumTask, ShTvwInsertSingleItemTask, SHTVWBACKGROUNDITEMENUMINFO
			CAtlList<LPSHTVWBACKGROUNDITEMENUMINFO> backgroundEnumResults;
		#endif
		/// \brief <em>The critical section used to synchronize access to \c backgroundEnumResults</em>
		///
		/// \sa backgroundEnumResults
		CRITICAL_SECTION backgroundEnumResultsCritSection;
		/// \brief <em>The default icon index for files in the system imagelist</em>
		int DEFAULTICON_FILE;
		/// \brief <em>The default icon index for folders in the system imagelist</em>
		int DEFAULTICON_FOLDER;
		/// \brief <em>Holds the handle of the unified image list that contains the items' images</em>
		///
		/// \sa UnifiedImageList_CreateInstance
		CComPtr<IImageList> pUnifiedImageList;

		#ifdef USE_STL
			/// \brief <em>Holds the \c LPSHELLTREEVIEWITEMDATA pointers that shall be stored in \c OnInternalInsertedItem</em>
			///
			/// \sa BufferShellItemData, OnInternalInsertedItem, SHELLTREEVIEWITEMDATA, ShellTreeViewItems::Add,
			///     namespacesOfInsertedItems
			std::stack<LPSHELLTREEVIEWITEMDATA> shellItemDataOfInsertedItems;
			/// \brief <em>Holds the namespace pIDLs of items that are going to be inserted</em>
			///
			/// \sa BufferShellItemNamespace, OnInternalInsertedItem, ShellTreeViewNamespaces::Add,
			///     shellItemDataOfInsertedItems
			std::unordered_map<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacesOfInsertedItems;
			/// \brief <em>Holds details about currently running background item enumerations</em>
			///
			/// Holds details about currently running background item enumerations. These details include the
			/// time the task has been started and whether to check for timeouts.\n
			/// The enumeration task's unique ID is stored as key, the details are stored as value.
			///
			/// \sa EnqueueTask, BACKGROUNDENUMERATION
			std::unordered_map<ULONGLONG, LPBACKGROUNDENUMERATION> backgroundEnums;
		#else
			/// \brief <em>Holds the \c LPSHELLTREEVIEWITEMDATA pointers that shall be stored in \c OnInternalInsertedItem</em>
			///
			/// \sa BufferShellItemData, OnInternalInsertedItem, SHELLTREEVIEWITEMDATA, ShellTreeViewItems::Add,
			///     namespacesOfInsertedItems
			CAtlList<LPSHELLTREEVIEWITEMDATA> shellItemDataOfInsertedItems;
			/// \brief <em>Holds the namespace pIDLs of items that are going to be inserted</em>
			///
			/// \sa BufferShellItemNamespace, OnInternalInsertedItem, ShellTreeViewNamespaces::Add,
			///     shellItemDataOfInsertedItems
			CAtlMap<PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE> namespacesOfInsertedItems;
			/// \brief <em>Holds details about currently running background item enumerations</em>
			///
			/// Holds details about currently running background item enumerations. These details include the
			/// time the task has been started and whether to check for timeouts.\n
			/// The enumeration task's unique ID is stored as key, the details are stored as value.
			///
			/// \sa EnqueueTask, BACKGROUNDENUMERATION
			CAtlMap<ULONGLONG, LPBACKGROUNDENUMERATION> backgroundEnums;
		#endif
		/// \brief <em>Holds the unique ID of the treeview item currently being inserted</em>
		///
		/// Holds the unique ID of the treeview item currently being inserted. We need this ID to be able to
		/// identify the item as shell item when handling the \c TVN_GETDISPINFO notification that is sent
		/// before we know the handle of the new item.
		///
		/// \sa OnInternalInsertingItem, OnGetDispInfoNotification
		LONG IDOfItemBeingInserted;

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

		Properties();
		~Properties();

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			autoEditNewItems = TRUE;
			colorCompressedItems = TRUE;
			colorEncryptedItems = TRUE;
			defaultManagedItemProperties = stmipAll;
			disabledEvents = static_cast<DisabledEventsConstants>(deNamespacePIDLChangeEvents | deNamespaceInsertionEvents | deNamespaceDeletionEvents | deItemPIDLChangeEvents | deItemInsertionEvents | deItemDeletionEvents);
			displayElevationShieldOverlays = TRUE;
			handleOLEDragDrop = static_cast<HandleOLEDragDropConstants>(hoddSourcePart | hoddTargetPart | hoddTargetPartWithDropHilite);
			hiddenItemsStyle = hisGhostedOnDemand;
			hImageList[0] = NULL;
			infoTipFlags = itfNoInfoTip;
			itemEnumerationTimeout = 3000;
			itemTypeSortOrder = itsoShellItemsFirst;
			limitLabelEditInput = TRUE;
			loadOverlaysOnDemand = TRUE;
			preselectBasenameOnLabelEdit = TRUE;
			processShellNotifications = TRUE;
			useGenericIcons = ugiOnlyForSlowItems;
			useSystemImageList = usilSmallImageList;

			pUnifiedImageList = NULL;
			timeOfLastItemCreatorInvocation = 0;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

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

		Flags()
		{
			detaching = FALSE;
			acceptNewTasks = TRUE;
		}
	} flags;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>Holds the handle of the treeview item currently being the drop target</em>
		///
		/// \sa OnInternalHandleDragEvent, pCurrentDropTarget
		HTREEITEM hCurrentDropTarget;
		/// \brief <em>Holds the \c IDropTarget implementation of the shell treeview item currently being the drop target</em>
		///
		/// \sa OnInternalHandleDragEvent, hCurrentDropTarget,
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

		DragDropStatus()
		{
			hCurrentDropTarget = NULL;
			pCurrentDropTarget = NULL;
			lastKeyState = 0;
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			hCurrentDropTarget = NULL;
			pCurrentDropTarget = NULL;
			lastKeyState = 0;
		}
	} dragDropStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to raise the \c ItemEnumerationTimedOut event</em>
		///
		/// \sa get_ItemEnumerationTimeout, Raise_ItemEnumerationTimedOut
		static const UINT_PTR ID_BACKGROUNDENUMTIMEOUT = 20;

		/// \brief <em>The interval of the timer that is used to raise the \c ItemEnumerationTimedOut event</em>
		///
		/// \sa get_ItemEnumerationTimeout, Raise_ItemEnumerationTimedOut
		static const UINT INT_BACKGROUNDENUMTIMEOUT = 100;
	} timers;

	/// \brief <em>Holds the imagelist that contains the icon that is drawn as the control's logo during design time</em>
	///
	/// \sa OnDraw, ShellTreeView
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
};     // ShellTreeView

OBJECT_ENTRY_AUTO(__uuidof(ShellTreeView), ShellTreeView)