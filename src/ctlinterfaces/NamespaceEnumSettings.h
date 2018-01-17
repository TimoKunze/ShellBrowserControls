//////////////////////////////////////////////////////////////////////
/// \class NamespaceEnumSettings
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing shell listview item</em>
///
/// This class is a container class used to store namespace enumeration settings.
///
/// \if UNICODE
///   \sa ShBrowserCtlsLibU::INamespaceEnumSettings, NamespaceEnumSettings,
///       ShellTreeViewNamespace::get_NamespaceEnumSettings, ShellTreeViewItem::get_NamespaceEnumSettings
/// \else
///   \sa ShBrowserCtlsLibA::INamespaceEnumSettings, NamespaceEnumSettings,
///       ShellTreeViewNamespace::get_NamespaceEnumSettings, ShellTreeViewItem::get_NamespaceEnumSettings
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../res/resource.h"
#ifdef UNICODE
	#include "../ShBrowserCtlsU.h"
#else
	#include "../ShBrowserCtlsA.h"
#endif
#include "IPersistentChildObject.h"
#include "_INamespaceEnumSettingsEvents_CP.h"
#include "../helpers.h"


class ATL_NO_VTABLE NamespaceEnumSettings : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<NamespaceEnumSettings, &CLSID_NamespaceEnumSettings>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<NamespaceEnumSettings>,
    public Proxy_INamespaceEnumSettingsEvents<NamespaceEnumSettings>,
    #ifdef UNICODE
    	public IDispatchImpl<INamespaceEnumSettings, &IID_INamespaceEnumSettings, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<INamespaceEnumSettings, &IID_INamespaceEnumSettings, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<NamespaceEnumSettings>,
    public IPersistStorageImpl<NamespaceEnumSettings>,
    public IPersistPropertyBagImpl<NamespaceEnumSettings>,
    public IPropertyNotifySinkCP<NamespaceEnumSettings>,
    public IPersistentChildObject
{
	friend class ClassFactory;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	NamespaceEnumSettings();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		// required for persistence
		BOOL m_bRequiresSave;

		DECLARE_REGISTRY_RESOURCEID(IDR_NAMESPACEENUMSETTINGS)

		BEGIN_COM_MAP(NamespaceEnumSettings)
			COM_INTERFACE_ENTRY(INamespaceEnumSettings)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY_IID(IID_IPersistentChildObject, IPersistentChildObject)
		END_COM_MAP()

		BEGIN_PROP_MAP(NamespaceEnumSettings)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_ENTRY_TYPE("EnumerationFlags", DISPID_NSES_ENUMERATIONFLAGS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ExcludedFileItemFileAttributes", DISPID_NSES_EXCLUDEDFILEITEMFILEATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ExcludedFileItemShellAttributes", DISPID_NSES_EXCLUDEDFILEITEMSHELLATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ExcludedFolderItemFileAttributes", DISPID_NSES_EXCLUDEDFOLDERITEMFILEATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ExcludedFolderItemShellAttributes", DISPID_NSES_EXCLUDEDFOLDERITEMSHELLATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IncludedFileItemFileAttributes", DISPID_NSES_INCLUDEDFILEITEMFILEATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IncludedFileItemShellAttributes", DISPID_NSES_INCLUDEDFILEITEMSHELLATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IncludedFolderItemFileAttributes", DISPID_NSES_INCLUDEDFOLDERITEMFILEATTRIBUTES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IncludedFolderItemShellAttributes", DISPID_NSES_INCLUDEDFOLDERITEMSHELLATTRIBUTES, CLSID_NULL, VT_I4)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(NamespaceEnumSettings)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_INamespaceEnumSettingsEvents))
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
	/// \name Implementation of persistance
	///
	//@{
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
	/// \name Implementation of IPersistentChildObject
	///
	//@{
	/// \brief <em>Sets the owner control</em>
	///
	/// \param[in] pOwner The owner control to set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IPersistentChildObject::SetOwnerControl
	virtual HRESULT STDMETHODCALLTYPE SetOwnerControl(CComControlBase* pOwner);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of INamespaceEnumSettings
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c EnumerationFlags property</em>
	///
	/// Retrieves a bit field indicating indicating which kinds of items shall be displayed and how the items
	/// shall be enumerated. Any combination of the values defined by the \c ShNamespaceEnumerationConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_EnumerationFlags, ShBrowserCtlsLibU::ShNamespaceEnumerationConstants
	/// \else
	///   \sa put_EnumerationFlags, ShBrowserCtlsLibA::ShNamespaceEnumerationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_EnumerationFlags(ShNamespaceEnumerationConstants* pValue);
	/// \brief <em>Sets the \c EnumerationFlags property</em>
	///
	/// Sets a bit field indicating indicating which kinds of items shall be displayed and how the items
	/// shall be enumerated. Any combination of the values defined by the \c ShNamespaceEnumerationConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_EnumerationFlags, ShBrowserCtlsLibU::ShNamespaceEnumerationConstants
	/// \else
	///   \sa get_EnumerationFlags, ShBrowserCtlsLibA::ShNamespaceEnumerationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_EnumerationFlags(ShNamespaceEnumerationConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ExcludedFileItemFileAttributes property</em>
	///
	/// Retrieves a bit field specifying the file attributes that mustn't be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa put_ExcludedFileItemFileAttributes, get_ExcludedFolderItemFileAttributes,
	///       get_IncludedFileItemFileAttributes, get_ExcludedFileItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa put_ExcludedFileItemFileAttributes, get_ExcludedFolderItemFileAttributes,
	///       get_IncludedFileItemFileAttributes, get_ExcludedFileItemShellAttributes
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExcludedFileItemFileAttributes(ItemFileAttributesConstants* pValue);
	/// \brief <em>Sets the \c ExcludedFileItemFileAttributes property</em>
	///
	/// Sets a bit field specifying the file attributes that mustn't be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa get_ExcludedFileItemFileAttributes, put_ExcludedFolderItemFileAttributes,
	///       put_IncludedFileItemFileAttributes, put_ExcludedFileItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa get_ExcludedFileItemFileAttributes, put_ExcludedFolderItemFileAttributes,
	///       put_IncludedFileItemFileAttributes, put_ExcludedFileItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExcludedFileItemFileAttributes(ItemFileAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ExcludedFileItemShellAttributes property</em>
	///
	/// Retrieves a bit field specifying the shell attributes that mustn't be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ExcludedFileItemShellAttributes, get_ExcludedFolderItemShellAttributes,
	///       get_IncludedFileItemShellAttributes, get_ExcludedFileItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa put_ExcludedFileItemShellAttributes, get_ExcludedFolderItemShellAttributes,
	///       get_IncludedFileItemShellAttributes, get_ExcludedFileItemFileAttributes
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExcludedFileItemShellAttributes(ItemShellAttributesConstants* pValue);
	/// \brief <em>Sets the \c ExcludedFileItemShellAttributes property</em>
	///
	/// Sets a bit field specifying the shell attributes that mustn't be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ExcludedFileItemShellAttributes, put_ExcludedFolderItemShellAttributes,
	///       put_IncludedFileItemShellAttributes, put_ExcludedFileItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa get_ExcludedFileItemShellAttributes, put_ExcludedFolderItemShellAttributes,
	///       put_IncludedFileItemShellAttributes, put_ExcludedFileItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExcludedFileItemShellAttributes(ItemShellAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ExcludedFileItemFileAttributes property</em>
	///
	/// Retrieves a bit field specifying the file attributes that mustn't be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa put_ExcludedFolderItemFileAttributes, get_ExcludedFileItemFileAttributes,
	///       get_IncludedFolderItemFileAttributes, get_ExcludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa put_ExcludedFolderItemFileAttributes, get_ExcludedFileItemFileAttributes,
	///       get_IncludedFolderItemFileAttributes, get_ExcludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExcludedFolderItemFileAttributes(ItemFileAttributesConstants* pValue);
	/// \brief <em>Sets the \c ExcludedFileItemFileAttributes property</em>
	///
	/// Sets a bit field specifying the file attributes that mustn't be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa get_ExcludedFolderItemFileAttributes, put_ExcludedFileItemFileAttributes,
	///       put_IncludedFolderItemFileAttributes, put_ExcludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa get_ExcludedFolderItemFileAttributes, put_ExcludedFileItemFileAttributes,
	///       put_IncludedFolderItemFileAttributes, put_ExcludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExcludedFolderItemFileAttributes(ItemFileAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ExcludedFolderItemShellAttributes property</em>
	///
	/// Retrieves a bit field specifying the shell attributes that mustn't be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ExcludedFolderItemShellAttributes, get_ExcludedFileItemShellAttributes,
	///       get_IncludedFolderItemShellAttributes, get_ExcludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa put_ExcludedFolderItemShellAttributes, get_ExcludedFileItemShellAttributes,
	///       get_IncludedFolderItemShellAttributes, get_ExcludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ExcludedFolderItemShellAttributes(ItemShellAttributesConstants* pValue);
	/// \brief <em>Sets the \c ExcludedFileItemShellAttributes property</em>
	///
	/// Sets a bit field specifying the shell attributes that mustn't be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item has any of the
	/// specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ExcludedFolderItemShellAttributes, put_ExcludedFileItemShellAttributes,
	///       put_IncludedFolderItemShellAttributes, put_ExcludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa get_ExcludedFolderItemShellAttributes, put_ExcludedFileItemShellAttributes,
	///       put_IncludedFolderItemShellAttributes, put_ExcludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ExcludedFolderItemShellAttributes(ItemShellAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IncludedFileItemFileAttributes property</em>
	///
	/// Retrieves a bit field specifying the file attributes that must be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all
	/// of the specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa put_IncludedFileItemFileAttributes, get_IncludedFolderItemFileAttributes,
	///       get_ExcludedFileItemFileAttributes, get_IncludedFileItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa put_IncludedFileItemFileAttributes, get_IncludedFolderItemFileAttributes,
	///       get_ExcludedFileItemFileAttributes, get_IncludedFileItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IncludedFileItemFileAttributes(ItemFileAttributesConstants* pValue);
	/// \brief <em>Sets the \c IncludedFileItemFileAttributes property</em>
	///
	/// Sets a bit field specifying the file attributes that must be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all of
	/// the specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa get_IncludedFileItemFileAttributes, put_IncludedFolderItemFileAttributes,
	///       put_ExcludedFileItemFileAttributes, put_IncludedFileItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa get_IncludedFileItemFileAttributes, put_IncludedFolderItemFileAttributes,
	///       put_ExcludedFileItemFileAttributes, put_IncludedFileItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IncludedFileItemFileAttributes(ItemFileAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IncludedFileItemShellAttributes property</em>
	///
	/// Retrieves a bit field specifying the shell attributes that must be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all
	/// of the specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IncludedFileItemShellAttributes, get_IncludedFolderItemShellAttributes,
	///       get_ExcludedFileItemShellAttributes, get_IncludedFileItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa put_IncludedFileItemShellAttributes, get_IncludedFolderItemShellAttributes,
	///       get_ExcludedFileItemShellAttributes, get_IncludedFileItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IncludedFileItemShellAttributes(ItemShellAttributesConstants* pValue);
	/// \brief <em>Sets the \c IncludedFileItemShellAttributes property</em>
	///
	/// Sets a bit field specifying the shell attributes that must be set for any file item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all of
	/// the specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IncludedFileItemShellAttributes, put_IncludedFolderItemShellAttributes,
	///       put_ExcludedFileItemShellAttributes, put_IncludedFileItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa get_IncludedFileItemShellAttributes, put_IncludedFolderItemShellAttributes,
	///       put_ExcludedFileItemShellAttributes, put_IncludedFileItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IncludedFileItemShellAttributes(ItemShellAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IncludedFolderItemFileAttributes property</em>
	///
	/// Retrieves a bit field specifying the file attributes that must be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all
	/// of the specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa put_IncludedFolderItemFileAttributes, get_IncludedFileItemFileAttributes,
	///       get_ExcludedFolderItemFileAttributes, get_IncludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa put_IncludedFolderItemFileAttributes, get_IncludedFileItemFileAttributes,
	///       get_ExcludedFolderItemFileAttributes, get_IncludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IncludedFolderItemFileAttributes(ItemFileAttributesConstants* pValue);
	/// \brief <em>Sets the \c IncludedFolderItemFileAttributes property</em>
	///
	/// Sets a bit field specifying the file attributes that must be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all of
	/// the specified file attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemFileAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property affects filesystem items only.
	///
	/// \if UNICODE
	///   \sa get_IncludedFolderItemFileAttributes, put_IncludedFileItemFileAttributes,
	///       put_ExcludedFolderItemFileAttributes, put_IncludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibU::ItemFileAttributesConstants
	/// \else
	///   \sa get_IncludedFolderItemFileAttributes, put_IncludedFileItemFileAttributes,
	///       put_ExcludedFolderItemFileAttributes, put_IncludedFolderItemShellAttributes,
	///       ShBrowserCtlsLibA::ItemFileAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IncludedFolderItemFileAttributes(ItemFileAttributesConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IncludedFolderItemShellAttributes property</em>
	///
	/// Retrieves a bit field specifying the shell attributes that must be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all
	/// of the specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IncludedFolderItemShellAttributes, get_IncludedFileItemShellAttributes,
	///       get_ExcludedFolderItemShellAttributes, get_IncludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa put_IncludedFolderItemShellAttributes, get_IncludedFileItemShellAttributes,
	///       get_ExcludedFolderItemShellAttributes, get_IncludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IncludedFolderItemShellAttributes(ItemShellAttributesConstants* pValue);
	/// \brief <em>Sets the \c IncludedFolderItemShellAttributes property</em>
	///
	/// Sets a bit field specifying the shell attributes that must be set for any folder item of the
	/// enumeration in order to have it inserted into the control. If an enumerated item doesn't have all of
	/// the specified shell attributes set, this item isn't inserted into the control.\n
	/// Any combination of the values defined by the \c ItemShellAttributesConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IncludedFolderItemShellAttributes, put_IncludedFileItemShellAttributes,
	///       put_ExcludedFolderItemShellAttributes, put_IncludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibU::ItemShellAttributesConstants
	/// \else
	///   \sa get_IncludedFolderItemShellAttributes, put_IncludedFileItemShellAttributes,
	///       put_ExcludedFolderItemShellAttributes, put_IncludedFolderItemFileAttributes,
	///       ShBrowserCtlsLibA::ItemShellAttributesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IncludedFolderItemShellAttributes(ItemShellAttributesConstants newValue);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c CComControlBase object that owns this object</em>
		///
		/// \sa SetOwnerControl, pOwnerUnknown
		CComControlBase* pOwner;
		/// \brief <em>Holds a reference to the owner of this object</em>
		///
		/// \sa SetOwnerControl, pOwner
		LPUNKNOWN pOwnerUnknown;
		/// \brief <em>Holds the \c EnumerationFlags property's setting</em>
		///
		/// \sa get_EnumerationFlags, put_EnumerationFlags
		ShNamespaceEnumerationConstants enumerationFlags;
		/// \brief <em>Holds the \c ExcludedFileItemFileAttributes property's setting</em>
		///
		/// \sa get_ExcludedFileItemFileAttributes, put_ExcludedFileItemFileAttributes
		ItemFileAttributesConstants excludedFileItemFileAttributes;
		/// \brief <em>Holds the \c ExcludedFileItemShellAttributes property's setting</em>
		///
		/// \sa get_ExcludedFileItemShellAttributes, put_ExcludedFileItemShellAttributes
		ItemShellAttributesConstants excludedFileItemShellAttributes;
		/// \brief <em>Holds the \c ExcludedFolderItemFileAttributes property's setting</em>
		///
		/// \sa get_ExcludedFolderItemFileAttributes, put_ExcludedFolderItemFileAttributes
		ItemFileAttributesConstants excludedFolderItemFileAttributes;
		/// \brief <em>Holds the \c ExcludedFolderItemShellAttributes property's setting</em>
		///
		/// \sa get_ExcludedFolderItemShellAttributes, put_ExcludedFolderItemShellAttributes
		ItemShellAttributesConstants excludedFolderItemShellAttributes;
		/// \brief <em>Holds the \c IncludedFileItemFileAttributes property's setting</em>
		///
		/// \sa get_IncludedFileItemFileAttributes, put_IncludedFileItemFileAttributes
		ItemFileAttributesConstants includedFileItemFileAttributes;
		/// \brief <em>Holds the \c IncludedFileItemShellAttributes property's setting</em>
		///
		/// \sa get_IncludedFileItemShellAttributes, put_IncludedFileItemShellAttributes
		ItemShellAttributesConstants includedFileItemShellAttributes;
		/// \brief <em>Holds the \c IncludedFolderItemFileAttributes property's setting</em>
		///
		/// \sa get_IncludedFolderItemFileAttributes, put_IncludedFolderItemFileAttributes
		ItemFileAttributesConstants includedFolderItemFileAttributes;
		/// \brief <em>Holds the \c IncludedFolderItemShellAttributes property's setting</em>
		///
		/// \sa get_IncludedFolderItemShellAttributes, put_IncludedFolderItemShellAttributes
		ItemShellAttributesConstants includedFolderItemShellAttributes;

		Properties()
		{
			pOwnerUnknown = NULL;
			ResetToDefaults();
		}

		~Properties();

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void);
	} properties;
};     // NamespaceEnumSettings

OBJECT_ENTRY_AUTO(__uuidof(NamespaceEnumSettings), NamespaceEnumSettings)