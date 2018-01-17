#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	/// \class SecurityZone
	/// \author Timo "TimoSoft" Kunze
	/// \brief <em>Wraps an Internet Explorer security zone</em>
	///
	/// This class is a wrapper around an Internet Explorer security zone.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::ISecurityZone, SecurityZone
	/// \else
	///   \sa ShBrowserCtlsLibA::ISecurityZone, SecurityZone
	/// \endif
	//////////////////////////////////////////////////////////////////////
#endif


#pragma once

#include "../res/resource.h"
#ifdef UNICODE
	#include "../ShBrowserCtlsU.h"
#else
	#include "../ShBrowserCtlsA.h"
#endif
#ifdef ACTIVATE_SECZONE_FEATURES
	#include "definitions.h"
	#include "_ISecurityZoneEvents_CP.h"
	#include "../helpers.h"


	class ATL_NO_VTABLE SecurityZone : 
	    public CComObjectRootEx<CComSingleThreadModel>,
	    public CComCoClass<SecurityZone, &CLSID_SecurityZone>,
	    public ISupportErrorInfo,
	    public IConnectionPointContainerImpl<SecurityZone>,
	    public Proxy_ISecurityZoneEvents<SecurityZone>,
	    #ifdef UNICODE
	    	public IDispatchImpl<ISecurityZone, &IID_ISecurityZone, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
	    #else
	    	public IDispatchImpl<ISecurityZone, &IID_ISecurityZone, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
	    #endif
	{
		friend class ClassFactory;

	public:
		#ifndef DOXYGEN_SHOULD_SKIP_THIS
			DECLARE_REGISTRY_RESOURCEID(IDR_SECURITYZONE)

			BEGIN_COM_MAP(SecurityZone)
				COM_INTERFACE_ENTRY(ISecurityZone)
				COM_INTERFACE_ENTRY(IDispatch)
				COM_INTERFACE_ENTRY(ISupportErrorInfo)
				COM_INTERFACE_ENTRY(IConnectionPointContainer)
			END_COM_MAP()

			BEGIN_CONNECTION_POINT_MAP(SecurityZone)
				CONNECTION_POINT_ENTRY(__uuidof(_ISecurityZoneEvents))
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
		/// \name Implementation of ISecurityZone
		///
		//@{
		/// \brief <em>Retrieves the current setting of the \c Description property</em>
		///
		/// Retrieves the security zone's description as provided by the shell.
		///
		/// \param[out] pValue The property's current setting.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		///
		/// \sa get_DisplayName
		virtual HRESULT STDMETHODCALLTYPE get_Description(BSTR* pValue);
		/// \brief <em>Retrieves the current setting of the \c DisplayName property</em>
		///
		/// Retrieves the security zone's display name as provided by the shell.
		///
		/// \param[out] pValue The property's current setting.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This is the default property of the \c ISecurityZone interface.\n
		///          This property is read-only.
		///
		/// \sa get_Description
		virtual HRESULT STDMETHODCALLTYPE get_DisplayName(BSTR* pValue);
		/// \brief <em>Retrieves the current setting of the \c hIcon property</em>
		///
		/// Retrieves the security zone's icon as provided by the shell.
		///
		/// \param[out] pValue The property's current setting.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks The calling application must destroy the icon if it doesn't need it anymore.\n
		///          This property is read-only.
		///
		/// \sa get_IconPath
		virtual HRESULT STDMETHODCALLTYPE get_hIcon(OLE_HANDLE* pValue);
		/// \brief <em>Retrieves the current setting of the \c IconPath property</em>
		///
		/// Retrieves the path to the security zone's icon as provided by the shell.
		///
		/// \param[out] pValue The property's current setting.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		///
		/// \sa get_hIcon
		virtual HRESULT STDMETHODCALLTYPE get_IconPath(BSTR* pValue);
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Attaches this object to a given Internet Explorer security zone</em>
		///
		/// Attaches this object to a given Internet Explorer security zone, so that the zone's properties can be
		/// retrieved using this object's methods.
		///
		/// \param[in] zoneIndex The zero-based index of the security zone to attach to.
		/// \param[in] pZoneAttributes The security zone's attributes.
		///
		/// \sa Detach
		void Attach(LONG zoneIndex, __in LPZONEATTRIBUTES pZoneAttributes);
		/// \brief <em>Detaches this object from an Internet Explorer security zone</em>
		///
		/// Detaches this object from the Internet Explorer security zone it currently wraps, so that it doesn't
		/// wrap any security zone anymore.
		///
		/// \sa Attach
		void Detach(void);

	protected:
		/// \brief <em>Holds the object's properties' settings</em>
		struct Properties
		{
			/// \brief <em>The zero-based index of the security zone wrapped by this object</em>
			LONG zoneIndex;
			/// \brief <em>The wrapped security zone's attributes</em>
			ZONEATTRIBUTES zoneAttributes;

			Properties()
			{
				zoneIndex = -1;
				ZeroMemory(&zoneAttributes, sizeof(zoneAttributes));
			}
		} properties;
	};     // SecurityZone

	OBJECT_ENTRY_AUTO(__uuidof(SecurityZone), SecurityZone)
#endif