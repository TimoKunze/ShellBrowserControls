#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	/// \class SecurityZones
	/// \author Timo "TimoSoft" Kunze
	/// \brief <em>Manages a collection of \c SecurityZone objects</em>
	///
	/// This class wraps the Internet Explorer security zones.
	///
	/// \if UNICODE
	///   \sa ShBrowserCtlsLibU::ISecurityZones, SecurityZone, ShellListView::SecurityZones,
	///       ShellTreeView::SecurityZones
	/// \else
	///   \sa ShBrowserCtlsLibA::ISecurityZones, SecurityZone, ShellListView::SecurityZones,
	///       ShellTreeView::SecurityZones
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
	#include "_ISecurityZonesEvents_CP.h"
	#include "../helpers.h"
	#include "SecurityZone.h"


	class ATL_NO_VTABLE SecurityZones : 
	    public CComObjectRootEx<CComSingleThreadModel>,
	    public CComCoClass<SecurityZones, &CLSID_SecurityZones>,
	    public ISupportErrorInfo,
	    public IConnectionPointContainerImpl<SecurityZones>,
	    public Proxy_ISecurityZonesEvents<SecurityZones>,
	    public IEnumVARIANT,
	    #ifdef UNICODE
	    	public IDispatchImpl<ISecurityZones, &IID_ISecurityZones, &LIBID_ShBrowserCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
	    #else
	    	public IDispatchImpl<ISecurityZones, &IID_ISecurityZones, &LIBID_ShBrowserCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
	    #endif
	{

	public:
		#ifndef DOXYGEN_SHOULD_SKIP_THIS
			DECLARE_REGISTRY_RESOURCEID(IDR_SECURITYZONES)

			BEGIN_COM_MAP(SecurityZones)
				COM_INTERFACE_ENTRY(ISecurityZones)
				COM_INTERFACE_ENTRY(IDispatch)
				COM_INTERFACE_ENTRY(ISupportErrorInfo)
				COM_INTERFACE_ENTRY(IConnectionPointContainer)
				COM_INTERFACE_ENTRY(IEnumVARIANT)
			END_COM_MAP()

			BEGIN_CONNECTION_POINT_MAP(SecurityZones)
				CONNECTION_POINT_ENTRY(__uuidof(_ISecurityZonesEvents))
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
		/// \brief <em>Clones the \c VARIANT iterator used to iterate the zones</em>
		///
		/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
		/// the \c SecurityZone objects managed by this collection object.
		///
		/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Next, Reset, Skip,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
		virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
		/// \brief <em>Retrieves the next x zones</em>
		///
		/// Retrieves the next \c numberOfMaxItems zones from the iterator.
		///
		/// \param[in] numberOfMaxItems The maximum number of zones the array identified by \c pItems can
		///            contain.
		/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
		///                the pointer to a zone's \c ISecurityZone implementation.
		/// \param[out] pNumberOfItemsReturned The number of zones that actually were copied to the array
		///             identified by \c pItems.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Clone, Reset, Skip, SecurityZone,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
		virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
		/// \brief <em>Resets the \c VARIANT iterator</em>
		///
		/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
		/// zone in the collection.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Clone, Next, Skip,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
		virtual HRESULT STDMETHODCALLTYPE Reset(void);
		/// \brief <em>Skips the next x zones</em>
		///
		/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip zones.
		///
		/// \param[in] numberOfItemsToSkip The number of zones to skip.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa Clone, Next, Reset,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
		virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
		//@}
		//////////////////////////////////////////////////////////////////////

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of ISecurityZones
		///
		//@{
		/// \brief <em>Retrieves a \c SecurityZone object from the collection</em>
		///
		/// Retrieves a \c SecurityZone object from the collection that wraps the Internet Explorer security zone
		/// identified by \c zoneIndex.
		///
		/// \param[in] zoneIndex A value that identifies the security zone to be retrieved.
		/// \param[out] ppZone Receives the zone's \c ISecurityZone implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This is the default property of the \c ISecurityZones interface.\n
		///          This property is read-only.
		///
		/// \sa SecurityZone, Count
		virtual HRESULT STDMETHODCALLTYPE get_Item(LONG zoneIndex, ISecurityZone** ppZone);
		/// \brief <em>Retrieves a \c VARIANT enumerator</em>
		///
		/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c SecurityZone objects
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

		/// \brief <em>Counts the Internet Explorer security zones in the collection</em>
		///
		/// Retrieves the number of \c SecurityZone objects in the collection.
		///
		/// \param[out] pValue The number of elements in the collection.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa get_Item
		virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
		//@}
		//////////////////////////////////////////////////////////////////////

	protected:
		/// \brief <em>Holds the object's properties' settings</em>
		struct Properties
		{
			/// \brief <em>The \c IInternetZoneManager object used to enumerate the zones</em>
			CComPtr<IInternetZoneManager> pZoneManager;
			/// \brief <em>Holds the next enumerated security zone</em>
			DWORD nextEnumeratedZone;
			/// \brief <em>Holds the unique ID of the security zone enumerator</em>
			DWORD zoneEnumerator;
			/// \brief <em>Holds the number of security zones</em>
			DWORD zoneCount;

			Properties()
			{
				ATLVERIFY(SUCCEEDED(CoCreateInstance(CLSID_InternetZoneManager, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pZoneManager))));
				ATLASSUME(pZoneManager);
				ATLVERIFY(SUCCEEDED(pZoneManager->CreateZoneEnumerator(&zoneEnumerator, &zoneCount, 0)));

				nextEnumeratedZone = 0;
			}

			~Properties()
			{
				ATLVERIFY(SUCCEEDED(pZoneManager->DestroyZoneEnumerator(zoneEnumerator)));
			}

			/// \brief <em>Copies this struct's content to another \c Properties struct</em>
			void CopyTo(Properties* pTarget)
			{
				ATLASSERT_POINTER(pTarget, Properties);
				if(pTarget) {
					pTarget->nextEnumeratedZone = this->nextEnumeratedZone;
				}
			}
		} properties;
	};     // SecurityZones

	OBJECT_ENTRY_AUTO(__uuidof(SecurityZones), SecurityZones)
#endif