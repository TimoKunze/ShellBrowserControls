#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	/// \class Proxy_ISecurityZonesEvents
	/// \author Timo "TimoSoft" Kunze
	/// \brief <em>Fires events specified by the \c _ISecurityZonesEvents interface</em>
	///
	/// \if UNICODE
	///   \sa SecurityZones, ShBrowserCtlsLibU::_ISecurityZonesEvents
	/// \else
	///   \sa SecurityZones, ShBrowserCtlsLibA::_ISecurityZonesEvents
	/// \endif
	//////////////////////////////////////////////////////////////////////
#endif


#pragma once

#include "../DispIDs.h"


#ifdef ACTIVATE_SECZONE_FEATURES
	template<class TypeOfTrigger>
	class Proxy_ISecurityZonesEvents :
	    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ISecurityZonesEvents), CComDynamicUnkArray>
	{
	public:
	};     // Proxy_ISecurityZonesEvents
#endif