#ifdef ACTIVATE_SECZONE_FEATURES
	//////////////////////////////////////////////////////////////////////
	/// \class Proxy_ISecurityZoneEvents
	/// \author Timo "TimoSoft" Kunze
	/// \brief <em>Fires events specified by the \c _ISecurityZoneEvents interface</em>
	///
	/// \if UNICODE
	///   \sa SecurityZone, ShBrowserCtlsLibU::_ISecurityZoneEvents
	/// \else
	///   \sa SecurityZone, ShBrowserCtlsLibA::_ISecurityZoneEvents
	/// \endif
	//////////////////////////////////////////////////////////////////////
#endif


#pragma once

#include "../DispIDs.h"


#ifdef ACTIVATE_SECZONE_FEATURES
	template<class TypeOfTrigger>
	class Proxy_ISecurityZoneEvents :
	    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ISecurityZoneEvents), CComDynamicUnkArray>
	{
	public:
	};     // Proxy_ISecurityZoneEvents
#endif