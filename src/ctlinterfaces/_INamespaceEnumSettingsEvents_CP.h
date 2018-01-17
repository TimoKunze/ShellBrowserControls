//////////////////////////////////////////////////////////////////////
/// \class Proxy_INamespaceEnumSettingsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _INamespaceEnumSettingsEvents interface</em>
///
/// \if UNICODE
///   \sa NamespaceEnumSettings, ShBrowserCtlsLibU::_INamespaceEnumSettingsEvents
/// \else
///   \sa NamespaceEnumSettings, ShBrowserCtlsLibA::_INamespaceEnumSettingsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_INamespaceEnumSettingsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_INamespaceEnumSettingsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_INamespaceEnumSettingsEvents