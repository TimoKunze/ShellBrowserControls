//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewNamespacesEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewNamespacesEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewNamespaces, ShBrowserCtlsLibU::_IShListViewNamespacesEvents
/// \else
///   \sa ShellListViewNamespaces, ShBrowserCtlsLibA::_IShListViewNamespacesEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewNamespacesEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewNamespacesEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewNamespacesEvents