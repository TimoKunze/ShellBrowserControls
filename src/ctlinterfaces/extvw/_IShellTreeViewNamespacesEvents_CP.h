//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShTreeViewNamespacesEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShTreeViewNamespacesEvents interface</em>
///
/// \if UNICODE
///   \sa ShellTreeViewNamespaces, ShBrowserCtlsLibU::_IShTreeViewNamespacesEvents
/// \else
///   \sa ShellTreeViewNamespaces, ShBrowserCtlsLibA::_IShTreeViewNamespacesEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShTreeViewNamespacesEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShTreeViewNamespacesEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShTreeViewNamespacesEvents