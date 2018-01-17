//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShTreeViewNamespaceEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShTreeViewNamespaceEvents interface</em>
///
/// \if UNICODE
///   \sa ShellTreeViewNamespace, ShBrowserCtlsLibU::_IShTreeViewNamespaceEvents
/// \else
///   \sa ShellTreeViewNamespace, ShBrowserCtlsLibA::_IShTreeViewNamespaceEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShTreeViewNamespaceEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShTreeViewNamespaceEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShTreeViewNamespaceEvents