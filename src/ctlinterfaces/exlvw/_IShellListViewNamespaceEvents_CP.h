//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewNamespaceEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewNamespaceEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewNamespace, ShBrowserCtlsLibU::_IShListViewNamespaceEvents
/// \else
///   \sa ShellListViewNamespace, ShBrowserCtlsLibA::_IShListViewNamespaceEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewNamespaceEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewNamespaceEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewNamespaceEvents