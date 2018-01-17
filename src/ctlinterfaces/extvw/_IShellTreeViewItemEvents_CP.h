//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShTreeViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShTreeViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa ShellTreeViewItem, ShBrowserCtlsLibU::_IShTreeViewItemEvents
/// \else
///   \sa ShellTreeViewItem, ShBrowserCtlsLibA::_IShTreeViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShTreeViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShTreeViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShTreeViewItemEvents