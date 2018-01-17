//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShTreeViewItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShTreeViewItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ShellTreeViewItems, ShBrowserCtlsLibU::_IShTreeViewItemsEvents
/// \else
///   \sa ShellTreeViewItems, ShBrowserCtlsLibA::_IShTreeViewItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShTreeViewItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShTreeViewItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShTreeViewItemsEvents