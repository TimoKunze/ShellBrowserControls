//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewItems, ShBrowserCtlsLibU::_IShListViewItemsEvents
/// \else
///   \sa ShellListViewItems, ShBrowserCtlsLibA::_IShListViewItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewItemsEvents