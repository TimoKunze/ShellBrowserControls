//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewItem, ShBrowserCtlsLibU::_IShListViewItemEvents
/// \else
///   \sa ShellListViewItem, ShBrowserCtlsLibA::_IShListViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewItemEvents