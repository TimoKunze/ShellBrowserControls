//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewColumnEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewColumnEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewColumn, ShBrowserCtlsLibU::_IShListViewColumnEvents
/// \else
///   \sa ShellListViewColumn, ShBrowserCtlsLibA::_IShListViewColumnEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewColumnEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewColumnEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewColumnEvents