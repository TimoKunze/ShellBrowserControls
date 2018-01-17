//////////////////////////////////////////////////////////////////////
/// \class Proxy_IShListViewColumnsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IShListViewColumnsEvents interface</em>
///
/// \if UNICODE
///   \sa ShellListViewColumns, ShBrowserCtlsLibU::_IShListViewColumnsEvents
/// \else
///   \sa ShellListViewColumns, ShBrowserCtlsLibA::_IShListViewColumnsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "../../DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IShListViewColumnsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IShListViewColumnsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IShListViewColumnsEvents