//////////////////////////////////////////////////////////////////////
/// \namespace RunTimeHelperEx
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Extends WTL's \c RunTimeHelper namespace</em>
//////////////////////////////////////////////////////////////////////


#pragma once


namespace RunTimeHelperEx
{
	inline bool IsWin8()
	{
#ifdef _versionhelpers_H_INCLUDED_
		return ::IsWindows8OrGreater();
#else // !_versionhelpers_H_INCLUDED_
		OSVERSIONINFO ovi = { sizeof(OSVERSIONINFO) };
		BOOL bRet = ::GetVersionEx(&ovi);
		return ((bRet != FALSE) && (ovi.dwMajorVersion == 6) && (ovi.dwMinorVersion >= 2));
#endif // _versionhelpers_H_INCLUDED_
	}

	inline bool IsWin81()
	{
#ifdef _versionhelpers_H_INCLUDED_
		return ::IsWindows8Point1OrGreater();
#else // !_versionhelpers_H_INCLUDED_
		OSVERSIONINFO ovi = { sizeof(OSVERSIONINFO) };
		BOOL bRet = ::GetVersionEx(&ovi);
		return ((bRet != FALSE) && (ovi.dwMajorVersion == 6) && (ovi.dwMinorVersion >= 3));
#endif // _versionhelpers_H_INCLUDED_
	}

	inline bool IsWin10()
	{
#ifdef _versionhelpers_H_INCLUDED_
		return ::IsWindows10OrGreater();
#else // !_versionhelpers_H_INCLUDED_
		OSVERSIONINFO ovi = { sizeof(OSVERSIONINFO) };
		BOOL bRet = ::GetVersionEx(&ovi);
		return ((bRet != FALSE) && (ovi.dwMajorVersion == 10));
#endif // _versionhelpers_H_INCLUDED_
	}
}
