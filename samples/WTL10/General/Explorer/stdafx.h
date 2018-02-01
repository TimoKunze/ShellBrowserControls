// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#pragma once

#define NTDDI_VERSION NTDDI_WIN10

#include <winsdkver.h>

#ifdef _DEBUG
	#define _ATL_DEBUG_INTERFACES
	#include <windows.h>
	#include <atldbgmem.h>
#endif

#include <atlbase.h>

#define _ATL_MAX_VARTYPES 15
#include <atlcom.h>
#include <atlhost.h>
#include <atlwin.h>
#include <atlctl.h>

#include <atlstr.h>

#include <atlapp.h>
extern CAppModule _Module;

#include <atlframe.h>
#include <atlctrls.h>
#include <atlctrlx.h>
#include <atlsplit.h>
#include <atldlgs.h>
#include <atlmisc.h>
#include <atlsafe.h>

#include <atltheme.h>

#if defined _M_IX86
	#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
	#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
	#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
	#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif


#include <initguid.h>

#import <libid:1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5> version("2.0") raw_dispinterfaces
#import <libid:1F9B9092-BEE4-4caf-9C7B-9384AF087C63> version("1.5") raw_dispinterfaces
#import <libid:9FC6639B-4237-4fb5-93B8-24049D39DF74> version("1.0") raw_dispinterfaces

DEFINE_GUID(LIBID_ExTVwLibU, 0x1f8f0fe7, 0x2cfb, 0x4466, 0xa2, 0xbc, 0xab, 0xb4, 0x41, 0xad, 0xed, 0xd5);
DEFINE_GUID(LIBID_ShBrowserCtlsLibU, 0x1F9B9092, 0xBEE4, 0x4caf, 0x9c, 0x7b, 0x93, 0x84, 0xaf, 0x08, 0x7c, 0x63);
DEFINE_GUID(LIBID_ExLVwLibU, 0x9fc6639b, 0x4237, 0x4fb5, 0x93, 0xb8, 0x24, 0x04, 0x9d, 0x39, 0xdf, 0x74);
