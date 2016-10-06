// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if _MSC_VER > 1000
#   pragma once
#endif // _MSC_VER > 1000

#ifndef _CRT_SECURE_NO_WARNINGS
#   define _CRT_SECURE_NO_WARNINGS
#endif

#define _WIN32_WINNT 0x0501

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
#endif

#define _ATL_APARTMENT_THREADED

#include <windows.h>
#include <streams.h>
#include <strsafe.h>

#include "DShowUtil.h"
#include "smartptr.h"   // smart pointer class
