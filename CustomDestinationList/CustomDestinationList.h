#pragma once

#define WIN32_LEAN_AND_MEAN       // Exclude rarely-used stuff from Windows headers
#define STRICT_TYPED_ITEMIDS      // Utilize strictly typed IDLists
#define COBJMACROS
#define CINTERFACE
#include <windows.h>
#include <tchar.h>

#if defined CDLSTATICLIB
  #if defined CDLINLINE
    #define CDLAPI(type) __forceinline EXTERN_C type __cdecl
  #else
    #define CDLAPI(type) EXTERN_C type __cdecl
  #endif
#else
  #if defined _WINDLL
    #define CDLAPI(type) __declspec(dllexport) EXTERN_C type __cdecl
  #else
    #define CDLAPI(type) __declspec(dllimport) EXTERN_C type __cdecl
  #endif
#endif

#if defined _M_IX86
  #pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
  #pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
  #pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
  #pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

typedef struct CDL ICDL;

// C:\%UserProfile%\AppData\Roaming\Microsoft\Windows\Recent\CustomDestinations

CDLAPI(HRESULT) ICDL_BeginList(LPCWSTR pcwszAppId, ICDL** ppThis);
CDLAPI(HRESULT) ICDL_AddTask(ICDL* pThis, LPCTSTR pcszImage, LPCTSTR pcszArgs, LPCTSTR pcszDescription, LPCTSTR pcszTitle, int nIconIndex);
CDLAPI(HRESULT) ICDL_AddSeparator(ICDL* pThis);
CDLAPI(HRESULT) ICDL_CommitList(ICDL* pThis);
CDLAPI(VOID)    ICDL_Release(ICDL* pThis);
