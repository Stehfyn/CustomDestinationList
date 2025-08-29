#pragma once

#define NTDDI_VERSION NTDDI_WIN7  // Specifies that the minimum required platform is Windows 7.
#define WIN32_LEAN_AND_MEAN       // Exclude rarely-used stuff from Windows headers
#define STRICT_TYPED_ITEMIDS      // Utilize strictly typed IDLists
#define STRICT
#define COBJMACROS
#include <windows.h>

#if defined _WINDLL
  #define CDLAPI(type) __declspec(dllexport) EXTERN_C type APIENTRY
#else
  #define CDLAPI(type) __declspec(dllimport) EXTERN_C type APIENTRY
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
typedef struct Task ITask;

CDLAPI(HRESULT) ICDL_Initialize(ICDL** picdl);
CDLAPI(HRESULT) ICDL_CreateJumpList(ICDL* icdl, PCWSTR pcszAppId);
CDLAPI(HRESULT) ICDL_AddUserTasks(ICDL* icdl);