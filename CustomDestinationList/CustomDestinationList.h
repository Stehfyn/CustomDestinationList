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

/**************************************************************************************************************************\
* Excerpt from: https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-icustomdestinationlist  *
*   >   A custom Jump List is never truly updated in the sense of changing elements in an existing list. Rather, the old   *
*       list is replaced with a new list.                                                                                  *
*                                                                                                                          *
*       The basic sequence of ICustomDestinationList method calls to build and display a custom Jump List is as follows:   *
*                                                                                                                          *
*         1). SetAppID (required only if an application provides its own AppUserModelID)                                   *
*         2). BeginList                                                                                                    *
*         3). AppendCategory, AppendKnownCategory, AddUserTasks, or any combination of those three methods.                *
*         4). CommitList                                                                                                   *
\**************************************************************************************************************************/

typedef struct CDL ICDL;

CDLAPI(HRESULT) ICDL_SetWindowAppUserModelID(HWND hwnd, LPCWSTR pcwszAppId);
CDLAPI(HRESULT) ICDL_ClearWindowAppUserModelID(HWND hwnd);
CDLAPI(HRESULT) ICDL_BeginList(LPCWSTR pcwszAppId, ICDL** ppThis);
CDLAPI(HRESULT) ICDL_BeginCategoryA(ICDL* pThis, LPCSTR pcszCategory);
CDLAPI(HRESULT) ICDL_BeginCategoryW(ICDL* pThis, LPCWSTR pcwszCategory);
CDLAPI(HRESULT) ICDL_AddTaskA(ICDL* pThis, LPCSTR pcszImage, LPCSTR pcszArgs, LPCSTR pcszDescription, LPCSTR pcszTitle, LPCSTR pcszIconLocation, int nIconIndex);
CDLAPI(HRESULT) ICDL_AddTaskW(ICDL* pThis, LPCWSTR pcwszImage, LPCWSTR pcwszArgs, LPCWSTR pcwszDescription, LPCWSTR pcwszTitle, LPCWSTR pcwszIconLocation, int nIconIndex);
CDLAPI(HRESULT) ICDL_AddSeparator(ICDL* pThis);
CDLAPI(HRESULT) ICDL_CommitCategory(ICDL* pThis);
CDLAPI(HRESULT) ICDL_CommitList(ICDL* pThis); // C:\%UserProfile%\AppData\Roaming\Microsoft\Windows\Recent\CustomDestinations
CDLAPI(void)    ICDL_Release(ICDL* pThis);

#ifdef _UNICODE
  #define ICDL_BeginCategory ICDL_BeginCategoryW
  #define ICDL_AddTask ICDL_AddTaskW
#elif defined _MBCS
  #define ICDL_BeginCategory ICDL_BeginCategoryA
  #define ICDL_AddTask ICDL_AddTaskA
#endif