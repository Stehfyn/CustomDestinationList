#include "CustomDestinationList.h"

#include <unknwnbase.h>
#include <oleauto.h>
#include <objectarray.h>
#include <propkey.h>
#include <propsys.h>
#include <propvarutil.h>
#include <shlobj.h>

#pragma comment (lib, "Propsys")

#define __WARNING_PADDING_ADDED_AFTER_DATA_MEMBER    4820
#define __WARNING_SPECTRE_MITIGATION_FOR_MEMORY_LOAD 5045
#pragma warning(disable:__WARNING_PADDING_ADDED_AFTER_DATA_MEMBER)    // 4820 Pad this bud
#pragma warning(disable:__WARNING_SPECTRE_MITIGATION_FOR_MEMORY_LOAD) // 5045 Mitigate this bud

#ifdef _UNICODE
  #define VT_LPTSTR VT_LPWSTR
  #define IShellLink_SetPath IShellLinkW_SetPath
  #define IShellLink_SetArguments IShellLinkW_SetArguments
  #define IShellLink_SetIconLocation IShellLinkW_SetIconLocation
  #define IShellLink_SetDescription IShellLinkW_SetDescription
  #define IShellLink_Release IShellLinkW_Release
#elif defined _MBCS
  #define VT_LPTSTR VT_LPSTR
  #define IShellLink_SetPath IShellLinkA_SetPath
  #define IShellLink_SetArguments IShellLinkA_SetArguments
  #define IShellLink_SetIconLocation IShellLinkA_SetIconLocation
  #define IShellLink_SetDescription IShellLinkA_SetDescription
  #define IShellLink_Release IShellLinkA_Release
#else
  #error "Ensure a chosen character set in either _UNICODE or _MBCS"
#endif

#define CDLAPIPRIVATE(type) static __forceinline type __cdecl

#define VALIDATE_RETURN_HRESULT(expr, hr)                                      \
    {                                                                          \
        if (!(expr))                                                           \
        {                                                                      \
            return (hr);                                                       \
        }                                                                      \
    }

typedef struct TASK
{
  LPCTSTR pszImage;
  LPCTSTR pszArgs;
  LPCTSTR pszDescription;
  LPCTSTR pszTitle;
  LONG_PTR nIconIndex;
} TASK, *PTASK;

struct CDL
{
  TCHAR szImageName[_MAX_ENV];
  ICustomDestinationList* picdl;
  IObjectArray* pri;
  IObjectCollection* poc;
  UINT cMaxSlots;
};

CDLAPIPRIVATE(void)    SafeRelease(IUnknown** ppInterfaceToRelease);
CDLAPIPRIVATE(HRESULT) Initialize(ICDL** ppThis);
CDLAPIPRIVATE(HRESULT) BeginList(ICDL* pThis, PCWSTR pcwszAppId);
CDLAPIPRIVATE(HRESULT) AddUserTask(ICDL* pThis, PTASK pTask);
CDLAPIPRIVATE(HRESULT) AddSeparator(ICDL* pThis);
CDLAPIPRIVATE(HRESULT) CommitList(ICDL* pThis);

CDLAPIPRIVATE(void) SafeRelease(IUnknown** ppInterfaceToRelease)
{
    if ((*ppInterfaceToRelease))
    {
      IUnknown_Release((*ppInterfaceToRelease));
      (*ppInterfaceToRelease) = NULL;
    }
}

CDLAPIPRIVATE(HRESULT) Initialize(ICDL** ppThis)
{
    HRESULT hr;

    (*ppThis) = GlobalAlloc(GPTR, sizeof(ICDL));

    if (SUCCEEDED(hr = (*ppThis) ? S_OK : E_OUTOFMEMORY))
    {
      DWORD cchImageName;

      cchImageName = _countof((*ppThis)->szImageName);

      if (FAILED(hr = QueryFullProcessImageName(GetCurrentProcess(), 0, (*ppThis)->szImageName, &cchImageName) ? S_OK : E_FAIL))
      {
        GlobalFree((*ppThis));
      }
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) BeginList(ICDL* pThis, PCWSTR pcwszAppId)
{
    HRESULT hr;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pThis->picdl)))
    {
      if (SUCCEEDED(hr = CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pThis->poc)))
      {
        if (SUCCEEDED(hr = ICustomDestinationList_SetAppID(pThis->picdl, pcwszAppId)))
        {
          hr = ICustomDestinationList_BeginList(pThis->picdl, &pThis->cMaxSlots, &IID_IObjectArray, &pThis->pri);
        }
      }
    }

    if (FAILED(hr))
    {
      SafeRelease((IUnknown**)&pThis->pri);
      SafeRelease((IUnknown**)&pThis->poc);
      SafeRelease((IUnknown**)&pThis->picdl);
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) AddUserTask(ICDL* pThis, PTASK pTask)
{
    HRESULT hr;
    IShellLink* psl;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLink, &psl)))
    {
      IPropertyStore* pps;

      if (SUCCEEDED(hr = IPropertyStore_QueryInterface(psl, &IID_IPropertyStore, &pps)))
      {
        LPCTSTR pcszImage;

        pcszImage = pTask->pszImage ? pTask->pszImage : pThis->szImageName;

        if (SUCCEEDED(hr = IShellLink_SetPath(psl, pcszImage)))
        {
          if (SUCCEEDED(hr = IShellLink_SetArguments(psl, pTask->pszArgs)))
          {
            if (SUCCEEDED(hr = IShellLink_SetIconLocation(psl, pcszImage, (int)pTask->nIconIndex)))
            {
              if (SUCCEEDED(hr = IShellLink_SetDescription(psl, pTask->pszDescription)))
              {
                PROPVARIANT pv;
                SIZE_T cbTitle;

                PropVariantInit(&pv);
                cbTitle = (_tcslen(pTask->pszTitle) + 1) * sizeof(TCHAR);
                V_UNION(&pv, pwszVal) = CoTaskMemAlloc(cbTitle);

                if (SUCCEEDED(hr = V_UNION(&pv, pwszVal) ? S_OK : E_OUTOFMEMORY))
                {
                  CopyMemory(V_UNION(&pv, pwszVal), pTask->pszTitle, cbTitle);
                  V_VT(&pv) = VT_LPTSTR;

                  if (SUCCEEDED(hr = IPropertyStore_SetValue(pps, &PKEY_Title, &pv)))
                  {
                    hr = IPropertyStore_Commit(pps);
                    PropVariantClear(&pv);

                    if (SUCCEEDED(hr))
                    {
                      hr = IObjectCollection_AddObject(pThis->poc, (IUnknown*)psl);
                    }
                  }
                }
              }
            }
          }
        }

        SafeRelease((IUnknown**)&pps);
      }

      SafeRelease((IUnknown**)&psl);
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) AddSeparator(ICDL* pThis)
{
    HRESULT hr;
    IShellLink* psl;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLink, &psl)))
    {
      IPropertyStore* pps;

      if (SUCCEEDED(hr = IPropertyStore_QueryInterface(psl, &IID_IPropertyStore, &pps)))
      {
        PROPVARIANT pv;

        PropVariantInit(&pv);
        V_VT(&pv) = VT_BOOL;
        V_UNION(&pv, boolVal) = VARIANT_TRUE;

        if (SUCCEEDED(hr = IPropertyStore_SetValue(pps, &PKEY_AppUserModel_IsDestListSeparator, &pv)))
        {
          hr = IPropertyStore_Commit(pps);
          PropVariantClear(&pv);

          if (SUCCEEDED(hr))
          {
            hr = IObjectCollection_AddObject(pThis->poc, (IUnknown*)psl);
          }
        }
      }

      SafeRelease((IUnknown**)&pps);
    }

    SafeRelease((IUnknown**)&psl);

    return hr;
}

CDLAPIPRIVATE(HRESULT) CommitList(ICDL* pThis)
{
    HRESULT hr;
    IObjectArray* poa;

    if (SUCCEEDED(hr = IObjectArray_QueryInterface(pThis->poc, &IID_IObjectArray, &poa)))
    {
      if (SUCCEEDED(hr = ICustomDestinationList_AddUserTasks(pThis->picdl, poa)))
      {
        hr = ICustomDestinationList_CommitList(pThis->picdl);
      }
    }

    SafeRelease((IUnknown**)&poa);

    return hr;
}

CDLAPI(HRESULT) ICDL_BeginList(LPCWSTR pcwszAppId, ICDL** ppThis)
{
    HRESULT hr;

    VALIDATE_RETURN_HRESULT(ppThis, E_POINTER)

    if (SUCCEEDED(hr = Initialize(ppThis)))
    {
      hr = BeginList((*ppThis), pcwszAppId);
    }

    return hr;
}

CDLAPI(HRESULT) ICDL_AddTask(ICDL* pThis, LPCTSTR pcszImage, LPCTSTR pcszArgs, LPCTSTR pcszDescription, LPCTSTR pcszTitle, int nIconIndex)
{
    TASK task;

    VALIDATE_RETURN_HRESULT(pThis, E_POINTER)

    task.pszImage       = pcszImage;
    task.pszArgs        = pcszArgs;
    task.pszDescription = pcszDescription;
    task.pszTitle       = pcszTitle;
    task.nIconIndex     = nIconIndex;

    return AddUserTask(pThis, &task);
}

CDLAPI(HRESULT) ICDL_AddSeparator(ICDL* pThis)
{
    VALIDATE_RETURN_HRESULT(pThis, E_POINTER)

    return AddSeparator(pThis);
}

CDLAPI(HRESULT) ICDL_CommitList(ICDL* pThis)
{
    VALIDATE_RETURN_HRESULT(pThis, E_POINTER)

    return CommitList(pThis);
}

CDLAPI(VOID) ICDL_Release(ICDL* pThis)
{
    if (pThis)
    {
      SafeRelease((IUnknown**)&pThis->poc);
      SafeRelease((IUnknown**)&pThis->pri);
      SafeRelease((IUnknown**)&pThis->picdl);
      GlobalFree(pThis);
    }
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    UNREFERENCED_PARAMETER(lpReserved);

    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
      DisableThreadLibraryCalls(hModule);
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

