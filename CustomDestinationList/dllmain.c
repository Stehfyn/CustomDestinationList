#include "CustomDestinationList.h"

#include <unknwnbase.h>
#include <oleauto.h>
#include <objectarray.h>
#include <propkey.h>
#include <propsys.h>
#include <propvarutil.h>
#include <shlobj.h>
#include <winbase.h>

#pragma comment (lib, "Propsys")

// Adding macros <suppress.h> _conveniently_ does not provide
#define __WARNING_EXPRESSION_BEFORE_COMMA_HAS_NO_EFFECT 4548
#define __WARNING_PADDING_ADDED_AFTER_DATA_MEMBER       4820
#define __WARNING_SPECTRE_MITIGATION_FOR_MEMORY_LOAD    5045
#pragma warning(disable:__WARNING_EXPRESSION_BEFORE_COMMA_HAS_NO_EFFECT) // 4548 Sure bud
#pragma warning(disable:__WARNING_PADDING_ADDED_AFTER_DATA_MEMBER)       // 4820 Pad this bud
#pragma warning(disable:__WARNING_SPECTRE_MITIGATION_FOR_MEMORY_LOAD)    // 5045 Mitigate this bud

#define CDLAPIPRIVATE(type) static __forceinline type __cdecl

#define VALIDATE_RETURN_HRESULT(expr, hr, msg)                                 \
    do {                                                                       \
        if (!(expr))                                                           \
        {                                                                      \
            return (hr);                                                       \
        }                                                                      \
    } while(0)

#define CDL_ANSI     (1)
#define CDL_NOT_ANSI (0)

typedef struct CTASKW
{
  WCHAR wszImage[_MAX_ENV];
  WCHAR wszArgs[_MAX_ENV]; // Don't assume how this makes its way to NtCreateProcess https://stackoverflow.com/a/28452546
  WCHAR wszDescription[INFOTIPSIZE];
  WCHAR wszTitle[INFOTIPSIZE];
  WCHAR wszIconLocation[_MAX_ENV];
  LONG_PTR nIconIndex;
} CTASKW, *PCTASKW;

typedef struct TASKV
{
  LPCVOID pcvImage;
  LPCVOID pcvArgs;
  LPCVOID pcvDescription;
  LPCVOID pcvTitle;
  LPCVOID pcvIconLocation;
  LONG_PTR nIconIndex;
} TASKV, *PTASKV;

typedef struct TASKW
{
  LPCWSTR pcwszImage;
  LPCWSTR pcwszArgs;
  LPCWSTR pcwszDescription;
  LPCWSTR pcwszTitle;
  LPCWSTR pcwszIconLocation;
  LONG_PTR nIconIndex;
} TASKW, *PTASKW;

C_ASSERT(sizeof(TASKW) == sizeof(TASKV));

struct CDL
{
  WCHAR wszImageName[_MAX_ENV];
  ICustomDestinationList* pcdl;
  IObjectArray* poaRemovedDestinations;
  IObjectCollection* pocUserTask;
  IObjectCollection* pocCategory;
  UINT cMaxSlots;
  WCHAR wszCurrentCategory[60];
  BOOL fInBeginList;
  BOOL fInBeginCategory;
};

CDLAPIPRIVATE(void)    SafeRelease(IUnknown** ppInterfaceToRelease);
CDLAPIPRIVATE(BOOL)    EnsurePointer(void* lp, UINT_PTR ucb);
CDLAPIPRIVATE(BOOL)    IsSystemCategory(PCWSTR pcwszCategory);
CDLAPIPRIVATE(HRESULT) Initialize(ICDL** ppThis);
CDLAPIPRIVATE(HRESULT) BeginList(ICDL* pThis, PCWSTR pcwszAppId);
CDLAPIPRIVATE(HRESULT) BeginCategory(ICDL* pThis, LPCVOID pcvCustomCategory, BOOL fAnsi);
CDLAPIPRIVATE(HRESULT) AddUserTask(ICDL* pThis, PTASKV pvTask, BOOL fAnsi);
CDLAPIPRIVATE(HRESULT) AddSeparator(ICDL* pThis);
CDLAPIPRIVATE(HRESULT) CommitCategory(ICDL* pThis);
CDLAPIPRIVATE(HRESULT) CommitList(ICDL* pThis);

CDLAPIPRIVATE(void) SafeRelease(IUnknown** ppInterfaceToRelease)
{
    if ((*ppInterfaceToRelease))
    {
      IUnknown_Release((*ppInterfaceToRelease));
      (*ppInterfaceToRelease) = NULL;
    }
}

// Four options:
// 1). Use deprecated IsBadHugeReadPtr()/ IsBadHugeWritePtr, and open self up to certain uncaught SEH exceptions for things like page guards
// 2). Use VirtualQuery, and tank kernel traps to consult the memory manager
// 3). Roll our own userspace validation, leveraging optimizations-disabled idempotent memcpy->memset->memcpy (swap-zero-swap) via a temporary buffer
// 4). Choose to not care--just evaluate the pointer to a boolean, because the above only matters in a true arbitrary ASTA or MTA COM apartment evnvironment
CDLAPIPRIVATE(BOOL) EnsurePointer(void* lp, UINT_PTR ucb)
{
    UNREFERENCED_PARAMETER(ucb);
    
    return !!lp;
}

CDLAPIPRIVATE(BOOL) IsSystemCategory(PCWSTR pcwszCategory)
{
    static PCWSTR c_ppcwszSystemCategories[] =
    {
      L"Frequent",
      L"Recent",
      L"Tasks",
    };

    SIZE_T nIndex;

    for (nIndex = 0; nIndex < _countof(c_ppcwszSystemCategories); ++nIndex)
    {
        PCWSTR pcwszSystemCategory;

        pcwszSystemCategory = c_ppcwszSystemCategories[nIndex];
        
        if (0 == _wcsicmp(pcwszCategory, pcwszSystemCategory))
        {
          return TRUE;
        }
    }

    return FALSE;
}

CDLAPIPRIVATE(HRESULT) Initialize(ICDL** ppThis)
{
    HRESULT hr;

    (*ppThis) = GlobalAlloc(GPTR, sizeof(ICDL));

    if (SUCCEEDED(hr = (*ppThis) ? S_OK : E_OUTOFMEMORY))
    {
      DWORD cchImageName;

      cchImageName = _countof((*ppThis)->wszImageName);

      if (FAILED(hr = QueryFullProcessImageNameW(GetCurrentProcess(), 0, (*ppThis)->wszImageName, &cchImageName) ? S_OK : E_FAIL))
      {
        GlobalFree((*ppThis));
      }
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) BeginList(ICDL* pThis, PCWSTR pcwszAppId)
{
    HRESULT hr;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pThis->pcdl)))
    {
      if (SUCCEEDED(hr = CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pThis->pocUserTask)))
      {
        if (SUCCEEDED(hr = ICustomDestinationList_SetAppID(pThis->pcdl, pcwszAppId)))
        {
          if (SUCCEEDED(hr = ICustomDestinationList_BeginList(pThis->pcdl, &pThis->cMaxSlots, &IID_IObjectArray, &pThis->poaRemovedDestinations)))
          {
            pThis->fInBeginList = TRUE;
          }
        }
      }
    }

    if (FAILED(hr))
    {
      SafeRelease((IUnknown**)&pThis->poaRemovedDestinations);
      SafeRelease((IUnknown**)&pThis->pocUserTask);
      SafeRelease((IUnknown**)&pThis->pcdl);
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) BeginCategory(ICDL* pThis, LPCVOID pcvCustomCategory, BOOL fAnsi)
{
    HRESULT hr;
    LPCWSTR lpcwszCustomCategory;
    WCHAR   wszFromAnsiCustomCategory[60];
    SIZE_T  cbCustomCategory;

    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_UNEXPECTED, "Must be currently creating a jump list");
    VALIDATE_RETURN_HRESULT(!pThis->fInBeginCategory, E_UNEXPECTED, "Must not be currently creating a custom category");
    VALIDATE_RETURN_HRESULT(pcvCustomCategory, E_INVALIDARG, "Must pass a valid pointer to category name");
    
    if (fAnsi)
    {
      MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pcvCustomCategory, -1, wszFromAnsiCustomCategory, (int)_countof(wszFromAnsiCustomCategory));

      lpcwszCustomCategory = &wszFromAnsiCustomCategory[0];
    }
    else
    {
      lpcwszCustomCategory = (LPCWSTR)pcvCustomCategory;
    }

    cbCustomCategory = sizeof(WCHAR) * (wcsnlen(lpcwszCustomCategory, _countof(wszFromAnsiCustomCategory) - 1) + 1);

    if (SUCCEEDED(hr = !IsSystemCategory(lpcwszCustomCategory) ? S_OK : E_INVALIDARG))
    {
      if (SUCCEEDED(hr = CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pThis->pocCategory)))
      {
        CopyMemory(pThis->wszCurrentCategory, lpcwszCustomCategory, cbCustomCategory);
        
        pThis->fInBeginCategory = TRUE;
      }
    }

    return hr;
}

CDLAPIPRIVATE(HRESULT) AddUserTask(ICDL* pThis, PTASKV pvTask, BOOL fAnsi)
{
    HRESULT hr;
    IShellLinkW* psl;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLinkW, &psl)))
    {
      IPropertyStore* pps;

      if (SUCCEEDED(hr = IPropertyStore_QueryInterface(psl, &IID_IPropertyStore, &pps)))
      {
        PTASKW  pwTask;
        PCTASKW pwCTask;

        pwCTask = NULL;

        if (fAnsi)
        {
          pwCTask = CoTaskMemAlloc(sizeof(CTASKW));

          if (SUCCEEDED(hr = pwCTask ? S_OK : E_OUTOFMEMORY))
          {
            MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pvTask->pcvImage,        -1, pwCTask->wszImage,        _countof(pwCTask->wszImage));
            MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pvTask->pcvArgs,         -1, pwCTask->wszArgs,         _countof(pwCTask->wszArgs));
            MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pvTask->pcvDescription,  -1, pwCTask->wszDescription,  _countof(pwCTask->wszDescription));
            MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pvTask->pcvTitle,        -1, pwCTask->wszTitle,        _countof(pwCTask->wszTitle));
            MultiByteToWideChar(CP_UTF8, 0, (LPCCH)pvTask->pcvIconLocation, -1, pwCTask->wszIconLocation, _countof(pwCTask->wszIconLocation));
            pvTask->pcvImage        = pwCTask->wszImage;
            pvTask->pcvArgs         = pwCTask->wszArgs;
            pvTask->pcvDescription  = pwCTask->wszDescription;
            pvTask->pcvTitle        = pwCTask->wszTitle;
            pvTask->pcvIconLocation = pwCTask->wszIconLocation;
          }
        }

        pwTask = (PTASKW)pvTask;

        if (SUCCEEDED(hr))
        {
          LPCWSTR pcwszImage;

          pcwszImage = pwTask->pcwszImage ? pwTask->pcwszImage : pThis->wszImageName;

          if (SUCCEEDED(hr = IShellLinkW_SetPath(psl, pcwszImage)))
          {
            if (SUCCEEDED(hr = IShellLinkW_SetArguments(psl, pwTask->pcwszArgs)))
            {
              LPCWSTR pcwszIconLocation;

              pcwszIconLocation = (pwTask->pcwszIconLocation) ? pwTask->pcwszIconLocation : pwTask->pcwszImage;

              if (SUCCEEDED(hr = IShellLinkW_SetIconLocation(psl, pcwszIconLocation, (int)pwTask->nIconIndex)))
              {
                if (SUCCEEDED(hr = IShellLinkW_SetDescription(psl, pwTask->pcwszDescription)))
                {
                  PROPVARIANT pv;
                  SIZE_T cbTitle;

                  PropVariantInit(&pv);
                  cbTitle = sizeof(WCHAR) * (wcslen(pwTask->pcwszTitle) + 1);
                  V_UNION(&pv, pwszVal) = CoTaskMemAlloc(cbTitle);

                  if (SUCCEEDED(hr = V_UNION(&pv, pwszVal) ? S_OK : E_OUTOFMEMORY))
                  {
                    CopyMemory(V_UNION(&pv, pwszVal), pwTask->pcwszTitle, cbTitle);
                    V_VT(&pv) = VT_LPWSTR;

                    if (SUCCEEDED(hr = IPropertyStore_SetValue(pps, &PKEY_Title, &pv)))
                    {
                      hr = IPropertyStore_Commit(pps);
                      PropVariantClear(&pv);

                      if (SUCCEEDED(hr))
                      {
                        IObjectCollection* poc;

                        poc = (pThis->fInBeginCategory) ? pThis->pocCategory : pThis->pocUserTask;

                        hr = IObjectCollection_AddObject(poc, (IUnknown*)psl);
                      }
                    }
                  }
                }
              }
            }
          }
        }

        if (pwCTask)
        {
          CoTaskMemFree(pwCTask);
          pwCTask = NULL;
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
    IShellLinkW* psl;

    if (SUCCEEDED(hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLinkW, &psl)))
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
            hr = IObjectCollection_AddObject(pThis->pocUserTask, (IUnknown*)psl);
          }
        }
      }

      SafeRelease((IUnknown**)&pps);
    }

    SafeRelease((IUnknown**)&psl);

    return hr;
}

CDLAPIPRIVATE(HRESULT) CommitCategory(ICDL* pThis)
{
    HRESULT hr;
    IObjectArray* poa;

    if (SUCCEEDED(hr = IObjectArray_QueryInterface(pThis->pocCategory, &IID_IObjectArray, &poa)))
    {
      if (SUCCEEDED(hr = ICustomDestinationList_AppendCategory(pThis->pcdl, pThis->wszCurrentCategory, poa)))
      {
          pThis->wszCurrentCategory[0] = 0;
          pThis->fInBeginCategory = FALSE;

          SafeRelease((IUnknown**)&pThis->pocCategory);
      }
    }

    SafeRelease((IUnknown**)&poa);

    return hr;
}

CDLAPIPRIVATE(HRESULT) CommitList(ICDL* pThis)
{
    HRESULT hr;
    IObjectArray* poa;

    if (SUCCEEDED(hr = IObjectArray_QueryInterface(pThis->pocUserTask, &IID_IObjectArray, &poa)))
    {
      UINT cObjects;

      if (SUCCEEDED(hr = IObjectArray_GetCount(pThis->pocUserTask, &cObjects)))
      {
        // Only need to _AddUserTasks if there are some to add

        if ((0 == cObjects) || SUCCEEDED(hr = ICustomDestinationList_AddUserTasks(pThis->pcdl, poa))) 
        {
          hr = ICustomDestinationList_CommitList(pThis->pcdl);
        }
      }
    }

    SafeRelease((IUnknown**)&poa);

    return hr;
}

CDLAPI(HRESULT) ICDL_BeginList(LPCWSTR pcwszAppId, ICDL** ppThis)
{
    HRESULT hr;

    VALIDATE_RETURN_HRESULT(ppThis, E_POINTER, "ICDL** pointer is invalid");

    if (SUCCEEDED(hr = Initialize(ppThis)))
    {
      hr = BeginList((*ppThis), pcwszAppId);
    }

    return hr;
}

CDLAPI(HRESULT) ICDL_BeginCategoryA(ICDL* pThis, LPCSTR pcszCategory)
{
    return BeginCategory(pThis, pcszCategory, CDL_ANSI);
}

CDLAPI(HRESULT) ICDL_BeginCategoryW(ICDL* pThis, LPCWSTR pcwszCategory)
{
    return BeginCategory(pThis, pcwszCategory, CDL_NOT_ANSI);
}

CDLAPI(HRESULT) ICDL_AddTaskA(ICDL* pThis, LPCSTR pcszImage, LPCSTR pcszArgs, LPCSTR pcszDescription, LPCSTR pcszTitle, LPCSTR pcszIconLocation, int nIconIndex)
{
    TASKV vtask;

    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_INVALIDARG, "Must be currently creating a jump list");

    vtask.pcvImage        = pcszImage;
    vtask.pcvArgs         = pcszArgs;
    vtask.pcvDescription  = pcszDescription;
    vtask.pcvTitle        = pcszTitle;
    vtask.pcvIconLocation = pcszIconLocation;
    vtask.nIconIndex      = nIconIndex;

    return AddUserTask(pThis, &vtask, CDL_ANSI);
}

CDLAPI(HRESULT) ICDL_AddTaskW(ICDL* pThis, LPCWSTR pcwszImage, LPCWSTR pcwszArgs, LPCWSTR pcwszDescription, LPCWSTR pcwszTitle, LPCWSTR pcwszIconLocation, int nIconIndex)
{
    TASKV vtask;

    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_INVALIDARG, "Must be currently creating a jump list");

    vtask.pcvImage        = pcwszImage;
    vtask.pcvArgs         = pcwszArgs;
    vtask.pcvDescription  = pcwszDescription;
    vtask.pcvTitle        = pcwszTitle;
    vtask.pcvIconLocation = pcwszIconLocation;
    vtask.nIconIndex      = nIconIndex;

    return AddUserTask(pThis, &vtask, CDL_NOT_ANSI);
}

CDLAPI(HRESULT) ICDL_AddSeparator(ICDL* pThis)
{
    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_INVALIDARG, "Must be currently creating a jump list");
    VALIDATE_RETURN_HRESULT(!pThis->fInBeginCategory, E_INVALIDARG, "Separators can't be added to custom categories");

    return AddSeparator(pThis);
}

CDLAPI(HRESULT) ICDL_CommitCategory(ICDL* pThis)
{
    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_INVALIDARG, "Must be currently creating a jump list");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginCategory, E_INVALIDARG, "Must be currently creating a custom category");
    
    return CommitCategory(pThis);
}

CDLAPI(HRESULT) ICDL_CommitList(ICDL* pThis)
{
    VALIDATE_RETURN_HRESULT(EnsurePointer(pThis, sizeof(ICDL)), E_POINTER, "ICDL* pointer is invalid");
    VALIDATE_RETURN_HRESULT(pThis->fInBeginList, E_INVALIDARG, "Must be currently creating a jump list");
    VALIDATE_RETURN_HRESULT(!pThis->fInBeginCategory, E_INVALIDARG, "Must not be currently creating a custom category");

    return CommitList(pThis);
}

CDLAPI(VOID) ICDL_Release(ICDL* pThis)
{
    if (pThis)
    {
      SafeRelease((IUnknown**)&pThis->pocUserTask);
      SafeRelease((IUnknown**)&pThis->pocCategory);
      SafeRelease((IUnknown**)&pThis->poaRemovedDestinations);
      SafeRelease((IUnknown**)&pThis->pcdl);
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

