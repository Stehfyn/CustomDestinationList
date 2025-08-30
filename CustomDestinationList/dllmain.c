#include "CustomDestinationList.h"

#include <windowsx.h>

#include <psapi.h>
#include <shlwapi.h>
#include <strsafe.h>

#include <objectarray.h>
#include <shobjidl.h>
#include <propkey.h>
#include <propsys.h>
#include <propvarutil.h>
#include <knownfolders.h>
#include <shlobj.h>
#include <tchar.h>

#pragma comment (lib, "Propsys")

#define EXIT_SUCCESS      (0)
#define EXIT_FAILURE      (1)

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

struct Task
{
  LPCTSTR szArgs;
  LPCTSTR szDescription;
  LPCTSTR szTitle;
  int nIconIndex;
};

struct CDL
{
  HANDLE hThread;
  DWORD  dwThreadId;
  HWND   hwnd;
  DWORD  dwLaunchThreadId;
  ICustomDestinationList* picdl;
  IObjectCollection* poc;
  UINT cMaxSlots;
  TCHAR szImageName[MAX_PATH];
};

EXTERN_C TCHAR c_szClassName[]   = _T("CDLWndClass");
EXTERN_C TCHAR c_szWindowTitle[] = _T("CDLWindow");

#define WM_CREATEJUMPLIST (WM_USER + 0x0)
#define WM_ADDUSERTASKS   (WM_USER + 0x1)
#define WM_ADDSEPARATOR   (WM_USER + 0x2)
#define WM_COMMITJUMPLIST (WM_USER + 0x3)

#define HANDLE_WM_CREATEJUMPLIST(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd), (PCWSTR)(wParam)))

#define HANDLE_WM_ADDUSERTASKS(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd), (ITask*)(wParam), (UINT)(lParam)))

#define HANDLE_WM_ADDSEPARATOR(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd)))

#define HANDLE_WM_COMMITJUMPLIST(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd)))

static FORCEINLINE HRESULT CALLBACK OnCreateJumpList(HWND hwnd, PCWSTR pszAppId)
{
    HRESULT hr;
    ICDL* pThis;
    IObjectArray* pri;

    pThis = GetWindowLongPtr(hwnd, GWLP_USERDATA);

    if (!pThis)
    {
      return E_FAIL;
    }

    hr = CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pThis->picdl);

    if (FAILED(hr))
    {
      return hr;
    }

    hr = ICustomDestinationList_SetAppID(pThis->picdl, pszAppId);

    if (FAILED(hr))
    {
      ICustomDestinationList_Release(pThis->picdl);

      pThis->picdl = NULL;

      return hr;
    }

    hr = ICustomDestinationList_BeginList(pThis->picdl, &pThis->cMaxSlots, &IID_IObjectArray, &pri);

    if (FAILED(hr))
    {
      ICustomDestinationList_Release(pThis->picdl);

      pThis->picdl = NULL;
    }

    return CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pThis->poc);
}

static FORCEINLINE HRESULT CALLBACK OnAddUserTasks(HWND hwnd, ITask* pTasks, UINT elems)
{
    HRESULT  hr;
    ICDL* pThis;
    UINT elm;
    IShellLink* psl;
    IPropertyStore* pps;
    PROPVARIANT pv;

    pThis = GetWindowLongPtr(hwnd, GWLP_USERDATA);

    if (!pThis)
    {
      return E_FAIL;
    }

    for (elm = 0; elm < elems; ++elm)
    {
      const ITask* c_pTask = &pTasks[elm];

      PCWSTR ppszTitle[] = { c_pTask->szTitle };

      hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLink, &psl);

      if (FAILED(hr))
      {
        return hr;
      }

      hr = IPropertyStore_QueryInterface(psl, &IID_IPropertyStore, &pps);

      if (!pps)
      {
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetPath(psl, pThis->szImageName);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetArguments(psl, c_pTask->szArgs);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetIconLocation(psl, pThis->szImageName, c_pTask->nIconIndex);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetDescription(psl, c_pTask->szDescription);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      SecureZeroMemory(&pv, sizeof(pv));
      hr = InitVariantFromStringArray(ppszTitle, 1, &pv);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IPropertyStore_SetValue(pps, &PKEY_Title, &pv);

      PropVariantClear(&pv);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IPropertyStore_Commit(pps);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IObjectCollection_AddObject(pThis->poc, psl);

      IPropertyStore_Release(pps);
      IShellLinkW_Release(psl);

      if (FAILED(hr))
      {
        return hr;
      }
    }

    return S_OK;
}

static FORCEINLINE HRESULT CALLBACK OnAddSeparator(HWND hwnd)
{
    HRESULT hr;
    IShellLink* psl;
    IPropertyStore* pps;
    PROPVARIANT pv;
    ICDL* icdl;
    BOOL flag;

    hr = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLink, &psl);

    if (FAILED(hr))
    {
      return E_FAIL;
    }

    IPropertyStore_QueryInterface(psl, &IID_IPropertyStore, &pps);

    if (!pps)
    {
      IShellLinkW_Release(psl);
      return E_FAIL;
    }

    SecureZeroMemory(&pv, sizeof(pv));
    flag = TRUE;
    hr = InitVariantFromBooleanArray(&flag, 1, &pv);
    if (FAILED(hr))
    {
      IShellLinkW_Release(psl);
      return E_FAIL;
    }

    hr = IPropertyStore_SetValue(pps, &PKEY_AppUserModel_IsDestListSeparator, &pv);
    PropVariantClear ( &pv );

    if ( FAILED(hr) )
    {
      IShellLinkW_Release(psl);
      return E_FAIL;
    }

    hr = IPropertyStore_Commit(pps);

    if ( FAILED(hr) )
    {
      IShellLinkW_Release(psl);
      return E_FAIL;
    }

    icdl = GetWindowLongPtr(hwnd, GWLP_USERDATA);

    if (!icdl)
    {
      IShellLinkW_Release(psl);
      return E_FAIL;
    }

    hr = IObjectCollection_AddObject(icdl->poc, psl);
    hr = ICustomDestinationList_AddUserTasks(icdl->picdl, icdl->poc);
    hr = ICustomDestinationList_CommitList(icdl->picdl);

    return SUCCEEDED(hr);
}

static FORCEINLINE HRESULT CALLBACK OnCommitJumpList(HWND hwnd)
{
    HRESULT  hr;
    ICDL* pThis;

    pThis = GetWindowLongPtr(hwnd, GWLP_USERDATA);

    if (!pThis)
    {
      return E_FAIL;
    }

    hr = ICustomDestinationList_AddUserTasks(pThis->picdl, pThis->poc);

    if (FAILED(hr))
    {
      return hr;
    }

    return ICustomDestinationList_CommitList(pThis->picdl);
}

static FORCEINLINE BOOL APIPRIVATE PumpMessages(void)
{
    MSG  msg;
    BOOL fQuit = FALSE;
    SecureZeroMemory(&msg, sizeof(msg));

    while (PeekMessage(&msg, 0, 0, 0, PM_REMOVE))
    {
      DispatchMessage(&msg);

      fQuit |= (WM_QUIT == msg.message);
    }

    return !fQuit;
}

static FORCEINLINE BOOL CALLBACK OnNCCreate(HWND hwnd, LPCREATESTRUCT lpCreateStruct)
{
    LONG_PTR offset;

    SetLastError(NO_ERROR);
    offset = SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)lpCreateStruct->lpCreateParams);
    if ((0 == offset) && (NO_ERROR != GetLastError()))
    {
      return FALSE;
    }

    return FORWARD_WM_NCCREATE(hwnd, lpCreateStruct, DefWindowProc);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg) {
    HANDLE_MSG(hwnd, WM_NCCREATE,       OnNCCreate);
    HANDLE_MSG(hwnd, WM_CREATEJUMPLIST, OnCreateJumpList);
    HANDLE_MSG(hwnd, WM_ADDUSERTASKS,   OnAddUserTasks);
    HANDLE_MSG(hwnd, WM_ADDSEPARATOR,   OnAddSeparator);
    HANDLE_MSG(hwnd, WM_COMMITJUMPLIST, OnCommitJumpList);
    default:
      return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

static HWND APIPRIVATE CreateMessageWindow(LPARAM lpParam)
{
    WNDCLASS wc;
    ATOM    atm;

    SecureZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = &__ImageBase;
    wc.lpszClassName = c_szClassName;
    
    atm = RegisterClass(&wc);

    if (!atm)
    {
      return NULL;
    }

    return CreateWindow(atm, c_szWindowTitle, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, &__ImageBase, lpParam);
}

static DWORD WINAPI CDLThread(LPVOID lpThreadParameter)
{
    HRESULT hr;
    ICDL* icdl;

    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    if (!SUCCEEDED(hr))
    {
      return EXIT_FAILURE;
    }

    icdl = (ICDL*)lpThreadParameter;
    icdl->hwnd = CreateMessageWindow(icdl);

    if (!icdl->hwnd)
    {
      return EXIT_FAILURE;
    }

    PostThreadMessage(icdl->dwLaunchThreadId, 0, 0, 0);

    while (PumpMessages());

    return EXIT_SUCCESS;
}

static void CreateThreadMessageQueue(void)
{
    MSG msg;
    PostThreadMessage(GetCurrentThreadId(), 0, 0, 0);
    GetMessage(&msg, 0, 0, 0);
}

static DWORD WINAPI CDLLaunch(LPVOID lpThreadParameter)
{
    ICDL* icdl;

    icdl = (ICDL*)lpThreadParameter;
    icdl->hThread = CreateThread(0, 0, CDLThread, icdl, CREATE_SUSPENDED, &icdl->dwThreadId);

    if (!icdl->hThread)
    {
      return EXIT_FAILURE;
    }

    CreateThreadMessageQueue();

    ResumeThread(icdl->hThread);
    
    WaitMessage();

    return EXIT_SUCCESS;
}

static HRESULT APIPRIVATE Initialize(ICDL* pThis)
{
    HANDLE hThread;
    DWORD  cchImageName;

    cchImageName = _countof(pThis->szImageName);

    if (!QueryFullProcessImageName(GetCurrentProcess(), 0, pThis->szImageName, &cchImageName))
    {
      return E_FAIL;
    }

    hThread = CreateThread(0, 0, CDLLaunch, pThis, CREATE_SUSPENDED, &pThis->dwLaunchThreadId);

    if (!hThread)
    {
      return E_FAIL;
    }

    ResumeThread(hThread);

    WaitForSingleObject(hThread, INFINITE);

    return S_OK;
}

CDLAPI(HRESULT) ICDL_CreateTaskList(ICDL* pThis, UINT elems, ITask** ppTasks)
{
    (*ppTasks) = GlobalAlloc(GPTR, elems * sizeof(ITask));

    if (!(*ppTasks))
    {
      return E_OUTOFMEMORY;
    }

    return S_OK;
}

CDLAPI(HRESULT) ICDL_SetTask(ICDL* pThis, ITask* pTasks, UINT elem, const ICDL_TASK_PARAMS* params)
{
    ITask* pTask = &pTasks[elem];

    pTask->szArgs        = params->szArgs;
    pTask->szDescription = params->szDescription;
    pTask->szTitle       = params->szTitle;
    pTask->nIconIndex    = params->nIconIndex;

    return S_OK;
}

CDLAPI(HRESULT) ICDL_Initialize(ICDL** ppThis)
{
  (*ppThis) = GlobalAlloc(GPTR, sizeof(ICDL));

  if (!(*ppThis))
  {
    return E_OUTOFMEMORY;
  }

  return Initialize((*ppThis));
}

CDLAPI(HRESULT) ICDL_CreateJumpList(ICDL* pThis, PCWSTR pcszAppId)
{
  return (HRESULT)SendMessage(pThis->hwnd, WM_CREATEJUMPLIST, pcszAppId, 0);
}

CDLAPI(HRESULT) ICDL_AddUserTasks(ICDL* pThis, ITask* pTasks, UINT elems)
{
  return (HRESULT)SendMessage(pThis->hwnd, WM_ADDUSERTASKS, pTasks, elems);
}

CDLAPI(HRESULT) ICDL_AddSeparator(ICDL* pThis)
{
  return (HRESULT)SendMessage(pThis->hwnd, WM_ADDSEPARATOR, 0, 0);
}

CDLAPI(HRESULT) ICDL_CommitJumpList(ICDL* pThis)
{
  return (HRESULT)SendMessage(pThis->hwnd, WM_COMMITJUMPLIST, 0, 0);
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
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

