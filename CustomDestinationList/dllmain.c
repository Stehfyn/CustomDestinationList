#include "CustomDestinationList.h"

#include <oleauto.h>
#include <objectarray.h>
#include <propkey.h>
#include <propsys.h>
#include <propvarutil.h>

#pragma comment (lib, "Propsys")

#define EXIT_SUCCESS      (0)
#define EXIT_FAILURE      (1)
#define PATHLEN           (256)

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

struct Task
{
  LPCTSTR pszImage;
  LPCTSTR pszArgs;
  LPCTSTR pszDescription;
  LPCTSTR pszTitle;
  LONG_PTR nIconIndex;
};

struct CDL
{
  TCHAR szImageName[PATHLEN];
  ICustomDestinationList* picdl;
  IObjectCollection* poc;
  HWND   hwnd;
  HANDLE hThread;
  DWORD  dwThreadId;
  DWORD  dwLaunchThreadId;
  UINT_PTR cMaxSlots;
};

EXTERN_C TCHAR c_szClassName[]   = _T("CDLWndClass");
EXTERN_C TCHAR c_szWindowTitle[] = _T("CDLWindow");

#define GetInterfaceHandle(hwnd) \
        ((ICDL*)GetWindowLongPtr(hwnd, GWLP_USERDATA))

#define WM_CREATEJUMPLIST (WM_USER + 0x0)
#define WM_ADDUSERTASKS   (WM_USER + 0x1)
#define WM_ADDSEPARATOR   (WM_USER + 0x2)
#define WM_COMMITJUMPLIST (WM_USER + 0x3)

#define HANDLE_WM_CREATEJUMPLIST(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd), (PCWSTR)(wParam)))

#define HANDLE_WM_ADDUSERTASKS(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd), (ITask*)(wParam), (UINT16)(HIWORD((lParam))), (UINT16)(LOWORD((lParam)))))

#define HANDLE_WM_ADDSEPARATOR(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd)))

#define HANDLE_WM_COMMITJUMPLIST(hwnd, wParam, lParam, fn) \
        (LRESULT)((fn)((hwnd)))

static FORCEINLINE HRESULT CALLBACK OnCreateJumpList(HWND hwnd, PCWSTR pszAppId)
{
    HRESULT hr;
    ICDL* pThis;
    IObjectArray* pri;

    pThis = GetInterfaceHandle(hwnd);

    if (!pThis)
    {
      return E_FAIL;
    }
    
    hr = CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pThis->picdl);

    if (FAILED(hr))
    {
      return hr;
    }

    hr = CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pThis->poc);

    if (FAILED(hr))
    {
      ICustomDestinationList_Release(pThis->picdl);

      pThis->picdl = NULL;

      return hr;
    }

    hr = ICustomDestinationList_SetAppID(pThis->picdl, pszAppId);

    if (FAILED(hr))
    {
      ICustomDestinationList_Release(pThis->picdl);
      IObjectCollection_Release(pThis->poc);

      pThis->picdl = NULL;

      return hr;
    }
    
    hr = ICustomDestinationList_BeginList(pThis->picdl, (UINT*)&pThis->cMaxSlots, &IID_IObjectArray, &pri);

    if (FAILED(hr))
    {
      ICustomDestinationList_Release(pThis->picdl);
      IObjectCollection_Release(pThis->poc);

      pThis->picdl = NULL;
    }

    return hr;
}

static FORCEINLINE HRESULT CALLBACK OnAddUserTasks(HWND hwnd, ITask* pTasks, UINT16 elems, UINT16 offset)
{
    HRESULT  hr;
    ICDL* pThis;
    UINT elm;
    IShellLink* psl;
    IPropertyStore* pps;
    PROPVARIANT pv;
    SIZE_T cbTitle;
    LPCTSTR pszImage;

    pThis = GetInterfaceHandle(hwnd);

    if (!pThis)
    {
      return E_FAIL;
    }


    for (elm = 0; elm < elems; ++elm)
    {
      const ITask* c_pTask = &pTasks[elm + offset];

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

      pszImage = c_pTask->pszImage ? c_pTask->pszImage : pThis->szImageName;

      hr = IShellLinkW_SetPath(psl, pszImage);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetArguments(psl, c_pTask->pszArgs);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetIconLocation(psl, pszImage, (int)c_pTask->nIconIndex);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IShellLinkW_SetDescription(psl, c_pTask->pszDescription);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      PropVariantInit(&pv);
      cbTitle = (wcslen(c_pTask->pszTitle) + 1) * sizeof(TCHAR);
      V_UNION(&pv, pwszVal) = CoTaskMemAlloc(cbTitle);
      hr = V_UNION(&pv, pwszVal) ? S_OK : E_OUTOFMEMORY;

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      CopyMemory(V_UNION(&pv, pwszVal), c_pTask->pszTitle, cbTitle);
      V_VT(&pv) = VT_LPWSTR;
      hr = IPropertyStore_SetValue(pps, &PKEY_Title, &pv);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IPropertyStore_Commit(pps);
      PropVariantClear(&pv);

      if (FAILED(hr))
      {
        IPropertyStore_Release(pps);
        IShellLinkW_Release(psl);

        return hr;
      }

      hr = IObjectCollection_AddObject(pThis->poc, (IUnknown*)psl);

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
    ICDL* pThis;

    pThis = GetInterfaceHandle(hwnd);

    if (!pThis)
    {
      return E_FAIL;
    }

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

    PropVariantInit(&pv);
    V_VT(&pv) = VT_BOOL;
    V_UNION(&pv, boolVal) = VARIANT_TRUE;
    hr = IPropertyStore_SetValue(pps, &PKEY_AppUserModel_IsDestListSeparator, &pv);

    if (FAILED(hr))
    {
      IShellLinkW_Release(psl);
      return hr;
    }

    hr = IPropertyStore_Commit(pps);
    PropVariantClear(&pv);

    if (FAILED(hr))
    {
      IPropertyStore_Release(pps);
      IShellLinkW_Release(psl);
      return hr;
    }

    IPropertyStore_Release(pps);

    hr = IObjectCollection_AddObject(pThis->poc, (IUnknown*)psl);

    IShellLinkW_Release(psl);

    return hr;
}

static FORCEINLINE HRESULT CALLBACK OnCommitJumpList(HWND hwnd)
{
    HRESULT  hr;
    ICDL* pThis;
    IObjectArray* poa;

    pThis = GetInterfaceHandle(hwnd);

    if (!pThis)
    {
      return E_FAIL;
    }

    hr = IObjectArray_QueryInterface(pThis->poc, &IID_IObjectArray, &poa);

    if (FAILED(hr))
    {
      return hr;
    }

    hr = ICustomDestinationList_AddUserTasks(pThis->picdl, poa);

    if (FAILED(hr))
    {
      IObjectArray_Release(poa);

      return hr;
    }

    hr = ICustomDestinationList_CommitList(pThis->picdl);

    IObjectArray_Release(poa);
    IObjectCollection_Release(pThis->poc);

    return hr;
}

static FORCEINLINE BOOL APIPRIVATE PumpMessages(void)
{
    MSG  msg;
    BOOL fQuit = FALSE;
    SecureZeroMemory(&msg, sizeof(msg));

    while (0 < GetMessage(&msg, 0, 0, 0))
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
    }
    
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

static HWND APIPRIVATE CreateMessageWindow(ICDL* pThis)
{
    WNDCLASS wc;

    SecureZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = (HINSTANCE)&__ImageBase;
    wc.lpszClassName = c_szClassName;

    return CreateWindow(MAKEINTATOM(RegisterClass(&wc)), c_szWindowTitle, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, (HINSTANCE)&__ImageBase, (LPVOID)pThis);
}

static DWORD WINAPI CDLThread(LPVOID lpThreadParameter)
{
    HRESULT hr;
    ICDL* icdl;

    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

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
    UNREFERENCED_PARAMETER(pThis);

    (*ppTasks) = GlobalAlloc(GPTR, elems * sizeof(ITask));

    if (!(*ppTasks))
    {
      return E_OUTOFMEMORY;
    }

    return S_OK;
}

CDLAPI(HRESULT) ICDL_SetTask(ICDL* pThis, ITask* pTasks, UINT elem, LPCTSTR pszImage, LPCTSTR pszArgs, LPCTSTR pszDescription, LPCTSTR pszTitle, int nIconIndex)
{
    ITask* pTask = &pTasks[elem];

    UNREFERENCED_PARAMETER(pThis);

    pTask->pszImage       = pszImage;
    pTask->pszArgs        = pszArgs;
    pTask->pszDescription = pszDescription;
    pTask->pszTitle       = pszTitle;
    pTask->nIconIndex     = nIconIndex;

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
  return (HRESULT)SendMessage(pThis->hwnd, WM_CREATEJUMPLIST, (WPARAM)pcszAppId, 0);
}

CDLAPI(HRESULT) ICDL_AddUserTasks(ICDL* pThis, ITask* pTasks, UINT16 elems, UINT16 offset)
{
  return (HRESULT)SendMessage(pThis->hwnd, WM_ADDUSERTASKS, (WPARAM)pTasks, MAKELPARAM(offset, elems));
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

