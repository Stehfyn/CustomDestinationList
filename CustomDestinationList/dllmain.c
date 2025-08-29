#include "CustomDestinationList.h"

#include <windowsx.h>

#include <psapi.h>
#include <shlwapi.h>
#include <strsafe.h>

#include <objectarray.h>
#include <shobjidl.h>
#include <propkey.h>
#include <propvarutil.h>
#include <knownfolders.h>
#include <shlobj.h>
#include <tchar.h>

#define EXIT_SUCCESS      (0)
#define EXIT_FAILURE      (1)

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

struct CDL
{
  HANDLE hThread;
  DWORD  dwThreadId;
  HWND   hwnd;
  ICustomDestinationList* picdl;
};

EXTERN_C TCHAR c_szClassName[]   = _T("CDLWndClass");
EXTERN_C TCHAR c_szWindowTitle[] = _T("CDLWindow");

static FORCEINLINE BOOL APIPRIVATE PumpMessages(void)
{
    MSG  msg;
    BOOL fQuit = FALSE;
    SecureZeroMemory(&msg, sizeof(msg));

    while (GetMessage(&msg, 0, 0, 0))
    {
      TranslateMessage(&msg);
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
    HANDLE_MSG(hwnd, WM_NCCREATE, OnNCCreate);
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

    while (PumpMessages());

    return EXIT_SUCCESS;
}

static HRESULT APIPRIVATE Initialize(ICDL* icdl)
{
    icdl->hThread = CreateThread(0, 0, CDLThread, icdl, CREATE_SUSPENDED, &icdl->dwThreadId);

    if (!icdl->hThread)
    {
      return E_FAIL;
    }

    ResumeThread(icdl->hThread);

    return S_OK;
}

CDLAPI(HRESULT) ICDL_Initialize(ICDL** picdl)
{
  (*picdl) = GlobalAlloc(GPTR, sizeof(ICDL));

  if (!(*picdl))
  {
    return E_OUTOFMEMORY;
  }

  return Initialize((*picdl));
}

CDLAPI(HRESULT) ICDL_AddTask(ICDL* icdl)
{
  HRESULT hr;

  hr = CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &icdl->picdl);

  return hr;
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

