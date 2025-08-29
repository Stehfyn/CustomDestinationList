#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <tchar.h>
#include <objectarray.h>
#include <shobjidl.h>
#include <propkey.h>
#include <propvarutil.h>
#include <knownfolders.h>
#include <shlobj.h>
#include "CustomDestinationList.h"

#define EXIT_SUCCESS      (0)
#define EXIT_FAILURE      (1)

#if (defined NDEBUG)
int
_tWinMain(
  _In_ HINSTANCE hInstance,
  _In_opt_ HINSTANCE hPrevInstance,
  _In_ LPWSTR lpCmdLine,
  _In_ int nShowCmd
)
{
  UNREFERENCED_PARAMETER(hInstance);
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);
  UNREFERENCED_PARAMETER(nShowCmd);
#else
void _tmain(void)
{
#endif
    HRESULT hr;
    ICDL* icdl;

    hr = ICDL_Initialize(&icdl);

    if (FAILED(hr))
    {
      return EXIT_FAILURE;
    }
    PWSTR pszAppId = 0;
    hr = SetCurrentProcessExplicitAppUserModelID(L"Yum");
    hr = GetCurrentProcessExplicitAppUserModelID(&pszAppId);
    hr = ICDL_CreateJumpList(icdl, pszAppId);
    hr = ICDL_AddUserTasks(icdl);
    hr = ICDL_AddSeparator(icdl);
    WNDCLASS wc;
    ATOM    atm;

    SecureZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = DefWindowProc;
    wc.hInstance     = GetModuleHandle( NULL );
    wc.lpszClassName = L"wow";
    
    atm = RegisterClass(&wc);

    if (!atm)
    {
      return NULL;
    }

    HWND hwnd = CreateWindow(atm, L"Demo", WS_OVERLAPPEDWINDOW, 0, 0, 640, 480, HWND_TOP, NULL, GetModuleHandle( NULL ), NULL);

    MSG  msg;
    BOOL fQuit = FALSE;
    SecureZeroMemory(&msg, sizeof(msg));

    UpdateWindow(hwnd);
    ShowWindow(hwnd, SW_SHOWDEFAULT);

    for (;;)
    {
      while (PeekMessage(&msg, 0, 0, 0, PM_REMOVE))
      {
        TranslateMessage(&msg);
        DispatchMessage(&msg);

        fQuit |= (WM_QUIT == msg.message);
      }

      if (fQuit)
        break;
    }

    ExitProcess(EXIT_SUCCESS);
}