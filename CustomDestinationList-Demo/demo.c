#include "CustomDestinationList.h"

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

#if (defined NDEBUG)
int
_tWinMain(
  _In_ HINSTANCE hInstance,
  _In_opt_ HINSTANCE hPrevInstance,
  _In_ LPTSTR lpCmdLine,
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
    HRESULT  hr;
    ICDL*    icdl;
    WNDCLASS wc;
    LPCTSTR  szClassAtom;
    HWND     hwnd;
    MSG      msg;
    BOOL     fQuit;
    PWSTR    pszAppId;
    ITask*   pTasks;
    int      elm;

    hr = CoInitializeEx(0,
      COINIT_APARTMENTTHREADED |
      COINIT_SPEED_OVER_MEMORY |
      COINIT_DISABLE_OLE1DDE
    );

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    if (!SetProcessDpiAwarenessContext(
      DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_Initialize(&icdl);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = SetCurrentProcessExplicitAppUserModelID(L"CustomDestinationList-Demo");

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    pszAppId = NULL;
    hr = GetCurrentProcessExplicitAppUserModelID(&pszAppId);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_CreateTaskList(icdl, 4, &pTasks);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    for (elm = 0; elm < 4; ++elm)
    {
      ICDL_TASK_PARAMS params;
      params.nIconIndex    = 0;
      params.szArgs        = _T("");
      params.szDescription = _T("Yerrrrp");
      params.szTitle       = _T("Yuh");

      hr = ICDL_SetTask(icdl, pTasks, elm, &params);

      if (FAILED(hr))
      {
        ExitProcess(EXIT_FAILURE);
      }
    }

    hr = ICDL_CreateJumpList(icdl, pszAppId);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_AddUserTasks(icdl, pTasks, 4);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_AddSeparator(icdl);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    SecureZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = (HINSTANCE)&__ImageBase;
    wc.lpszClassName = _T("CustomDestinationList-DemoWndClass");
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor       = LoadCursor(NULL, OCR_NORMAL);
    
    szClassAtom = (LPCTSTR)RegisterClass(&wc);

    if (!szClassAtom)
    {
      ExitProcess(EXIT_FAILURE);
    }

    hwnd = CreateWindow(szClassAtom, _T("CustomDestinationList-Demo"), WS_OVERLAPPEDWINDOW, 0, 0, 640, 480, HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);

    if (!hwnd)
    {
      ExitProcess(EXIT_FAILURE);
    }

    SecureZeroMemory(&msg, sizeof(msg));

    UpdateWindow(hwnd);
    ShowWindow(hwnd, SW_SHOWDEFAULT);

    fQuit = FALSE;

    for (;;)
    {
      while (PeekMessage(&msg, 0, 0, 0, PM_REMOVE))
      {
        DispatchMessage(&msg);

        fQuit |= (WM_QUIT == msg.message);
      }

      if (fQuit)
        break;

      MsgWaitForMultipleObjects(0, 0, 0, USER_TIMER_MINIMUM, QS_INPUT);
    }

    ExitProcess(EXIT_SUCCESS);
}

__declspec(dllimport) EXTERN_C BOOL DwmFlush(void);
#pragma comment (lib, "Dwmapi")

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  switch (uMsg) {
  case WM_PAINT:
  {
    LRESULT lResult;

    GdiFlush();
    DwmFlush();
    
    lResult = DefWindowProc(hwnd, uMsg, wParam, lParam);
    GdiFlush();
    
    return lResult;
  }
  case WM_DESTROY:
    PostQuitMessage(0);
    return 0;
  default:
    break;
  }

  return DefWindowProc(hwnd, uMsg, wParam, lParam);
}
