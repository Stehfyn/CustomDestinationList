#include "CustomDestinationList.h"

#include <dwmapi.h>
#include <UxTheme.h>
#pragma comment (lib, "dwmapi")
#pragma comment (lib, "UxTheme")

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
    WNDCLASSEX wcx;
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

    hr = ICDL_Initialize(&icdl);

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
      typedef struct TASK
      {
        LPCTSTR szArgs;
        LPCTSTR szDescription;
        LPCTSTR szTitle;
        int nIconIndex;
      } TASK;

      static const TASK c_tasks[] =
      {
        {
          _T(" "),
          _T("Hovered1"),
          _T("Task1"),
          0
        },
        {
          _T(" "),
          _T("Hovered2"),
          _T("Task2"),
          0
        },
        {
          _T(" "),
          _T("Hovered3"),
          _T("Task3"),
          0
        },
        {
          _T(" "),
          _T("Hovered4"),
          _T("Task4"),
          0
        }
      };

      hr = ICDL_SetTask(icdl, pTasks, elm, NULL, c_tasks[elm].szArgs, c_tasks[elm].szDescription, c_tasks[elm].szTitle, c_tasks[elm].nIconIndex);

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

    CoTaskMemFree(pszAppId);

    hr = ICDL_AddSeparator(icdl);

    if (FAILED(hr))
    {
      return hr;
    }

    hr = ICDL_AddUserTasks(icdl, pTasks, 2, 0);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_AddSeparator(icdl);

    if (FAILED(hr))
    {
      return hr;
    }
    
    hr = ICDL_AddUserTasks(icdl, pTasks, 2, 2);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = ICDL_AddSeparator(icdl);

    if (FAILED(hr))
    {
      return hr;
    }

    hr = ICDL_CommitJumpList(icdl);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    SecureZeroMemory(&wcx, sizeof(wcx));
    wcx.cbSize        = sizeof(wcx);
    wcx.style         = CS_OWNDC | CS_BYTEALIGNCLIENT | CS_BYTEALIGNWINDOW;
    wcx.lpfnWndProc   = WndProc;
    wcx.hInstance     = (HINSTANCE)&__ImageBase;
    wcx.hbrBackground = CreateSolidBrush(RGB(43, 43, 43));
    wcx.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wcx.lpszClassName = _T("CustomDestinationList-DemoWndClass");
    szClassAtom       = (LPCTSTR)RegisterClassEx(&wcx);

    if (!szClassAtom)
    {
      ExitProcess(EXIT_FAILURE);
    }

    hwnd = CreateWindowEx(WS_EX_WINDOWEDGE | WS_EX_APPWINDOW, szClassAtom, _T("CustomDestinationList-Demo"), WS_OVERLAPPEDWINDOW, 0, 0, 640, 480, HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);

    if (!hwnd)
    {
      ExitProcess(EXIT_FAILURE);
    }

    hr = SetWindowTheme(hwnd, L"DarkMode_Explorer", NULL);

    if (FAILED(hr))
    {
      ExitProcess(EXIT_FAILURE);
    }

    UpdateWindow(hwnd);
    ShowWindow(hwnd, SW_SHOWDEFAULT);

    fQuit = FALSE;

    for (;;)
    {
      SecureZeroMemory(&msg, sizeof(msg));
      while (PeekMessage(&msg, 0, 0, 0, PM_REMOVE))
      {
        DispatchMessage(&msg);

        fQuit |= (WM_QUIT == msg.message);
      }

      if (fQuit)
        break;

      MsgWaitForMultipleObjects(0, 0, 0, 0, QS_ALLINPUT);
      SleepEx(USER_TIMER_MINIMUM, 0);

      WaitMessage();
    }

    ExitProcess(EXIT_SUCCESS);
}

// Affects the rendering of the background of a window. 
typedef enum tagACCENT_STATE {
  ACCENT_DISABLED = 0,  // Default value. Background is black.
  ACCENT_ENABLE_GRADIENT = 1,  // Background is GradientColor, alpha channel ignored.
  ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,	// Background is GradientColor.
  ACCENT_ENABLE_BLURBEHIND = 3,  // Background is GradientColor, with blur effect.
  ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,  // Background is GradientColor, with acrylic blur effect.
  ACCENT_ENABLE_HOSTBACKDROP = 5,  // Unknown.
  ACCENT_INVALID_STATE = 6,  // Unknown. Seems to draw background fully transparent.
} ACCENT_STATE;

// Determines how a window's background is rendered.
typedef struct tagACCENT_POLICY {
  ACCENT_STATE	AccentState;	            // Background effect.
  UINT			    AccentFlags;	            // Flags. Set to 2 to tell GradientColor is used, rest is unknown.
  COLORREF		  GradientColor;            // Background color.
  LONG			    AnimationId;	            // Unknown
} ACCENT_POLICY;

// Determines what attribute is being manipulated.
typedef enum tagWINDOWCOMPOSITIONATTRIB {
  WCA_ACCENT_POLICY = 0x13				        // The attribute being get or set is an accent policy.
} WINDOWCOMPOSITIONATTRIB;

// Options for [Get/Set]WindowCompositionAttribute.
typedef struct tagWINDOWCOMPOSITIONATTRIBDATA {
  WINDOWCOMPOSITIONATTRIB	Attrib;			    // Type of what is being get or set.
  LPVOID					        pvData;			    // Pointer to memory that will receive what is get or that contains what will be set.
  UINT					          cbData;			    // Size of the data being pointed to by pvData.
} WINDOWCOMPOSITIONATTRIBDATA;

typedef BOOL(WINAPI* PFNSETWINDOWCOMPOSITIONATTRIBUTE)(HWND, const WINDOWCOMPOSITIONATTRIBDATA*);

static PFNSETWINDOWCOMPOSITIONATTRIBUTE SetWindowCompositionAttribute;
typedef enum tagPREFERREDAPPMODE {
  PAM_DEFAULT,
  PAM_ALLOWDARK,
  PAM_FORCEDARK,
  PAM_FORCELIGHT,
  PAM_LAST
} PREFERREDAPPMODE;

typedef PREFERREDAPPMODE(WINAPI* PFNSETPREFERREDAPPMODE)(PREFERREDAPPMODE eAppMode);
typedef VOID(WINAPI* PFNFLUSHMENUTHEME)(VOID);
typedef VOID(WINAPI* PFNINVALIDATEAPPTHEME)(VOID);
typedef VOID(WINAPI* PFNREFRESHIMMERSIVECOLORPOLICYSTATE)(VOID);

static PFNSETPREFERREDAPPMODE SetPreferredAppMode;
static PFNFLUSHMENUTHEME FlushMenuTheme;
static PFNINVALIDATEAPPTHEME InvalidateAppTheme;
static PFNREFRESHIMMERSIVECOLORPOLICYSTATE RefreshImmersiveColorPolicyState;

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static BOOL fMoving = FALSE;
  static BOOL fSizing = FALSE;

  switch (uMsg) {
  case WM_NCCREATE:
  {
    HMODULE __uxtheme = GetModuleHandle(TEXT("UxTheme.dll"));
    HMODULE __user32 = GetModuleHandle(TEXT("user32.dll"));
    SetWindowCompositionAttribute = (PFNSETWINDOWCOMPOSITIONATTRIBUTE)GetProcAddress(__user32, "SetWindowCompositionAttribute");
    SetPreferredAppMode = (PFNSETPREFERREDAPPMODE)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(135));
    FlushMenuTheme = (PFNFLUSHMENUTHEME)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(136));
    InvalidateAppTheme = (PFNINVALIDATEAPPTHEME)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(115));
    RefreshImmersiveColorPolicyState = (PFNREFRESHIMMERSIVECOLORPOLICYSTATE)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(104));
    SetPreferredAppMode(PAM_ALLOWDARK);
    SetProp(hwnd, TEXT("UseImmersiveDarkModeColors"), (HANDLE)1);
    RefreshImmersiveColorPolicyState();
    InvalidateAppTheme();
    FlushMenuTheme();
    BOOL fDark = TRUE;
    WINDOWCOMPOSITIONATTRIBDATA data = { 26, &fDark, sizeof(fDark) };
    SetWindowCompositionAttribute(hwnd, &data);
    SetWindowTheme(hwnd, L"DarkMode_Explorer", 0);
    EnableNonClientDpiScaling(hwnd);
    break;
  }
  case WM_SYSCOMMAND:
  {
    fMoving = 0;
    fSizing = 0;
    switch (GET_SC_WPARAM(wParam))
    {
    case SC_SIZE:
      fSizing = 1;
      break;
    case SC_MOVE:
      fMoving = 1;
      break;
    }
    break;
  }
  case WM_NCCALCSIZE:
  {
    if (wParam)
    {
      GUITHREADINFO gti;
      gti.cbSize = sizeof(gti);

      if (GetGUIThreadInfo(GetCurrentThreadId(), &gti))
      {
        if (!(GUI_INMOVESIZE & gti.flags) || !fMoving)
        {
          DwmFlush();
        }
      }
    }
    break;
  }
  case WM_WINDOWPOSCHANGING:
  {
    ((LPWINDOWPOS)lParam)->flags |= SWP_DEFERERASE;
    ((LPWINDOWPOS)lParam)->flags &= ~SWP_DRAWFRAME;
    break;
  }
  case WM_DESTROY:
    PostQuitMessage(0);
    return 0;
  default:
    break;
  }

  return DefWindowProc(hwnd, uMsg, wParam, lParam);
}
