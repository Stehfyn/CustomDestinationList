#include "CustomDestinationList.h"

#include <windows.h>
#include <windowsx.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <dwmapi.h>
#include <UxTheme.h>
#include <vsstyle.h>
#include <Vssym32.h>
#include "resource.h"
#include <ShellScalingApi.h>

#pragma comment (lib, "dwmapi")
#pragma comment (lib, "Shcore")
#pragma comment (lib, "UxTheme")
#pragma comment (lib, "Version")

#define RECTWIDTH(rc) \
        (labs(rc.right - rc.left))
#define RECTHEIGHT(rc) \
        (labs(rc.bottom - rc.top))
#define CLAMP(x, lo, hi) \
        (max(lo, min(x, hi)))

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
static void GetStartupRect(int nWidth, int nHeight, LPRECT lprc);

#if (defined NDEBUG)
int
__stdcall
_tWinMain(
  _In_ HINSTANCE hInstance,
  _In_opt_ HINSTANCE hPrevInstance,
  _In_ LPTSTR lpCmdLine,
  _In_ int nShowCmd)
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
    ICDL* pcdl;
    WNDCLASSEX wcx;
    LPCTSTR szClassAtom;
    HWND hwnd;
    MSG msg;
    BOOL fQuit;
    PWSTR pszAppId;
    STARTUPINFO si;
    RECT rc;

    if (FAILED(hr = CoInitializeEx(0, COINIT_MULTITHREADED | COINIT_SPEED_OVER_MEMORY | COINIT_DISABLE_OLE1DDE)))
    {
      ExitProcess(EXIT_FAILURE);
    }

    if (S_OK != SetProcessDpiAwareness(DPI_AWARENESS_PER_MONITOR_AWARE))
    {
      ExitProcess(EXIT_FAILURE);
    }

    if (FAILED(hr = SetCurrentProcessExplicitAppUserModelID(L"CustomDestinationList-Demo")))
    {
      ExitProcess(EXIT_FAILURE);
    }

    pszAppId = NULL;
    if (FAILED(hr = GetCurrentProcessExplicitAppUserModelID(&pszAppId)))
    {
      ExitProcess(EXIT_FAILURE);
    }

    pcdl = NULL;
    if (SUCCEEDED(hr = ICDL_BeginList(pszAppId, &pcdl)))
    {
      CoTaskMemFree(pszAppId);

      if (SUCCEEDED(hr = ICDL_BeginCategory(pcdl, _T("Custom Category"))))
      {
        int elm;

        for (elm = 0; elm < 5; ++elm)
        {
          typedef struct TASKV
          {
            LPCTSTR szImage;
            LPCTSTR szArgs;
            LPCTSTR szDescription;
            LPCTSTR szTitle;
            int nIconIndex;
          } TASKV;

          static const TASKV c_tasks[] =
          {
            {
              _T("C:\\Windows\\system32\\cmd.exe"),
              _T(" "),
              _T("Hovered1"),
              _T("Command Prompt"),
              0
            },
            {
              _T("C:\\Windows\\explorer.exe"),
              _T(" "),
              _T("Hovered2"),
              _T("File Explorer"),
              0
            },
            {
              _T("C:\\Windows\\system32\\notepad.exe"),
              _T(" "),
              _T("Hovered3"),
              _T("Notepad"),
              0
            },
            {
              _T("C:\\Windows\\regedit.exe"),
              _T(" "),
              _T("Hovered4"),
              _T("Registry Editor"),
              0
            },

            // Won't work:
            // 1. Need to know it's a shell URI
            // 2. Then use SHParseDisplayName
            //{
            //  _T("%SystemRoot%\\explorer.exe"),
            //  _T("/e,\"shell:Recent\CustomDestinations\""),
            //  _T("Open Custom Destinations in Explorer"),
            //  _T("Open Custom Destinations"),
            //  0
            //}
            
            // But this will
            {
              _T("%SystemRoot%\\system32\\cmd.exe"),
              _T("/c start explorer shell:Recent\\CustomDestinations"),
              _T("Open Custom Destinations in Explorer"),
              _T("Open Custom Destinations"),
              0
            }
          };

          if (FAILED(hr = ICDL_AddTask(pcdl,
                                       c_tasks[elm].szImage,
                                       c_tasks[elm].szArgs,
                                       c_tasks[elm].szDescription,
                                       c_tasks[elm].szTitle,
                                       c_tasks[elm].nIconIndex)))
          {
            ExitProcess(EXIT_FAILURE);
          }
        }

        if (FAILED(hr = ICDL_CommitCategory(pcdl)))
        {
          ExitProcess(EXIT_FAILURE);
        }
      }

      if (FAILED(hr = ICDL_CommitList(pcdl)))
      {
        ExitProcess(EXIT_FAILURE);
      }

      ICDL_Release(pcdl);
    }

    SecureZeroMemory(&wcx, sizeof(wcx));
    wcx.cbSize        = sizeof(wcx);
    wcx.style         = CS_BYTEALIGNCLIENT | CS_BYTEALIGNWINDOW;
    wcx.lpfnWndProc   = WndProc;
    wcx.hInstance     = (HINSTANCE)&__ImageBase;
    wcx.hbrBackground = CreateSolidBrush(RGB(43, 43, 43));
    wcx.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wcx.lpszClassName = _T("CustomDestinationList-DemoWndClass");
    wcx.hIcon         = LoadIcon((HINSTANCE)&__ImageBase, MAKEINTRESOURCE(IDI_ICON1));
    szClassAtom       = (LPCTSTR)RegisterClassEx(&wcx);

    if (!szClassAtom)
    {
      ExitProcess(EXIT_FAILURE);
    }

    GetStartupInfo(&si);
    GetStartupRect(640, 480, &rc);

    if (!(STARTF_USEPOSITION & si.dwFlags) && !si.hStdOutput)
    {
      hwnd = CreateWindowEx(
        WS_EX_APPWINDOW | WS_EX_COMPOSITED,
        szClassAtom, _T("CustomDestinationList-Demo"),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }
    else
    {
      hwnd = CreateWindowEx(
        WS_EX_APPWINDOW | WS_EX_COMPOSITED,
        szClassAtom, _T("CustomDestinationList-Demo"), 
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }

    if (!hwnd)
    {
      ExitProcess(EXIT_FAILURE);
    }

    // Win11 24H2: Default theming just tosses NONCLIENT operability out the window--
    // This is needed for maximize/restore down to __just work__, just shameful
    SetWindowTheme(hwnd, L"WINDOW", L"");

    UpdateWindow(hwnd);
    ShowWindow(hwnd, SW_SHOWDEFAULT);

    fQuit = FALSE;

    SecureZeroMemory(&msg, sizeof(msg));
    while (0 < GetMessage(&msg, 0, 0, 0))
    {
      TranslateMessage(&msg);
      DispatchMessage(&msg);

      fQuit |= (WM_QUIT == msg.message);

      if (fQuit)
        break;
    }

    ExitProcess(EXIT_SUCCESS);
}

#pragma region dontcare
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

static void EnableDarkMode(HWND hwnd)
{
  static HMODULE __uxtheme; 
  static HMODULE __user32; 

  if (!__uxtheme)
  {
    __uxtheme = GetModuleHandle(TEXT("UxTheme.dll"));
    SetPreferredAppMode = (PFNSETPREFERREDAPPMODE)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(135));
    FlushMenuTheme = (PFNFLUSHMENUTHEME)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(136));
    InvalidateAppTheme = (PFNINVALIDATEAPPTHEME)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(115));
    RefreshImmersiveColorPolicyState = (PFNREFRESHIMMERSIVECOLORPOLICYSTATE)GetProcAddress(__uxtheme, MAKEINTRESOURCEA(104));
  }

  if (!__user32)
  {
    __user32 = GetModuleHandle(TEXT("user32.dll"));
    SetWindowCompositionAttribute = (PFNSETWINDOWCOMPOSITIONATTRIBUTE)GetProcAddress(__user32, "SetWindowCompositionAttribute");
  }

  SetPreferredAppMode(PAM_ALLOWDARK);
  SetProp(hwnd, TEXT("UseImmersiveDarkModeColors"), (HANDLE)1);
  RefreshImmersiveColorPolicyState();
  InvalidateAppTheme();
  FlushMenuTheme();
  BOOL fDark = TRUE;
  WINDOWCOMPOSITIONATTRIBDATA data = { 26, &fDark, sizeof(fDark) };
  SetWindowCompositionAttribute(hwnd, &data);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;
  if (DwmDefWindowProc(hwnd, uMsg, wParam, lParam, &lResult))
  {
    return lResult;
  }
    
  switch (uMsg) {
  case WM_COMMAND:
  {
    int wmId = LOWORD(wParam);
    // Parse the menu selections:
    switch (wmId)
    {
    case IDM_EXIT:
      DestroyWindow(hwnd);
      break;
    default:
      return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    break;
  }
  case WM_NCCREATE:
  {
    EnableDarkMode(hwnd);
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

static void GetStartupRect(int nWidth, int nHeight, LPRECT lprc)
{
    RECT        rc;
    STARTUPINFO si;
    
    GetStartupInfo(&si);
    rc.left   = 0;
    rc.top    = 0;
    rc.right  = nWidth;
    rc.bottom = nHeight;

    if (STARTF_USESIZE & si.dwFlags)
    {
      rc.right  = si.dwXSize;
      rc.bottom = si.dwYSize;
    }

    if (si.hStdOutput)
    {
      MONITORINFO mi;
      mi.cbSize = sizeof(mi);

      if (GetMonitorInfo((HMONITOR)si.hStdOutput, &mi))
      {
        POINT pt;

        pt.x = (LONG)((mi.rcWork.left) + (0.5f * RECTWIDTH(mi.rcWork)) + (-0.5f * RECTWIDTH(rc)));
        pt.y = (LONG)((mi.rcWork.top) + (0.5f * RECTHEIGHT(mi.rcWork)) + (-0.5f * RECTHEIGHT(rc)));

        rc.right  = pt.x + RECTWIDTH(rc);
        rc.bottom = pt.y + RECTHEIGHT(rc);
        rc.left   = pt.x;
        rc.top    = pt.y;
      }
    }

    if (STARTF_USEPOSITION & si.dwFlags)
    {
      rc.right  = (LONG)(si.dwX + RECTWIDTH(rc));
      rc.bottom = (LONG)(si.dwY + RECTHEIGHT(rc));
      rc.left   = (LONG)si.dwX;
      rc.top    = (LONG)si.dwY;
    }
    
    if (!(STARTF_USEPOSITION & si.dwFlags) && !si.hStdOutput)
    {
        //CW_USEDEFAULT, CW_USEDEFAULT, RECTWIDTH(rc), RECTHEIGHT(rc);
    }
    else
    {
      HMONITOR hMonitor;
      MONITORINFO mi;
      mi.cbSize = sizeof(mi);

      hMonitor = MonitorFromRect(&rc, MONITOR_DEFAULTTONEAREST);

      if (hMonitor && GetMonitorInfo(hMonitor, &mi))
      {
        LONG w;
        LONG h;

        w = RECTWIDTH(rc);
        h = RECTHEIGHT(rc);

        rc.left  = CLAMP(rc.left, mi.rcWork.left + 1, mi.rcWork.right - 1);
        rc.right = rc.left + w;
        rc.right = CLAMP(rc.right, mi.rcWork.left + 1, mi.rcWork.right - 1);
        rc.left  = rc.right - w;

        rc.top    = CLAMP(rc.top, mi.rcWork.top + 1, mi.rcWork.bottom - 1);
        rc.bottom = rc.top + h;
        rc.bottom = CLAMP(rc.bottom, mi.rcWork.top + 1, mi.rcWork.bottom - 1);
        rc.top    = rc.bottom - h;
      }
      //rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc);
    }

    (*lprc) = rc;
}
#pragma endregion