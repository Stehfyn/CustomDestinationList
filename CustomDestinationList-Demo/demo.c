#include "CustomDestinationList.h"
//#include "ImmersiveFlyouts.h"

#include <dwmapi.h>
#include <UxTheme.h>
#include <vsstyle.h>
#include <Vssym32.h>
#include "resource.h"

#pragma comment (lib, "dwmapi")
#pragma comment (lib, "UxTheme")
#pragma comment (lib, "Version")

//#pragma comment (lib, "user32")
//#pragma comment (lib, "windowsapp")

#define RECTWIDTH(rc) \
        (labs(rc.right - rc.left))
#define RECTHEIGHT(rc) \
        (labs(rc.bottom - rc.top))
#define CLAMP(x, lo, hi) \
        (max(lo, min(x, hi)))

EXTERN_C IMAGE_DOS_HEADER __ImageBase;
#define HINST_THISCOMPONENT ((HINSTANCE)&__ImageBase)

#define GetHeaders(base) \
        ((PIMAGE_NT_HEADERS)((CONST LPBYTE)(base) + (base)->e_lfanew))

#define GetSubsystem(base) \
        ((GetHeaders((base)))->OptionalHeader.Subsystem)

#define WS_SERVERSIDEPROC 0x0204

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
static BOOL UAHWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* lr);
static HWND g_hWndListView = NULL;

static void GetStartupRect(int nWidth, int nHeight, LPRECT lprc);
static void da(void);
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
    STARTUPINFO si;
    POINT    pt;
    RECT     rc;
    int      elm;
    WORD sub;
    hr = CoInitializeEx(0,
      COINIT_APARTMENTTHREADED |
      COINIT_SPEED_OVER_MEMORY |
      COINIT_DISABLE_OLE1DDE
    );

    sub = GetSubsystem(&__ImageBase);
    da();
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
        LPCTSTR szImage;
        LPCTSTR szArgs;
        LPCTSTR szDescription;
        LPCTSTR szTitle;
        int nIconIndex;
      } TASK;

      static const TASK c_tasks[] =
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
        }
      };

      hr = ICDL_SetTask(icdl, pTasks, elm, c_tasks[elm].szImage, c_tasks[elm].szArgs, c_tasks[elm].szDescription, c_tasks[elm].szTitle, c_tasks[elm].nIconIndex);

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
    wcx.lpszMenuName  = MAKEINTRESOURCEW(IDC_WIN32CUSTOMMENUBARAEROTHEME);
    wcx.lpszClassName = _T("CustomDestinationList-DemoWndClass");
    wcx.hIcon         = LoadIcon((HINSTANCE)&__ImageBase, MAKEINTRESOURCE(101));
    szClassAtom       = (LPCTSTR)RegisterClassEx(&wcx);

    if (!szClassAtom)
    {
      ExitProcess(EXIT_FAILURE);
    }

    GetStartupInfo(&si);
    GetStartupRect(640, 480, &rc);

    wprintf_s(L"__app_type: %d", _query_app_type());

    if (!(STARTF_USEPOSITION & si.dwFlags) && !si.hStdOutput)
    {
      hwnd = CreateWindowEx(
        WS_EX_WINDOWEDGE | WS_EX_APPWINDOW,
        szClassAtom, _T("CustomDestinationList-Demo"),
        WS_OVERLAPPEDWINDOW
        | WS_SERVERSIDEPROC
        ,

        CW_USEDEFAULT, CW_USEDEFAULT, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }
    else
    {
      hwnd = CreateWindowEx(
        WS_EX_WINDOWEDGE | WS_EX_APPWINDOW,
        szClassAtom, _T("CustomDestinationList-Demo"), 
        WS_OVERLAPPEDWINDOW
        | WS_SERVERSIDEPROC
        ,
        rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }

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

static void MakeYummy(HWND hwnd)
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
  static BOOL fMoving = FALSE;
  static BOOL fSizing = FALSE;
  LRESULT lr = 0;

  if (UAHWndProc(hwnd, uMsg, wParam, lParam, &lr))
  {
    return lr;
  }

  switch (uMsg) {
  case WM_COMMAND:
  {
    int wmId = LOWORD(wParam);
    // Parse the menu selections:
    switch (wmId)
    {
    case ID_ANOTHER_ONE:
    {
      break;
    }
    case IDM_EXIT:
      DestroyWindow(hwnd);
      break;
    default:
      return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
  }
  break;
  case WM_NCCREATE:
  {
    MakeYummy(hwnd);
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
  case WM_NCACTIVATE:
  {
    if (wParam)
      wParam = 0;
    break;
  }
  case WM_WINDOWPOSCHANGING:
  {
    //((LPWINDOWPOS)lParam)->flags |= SWP_DEFERERASE;
    //((LPWINDOWPOS)lParam)->flags &= ~SWP_DRAWFRAME;
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

//static HBRUSH g_brBarBackground = CreateSolidBrush(RGB(0x80, 0x80, 0xFF));
// window messages related to menu bar drawing
#define WM_UAHDESTROYWINDOW    0x0090	// handled by DefWindowProc
#define WM_UAHDRAWMENU         0x0091	// lParam is UAHMENU
#define WM_UAHDRAWMENUITEM     0x0092	// lParam is UAHDRAWMENUITEM
#define WM_UAHINITMENU         0x0093	// handled by DefWindowProc
#define WM_UAHMEASUREMENUITEM  0x0094	// lParam is UAHMEASUREMENUITEM
#define WM_UAHNCPAINTMENUPOPUP 0x0095	// handled by DefWindowProc

// describes the sizes of the menu bar or menu item
typedef union tagUAHMENUITEMMETRICS
{
  // cx appears to be 14 / 0xE less than rcItem's width!
  // cy 0x14 seems stable, i wonder if it is 4 less than rcItem's height which is always 24 atm
  struct {
    DWORD cx;
    DWORD cy;
  } rgsizeBar[2];
  struct {
    DWORD cx;
    DWORD cy;
  } rgsizePopup[4];
} UAHMENUITEMMETRICS;

// not really used in our case but part of the other structures
typedef struct tagUAHMENUPOPUPMETRICS
{
  DWORD rgcx[4];
  DWORD fUpdateMaxWidths : 2; // from kernel symbols, padded to full dword
} UAHMENUPOPUPMETRICS;

// hmenu is the main window menu; hdc is the context to draw in
typedef struct tagUAHMENU
{
  HMENU hmenu;
  HDC hdc;
  DWORD dwFlags; // no idea what these mean, in my testing it's either 0x00000a00 or sometimes 0x00000a10
} UAHMENU;

// menu items are always referred to by iPosition here
typedef struct tagUAHMENUITEM
{
  int iPosition; // 0-based position of menu item in menubar
  UAHMENUITEMMETRICS umim;
  UAHMENUPOPUPMETRICS umpm;
} UAHMENUITEM;

// the DRAWITEMSTRUCT contains the states of the menu items, as well as
// the position index of the item in the menu, which is duplicated in
// the UAHMENUITEM's iPosition as well
typedef struct UAHDRAWMENUITEM
{
  DRAWITEMSTRUCT dis; // itemID looks uninitialized
  UAHMENU um;
  UAHMENUITEM umi;
} UAHDRAWMENUITEM;

// the MEASUREITEMSTRUCT is intended to be filled with the size of the item
// height appears to be ignored, but width can be modified
typedef struct tagUAHMEASUREMENUITEM
{
  MEASUREITEMSTRUCT mis;
  UAHMENU um;
  UAHMENUITEM umi;
} UAHMEASUREMENUITEM;
void UAHDrawMenuNCBottomLine(HWND hWnd)
{
    MENUBARINFO mbi = { sizeof(mbi) };
    if (!GetMenuBarInfo(hWnd, OBJID_MENU, 0, &mbi))
    {
        return;
    }

    RECT rcClient = { 0 };
    GetClientRect(hWnd, &rcClient);
    MapWindowPoints(hWnd, NULL, (POINT*)&rcClient, 2);

    RECT rcWindow = { 0 };
    GetWindowRect(hWnd, &rcWindow);

    OffsetRect(&rcClient, -rcWindow.left, -rcWindow.top);

    // the rcBar is offset by the window rect
    RECT rcAnnoyingLine = rcClient;
    rcAnnoyingLine.bottom = rcAnnoyingLine.top;
    rcAnnoyingLine.top--;

    
    HDC hdc = GetWindowDC(hWnd);
    HBRUSH hbr = CreateSolidBrush(RGB(43, 43, 43));
    FillRect(hdc, &rcAnnoyingLine, hbr);
    DeleteBrush(hbr);
    ReleaseDC(hWnd, hdc);
}

static HTHEME g_menuTheme = NULL;
static HBRUSH g_brItemBackground;
static HBRUSH g_brItemBackgroundHot;
static HBRUSH g_brItemBackgroundSelected;
static HBRUSH g_brItemBorder;
// processes messages related to UAH / custom menubar drawing.
// return true if handled, false to continue with normal processing in your wndproc
static BOOL UAHWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* lr)
{
    switch (message)
    {
    case WM_UAHDRAWMENU:
    {
        UAHMENU* pUDM = (UAHMENU*)lParam;
        RECT rc = { 0 };

        // get the menubar rect
        {
            MENUBARINFO mbi = { sizeof(mbi) };
            GetMenuBarInfo(hWnd, OBJID_MENU, 0, &mbi);

            RECT rcWindow;
            GetWindowRect(hWnd, &rcWindow);

            // the rcBar is offset by the window rect
            rc = mbi.rcBar;
            OffsetRect(&rc, -rcWindow.left, -rcWindow.top);
        }
        
        HBRUSH hbr = CreateSolidBrush(RGB(43, 43, 43));
        FillRect(pUDM->hdc, &rc, hbr);
        DeleteBrush(hbr);

        return 1;
    }
    case WM_UAHDRAWMENUITEM:
    {
        UAHDRAWMENUITEM* pUDMI = (UAHDRAWMENUITEM*)lParam;

        // ugly colors for illustration purposes


        HBRUSH* pbrBackground = &g_brItemBackground;
        HBRUSH* pbrBorder = &g_brItemBackground;

        // get the menu item string
        wchar_t menuString[256] = { 0 };
        MENUITEMINFO mii = { sizeof(mii), MIIM_STRING };
        {
            mii.dwTypeData = menuString;
            mii.cch = (sizeof(menuString) / 2) - 1;

            GetMenuItemInfo(pUDMI->um.hmenu, pUDMI->umi.iPosition, TRUE, &mii);
        }

        // get the item state for drawing

        DWORD dwFlags = DT_CENTER | DT_SINGLELINE | DT_VCENTER;

        int iTextStateID = 0;
        int iBackgroundStateID = 0;
        {
            if ((pUDMI->dis.itemState & ODS_INACTIVE) || (pUDMI->dis.itemState & ODS_DEFAULT)) {
                // normal display
                iTextStateID = MBI_NORMAL;
                iBackgroundStateID = MBI_NORMAL;
            }
            if ((pUDMI->dis.itemState & ODS_HOTLIGHT) && !(pUDMI->dis.itemState & ODS_INACTIVE)) {
                // hot tracking
                iTextStateID = MBI_HOT;
                iBackgroundStateID = MBI_HOT;

                pbrBackground = &g_brItemBackgroundHot;
                pbrBorder = &g_brItemBorder;
            }
            if (pUDMI->dis.itemState & ODS_SELECTED) {
                // clicked
                iTextStateID = MBI_PUSHED;
                iBackgroundStateID = MBI_PUSHED;

                pbrBackground = &g_brItemBackgroundSelected;
                pbrBorder = &g_brItemBorder;
            }
            if ((pUDMI->dis.itemState & ODS_GRAYED) || (pUDMI->dis.itemState & ODS_DISABLED) || (pUDMI->dis.itemState & ODS_INACTIVE)) {
                // disabled / grey text / inactive
                iTextStateID = MBI_DISABLED;
                iBackgroundStateID = MBI_DISABLED;
            }
            if (pUDMI->dis.itemState & ODS_NOACCEL) {
                dwFlags |= DT_HIDEPREFIX;
            }
        }

        if (!g_menuTheme) {
            g_menuTheme = OpenThemeData(hWnd, L"Menu");
            g_brItemBackground = CreateSolidBrush(RGB(43, 43, 43));
            g_brItemBackgroundHot = CreateSolidBrush(RGB(43, 43, 43));
            g_brItemBackgroundSelected = CreateSolidBrush(RGB(0xE0, 0xE0, 0xFF));
            g_brItemBorder = CreateSolidBrush(RGB(0xB0, 0xB0, 0xFF));
        }

        DTTOPTS opts = { sizeof(opts), DTT_TEXTCOLOR, iTextStateID != MBI_DISABLED ? RGB(0xFF, 0xFF, 0xFF) : RGB(0x40, 0x40, 0x40) };

        FillRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, *pbrBackground);
        FrameRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, *pbrBorder);
        DrawThemeTextEx(g_menuTheme, pUDMI->um.hdc, MENU_BARITEM, MBI_NORMAL, menuString, mii.cch, dwFlags, &pUDMI->dis.rcItem, &opts);

        return 1;
    }
    case WM_UAHMEASUREMENUITEM:
    {
        UAHMEASUREMENUITEM* pMmi = (UAHMEASUREMENUITEM*)lParam;

        // allow the default window procedure to handle the message
        // since we don't really care about changing the width
        *lr = DefWindowProc(hWnd, message, wParam, lParam);

        // but we can modify it here to make it 1/3rd wider for example
        pMmi->mis.itemWidth = (pMmi->mis.itemWidth * 4) / 3;

        return 1;
    }
    case WM_THEMECHANGED:
    {
        if (g_menuTheme) {
            CloseThemeData(g_menuTheme);
            g_menuTheme = NULL;
        }
        // continue processing in main wndproc
        return 0;
    }
    case WM_NCPAINT:
    case WM_NCACTIVATE:
        *lr = DefWindowProc(hWnd, message, wParam, lParam);
        UAHDrawMenuNCBottomLine(hWnd);
        return 1;
        break;
    default:
        return 0;
    }
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

static void da(void)
{
  VS_FIXEDFILEINFO* pFixedInfo = NULL;
  UINT len = 0;

  TCHAR szImageName[MAX_PATH];
  DWORD cchImageName = _countof(szImageName);
  if (!QueryFullProcessImageName(GetCurrentProcess(), 0, szImageName, &cchImageName))
    __debugbreak();

  DWORD old;
  BOOL fSuccess;
  HRSRC hResInfo = FindResourceW(HINST_THISCOMPONENT, MAKEINTRESOURCE(1), RT_VERSION);
  HGLOBAL hResData = LoadResource(HINST_THISCOMPONENT, hResInfo);
  void* pRes = LockResource(hResData);
  fSuccess = VirtualProtect(pFixedInfo, sizeof(VS_FIXEDFILEINFO), PAGE_EXECUTE_READWRITE, &old);

  fSuccess = VerQueryValueW(pRes, L"\\", (LPVOID*)(&pFixedInfo), &len);

  VS_FIXEDFILEINFO fixed = *pFixedInfo;
  fixed.dwFileVersionMS = (2 << 16) | 0;  // 2.0
  fixed.dwFileVersionLS = (0 << 16) | 0;  // .0

  HANDLE hRes = BeginUpdateResourceW(szImageName, FALSE);
  fSuccess = UpdateResourceW(hRes, RT_VERSION, MAKEINTRESOURCE(1), MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL),
    &fixed, sizeof(fixed));
  fSuccess = EndUpdateResourceW(hRes, FALSE);
}