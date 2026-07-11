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
#include <propkey.h>
#include <propsys.h>
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
static INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
static void GetStartupRect(int nWidth, int nHeight, LPRECT lprc);

/* Dark-mode + custom-menu-bar state (defined with the theming helpers, below WndProc). */
static void ThemeInitProcess(void);
static void ThemeApplyWindow(HWND hwnd, BOOL fDark);
static void ThemeCommitFrame(HWND hwnd);
static BOOL   g_fDark;                /* current effective dark mode (follows the system setting live)  */
static HBRUSH g_hbrDark;              /* cached dark client-background brush                            */
static int    g_policy;               /* 0 = no app dark mode, 1 = Win10 1809, 2 = Win10 1903+ / Win11  */
static BOOL   g_fThemeChangePending;  /* an ImmersiveColorSet broadcast is queued for deferred handling */
static BOOL   g_fActiveItem;          /* light mode: a top-level menu item is hot/pressed              */
static RECT   g_rcActiveItem;         /* ...its rect (window coords), so the seam can spare its bottom  */
static HWND   g_hwndAbout;            /* live About box, so a system theme switch can re-dress it       */

/* Private message: a system light/dark switch, re-read on the next message-loop turn (see WndProc). */
#define WMAPP_THEMECHANGED (WM_APP + 1)

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

    if (FAILED(hr = CoInitializeEx(0, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE | COINIT_SPEED_OVER_MEMORY)))
    {
      ExitProcess(EXIT_FAILURE);
    }

    if (FAILED(SetProcessDpiAwareness(DPI_AWARENESS_PER_MONITOR_AWARE)))
    {
      ExitProcess(EXIT_FAILURE);
    }

    // Opt the process into dark mode and read the live system setting before the class is registered.
    ThemeInitProcess();

    SecureZeroMemory(&wcx, sizeof(wcx));
    wcx.cbSize        = sizeof(wcx);
    wcx.style         = CS_BYTEALIGNCLIENT | CS_BYTEALIGNWINDOW;
    wcx.lpfnWndProc   = WndProc;
    wcx.hInstance     = (HINSTANCE)&__ImageBase;
    wcx.hbrBackground = NULL; // WM_ERASEBKGND paints the themed client color
    wcx.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wcx.lpszMenuName  = MAKEINTRESOURCE(IDC_WIN32CUSTOMMENUBARAEROTHEME);
    wcx.lpszClassName = TEXT("CustomDestinationList-DemoWndClass");
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
        szClassAtom, TEXT("CustomDestinationList-Demo"),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }
    else
    {
      hwnd = CreateWindowEx(
        WS_EX_APPWINDOW | WS_EX_COMPOSITED,
        szClassAtom, TEXT("CustomDestinationList-Demo"), 
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc),
        HWND_TOP, NULL, (HINSTANCE)&__ImageBase, NULL);
    }

    if (!hwnd)
    {
      ExitProcess(EXIT_FAILURE);
    }

    // Dress the title-bar frame, the window's control sub-theme, and (via WM_UAHDRAWMENU*) the menu
    // bar to match the current system theme. Win11 24H2 also needs the DarkMode_Explorer sub-theme
    // applied here for maximize/restore-down non-client operability to work.
    ThemeApplyWindow(hwnd, g_fDark);
    ThemeCommitFrame(hwnd);  // make the first composition pick up the immersive-dark attribute

    UpdateWindow(hwnd);
    ShowWindow(hwnd, SW_SHOWDEFAULT);

    if (FAILED(hr = ICDL_SetWindowAppUserModelID(hwnd, pszAppId = L"CustomDestinationList-Demo")))
    {
      ExitProcess(EXIT_FAILURE);
    }

    pcdl = NULL;
    if (SUCCEEDED(hr = ICDL_BeginList(pszAppId, &pcdl)))
    {
      if (SUCCEEDED(hr = ICDL_BeginCategory(pcdl, TEXT("Quick Launch"))))
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
              TEXT("%SystemRoot%\\system32\\cmd.exe"),
              TEXT(" "),
              TEXT("Hovered69"),
              TEXT("Command Prompt"),
              0
            },
            {
              TEXT("%SystemRoot%\\explorer.exe"),
              TEXT(" "),
              TEXT("Hovered2"),
              TEXT("File Explorer"),
              0
            },
            {
              TEXT("%SystemRoot%\\system32\\notepad.exe"),
              TEXT(" "),
              TEXT("Hovered3"),
              TEXT("Notepad"),
              0
            },
            {
              TEXT("%SystemRoot%\\regedit.exe"),
              TEXT(" "),
              TEXT("Hovered4"),
              TEXT("Registry Editor"),
              0
            },

            // Won't work:
            // 1. Need to know it's a shell URI
            // 2. Then use SHParseDisplayName or whatever
            //{
            //  TEXT("%SystemRoot%\\explorer.exe"),
            //  TEXT("/e,\"shell:Recent\CustomDestinations\""),
            //  TEXT("Open Custom Destinations in Explorer"),
            //  TEXT("Open Custom Destinations"),
            //  0
            //}
            
            // But this will
            {
              TEXT("%SystemRoot%\\system32\\cmd.exe"),
              TEXT("/c explorer shell:Recent\\CustomDestinations"),
              TEXT("Open Custom Destinations in Explorer"),
              TEXT("Open Custom Destinations"),
              0
            }
          };

          if (FAILED(hr = ICDL_AddTask(pcdl,
                                       c_tasks[elm].szImage,
                                       c_tasks[elm].szArgs,
                                       c_tasks[elm].szDescription,
                                       c_tasks[elm].szTitle,
                                       (4 == elm) ? TEXT("%SystemRoot%\\explorer.exe") : NULL, // Make the "Open Custom Destinations" use the File Explorer icon
                                       c_tasks[elm].nIconIndex)))
          {
            ExitProcess(EXIT_FAILURE);
          }
        }
      }

      if (FAILED(hr = ICDL_CommitCategory(pcdl)))
      {
        ExitProcess(EXIT_FAILURE);
      }

      if (FAILED(hr = ICDL_CommitList(pcdl)))
      {
        ExitProcess(EXIT_FAILURE);
      }

      ICDL_Release(pcdl);
    }

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

/* ============================================================================================== *
 *  Win32 dark mode + owner-drawn ("aero theme") custom menu bar.
 *
 *  Every external function this feature needs is delay-bound through the two macros below (this file
 *  is standalone -- no external headers). DELAYLOAD binds a named export; DELAYLOAD_ORDINAL is a
 *  byte-for-byte copy that resolves the export by ordinal instead of by name, for the five
 *  undocumented, ordinal-only uxtheme dark-mode exports that a name lookup cannot reach. Each wrapper
 *  resolves its target once via GetProcAddress, caches it, and returns a caller-supplied value if the
 *  DLL or export is missing (so the .exe still loads and runs light on a pre-1809 system).
 *
 *  The FARPROC-to-typed-pointer transfer uses a union, not a cast: GetProcAddress returns FARPROC,
 *  and a cast to the concrete signature is a function-pointer reinterpret that warns at full warning
 *  strength (MSVC C4191). This is a .c file, so reading a different union member is permitted type
 *  punning (C11 6.5.2.3p3); with no cast operator present, no such warning can fire.
 * ============================================================================================== */

/* Define a delay-bound wrapper _WrapperName that resolves _DllName!_ExportName (by name) on first use
   and forwards to it. _hInst is a caller-owned HMODULE lvalue (shared across wrappers for the same
   DLL), LoadLibrary'd on demand; _ErrVal (of type _RetType) is returned if the DLL or export is
   missing. */
#define DELAYLOAD(_hInst, _DllName, _CallConv, _WrapperName, _ExportName, _Args1, _Args2, _RetType, _ErrVal) \
    typedef _RetType(_CallConv* PFN_##_WrapperName) _Args1;                                                  \
    _RetType _CallConv _WrapperName _Args1                                                                   \
    {                                                                                                       \
        static PFN_##_WrapperName s_pfn;                                                                     \
        union { FARPROC fp; PFN_##_WrapperName pfn; } u;                                                     \
        if (!s_pfn)                                                                                          \
        {                                                                                                   \
            if (!(_hInst)) { (_hInst) = LoadLibrary(_DllName); }                                             \
            if (!(_hInst)) { return (_ErrVal); }                                                             \
            u.fp = GetProcAddress((_hInst), #_ExportName);                                                   \
            if (!u.fp) { return (_ErrVal); }                                                                 \
            s_pfn = u.pfn;                                                                                   \
        }                                                                                                   \
        return s_pfn _Args2;                                                                                 \
    }

/* DELAYLOAD, but GetProcAddress resolves MAKEINTRESOURCEA(_Ordinal) rather than the stringized name. */
#define DELAYLOAD_ORDINAL(_hInst, _DllName, _CallConv, _WrapperName, _Ordinal, _Args1, _Args2, _RetType, _ErrVal) \
    typedef _RetType(_CallConv* PFN_##_WrapperName) _Args1;                                                       \
    _RetType _CallConv _WrapperName _Args1                                                                        \
    {                                                                                                            \
        static PFN_##_WrapperName s_pfn;                                                                          \
        union { FARPROC fp; PFN_##_WrapperName pfn; } u;                                                          \
        if (!s_pfn)                                                                                               \
        {                                                                                                        \
            if (!(_hInst)) { (_hInst) = LoadLibrary(_DllName); }                                                  \
            if (!(_hInst)) { return (_ErrVal); }                                                                  \
            u.fp = GetProcAddress((_hInst), MAKEINTRESOURCEA(_Ordinal));                                          \
            if (!u.fp) { return (_ErrVal); }                                                                      \
            s_pfn = u.pfn;                                                                                        \
        }                                                                                                        \
        return s_pfn _Args2;                                                                                      \
    }

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif
#define DWMWA_USE_IMMERSIVE_DARK_MODE_OLD 19  /* Win10 1809 used attribute 19 before it settled on 20 */

/* PREFERRED_APP_MODE argument to SetPreferredAppMode (ordinal 135, Win10 1903+). */
#define PAM_DEFAULT   0
#define PAM_ALLOWDARK 1

/* uxtheme dark-mode ordinals + Win10 dark-mode build boundaries. */
#define ORD_REFRESH_IMMERSIVE     104
#define ORD_SHOULD_APPS_USE_DARK  132
#define ORD_ALLOW_DARK_FOR_WINDOW 133
#define ORD_APP_MODE_135          135  /* AllowDarkModeForApp (1809) / SetPreferredAppMode (1903+) */
#define ORD_FLUSH_MENU_THEMES     136
#define BUILD_WIN10_1809 17763u
#define BUILD_WIN10_1903 18362u

static HMODULE g_hUxtheme;
static HMODULE g_hDwmapi;
static HMODULE g_hNtdll;
static HMODULE g_hAdvapi32;

/* Named exports -- delay-bound with the header's DELAYLOAD. */
DELAYLOAD(g_hNtdll,  TEXT("ntdll.dll"),   WINAPI, DlgRtlGetVersion,        RtlGetVersion,
          (OSVERSIONINFOEXW* p), (p), LONG, -1L)
DELAYLOAD(g_hDwmapi, TEXT("dwmapi.dll"),  WINAPI, DlgDwmSetWindowAttribute, DwmSetWindowAttribute,
          (HWND h, DWORD a, LPCVOID pv, DWORD cb), (h, a, pv, cb), HRESULT, E_NOTIMPL)
DELAYLOAD(g_hDwmapi, TEXT("dwmapi.dll"),  WINAPI, DlgDwmDefWindowProc,      DwmDefWindowProc,
          (HWND h, UINT m, WPARAM w, LPARAM l, LRESULT* pr), (h, m, w, l, pr), BOOL, FALSE)
DELAYLOAD(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgSetWindowTheme,       SetWindowTheme,
          (HWND h, LPCWSTR a, LPCWSTR b), (h, a, b), HRESULT, E_NOTIMPL)
DELAYLOAD(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgOpenThemeData,        OpenThemeData,
          (HWND h, LPCWSTR c), (h, c), HTHEME, NULL)
DELAYLOAD(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgCloseThemeData,       CloseThemeData,
          (HTHEME t), (t), HRESULT, E_NOTIMPL)
DELAYLOAD(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgDrawThemeTextEx,      DrawThemeTextEx,
          (HTHEME t, HDC dc, int iPart, int iState, LPCWSTR psz, int cch, DWORD fl, LPRECT prc, const DTTOPTS* po),
          (t, dc, iPart, iState, psz, cch, fl, prc, po), HRESULT, E_NOTIMPL)
DELAYLOAD(g_hAdvapi32, TEXT("advapi32.dll"), WINAPI, DlgRegOpenKeyExW,    RegOpenKeyExW,
          (HKEY h, LPCWSTR sub, DWORD opt, REGSAM sam, PHKEY pres), (h, sub, opt, sam, pres), LONG, ERROR_FILE_NOT_FOUND)
DELAYLOAD(g_hAdvapi32, TEXT("advapi32.dll"), WINAPI, DlgRegQueryValueExW, RegQueryValueExW,
          (HKEY h, LPCWSTR val, LPDWORD res, LPDWORD pType, LPBYTE pData, LPDWORD pcb),
          (h, val, res, pType, pData, pcb), LONG, ERROR_FILE_NOT_FOUND)
DELAYLOAD(g_hAdvapi32, TEXT("advapi32.dll"), WINAPI, DlgRegCloseKey,      RegCloseKey,
          (HKEY h), (h), LONG, 0)

/* Ordinal-only uxtheme dark-mode exports -- delay-bound with the local DELAYLOAD_ORDINAL. The two
   FlushMenuThemes/RefreshImmersiveColorPolicyState exports are void; they are typed BOOL here only so
   the macro's fallback branch has a value to return (the x64 ABI leaves the ignored result register
   undefined -- the callers discard it). Ordinal 135 is bound twice, once per ABI; only the wrapper
   matching the running OS is ever called. */
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgSetPreferredAppMode,   ORD_APP_MODE_135,
                  (int appMode), (appMode), int, PAM_DEFAULT)
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgAllowDarkModeForApp,   ORD_APP_MODE_135,
                  (BOOL allow), (allow), BOOL, FALSE)
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgAllowDarkModeForWindow, ORD_ALLOW_DARK_FOR_WINDOW,
                  (HWND h, BOOL allow), (h, allow), BOOL, FALSE)
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgShouldAppsUseDarkMode, ORD_SHOULD_APPS_USE_DARK,
                  (void), (), BOOL, FALSE)
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgFlushMenuThemes,       ORD_FLUSH_MENU_THEMES,
                  (void), (), BOOL, FALSE)
DELAYLOAD_ORDINAL(g_hUxtheme, TEXT("uxtheme.dll"), WINAPI, DlgRefreshImmersiveColorPolicyState, ORD_REFRESH_IMMERSIVE,
                  (void), (), BOOL, FALSE)

/* ---- undocumented "UAH" menu-bar paint messages (stable on Win10/11, in no SDK header) ---------- */
#ifndef WM_UAHDRAWMENU
#define WM_UAHDRAWMENU     0x0091  /* lParam -> UAHMENU:         paint the whole bar background */
#endif
#ifndef WM_UAHDRAWMENUITEM
#define WM_UAHDRAWMENUITEM 0x0092  /* lParam -> UAHDRAWMENUITEM:  paint one bar item             */
#endif

typedef union tagUAHMENUITEMMETRICS
{
    struct { DWORD cx; DWORD cy; } rgsizeBar[2];
    struct { DWORD cx; DWORD cy; } rgsizePopup[4];
} UAHMENUITEMMETRICS;

typedef struct tagUAHMENUPOPUPMETRICS { DWORD rgcx[4]; DWORD fUpdateMaxWidths : 2; } UAHMENUPOPUPMETRICS;

typedef struct tagUAHMENU { HMENU hmenu; HDC hdc; DWORD dwFlags; DWORD dwReserved0; } UAHMENU;

typedef struct tagUAHMENUITEM
{
    int                 iPosition;
    UAHMENUITEMMETRICS  umim;
    UAHMENUPOPUPMETRICS umpm;
} UAHMENUITEM;

typedef struct tagUAHDRAWMENUITEM { DRAWITEMSTRUCT dis; UAHMENU um; UAHMENUITEM umi; } UAHDRAWMENUITEM;

/* ---- palette -------------------------------------------------------------------------------- */
#define DARK_BG        RGB(43, 43, 43)  // matches the system's dark title-bar color
#define DARK_BG_HOT    RGB(64, 64, 64)
#define DARK_BG_PUSHED RGB(80, 80, 80)
#define DARK_TEXT      RGB(240, 240, 240)
#define DARK_TEXT_DIM  RGB(150, 150, 150)

typedef struct MENUBAR_PALETTE
{
    COLORREF clrBar;
    COLORREF clrText;
    COLORREF clrTextDim;
    COLORREF clrItemHot;
    COLORREF clrItemPushed;
} MENUBAR_PALETTE;

static void MenuBarPalette(BOOL fDark, MENUBAR_PALETTE* pPalette)
{
    if (fDark && (0 != g_policy))
    {
        pPalette->clrBar        = DARK_BG;
        pPalette->clrText       = DARK_TEXT;
        pPalette->clrTextDim    = DARK_TEXT_DIM;
        pPalette->clrItemHot    = DARK_BG_HOT;
        pPalette->clrItemPushed = DARK_BG_PUSHED;
    }
    else
    {
        /* Light, active chrome: match the window/caption surface (COLOR_WINDOW), not the darker
           legacy COLOR_MENUBAR, so the owner-drawn bar blends with the caption and client. */
        pPalette->clrBar        = GetSysColor(COLOR_WINDOW);
        pPalette->clrText       = GetSysColor(COLOR_MENUTEXT);
        pPalette->clrTextDim    = GetSysColor(COLOR_GRAYTEXT);
        pPalette->clrItemHot    = GetSysColor(COLOR_MENUHILIGHT);
        pPalette->clrItemPushed = GetSysColor(COLOR_HIGHLIGHT);
    }
}

static void ThemePaintSolidColor(HDC hdc, const RECT* prc, COLORREF cr)
{
    HGDIOBJ hbrPrev;

    hbrPrev = SelectObject(hdc, GetStockObject(DC_BRUSH));
    if (hbrPrev)
    {
        SetDCBrushColor(hdc, cr);
        (void)PatBlt(hdc, prc->left, prc->top, prc->right - prc->left, prc->bottom - prc->top, PATCOPY);
        SelectObject(hdc, hbrPrev);
    }
}

/* WM_UAHDRAWMENU: fill the whole menu-bar background. */
static void MenuBarOnDrawMenu(HWND hwnd, const UAHMENU* pUDM, const MENUBAR_PALETTE* pPalette)
{
    MENUBARINFO mbi;
    RECT        rcWindow;
    RECT        rcBar;

    SecureZeroMemory(&mbi, sizeof(mbi));
    mbi.cbSize = sizeof(mbi);
    if (!GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi))
    {
        return;
    }
    GetWindowRect(hwnd, &rcWindow);
    rcBar        = mbi.rcBar;
    OffsetRect(&rcBar, -rcWindow.left, -rcWindow.top);
    rcBar.left   = 0;
    rcBar.right  = rcWindow.right - rcWindow.left;
    rcBar.bottom = rcBar.bottom + 1;
    ThemePaintSolidColor(pUDM->hdc, &rcBar, pPalette->clrBar);
}

/* WM_UAHDRAWMENUITEM: fill one bar item by state, then draw its label with the "Menu" theme part. */
static void MenuBarOnDrawMenuItem(HWND hwnd, const UAHDRAWMENUITEM* pUDMI, const MENUBAR_PALETTE* pPalette)
{
    WCHAR         szText[64];
    MENUITEMINFOW mii;
    RECT          rcItem;
    HTHEME        hTheme;
    COLORREF      clrBg;
    COLORREF      clrText;
    UINT          uFormat;
    int           iState;
    BOOL          fHot;
    BOOL          fPushed;
    BOOL          fDisabled;

    fHot      = !!(ODS_HOTLIGHT & pUDMI->dis.itemState);
    fPushed   = !!(ODS_SELECTED & pUDMI->dis.itemState);
    fDisabled = !!((ODS_GRAYED | ODS_DISABLED) & pUDMI->dis.itemState);

    clrBg   = pPalette->clrBar;
    clrText = pPalette->clrText;
    iState  = MBI_NORMAL;
    uFormat = DT_CENTER | DT_SINGLELINE | DT_VCENTER;
    if (fHot)      { clrBg = pPalette->clrItemHot;    iState = MBI_HOT; }
    if (fPushed)   { clrBg = pPalette->clrItemPushed; iState = MBI_PUSHED; }
    if (fDisabled) { clrText = pPalette->clrTextDim;  iState = MBI_DISABLED; }
    if (ODS_NOACCEL & pUDMI->dis.itemState) { uFormat |= DT_HIDEPREFIX; }

    szText[0] = 0;
    SecureZeroMemory(&mii, sizeof(mii));
    mii.cbSize     = sizeof(mii);
    mii.fMask      = MIIM_STRING;
    mii.dwTypeData = szText;
    mii.cch        = (UINT)(ARRAYSIZE(szText) - 1);
    (void)GetMenuItemInfoW(pUDMI->um.hmenu, (UINT)pUDMI->umi.iPosition, TRUE, &mii);

    rcItem = pUDMI->dis.rcItem;
    ThemePaintSolidColor(pUDMI->um.hdc, &rcItem, clrBg);

    hTheme = DlgOpenThemeData(hwnd, L"Menu");
    if (hTheme)
    {
        DTTOPTS opts;

        SecureZeroMemory(&opts, sizeof(opts));
        opts.dwSize  = (DWORD)sizeof(opts);
        opts.dwFlags = DTT_TEXTCOLOR;
        opts.crText  = clrText;
        (void)DlgDrawThemeTextEx(hTheme, pUDMI->um.hdc, MENU_BARITEM, iState,
                                 szText, -1, uFormat, &rcItem, &opts);
        (void)DlgCloseThemeData(hTheme);
        return;
    }

    SetBkMode(pUDMI->um.hdc, TRANSPARENT);
    SetTextColor(pUDMI->um.hdc, clrText);
    DrawTextW(pUDMI->um.hdc, szText, -1, &rcItem, uFormat);
}

/* Paint the 1px seam the standard frame draws between the menu bar and the client in the bar color,
   so no light line survives under a dark bar (the win32-custom-menubar-aero-theme seam fix). */
static void MenuBarPaintSeam(HWND hwnd, const MENUBAR_PALETTE* pPalette)
{
    MENUBARINFO mbi;
    RECT        rcClient;
    RECT        rcWindow;
    RECT        rcLine;
    HDC         hdc;

    SecureZeroMemory(&mbi, sizeof(mbi));
    mbi.cbSize = sizeof(mbi);
    if (!GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi))
    {
        return;
    }
    hdc = GetWindowDC(hwnd);
    if (!hdc)
    {
        return;
    }
    GetClientRect(hwnd, &rcClient);
    MapWindowPoints(hwnd, NULL, (POINT*)&rcClient, 2);
    GetWindowRect(hwnd, &rcWindow);
    OffsetRect(&rcClient, -rcWindow.left, -rcWindow.top);

    rcLine        = rcClient;
    rcLine.bottom = rcLine.top;
    rcLine.top    = rcLine.top - 2;  // the frame's menu/client separator is 2px; a 1px cover leaves half

    // If a top-level item is hot/pressed (light mode), paint the seam in two parts so its x-range is
    // left untouched -- the native aero box keeps its own bottom edge there, while the separator stays
    // covered everywhere else. Full-width single line otherwise.
    if (g_fActiveItem)
    {
        RECT rcLeft  = rcLine;
        RECT rcRight = rcLine;
        rcLeft.right = g_rcActiveItem.left;
        rcRight.left = g_rcActiveItem.right;
        if (rcLeft.right > rcLeft.left)   { ThemePaintSolidColor(hdc, &rcLeft, pPalette->clrBar); }
        if (rcRight.right > rcRight.left) { ThemePaintSolidColor(hdc, &rcRight, pPalette->clrBar); }
    }
    else
    {
        ThemePaintSolidColor(hdc, &rcLine, pPalette->clrBar);
    }
    ReleaseDC(hwnd, hdc);
}

/* Ground-truth "does the user want dark app chrome?" -- read straight from the registry value the
   Settings app writes (HKCU\...\Personalize\AppsUseLightTheme). The Settings app writes this BEFORE
   it broadcasts WM_SETTINGCHANGE, so it is already current when we handle the broadcast. This avoids
   ShouldAppsUseDarkMode (ordinal 132), whose value is a policy snapshot that can stay stale across a
   live switch. Falls back to the ordinal export if the key can't be read. */
static BOOL ThemeSystemUsesDarkMode(void)
{
    HKEY  hKey;
    DWORD dwValue;
    DWORD dwType;
    DWORD cbValue;

    if (0 == g_policy)
    {
        return FALSE;
    }
    hKey = NULL;
    if (ERROR_SUCCESS != DlgRegOpenKeyExW(HKEY_CURRENT_USER,
                                          L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                                          0, KEY_QUERY_VALUE, &hKey))
    {
        return DlgShouldAppsUseDarkMode();
    }
    dwValue = 1;
    dwType  = 0;
    cbValue = (DWORD)sizeof(dwValue);
    if ((ERROR_SUCCESS != DlgRegQueryValueExW(hKey, L"AppsUseLightTheme", NULL, &dwType, (LPBYTE)&dwValue, &cbValue))
        || (REG_DWORD != dwType))
    {
        (void)DlgRegCloseKey(hKey);
        return DlgShouldAppsUseDarkMode();
    }
    (void)DlgRegCloseKey(hKey);
    return (0 == dwValue);  /* AppsUseLightTheme == 0 means dark */
}

/* Read the current OS build and classify the app-dark-mode policy. */
static void ThemeInitProcess(void)
{
    OSVERSIONINFOEXW osvi;
    DWORD            dwBuild;

    SecureZeroMemory(&osvi, sizeof(osvi));
    osvi.dwOSVersionInfoSize = (DWORD)sizeof(osvi);
    dwBuild = 0;
    if (DlgRtlGetVersion(&osvi) >= 0)
    {
        dwBuild = osvi.dwBuildNumber;
    }

    if (dwBuild >= BUILD_WIN10_1903)      { g_policy = 2; }
    else if (dwBuild >= BUILD_WIN10_1809) { g_policy = 1; }
    else                                  { g_policy = 0; }

    if (2 == g_policy)      { (void)DlgSetPreferredAppMode(PAM_ALLOWDARK); }
    else if (1 == g_policy) { (void)DlgAllowDarkModeForApp(TRUE); }

    (void)DlgRefreshImmersiveColorPolicyState();
    (void)DlgFlushMenuThemes();

    g_fDark   = ThemeSystemUsesDarkMode();
    g_hbrDark = CreateSolidBrush(DARK_BG);
}

/* Force the DWM caption to recomposite so it re-reads the immersive-dark attribute just set. The
   in-process DwmSetWindowAttribute only marshals the value to dwm.exe (confirmed by disassembling
   dwmapi 10.0.19041 DwmSetWindowAttribute: it validates cb>=4, accepts attribute 20, and forwards);
   uDWM applies the new caption shade on the next composition, which an already-realized window does
   NOT get from an ordinary invalidate -- only from a frame change. SWP_FRAMECHANGED delivers the
   WM_NCCALCSIZE that makes uDWM recolor the caption. (Same technique WinMainEx uses.) */
static void ThemeCommitFrame(HWND hwnd)
{
    (void)SetWindowPos(hwnd, NULL, 0, 0, 0, 0,
                       SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

/* Dress the title-bar frame + the window's control sub-theme for the given mode. */
static void ThemeApplyWindow(HWND hwnd, BOOL fDark)
{
    BOOL fEffective;

    fEffective = fDark && (0 != g_policy);
    (void)DlgAllowDarkModeForWindow(hwnd, fEffective);
    (void)DlgSetWindowTheme(hwnd, fEffective ? L"DarkMode_Explorer" : L"Explorer", NULL);
    /* Set BOTH immersive-dark attribute ids unconditionally: 20 is the current id (Win10 20H1+/Win11),
       19 the pre-20H1 id (Win10 1809/1903/1909). Whichever the running build does not recognize is
       ignored, so this one call correctly darkens the title bar across every dark-capable build. */
    (void)DlgDwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &fEffective, sizeof(fEffective));
    (void)DlgDwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, &fEffective, sizeof(fEffective));
}

/* Dress (or re-dress) the About box for the given mode: title-bar frame + control sub-theme via
   ThemeApplyWindow, the OK button's push-button sub-theme, then a full repaint so the WM_CTLCOLOR*
   hooks in AboutDlgProc re-run with the new mode. */
static void AboutDlgApplyTheme(HWND hDlg, BOOL fDark)
{
    HWND hwndOK;
    BOOL fEffective;

    fEffective = fDark && (0 != g_policy);
    ThemeApplyWindow(hDlg, fDark);
    hwndOK = GetDlgItem(hDlg, IDOK);
    if (hwndOK)
    {
        (void)DlgAllowDarkModeForWindow(hwndOK, fEffective);
        (void)DlgSetWindowTheme(hwndOK, fEffective ? L"DarkMode_Explorer" : L"Explorer", NULL);
    }
    ThemeCommitFrame(hDlg);  /* make the caption pick up the immersive-dark attribute */
    (void)RedrawWindow(hDlg, NULL, NULL, RDW_ERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
}

/* Re-read the live system theme and re-dress the whole window. Runs on the deferred WMAPP_THEMECHANGED,
   NOT inline on WM_SETTINGCHANGE: the shell broadcasts the change before it has committed the new value,
   so an immediate read races and returns the old setting. ShouldAppsUseDarkMode (ordinal 132) also reads
   a cached policy snapshot -- RefreshImmersiveColorPolicyState must run first or it stays stale. */
static void ThemeRefreshFromSystem(HWND hwnd)
{
    MENUBAR_PALETTE pal;
    BOOL            fNew;

    /* IDEMPOTENT: one system light/dark switch makes the shell re-broadcast ImmersiveColorSet several
       times over ~1s. Re-running the whole retheme (frame change + menu redraw + FlushMenuThemes) on
       each duplicate is what makes the menu text flicker. Read the (registry-truth) new value and bail
       out unless it actually changed -- so exactly one repaint happens per real switch. */
    fNew = ThemeSystemUsesDarkMode();
    if (fNew == g_fDark)
    {
        return;
    }
    g_fDark = fNew;

    (void)DlgRefreshImmersiveColorPolicyState();
    (void)DlgFlushMenuThemes();

    ThemeApplyWindow(hwnd, g_fDark);
    ThemeCommitFrame(hwnd);  /* frame change: recolors the caption AND repaints the NC menu bar + seam */
    /* Client only: coalesced invalidate (NOT RDW_UPDATENOW, which forces a synchronous extra paint each
       message). WM_ERASEBKGND picks the new brush; the frame change above already repainted the bar. */
    InvalidateRect(hwnd, NULL, TRUE);
    DrawMenuBar(hwnd);
    MenuBarPalette(g_fDark && (0 != g_policy), &pal);
    MenuBarPaintSeam(hwnd, &pal);

    /* The About box is modal but the theme switch still lands here (its modal loop dispatches the
       posted WMAPP_THEMECHANGED to this window) -- re-dress it too or it keeps the stale mode. */
    if (g_hwndAbout)
    {
        AboutDlgApplyTheme(g_hwndAbout, g_fDark);
    }
}

/* About box. Themed to match the main window: the title bar + control sub-theme via
   AboutDlgApplyTheme, and the dialog surface + static text via the WM_CTLCOLOR* hooks (the dialog
   manager lets a DLGPROC return the brush directly for those). Light mode returns FALSE everywhere
   -> stock dialog rendering. A system light/dark switch while the box is open is handled by
   ThemeRefreshFromSystem, which re-dresses the dialog registered in g_hwndAbout. */
static INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);

    switch (uMsg)
    {
    case WM_INITDIALOG:
    {
        /* Center on the owner's window rect (the dialog manager only offsets from the owner's
           client origin, so an off-center owner position would carry over). */
        HWND hwndOwner;
        RECT rcOwner;
        RECT rcDlg;

        hwndOwner = GetParent(hDlg);
        if (!hwndOwner)
        {
            hwndOwner = GetDesktopWindow();
        }
        GetWindowRect(hwndOwner, &rcOwner);
        GetWindowRect(hDlg, &rcDlg);
        (void)SetWindowPos(hDlg, NULL,
                           rcOwner.left + (LONG)(RECTWIDTH(rcOwner) - RECTWIDTH(rcDlg)) / 2,
                           rcOwner.top + (LONG)(RECTHEIGHT(rcOwner) - RECTHEIGHT(rcDlg)) / 2,
                           0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);

        g_hwndAbout = hDlg;
        AboutDlgApplyTheme(hDlg, g_fDark);
        return TRUE;
    }

    case WM_DESTROY:
        g_hwndAbout = NULL;
        break;

    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN:
        if (g_fDark && (0 != g_policy))
        {
            SetTextColor((HDC)wParam, DARK_TEXT);
            SetBkColor((HDC)wParam, DARK_BG);
            return (INT_PTR)g_hbrDark;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            EndDialog(hDlg, LOWORD(wParam));
            return TRUE;
        default:
            break;
        }
        break;

    default:
        break;
    }

    return FALSE;
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;
  MENUBAR_PALETTE pal;

  if (DlgDwmDefWindowProc(hwnd, uMsg, wParam, lParam, &lResult))
  {
    return lResult;
  }

  switch (uMsg) {
  case WM_NCCREATE:
    ThemeApplyWindow(hwnd, g_fDark);
    break;

  case WM_UAHDRAWMENU:
    // Owner-draw the bar dark ONLY in dark mode. In light mode fall through to DefWindowProc, which
    // renders the native Win10 aero menu bar -- the canonical light look (correct aero blue hover).
    g_fActiveItem = FALSE;  // reset on a full bar repaint; hot item (if any) re-set below
    if (!(g_fDark && (0 != g_policy))) { break; }
    MenuBarPalette(TRUE, &pal);
    MenuBarOnDrawMenu(hwnd, (const UAHMENU*)lParam, &pal);
    return TRUE;

  case WM_UAHDRAWMENUITEM:
    if (g_fDark && (0 != g_policy))
    {
      MenuBarPalette(TRUE, &pal);
      MenuBarOnDrawMenuItem(hwnd, (const UAHDRAWMENUITEM*)lParam, &pal);
      return TRUE;
    }
    // Light: let DefWindowProc draw the native aero box, but remember the hot/pressed item's rect so
    // MenuBarPaintSeam can spare its bottom edge (the 2px seam would otherwise clip the box border).
    {
      const UAHDRAWMENUITEM* pmi = (const UAHDRAWMENUITEM*)lParam;
      if (pmi->dis.itemState & (ODS_HOTLIGHT | ODS_SELECTED))
      {
        g_rcActiveItem = pmi->dis.rcItem;
        g_fActiveItem  = TRUE;
      }
      else if (g_fActiveItem && EqualRect(&g_rcActiveItem, &pmi->dis.rcItem))
      {
        g_fActiveItem = FALSE;  // this item just went back to normal (hover-out)
      }
    }
    break;

  case WM_NCACTIVATE:
  case WM_NCPAINT:
    // Let the standard frame render, then paint over the 1px separator it draws between the menu bar
    // and the client -- in BOTH themes, using the current bar color (dark fill in dark mode,
    // COLOR_WINDOW in light mode) -- so no line shows in either theme.
    lResult = DefWindowProc(hwnd, uMsg, wParam, lParam);
    MenuBarPalette(g_fDark && (0 != g_policy), &pal);
    MenuBarPaintSeam(hwnd, &pal);
    return lResult;

  case WM_ERASEBKGND:
  {
    RECT rc;
    GetClientRect(hwnd, &rc);
    FillRect((HDC)wParam, &rc, g_fDark ? g_hbrDark : GetSysColorBrush(COLOR_WINDOW));
    return 1;
  }

  case WM_SETTINGCHANGE:
    // The shell broadcasts "ImmersiveColorSet" when the user flips the system light/dark setting.
    // This is an ANSI (MultiByte) window, so USER32 delivers lParam as an ANSI string -- compare with
    // the TCHAR-mapped lstrcmpi/TEXT() (NOT lstrcmpiW/L"...", which would read the ANSI bytes as wide
    // and never match, silently dropping every theme change). Defer the re-read (coalesced) to the
    // next loop turn so it runs after the broadcast returns.
    if (lParam && (0 == lstrcmpi((LPCTSTR)lParam, TEXT("ImmersiveColorSet"))) && !g_fThemeChangePending)
    {
      g_fThemeChangePending = TRUE;
      PostMessage(hwnd, WMAPP_THEMECHANGED, 0, 0);
    }
    break;

  // NOTE: WM_THEMECHANGED is deliberately NOT handled. Our own SetWindowTheme (in ThemeApplyWindow)
  // raises it, and the system raises several during a switch; handling it (e.g. DrawMenuBar) adds
  // redundant menu repaints -> flicker. The single ImmersiveColorSet path below drives the retheme.

  case WMAPP_THEMECHANGED:
    g_fThemeChangePending = FALSE;
    ThemeRefreshFromSystem(hwnd);
    return 0;

  case WM_COMMAND:
    switch (LOWORD(wParam)) {
    case IDM_EXIT:
      DestroyWindow(hwnd);
      return 0;
    case IDM_ABOUT:
      DialogBox((HINSTANCE)&__ImageBase, MAKEINTRESOURCE(IDD_ABOUTBOX), hwnd, AboutDlgProc);
      return 0;
    default:
      break;
    }
    break;

  case WM_DESTROY:
    ICDL_ClearWindowAppUserModelID(hwnd);
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