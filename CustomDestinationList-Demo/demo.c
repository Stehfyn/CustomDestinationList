#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <tchar.h>

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

    SuspendThread(GetCurrentThread());
}