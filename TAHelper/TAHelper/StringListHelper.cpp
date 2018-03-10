#include <string>
#include <string.h>
#include <windows.h> // DllMain support
using namespace std;

const wchar_t Ellipsis = 0x2026; //8239

//int __stdcall DllMain(void* hDllHandle, long nReason, void* Reserved)
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved)
{
   int bSuccess = 1;

   ////  Perform global initialization.
   //switch (nReason)
   //{
   //case DLL_PROCESS_ATTACH:
   //   //  For optimization.
   //   DisableThreadLibraryCalls(hDllHandle);
   //   break;

   //case DLL_PROCESS_DETACH:
   //   break;
   //}

   //  Perform type-specific initialization.
   // TODO?

   return bSuccess;

}



extern "C" {
   __declspec(dllexport)
      wstring* _cdecl AddEllipses(std::wstring*  words, int n) {
      for (int i = 0; i < n; i++) {
         wstring word = words[i];
         word.append(&Ellipsis);
      }

      return words;
   }

   __declspec(dllexport)
      void _cdecl AddEllipses1(wchar_t* word) {
      size_t len = wcsnlen(word, 1024);
      word[len] = Ellipsis;
      word[len + 1] = 0;

      //wstring ws(word);
      //ws.append(&Ellipsis);
      //size_t size = (size_t)ws.size();
      //memcpy(word, ws.c_str(), (size+1)*sizeof(wchar_t));
   }

   __declspec(dllexport)
      int _cdecl Add5(int val) {
      return val + 5;
   }
}