// dllmain.cpp: DllMain 的实现。

#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "NativePPTAddin_i.h"
#include "dllmain.h"

CNativePPTAddinModule _AtlModule;

HINSTANCE g_hInstance;

// DLL 入口点
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    g_hInstance = hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved);
}


