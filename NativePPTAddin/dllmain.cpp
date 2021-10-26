// dllmain.cpp: DllMain 的实现。

#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "NativePPTAddin_i.h"
#include "dllmain.h"

CNativePPTAddinModule _AtlModule;

HINSTANCE g_hInstance;
HWND g_hWnd;

namespace {
    struct _TempHandleData {
        unsigned long processId;
        HWND windowHandle;
    };
   
    BOOL isMainWindow(HWND handle)
    {
        return ::GetWindow(handle, GW_OWNER) == (HWND)0 && ::IsWindowVisible(handle);
    }

    BOOL CALLBACK enumWindowsCallback(HWND handle, LPARAM lParam)
    {
        _TempHandleData& data = *(_TempHandleData*)lParam;
        unsigned long process_id = 0;
        ::GetWindowThreadProcessId(handle, &process_id);
        if (data.processId != process_id || !isMainWindow(handle))
            return TRUE;
        data.windowHandle = handle;
        return FALSE;
    }

    HWND GetMainWindow(unsigned long process_id)
    {
        _TempHandleData data;
        data.processId = process_id;
        data.windowHandle = 0;
        ::EnumWindows(enumWindowsCallback, (LPARAM)&data);
        return data.windowHandle;
    }
}

// DLL 入口点
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    g_hInstance = hInstance;
    g_hWnd = GetMainWindow(::GetCurrentProcessId());
    return _AtlModule.DllMain(dwReason, lpReserved);
}


