// pch.h: 这是预编译标头文件。
// 下方列出的文件仅编译一次，提高了将来生成的生成性能。
// 这还将影响 IntelliSense 性能，包括代码完成和许多代码浏览功能。
// 但是，如果此处列出的文件中的任何一个在生成之间有更新，它们全部都将被重新编译。
// 请勿在此处添加要频繁更新的文件，这将使得性能优势无效。

#ifndef PCH_H
#define PCH_H

// 添加要在此处预编译的标头
#include "framework.h"
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" \
auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")

// Office type library (i.e., mso.dll).
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
    auto_rename auto_search raw_interfaces_only rename_namespace("Office")
using namespace AddinDesign;
using namespace Office;

#pragma comment(lib, "gdiplus.lib")

#include "UIlib.h"
using namespace DuiLib;
#ifdef _DEBUG
#   ifdef _UNICODE
#       pragma comment(lib, "DuiLib_d.lib")
#   else
#       pragma comment(lib, "DuiLibA_d.lib")
#   endif
#else
#   ifdef _UNICODE
#       pragma comment(lib, "DuiLib.lib")
#   else
#       pragma comment(lib, "DuiLibA.lib")
#   endif
#endif

extern HINSTANCE GetApplicationHInstance();

#endif //PCH_H
