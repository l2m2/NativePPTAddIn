// Connect.cpp: CConnect 的实现

#include "pch.h"
#include "Connect.h"
#include <gdiplus.h>
#include "LoginDialog.h"

namespace {
    HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
    {
        HMODULE hModule = _AtlBaseModule.GetModuleInstance();
        if (!hModule)
            return E_UNEXPECTED;
        HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
        if (!hRsrc)
            return HRESULT_FROM_WIN32(GetLastError());
        HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
        if (!hGlobal)
            return HRESULT_FROM_WIN32(GetLastError());
        *pdwSizeInBytes = SizeofResource(hModule, hRsrc);
        *ppvResourceData = LockResource(hGlobal);
        return S_OK;
    }

    BSTR GetXMLResource(int nId)
    {
        LPVOID pResourceData = NULL;
        DWORD dwSizeInBytes = 0;
        HRESULT hr = HrGetResource(nId, TEXT("XML"),
            &pResourceData, &dwSizeInBytes);
        if (FAILED(hr))
            return NULL;
        // Assumes that the data is not stored in Unicode.
        CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
        return cbstr.Detach();
    }

    SAFEARRAY* GetOFSResource(int nId)
    {
        LPVOID pResourceData = NULL;
        DWORD dwSizeInBytes = 0;
        if (FAILED(HrGetResource(nId, TEXT("OFS"),
            &pResourceData, &dwSizeInBytes)))
            return NULL;
        SAFEARRAY* psa;
        SAFEARRAYBOUND dim = { dwSizeInBytes, 0 };
        psa = SafeArrayCreate(VT_UI1, 1, &dim);
        if (psa == NULL)
            return NULL;
        BYTE* pSafeArrayData;
        SafeArrayAccessData(psa, (void**)&pSafeArrayData);
        memcpy((void*)pSafeArrayData, pResourceData, dwSizeInBytes);
        SafeArrayUnaccessData(psa);
        return psa;
    }

    HRESULT HrGetImageFromResource(int nId, LPCTSTR lpType, IPictureDisp ** ppdispImage)
    {
        LPVOID pResourceData = NULL;
        DWORD len = 0;
        HRESULT hr = HrGetResource(nId, lpType, &pResourceData, &len);
        if (FAILED(hr)) {
            return E_UNEXPECTED;
        }

        IStream* pStream = nullptr;
        HGLOBAL hGlobal = nullptr;

        // copy image bytes into a real hglobal memory handle
        hGlobal = ::GlobalAlloc(GHND, len);
        if (hGlobal) {
            void* pBuffer = ::GlobalLock(hGlobal);
            if (pBuffer) {
                memcpy(pBuffer, reinterpret_cast<BYTE*>(pResourceData), len);
                HRESULT hr = CreateStreamOnHGlobal(hGlobal, TRUE, &pStream);
                if (SUCCEEDED(hr)) {
                    // pStream now owns the global handle and will invoke GlobalFree on release
                    hGlobal = nullptr;

                    PICTDESC pic;
                    memset(&pic, 0, sizeof pic);
                    Gdiplus::Bitmap *png = Gdiplus::Bitmap::FromStream(pStream);
                    HBITMAP hMap = NULL;
                    png->GetHBITMAP(Gdiplus::Color(), &hMap);
                    pic.picType = PICTYPE_BITMAP;
                    pic.bmp.hbitmap = hMap;

                    OleCreatePictureIndirect(&pic, IID_IPictureDisp, true, (LPVOID*)ppdispImage);
                }
            }
        }

        if (pStream) {
            pStream->Release();
            pStream = nullptr;
        }

        if (hGlobal) {
            GlobalFree(hGlobal);
            hGlobal = nullptr;
        }

        return S_OK;
    }

    HRESULT HrGetImageFromLocal(LPCWSTR pstrFilePath, IPictureDisp ** ppdispImage)
    {
        PICTDESC pic;
        memset(&pic, 0, sizeof pic);
        Gdiplus::Bitmap *png = Gdiplus::Bitmap::FromFile(pstrFilePath);
        HBITMAP hMap = NULL;
        png->GetHBITMAP(Gdiplus::Color(), &hMap);
        pic.picType = PICTYPE_BITMAP;
        pic.bmp.hbitmap = hMap;
        OleCreatePictureIndirect(&pic, IID_IPictureDisp, true, (LPVOID*)ppdispImage);
        return S_OK;
    }

    wstring GetDllPath()
    {
        wstring temp;
        temp.resize(1024 * 4);
        LPWSTR pcTemp = (LPWSTR)temp.data();
        GetModuleFileName(g_hInstance, pcTemp, 1024);
        *_tcsrchr(pcTemp, _T('\\')) = _T('\0');
        _tcscat_s(pcTemp, 1024, _T("\\"));
        temp.resize(_tcslen(pcTemp));
        return temp;
    }

    wstring GetImagesPath()
    {
        return GetDllPath() + L"images\\";
    }

    void ShowLoginDialog()
    {
        DuiLib::CPaintManagerUI::SetInstance(g_hInstance);
        TCHAR szPath[MAX_PATH] = { 0 };
        GetModuleFileName(g_hInstance, szPath, _countof(szPath));
        *_tcsrchr(szPath, _T('\\')) = 0;
        DuiLib::CPaintManagerUI::SetResourcePath(szPath);
        LoginDialog dialog;
        dialog.Create(g_hWnd, _T("登录"), UI_WNDSTYLE_FRAME, WS_EX_WINDOWEDGE);
        dialog.CenterWindow();
        dialog.ShowModal();
        //DuiLib::CPaintManagerUI::MessageLoop();
    }
} // End of anonymous namespace

// CConnect

CConnect::CConnect()
{
}

HRESULT CConnect::FinalConstruct()
{
    return S_OK;
}

void CConnect::FinalRelease()
{
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY ** custom)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY ** custom)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::OnAddInsUpdate(SAFEARRAY ** custom)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::OnStartupComplete(SAFEARRAY ** custom)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::OnBeginShutdown(SAFEARRAY ** custom)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
    if (!RibbonXml)
        return E_POINTER;
    *RibbonXml = GetXMLResource(IDR_XML1);
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::ButtonClicked(IDispatch * control)
{
    CComQIPtr<IRibbonControl> ribbonCtl(control);
    CComBSTR idStr;
    if (ribbonCtl->get_Id(&idStr) != S_OK)
        return S_FALSE;
    if (idStr == OLESTR("loginButton")) {
        ShowLoginDialog();
    } else if (idStr == OLESTR("uploadButton")) {
        WCHAR msg[64];
        swprintf_s(msg, L"I am uploadButton");
        MessageBoxW(NULL, msg, L"NativePPTAddin", MB_OK);
    }
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::GetLabel(IDispatch * control, BSTR * returnedVal)
{
    CComQIPtr<IRibbonControl> ribbonCtl(control);
    CComBSTR idStr;
    if (ribbonCtl->get_Id(&idStr) != S_OK)
        return S_FALSE;
    CComBSTR ret;
    if (idStr == OLESTR("loginButton")) {
        ret = OLESTR("登录");
    } else if (idStr == OLESTR("uploadButton")) {
        ret = OLESTR("上传");
    } else if (idStr == OLESTR("updateButton")) {
        ret = OLESTR("更新");
    }
    *returnedVal = ret.Detach();
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::GetVisible(IDispatch * control, VARIANT_BOOL * returnedVal)
{
    CComQIPtr<IRibbonControl> ribbonCtl(control);
    CComBSTR idStr;
    if (ribbonCtl->get_Id(&idStr) != S_OK)
        return S_FALSE;
    if (idStr == OLESTR("loginButton")) {
        *returnedVal = VARIANT_TRUE;
    } else if (idStr == OLESTR("uploadButton")) {
        *returnedVal = VARIANT_TRUE;
    } else {
        *returnedVal = VARIANT_TRUE;
    }
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::GetImage(IDispatch * control, IPictureDisp ** returnedVal)
{
    CComQIPtr<IRibbonControl> ribbonCtl(control);
    CComBSTR idStr;
    if (ribbonCtl->get_Id(&idStr) != S_OK)
        return S_FALSE;
    if (idStr == OLESTR("loginButton")) {
        // do nothing
        // 登录按钮使用image属性定义
    } else if (idStr == OLESTR("uploadButton")) {
        return HrGetImageFromResource(IDB_PNG_UPLOAD, TEXT("PNG"), returnedVal);
    } else if (idStr == OLESTR("updateButton")) {
        wstring png = GetImagesPath() + L"update.png";
        return HrGetImageFromLocal(png.c_str(), returnedVal);
    }
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CConnect::CustomUILoadImage(BSTR * imageId, IPictureDisp ** returnedVal)
{
    return HrGetImageFromResource(_wtoi(*imageId), TEXT("PNG"), returnedVal);
}
