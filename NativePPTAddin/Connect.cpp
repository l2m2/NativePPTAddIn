// Connect.cpp: CConnect 的实现

#include "pch.h"
#include "Connect.h"



namespace {
HRESULT HrGetResource(int nId,
    LPCTSTR lpType,
    LPVOID* ppvResourceData,
    DWORD* pdwSizeInBytes)
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
    WCHAR msg[64];
    if (ribbonCtl->get_Id(&idStr) != S_OK)
        return S_FALSE;
    if (idStr == L"loginButton") {
        swprintf_s(msg, L"I am loginButton");
    } else if (idStr == L"uploadButton") {
        swprintf_s(msg, L"I am uploadButton");
    }
    MessageBoxW(NULL, msg, L"NativePPTAddin", MB_OK);
    return S_OK;
}
