// Connect.h: CConnect 的声明

#pragma once
#include "resource.h"       // 主符号



#include "NativePPTAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>
IDTImpl;

typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
RibbonImpl;

typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>
CallbackImpl;

// CConnect

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_NativePPTAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IDTImpl,
    public RibbonImpl,
    public CallbackImpl
{
public:
    CConnect();

DECLARE_REGISTRY_RESOURCEID(106)


BEGIN_COM_MAP(CConnect)
    COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
	COM_INTERFACE_ENTRY(IConnect)
	COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
    COM_INTERFACE_ENTRY(_IDTExtensibility2)
    COM_INTERFACE_ENTRY(IRibbonExtensibility)
    COM_INTERFACE_ENTRY(IRibbonCallback)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct();

    void FinalRelease();

public:




// _IDTExtensibility2 Methods
public:
    STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
    STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
    STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
    STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
    STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);

    STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);

    STDMETHOD(ButtonClicked)(IDispatch* control);
    STDMETHOD(GetLabel)(IDispatch *control, BSTR *returnedVal);
    STDMETHOD(GetVisible)(IDispatch *control, VARIANT_BOOL *returnedVal);
    STDMETHOD(GetImage)(IDispatch *control, IPictureDisp **returnedVal);
    STDMETHOD(CustomUILoadImage)(BSTR *imageId, IPictureDisp **returnedVal);

};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
