// dllmain.h: 模块类的声明。

class CNativePPTAddinModule : public ATL::CAtlDllModuleT< CNativePPTAddinModule >
{
public :
	DECLARE_LIBID(LIBID_NativePPTAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_NATIVEPPTADDIN, "{88a0e8ac-0fbd-4973-a824-210f272d5240}")
};

extern class CNativePPTAddinModule _AtlModule;
