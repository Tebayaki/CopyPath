// dllmain.h: 模块类的声明。

class CCopyPathModule : public ATL::CAtlDllModuleT< CCopyPathModule >
{
public :
	DECLARE_LIBID(LIBID_CopyPathLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COPYPATH, "{ece137e8-a6e1-40ff-9566-fcb874f9670a}")
};

extern class CCopyPathModule _AtlModule;
