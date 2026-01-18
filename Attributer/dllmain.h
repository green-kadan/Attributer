// dllmain.h : モジュール クラスの宣言です。

class CAttrShlExtModule : public ATL::CAtlDllModuleT< CAttrShlExtModule >
{
public :
	DECLARE_LIBID(LIBID_AttrShlExtLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ATTRSHLEXT, "{dccd00e2-b94e-4646-b419-ee7c3b26fead}")
};

extern class CAttrShlExtModule _AtlModule;
