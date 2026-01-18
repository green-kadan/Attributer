#pragma once
#include "resource.h"
#include "AttrShlExt_i.h"
#include <vector>
using namespace ATL;

class ATL_NO_VTABLE CAttrShlExt :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAttrShlExt, &CLSID_CAttrShlExt>,
	public IShellExtInit,
	public IShellPropSheetExt
{
public:
	std::vector<CString> m_selectedFiles;

	CAttrShlExt()
	{
	}
	
	~CAttrShlExt()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ATTRSHLEXT)

BEGIN_COM_MAP(CAttrShlExt)
	COM_INTERFACE_ENTRY(IShellExtInit)
	COM_INTERFACE_ENTRY(IShellPropSheetExt)
END_COM_MAP()

public:
	// IShellExtInit
	STDMETHOD(Initialize)(LPCITEMIDLIST, LPDATAOBJECT, HKEY);

	// IShellPropSheetExt
	STDMETHOD(AddPages)(LPFNADDPROPSHEETPAGE, LPARAM);
	STDMETHOD(ReplacePage)(UINT, LPFNADDPROPSHEETPAGE, LPARAM)
	{
		return E_NOTIMPL;
	}
};

OBJECT_ENTRY_AUTO(__uuidof(CAttrShlExt), CAttrShlExt)
