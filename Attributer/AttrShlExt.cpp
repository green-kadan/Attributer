// AttrShlExt.cpp : DLL エクスポートの実装。


#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "AttrShlExt_i.h"
#include "dllmain.h"


using namespace ATL;

// DLL を OLE によってアンロードできるようにするかどうかを指定します。
_Use_decl_annotations_
STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

// 要求された型のオブジェクトを作成するクラス ファクトリを返します。
_Use_decl_annotations_
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - エントリをシステム レジストリに追加します。
STDAPI DllRegisterServer(void)
{
	CRegKey cReg;

	CString KeyName = _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved");
	if (ERROR_SUCCESS == cReg.Open(HKEY_LOCAL_MACHINE, KeyName, KEY_SET_VALUE))
	{
		LONG lStat;
		lStat = cReg.SetStringValue(_T("{470bf21b-4846-11d4-a53b-0000f87a6bf9}"), _T("Attributer Shell Extension"));
		cReg.Close();

		if (lStat != ERROR_SUCCESS)
		{
			// アクセス許可がない
			return E_ACCESSDENIED;
		}
	}
	else
	{
		// アクセス許可がない
		return E_ACCESSDENIED;
	}

	// オブジェクト、タイプ ライブラリおよびタイプ ライブラリ内のすべてのインターフェイスを登録します
	HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}

// DllUnregisterServer - エントリをシステム レジストリから削除します。
STDAPI DllUnregisterServer(void)
{
	CRegKey cReg;
	CString KeyName = _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved");
	if (ERROR_SUCCESS == cReg.Open(HKEY_LOCAL_MACHINE, KeyName, KEY_SET_VALUE))
	{
		cReg.DeleteValue(_T("{470bf21b-4846-11d4-a53b-0000f87a6bf9}"));
		cReg.Close();
	}
	else
	{
		// アクセス許可がない
		return E_ACCESSDENIED;
	}

	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

// DllInstall - ユーザーおよびコンピューターごとのシステム レジストリ エントリを追加または削除します。
STDAPI DllInstall(BOOL bInstall, _In_opt_  LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != nullptr)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			ATL::AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}


