#include "pch.h"
#include "AttrShlExt.h"
#include "dllmain.h"
#include "Logger.h"
#include <vector>
#include <mutex>

// 定数
#define NTFS_MAX_PATH 32768
const CString ADS_ZONE_ID = _T(":Zone.Identifier"); // ADS(Alternate Data Streams)

// C26454  演算のオーバーフロー: '-' の操作では、コンパイル時に負の符号なしの結果が生成されます(io.5)。
#pragma warning(disable:26454)

// 前方宣言
BOOL CALLBACK PropPageDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
UINT CALLBACK PropPageCallbackProc(HWND hWnd, UINT uMsg, LPPROPSHEETPAGE ppsp);
BOOL OnInitDialog(HWND hWnd, WPARAM wParam, LPARAM lParam);
BOOL OnLeftClicked(HWND hWnd, WPARAM wParam, LPARAM lParam);
BOOL OnDoubleClicked(HWND hWnd, WPARAM wParam, LPARAM lParam);
BOOL OnRightClicked(HWND hWnd, WPARAM wParam, LPARAM lParam);
BOOL OnApply(HWND hWnd, WPARAM wParam, LPARAM lParam);
BOOL OnDestroyDialog(HWND hWnd, WPARAM wParam, LPARAM lParam);
void OnChangedCheckSame(HWND hWnd, WPARAM wParam, LPARAM lParam);
void OnChangedCheckPrintClear(HWND hWnd, WPARAM wParam, LPARAM lParam);
void OnChangedCheckEditClear(HWND hWnd, WPARAM wParam, LPARAM lParam);
void OnApplyItem(HWND hWnd, CString csItemPath, WIN32_FILE_ATTRIBUTE_DATA& FileInfo, DWORD dwFileAttributes);
void RecursiveFindFile(HWND hWnd, CString csFolderPath, UINT &nFolderCount, UINT &nFileCount);
void RecursiveOnApply(HWND hWnd, CString csFolderPath, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML);
void SetIPropStgOfficeBinary(HWND hWnd, CString csFilePath, CString csExt, WIN32_FILE_ATTRIBUTE_DATA& FileInfo);
void GetIPropStgOfficeBinary(HWND hWnd, CString csFilePath, CString csExt);
void SetIPropStgOfficeOpenXML(HWND hWnd, CString csFilePath, CString csExt, WIN32_FILE_ATTRIBUTE_DATA& FileInfo, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML);
void GetIPropStgOfficeOpenXML(HWND hWnd, CString csFilePath, CString csExt, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML);
void QueryHandlerClassId(CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML);
void SetDTPControl(HWND hWnd, UINT nIDDlgDate, UINT nIDDlgTime, FILETIME FileTime);
void GetDTPControl(HWND hWnd, UINT nIDDlgDate, UINT nIDDlgTime, FILETIME& FileTime);
void SetEditControl(HWND hWnd, FILETIME FileTime);
void GetEditControl(HWND hWnd, FILETIME& FileTime);
void SetAttributeControl(HWND hWnd, UINT nIDDlgItem, DWORD dwFileAttribute, DWORD dwFileAttributes);
void GetAttributeControl(HWND hWnd, UINT nIDDlgItem, DWORD dwFileAttribute, DWORD* pdwFileAttributes);
void RemoveReadOnly(HWND hWnd, CString csFilePath, DWORD* pdwFileAttributes);
void DeleteZoneId(HWND hWnd, CString csFilePath);
void ShowMessageAndAppendLog(HWND hWnd, CStringW csLocation, CStringW csMessage, CStringW csItemname);
void GetErrorMessage(CStringW* csMessage);
void GetHRESULTValue(HRESULT hr, CStringW* csMessage);

// グローバル変数
CAtlMap<HWND, std::vector<CString>> cmFullPath;
CAtlMap<HWND, HICON> cmIconHandle;
CAtlMap<HWND, HMENU> cmContextSame;
CAtlMap<HWND, HMENU> cmContextSetTime;
CAtlMap<HWND, HMENU> cmContextOverwrite;
// CAtlMap<HWND, HFONT> cmFontHandle;
CAtlMap<HWND, UINT> cmOverwriteId;		// 「上書きする」オプションが選択されているか
CAtlMap<HWND, BOOL> cmControlChange;	// コントロールの内容が変更されたか
CAtlMap<HWND, BOOL> cmMessageShow;		// メッセージを表示するか
CString msg;							// 一時的使用メッセージ

// シェル拡張機能の初期化（データの受け取り）
STDMETHODIMP CAttrShlExt::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hProgID)
{
	HRESULT hr = S_OK;
	LOG(_T("[AttrShlExt] Initialize called\n"));

	// コモンコントロールを初期化する
	INITCOMMONCONTROLSEX iccex = {sizeof(INITCOMMONCONTROLSEX), ICC_DATE_CLASSES};
	InitCommonControlsEx(&iccex);

	// データオブジェクトはCF_HDROP 形式
	FORMATETC etc = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};

    // データオブジェクトを取得する
	STGMEDIUM stg;
	if (FAILED(pDataObj->GetData(&etc, &stg)))
	{
		// データオブジェクトの取得に失敗した
		return E_INVALIDARG;
	}

    // グローバルメモリオブジェクトをロックする
	HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
	if (!hDrop)
	{
		// グローバルメモリオブジェクトのロックに失敗した場合, STGMEDIUM 構造体を解放する
		ReleaseStgMedium(&stg);
		return E_INVALIDARG;
	}

	// グローバルメモリオブジェクトに格納されているパスの総数を取得する
	UINT nPathCount = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);

	m_selectedFiles.clear(); // グローバルではなくメンバ変数

	// グローバルメモリオブジェクトからフルパスを取得する
	for (UINT i = 0; i < nPathCount; i++)
	{
		// フルパスの長さを取得する（終端のNULL を含む）
		UINT nPathLen = DragQueryFile(hDrop, i, NULL, 0) + 1;

		// フルパスを取得する
		CString csPath;
		DragQueryFile(hDrop, i, csPath.GetBuffer(nPathLen), nPathLen);
		csPath.ReleaseBuffer();

		// パスがルートフォルダの場合, プロパティシートを追加（表示）しない
		if (CPath(csPath).IsRoot())
		{
			LOG(_T("[AttrShlExt] Initialize finished NG\n"));
			hr = E_FAIL;
			break;
		}

		// フルパスを動的配列クラスに格納する
		m_selectedFiles.push_back(csPath);
		
		msg.Format(_T("[AttrShlExt] Initialize: Path = %s\n"), csPath);
		LOG(msg);
	}

	// グローバルメモリオブジェクトとSTGMEDIUM 構造体を解放する
	GlobalUnlock(stg.hGlobal);
	ReleaseStgMedium(&stg);

	msg.Format(_T("[AttrShlExt] Initialize: this = %p\n"), this);
	LOG(msg);

	msg.Format(_T("[AttrShlExt] Initialize: m_selectedFiles size = %d\n"), (int)m_selectedFiles.size());
	LOG(msg);

	LOG(_T("[AttrShlExt] Initialize finished OK\n"));
	return hr;
}

// プロパティシートの追加
STDMETHODIMP CAttrShlExt::AddPages(LPFNADDPROPSHEETPAGE pfnAddPropSheet, LPARAM lParam)
{
	PROPSHEETPAGE psp = {0};
	HPROPSHEETPAGE hPropSheet;

	// PROPSHEELPAGE 構造体を設定する
	psp.dwSize = sizeof(PROPSHEETPAGE);
	psp.dwFlags = PSP_DEFAULT | PSP_USETITLE | PSP_USECALLBACK | PSP_USEREFPARENT;
	psp.pszTitle = MAKEINTRESOURCE(IDS_TABNAME);

	// Windows バージョンを確認する
	if (IsWindows10OrGreater())
	{
		// Windows 10 以降
		psp.pszTemplate = MAKEINTRESOURCE(IDD_ATTRIBUTER);
	}
	else
	{
		// Windows 7, 8
		psp.pszTemplate = MAKEINTRESOURCE(IDD_ATTRIBUTER78);
	}
	
	psp.hInstance = _AtlBaseModule.GetResourceInstance();
	psp.pcRefParent = (UINT*)&_AtlModule.m_nLockCnt;
	psp.pfnDlgProc = (DLGPROC)PropPageDlgProc;
	psp.pfnCallback = PropPageCallbackProc;			// ← 寿命管理用
	psp.lParam = (LPARAM)this;						// ← this を渡す

	msg.Format(_T("[AttrShlExt] AddPages: this = %p\n"), this);
	LOG(msg);

	msg.Format(_T("[AttrShlExt] AddPages: m_selectedFiles size = %d\n"),
		(int)this->m_selectedFiles.size());
	LOG(msg);

	// プロパティシートのハンドルを生成する
	hPropSheet = CreatePropertySheetPage(&psp);

	// ハンドルの生成に成功したことを確認する
	if (NULL != hPropSheet)
	{
		// 作成したプロパティシートを追加する
		if (!pfnAddPropSheet(hPropSheet, lParam))
		{
			// プロパティシートの追加に失敗した場合, ハンドルを廃棄する
			// （成功した場合はプロパティシートの破棄時にハンドルも自動廃棄される）
			DestroyPropertySheetPage(hPropSheet);
			return E_FAIL;
		}
	}
	else
	{
		// ハンドルの生成に失敗した
		return E_FAIL;
	}

	return S_OK;
}

// ダイアログボックスプロシージャ
BOOL CALLBACK PropPageDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	NMHDR* pHdr;
	LONG_PTR lpId;
	SYSTEMTIME st = {0};
	BOOL bRet = FALSE;

	switch (uMsg)
	{
	// ダイアログボックスが作成される
	case WM_INITDIALOG:
		bRet = OnInitDialog(hWnd, wParam, lParam);
		break;

	// 通知メッセージ
	case WM_COMMAND:
		switch (HIWORD(wParam))
		{
		// チェックボックスがクリックされた
		case BN_CLICKED:
			bRet = OnLeftClicked(hWnd, wParam, lParam);
			break;

		// スタティックがダブルクリックされた
		case STN_DBLCLK:
			bRet = OnDoubleClicked(hWnd, wParam, lParam);
			break;

		// エディットコントロールが変更された
		case EN_CHANGE:

			// 「適用」ボタンを有効にする
			PropSheet_Changed(GetParent(hWnd), hWnd);
			cmControlChange[hWnd] = TRUE;
			bRet = TRUE;
			break;
		}

		break;

	// 通知メッセージ（コモンコントロール）
	case WM_NOTIFY:
		pHdr = (NMHDR*)lParam;

		switch (pHdr->code)
		{
		// DateTimePicker コントロールの内容が変更された
		case DTN_DATETIMECHANGE:

			// 「日時を同じにする」の状態を確認する
			if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_SAME))
			{
				// 「日時を同じにする」がオンの状態でDateTimePicker コントロールが変更されたので
				// 「作成日時」を「更新（前回保存）日時」および「アクセス日時」にコピーする
				DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), &st);
				DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), GDT_VALID, &st);
				DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), GDT_VALID, &st);
				DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), &st);
				DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), GDT_VALID, &st);
				DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), GDT_VALID, &st);
			}

			// 「適用」ボタンを有効にする
			PropSheet_Changed(GetParent(hWnd), hWnd);
			cmControlChange[hWnd] = TRUE;
			bRet = TRUE;
			break;

		// 「適用」ボタンがクリックされた
		case PSN_APPLY:
			bRet = OnApply(hWnd, wParam, lParam);
			break;
		}

		break;

	// コンテキストメニュー
	case WM_CONTEXTMENU:
		bRet = OnRightClicked(hWnd, wParam, lParam);
		break;

	// スタティックコントロールが描画される（前景色および背景色の設定のみ可能）
	case WM_CTLCOLORSTATIC:
		lpId = GetWindowLongPtr((HWND)lParam, GWLP_ID);

		switch (lpId)
		{
		case IDC_RECT_TIMESTAMP:
		case IDC_RECT_ATTRIBUTE:
		case IDC_RECT_FOLDER:
		case IDC_RECT_FILE:

			// 背景を透明にする
			return (BOOL)(LRESULT)GetStockObject(NULL_BRUSH);

		case IDC_ITEMEDIT:

			// 背景を白くする
			return (BOOL)(LRESULT)GetStockObject(WHITE_BRUSH);
		}

		break;

	// ダイアログボックスが作成される
	case WM_DESTROY:
		bRet = OnDestroyDialog(hWnd, wParam, lParam);
		break;
	}

	return bRet;
}

// プロパティシートの作成と破棄時に呼び出されるコールバック関数
UINT CALLBACK PropPageCallbackProc(HWND hWnd, UINT uMsg, LPPROPSHEETPAGE ppsp)
{
	CAttrShlExt* pThis = (CAttrShlExt*)ppsp->lParam;

	switch (uMsg)
	{
	case PSPCB_ADDREF: // ここで初めて "実際のダイアログ hWnd" が来る
		msg.Format(_T("[AttrShlExt] PSPCB_ADDREF: this = %p\n"), (void*)ppsp->lParam);
		LOG(msg);
		
		if (pThis)
			pThis->AddRef();					// ★ ページが生きている間、COMオブジェクトを保持
		
		break;

	// プロパティシートの破棄時
	case PSPCB_RELEASE:
		if (pThis) 
			pThis->Release();					// ★ ページ破棄時に Release

		break;
	}

	return TRUE;
}

// ダイアログボックス初期化
BOOL OnInitDialog(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	PROPSHEETPAGE* psp = (PROPSHEETPAGE*)lParam;
	CAttrShlExt* pThis = (CAttrShlExt*)psp->lParam;

	msg.Format(_T("[AttrShlExt] OnInitDialog: pThis = %p\n"), pThis);
	LOG(msg);

	// DEBUG用
#ifdef _DEBUG
	if (!pThis) return TRUE;
	if (pThis->m_selectedFiles.size() == 0 || pThis->m_selectedFiles.size() > 1000) return TRUE;
#endif

	msg.Format(_T("[AttrShlExt] OnInitDialog: m_selectedFiles size = %d\n"),
		(int)pThis->m_selectedFiles.size());
	LOG(msg);

	// メンバ変数からファイルリストを取得
	cmFullPath[hWnd] = pThis->m_selectedFiles;

	// リソースインスタンスのハンドルを取得する
	HINSTANCE hInst;
	hInst = _AtlBaseModule.GetResourceInstance();

	// System フォルダのパスを取得する
	CString csPath;
	SHGetFolderPath(hWnd, CSIDL_SYSTEM, NULL, SHGFP_TYPE_CURRENT, csPath.GetBufferSetLength(NTFS_MAX_PATH + 1));
	csPath.ReleaseBuffer();

	// アイコンのハンドルを取得する
	HICON hIcon;
	hIcon = ExtractIcon(hInst, csPath + _T("\\Shell32.dll"), 269);

	// アイコンを表示する
	if (NULL != hIcon)
	{
		Static_SetIcon(GetDlgItem(hWnd, IDC_ITEMICON), hIcon);
		cmIconHandle[hWnd] = hIcon;
	}

	// コンテキストメニューを作成
	HMENU hContextSame, hContextSetTime, hContextOverwrite;
	hContextSame = GetSubMenu(LoadMenu(hInst, _T("CONTEXT_MENU")), 0);
	cmContextSame[hWnd] = hContextSame;
	hContextSetTime = GetSubMenu(LoadMenu(hInst, _T("CONTEXT_MENU")), 1);
	cmContextSetTime[hWnd] =  hContextSetTime;
	hContextOverwrite = GetSubMenu(LoadMenu(hInst, _T("CONTEXT_MENU")), 2);
	cmContextOverwrite[hWnd] = hContextOverwrite;

	// システムフォントの設定内容を取得する
	// NONCLIENTMETRICS ncm = {0};
	// ncm.cbSize = sizeof(ncm);
	// if (0 != SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
	// {
	//	   // フォントハンドルを取得する
	// 	   HFONT hFont;
	//     hFont = CreateFontIndirect(&ncm.lfMessageFont);
	//     cmFontHandle[hWnd] = hFont;
	// }

	// グループボックスのみテーマ（ビジュアルスタイル）を無効にする
	// （テーマが有効の場合, グループボックス無効時にグレイアウトしないため）
	SetWindowTheme(GetDlgItem(hWnd, IDC_GROUP_TIMESTAMP), _T(""), _T(""));
	SetWindowTheme(GetDlgItem(hWnd, IDC_GROUP_EDIT), _T(""), _T(""));
	SetWindowTheme(GetDlgItem(hWnd, IDC_GROUP_ATTRIBUTE), _T(""), _T(""));
	SetWindowTheme(GetDlgItem(hWnd, IDC_GROUP_FOLDER), _T(""), _T(""));
	SetWindowTheme(GetDlgItem(hWnd, IDC_GROUP_FILE), _T(""), _T(""));

	// スピンコントロールの範囲を設定する（編集日数）
	SendDlgItemMessage(hWnd, IDC_SPIN_EDIT, UDM_SETRANGE32, (WPARAM)0, (LPARAM)9999);

	// チェックボックスを初期化する
	CheckDlgButton(hWnd, IDC_CHECK_CREATION, BST_CHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_WRITE, BST_CHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_ACCESS, BST_CHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_SAME, BST_UNCHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_PRINT_CLEAR, BST_UNCHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_EDIT_CLEAR, BST_UNCHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_FOLDER, BST_UNCHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_FILE, BST_UNCHECKED);
	CheckDlgButton(hWnd, IDC_CHECK_ZONE, BST_UNCHECKED);

	// 「前回印刷日」をクリアする
	CheckDlgButton(hWnd, IDC_CHECK_PRINT, BST_UNCHECKED);
	OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_PRINT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_PRINT));

	// 「総編集時間」をクリアする
	SetWindowText(GetDlgItem(hWnd, IDC_EDIT_EDIT), _T("0"));
	SYSTEMTIME st = {0};
	st.wYear = 1601;
	st.wMonth = 1;
	st.wDay = 1;
	st.wHour = 0;
	st.wMinute = 0;
	st.wSecond = 0;
	st.wMilliseconds = 0;
	DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), GDT_VALID, &st);
	CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_UNCHECKED);
	OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_EDIT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_EDIT));

	// 選択されたアイテムの情報をコントロールに設定する
	std::vector<CString> svFullPath;
	svFullPath = cmFullPath[hWnd];
	UINT nFolderCount = 0;
	UINT nFileCount = 0;
	if (1 == svFullPath.size())
	{
		// 選択されたアイテムは1 つ
		CString csFilename;
		csFilename = svFullPath[0].Right(svFullPath[0].GetLength() - svFullPath[0].ReverseFind(_T('\\')) - 1);
		SetWindowText(GetDlgItem(hWnd, IDC_ITEMNAME), csFilename);

		// 選択されたアイテムの種類を確認する
		if (!CPath(svFullPath[0]).IsDirectory())
		{
			// プロパティハンドラのCLSID を取得
			CAtlMap<CString, CLSID> cmOfficeBin;
			CAtlMap<CString, CLSID> cmOfficeXML;
			QueryHandlerClassId(cmOfficeBin, cmOfficeXML);

			// 拡張子を取得する
			CString csExt;
			csExt = CPath(svFullPath[0]).GetExtension().MakeLower();

			CLSID classId;
			if (cmOfficeBin.Lookup(csExt, classId) || cmOfficeXML.Lookup(csExt, classId))
			{
				// Office Binary 形式
				if (cmOfficeBin.Lookup(csExt, classId))
				{
					// プロパティセットから日時を取得（Office Binary 形式）
					GetIPropStgOfficeBinary(hWnd, svFullPath[0], csExt);
				}

				// Office OpenXML 形式
				CLSID classId;
				if (cmOfficeXML.Lookup(csExt, classId))
				{
					// プロパティセットから日時を取得（Office OpenXML 形式）
					GetIPropStgOfficeOpenXML(hWnd, svFullPath[0], csExt, cmOfficeBin, cmOfficeXML);
				}
			}
			else
			{
				// 「前回印刷日」「印刷日クリア」を無効にする
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_PRINT), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_PRINT_CLEAR), FALSE);

				// 「総編集時間」「編集時間クリア」を無効にする
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_GROUP_EDIT), FALSE);
				ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_HIDE);
				EnableWindow(GetDlgItem(hWnd, IDC_SPIN_EDIT), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT_CLEAR), FALSE);
			}

			nFileCount = 1;
		}
		else
		{
			nFolderCount = 1;
		}
	}
	else
	{
		// 複数のアイテムが選択されている

		// 選択されたフォルダとファイルの数を取得する
		for (UINT i = 0; i < svFullPath.size(); i++)
		{
			// 選択されたアイテムの種類を確認する
			if (CPath(svFullPath[i]).IsDirectory())
			{
				// フォルダ
				RecursiveFindFile(hWnd, svFullPath[i], nFolderCount, nFileCount);
				nFolderCount++;
			}
			else
			{
				// ファイル
				nFileCount++;
			}
		}

		// 確認したフォルダとファイル数をコントロールに設定する
		CString csItemCount;
		CString csItemTitle = MAKEINTRESOURCE(IDS_MULTIPLE_TYPES);
		csItemCount.Format(_T("%d"), nFolderCount);
		csItemTitle.Replace(_T("%FOLDER_COUNT%"), csItemCount);
		csItemCount.Format(_T("%d"), nFileCount);
		csItemTitle.Replace(_T("%FILE_COUNT%"), csItemCount);
		SetWindowText(GetDlgItem(hWnd, IDC_ITEMNAME), csItemTitle);
	}

	//  タイムスタンプや属性を取得する
	WIN32_FILE_ATTRIBUTE_DATA FileInfo = {0};
	if (0 != GetFileAttributesEx((LPCTSTR)svFullPath[0], GetFileExInfoStandard, &FileInfo))
	{
		// タイムスタンプをコントロールに設定する
		SetDTPControl(hWnd, IDC_DTP_CREATION_DATE, IDC_DTP_CREATION_TIME, FileInfo.ftCreationTime);
		SetDTPControl(hWnd, IDC_DTP_WRITE_DATE, IDC_DTP_WRITE_TIME, FileInfo.ftLastWriteTime);
		SetDTPControl(hWnd, IDC_DTP_ACCESS_DATE, IDC_DTP_ACCESS_TIME, FileInfo.ftLastAccessTime);

		// 選択されているアイテムの数と種類を確認する
		if ((1 == svFullPath.size()) && !(FILE_ATTRIBUTE_DIRECTORY & FileInfo.dwFileAttributes))
		{
			// 1 ファイルのみ選択されているため, 属性をコントロールに設定する
			SetAttributeControl(hWnd, IDC_CHECK_READONLY, FILE_ATTRIBUTE_READONLY, FileInfo.dwFileAttributes);
			SetAttributeControl(hWnd, IDC_CHECK_HIDDEN, FILE_ATTRIBUTE_HIDDEN, FileInfo.dwFileAttributes);
			SetAttributeControl(hWnd, IDC_CHECK_ARCHIVE, FILE_ATTRIBUTE_ARCHIVE, FileInfo.dwFileAttributes);
			SetAttributeControl(hWnd, IDC_CHECK_INDEX, FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, FileInfo.dwFileAttributes);
			SetAttributeControl(hWnd, IDC_CHECK_SYSTEM, FILE_ATTRIBUTE_SYSTEM, FileInfo.dwFileAttributes);

			// 「選択したフォルダ配下の」グループボックスとその配下のコントロールを無効にする
			EnableWindow(GetDlgItem(hWnd, IDC_GROUP_FOLDER), FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_CHECK_FOLDER), FALSE);
			EnableWindow(GetDlgItem(hWnd, IDC_CHECK_FILE), FALSE);

			// ゾーン情報の有無を確認する
			if (CPath(svFullPath[0] + ADS_ZONE_ID).FileExists())
			{
				// ゾーン情報あり（「選択したフォルダ配下とファイルの」グループボックスとその配下を有効にする）
				EnableWindow(GetDlgItem(hWnd, IDC_GROUP_FILE), TRUE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_ZONE), TRUE);

			}
			else
			{
				// ゾーン情報なし（「選択したフォルダ配下とファイルの」グループボックスとその配下を無効にする）
				EnableWindow(GetDlgItem(hWnd, IDC_GROUP_FILE), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_ZONE), FALSE);
			}
		}
		else
		{
			// 複数のアイテムまたは1 フォルダが選択されている

			// 「属性」グループボックス配下のコントロールを不確定にする
			CheckDlgButton(hWnd, IDC_CHECK_READONLY, BST_INDETERMINATE);
			CheckDlgButton(hWnd, IDC_CHECK_HIDDEN, BST_INDETERMINATE);
			CheckDlgButton(hWnd, IDC_CHECK_ARCHIVE, BST_INDETERMINATE);
			CheckDlgButton(hWnd, IDC_CHECK_INDEX, BST_INDETERMINATE);
			CheckDlgButton(hWnd, IDC_CHECK_SYSTEM, BST_INDETERMINATE);

			// 選択されたアイテムにフォルダが含まれるか確認する
			if (0 == nFolderCount)
			{
				// 「選択したフォルダ配下の」グループボックス配下のコントロールを無効にする
				EnableWindow(GetDlgItem(hWnd, IDC_GROUP_FOLDER), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_FOLDER), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_FILE), FALSE);
			}
		}
	}
	else
	{
		// 属性の取得に失敗した
		CStringW csMessage;
		GetErrorMessage(&csMessage);
		ShowMessageAndAppendLog(hWnd, _T("OnInitDialog-GetAttribute"), csMessage, svFullPath[0]);
	}

	// グローバル変数を初期化する
	cmOverwriteId[hWnd] = 0;
	cmControlChange[hWnd] = FALSE;
	cmMessageShow[hWnd] = TRUE;

	return TRUE;
}

// 左クリックされた
BOOL OnLeftClicked(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	CString csTimeStr;
	BOOL bEnable = NULL;

	switch (LOWORD(wParam))
	{
	case IDC_CHECK_CREATION:
	case IDC_CHECK_WRITE:
	case IDC_CHECK_ACCESS:
	case IDC_CHECK_PRINT:
	case IDC_CHECK_EDIT:
		switch (Button_GetCheck((HWND)lParam))
		{
		// オン
		case BST_CHECKED:
			bEnable = TRUE;

			// 「前回印刷日」
			if (IDC_CHECK_PRINT == LOWORD(wParam))
			{
				// 「前回印刷日」DateTimePicker を編集可能にする
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_PRINT_DATE), NULL);
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_PRINT_TIME), NULL);
			}

			// 「総編集時間」
			if (IDC_CHECK_EDIT == LOWORD(wParam))
			{
				// 「総編集時間」DateTimePicker を編集可能にする
				ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_SHOW);
				ShowWindow(GetDlgItem(hWnd, IDC_ITEMEDIT), SW_SHOW);
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), NULL);
			}

			break;

		// オフ
		case BST_UNCHECKED:
			bEnable = FALSE;

			// 「前回印刷日」
			if (IDC_CHECK_PRINT == LOWORD(wParam))
			{
				// 「前回印刷日」DateTimePicker を編集不可にする
				csTimeStr = _T(" ");
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_PRINT_DATE), (LPCTSTR)csTimeStr);
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_PRINT_TIME), (LPCTSTR)csTimeStr);
			}

			// 「総編集時間」
			if (IDC_CHECK_EDIT == LOWORD(wParam))
			{
				// 「総編集時間」DateTimePicker を編集不可にする
				ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_HIDE);
				ShowWindow(GetDlgItem(hWnd, IDC_ITEMEDIT), SW_HIDE);
				csTimeStr = _T(" ");
				DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), (LPCTSTR)csTimeStr);
			}

			break;
		}

		// チェックボックスの状態に応じてDateTimePicker コントロールの有効／無効化を行う
		switch (LOWORD(wParam))
		{
		case IDC_CHECK_CREATION:
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), bEnable);
			break;
		case IDC_CHECK_WRITE:
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), bEnable);
			break;
		case IDC_CHECK_ACCESS:
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), bEnable);
			break;
		case IDC_CHECK_PRINT:
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_PRINT_DATE), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_PRINT_TIME), bEnable);
			break;
		case IDC_CHECK_EDIT:
			EnableWindow(GetDlgItem(hWnd, IDC_GROUP_EDIT), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_SPIN_EDIT), bEnable);
			EnableWindow(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), bEnable);
			break;
		}

		break;

	// 「日時を同じにする」
	case IDC_CHECK_SAME:
		OnChangedCheckSame(hWnd, wParam, lParam);
		break;

	// 「印刷日クリア」
	case IDC_CHECK_PRINT_CLEAR:
		OnChangedCheckPrintClear(hWnd, wParam, lParam);
		break;

	// 「編集時間クリア」
	case IDC_CHECK_EDIT_CLEAR:
		OnChangedCheckEditClear(hWnd, wParam, lParam);
		break;
	}

	// 「適用」ボタンを有効にする
	PropSheet_Changed(GetParent(hWnd), hWnd);
	cmControlChange[hWnd] = TRUE;

	return TRUE;
}

// ダブルクリックされた
BOOL OnDoubleClicked(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	switch (LOWORD(wParam))
	{
	// 「タイムスタンプ（フォルダ／ファイル）」
	case IDC_RECT_TIMESTAMP:

		// 「日時を同じにする」をオフにする
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_SAME), BST_UNCHECKED);
		OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_SAME, 0), BST_UNCHECKED);

		// 「印刷日クリア」
		if (IsWindowEnabled(GetDlgItem(hWnd, IDC_CHECK_PRINT_CLEAR)))
		{
			Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_PRINT_CLEAR), BST_UNCHECKED);
			OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_PRINT_CLEAR, 0), BST_UNCHECKED);
		}

		// 「編集時間クリア」
		if (IsWindowEnabled(GetDlgItem(hWnd, IDC_CHECK_EDIT_CLEAR)))
		{
			Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_EDIT_CLEAR), BST_UNCHECKED);
			OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_EDIT_CLEAR, 0), BST_UNCHECKED);
		}

		// 「作成日時」
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_CREATION), BST_UNCHECKED);
		OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_CREATION, 0), BST_UNCHECKED);

		// 「更新(前回保存)日時」
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_WRITE), BST_UNCHECKED);
		OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_WRITE, 0), BST_UNCHECKED);

		// 「アクセス日時」
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_ACCESS), BST_UNCHECKED);
		OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_ACCESS, 0), BST_UNCHECKED);

		break;

	// 「属性」
	case IDC_RECT_ATTRIBUTE:
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_READONLY), BST_INDETERMINATE);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_HIDDEN), BST_INDETERMINATE);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_ARCHIVE), BST_INDETERMINATE);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_INDEX), BST_INDETERMINATE);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_SYSTEM), BST_INDETERMINATE);
		break;

	// 「選択したフォルダ配下の」
	case IDC_RECT_FOLDER:

		// 「サブフォルダを更新する」
		if (IsWindowEnabled(GetDlgItem(hWnd, IDC_CHECK_FOLDER)))
		{
			Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_FOLDER), BST_UNCHECKED);
		}

		// 「ファイルを更新する」
		if (IsWindowEnabled(GetDlgItem(hWnd, IDC_CHECK_FILE)))
		{
			Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_FILE), BST_UNCHECKED);
		}

		break;

	// 「選択したフォルダ配下とファイルの」
	case IDC_RECT_FILE:

		// 「ゾーン情報を削除する」
		if (IsWindowEnabled(GetDlgItem(hWnd, IDC_CHECK_ZONE)))
		{
			Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_ZONE), BST_UNCHECKED);
		}

		break;
	}

	return TRUE;
}

// 右クリックされた
BOOL OnRightClicked(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	MENUITEMINFO mii = {0};
	CString csTimeStr;
	SYSTEMTIME st = {0};
	NMHDR pHdr = {0};
	UINT selectedId;
	
	//マウスポインタの座標を取得する
	POINT mousePoint = {0};
	mousePoint.x = LOWORD(lParam);
	mousePoint.y = HIWORD(lParam);
	
	UINT uFlags = TPM_LEFTALIGN | TPM_TOPALIGN | TPM_NONOTIFY | TPM_RETURNCMD | TPM_LEFTBUTTON;
	switch (GetDlgCtrlID((HWND)wParam))
	{
	// 「作成日時」「更新日時（前回保存）日時」「アクセス日時」
	case IDC_CHECK_CREATION:
	case IDC_CHECK_WRITE:
	case IDC_CHECK_ACCESS:

		// コンテキストメニューを表示する
		HMENU hContextSame;
		cmContextSame.Lookup(hWnd, hContextSame);
		selectedId = TrackPopupMenu(hContextSame, uFlags, mousePoint.x, mousePoint.y, 0, hWnd, NULL);

		switch (selectedId)
		{
		// 「作成日時にあわせる」
		case ID_SAME_CREATION:

			// 「作成日時」を「更新（前回保存）日時」および「アクセス日時」にコピーする
			DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), GDT_VALID, &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), GDT_VALID, &st);
			DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), GDT_VALID, &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), GDT_VALID, &st);
			break;

		// 「更新（前回保存）日時にあわせる」
		case ID_SAME_WRITE:

			// 「更新（前回保存）日時作成日時」を「作成日時」および「アクセス日時」にコピーする
			DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), GDT_VALID, &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), GDT_VALID, &st);
			DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), GDT_VALID, &st);
			DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), GDT_VALID, &st);
			break;
		}

		// 「適用」ボタンを有効にする
		PropSheet_Changed(GetParent(hWnd), hWnd);
		cmControlChange[hWnd] = TRUE;

		break;

	// 「作成時刻」「更新日時（前回保存）時刻」「アクセス時刻」「前回印刷時刻」「総編集時間」
	case IDC_DTP_CREATION_TIME:
	case IDC_DTP_WRITE_TIME:
	case IDC_DTP_ACCESS_TIME:
	case IDC_DTP_PRINT_TIME:
	case IDC_DTP_EDIT_TIME:

		// コンテキストメニューを表示する
		HMENU hContextSetTime;
		cmContextSetTime.Lookup(hWnd, hContextSetTime);
		selectedId = TrackPopupMenu(hContextSetTime, uFlags, mousePoint.x, mousePoint.y, 0, hWnd, NULL);

		// 選択されたメニューの文字列を取得（00:00:00または12:00:00）
		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask = MIIM_STRING;
		mii.dwTypeData = NULL;
		GetMenuItemInfo(hContextSetTime, selectedId, FALSE, &mii);
		mii.cch++;
		mii.dwTypeData = csTimeStr.GetBuffer(mii.cch);
		GetMenuItemInfo(hContextSetTime, selectedId, FALSE, &mii);
		csTimeStr.ReleaseBuffer();

		// DateTimePicker コントロールから時刻を取得する
		DateTime_GetSystemtime((HWND)wParam, (LPARAM)&st);

		// 選択されたメニューから時刻を取得し構造体に設定する
		_stscanf_s(csTimeStr, _T("%hd:%hd:%hd"), &st.wHour, &st.wMinute, &st.wSecond);

		// 時刻をDateTimePicker コントロールに設定する
		DateTime_SetSystemtime((HWND)wParam, GDT_VALID, (LPARAM)&st);

		// 「日時を同じにする」の状態を確認する
		if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_SAME))
		{
			// DateTimePicker コントロールの内容が変更されたイベントを発生させる
			pHdr.code = (UINT)DTN_DATETIMECHANGE;
			PropPageDlgProc(hWnd, WM_NOTIFY, wParam, (LPARAM)&pHdr);
		}

		// 「適用」ボタンを有効にする
		PropSheet_Changed(GetParent(hWnd), hWnd);
		cmControlChange[hWnd] = TRUE;

		break;

	// 「日時を同じにする」
	case IDC_CHECK_SAME:
		if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_SAME))
		{
			// コンテキストメニューを表示する
			HMENU hContextOverwrite;
			cmContextOverwrite.Lookup(hWnd, hContextOverwrite);
			selectedId = TrackPopupMenu(hContextOverwrite, uFlags, mousePoint.x, mousePoint.y, 0, hWnd, NULL);

			if (0 < selectedId)
			{
				// 選択されたメニューのチェックマークを確認する
				mii = {0};
				mii.cbSize = sizeof(mii);
				mii.fMask = MIIM_STATE;

				GetMenuItemInfo(hContextOverwrite, selectedId, FALSE, &mii);
				if (MFS_CHECKED == mii.fState)
				{
					// 選択されたメニューのチェックマークがオンのため, オフにする
					mii.fState = MFS_UNCHECKED;
					SetMenuItemInfo(hContextOverwrite, selectedId, FALSE, &mii);

					// 選択されたメニューは存在しない
					selectedId = 0;
				}
				else
				{
					// すべてのメニューでチェックマークをオフにする
					mii.fState = MFS_UNCHECKED;
					for (UINT i = 0; i < (UINT)GetMenuItemCount(hContextOverwrite); i++)
					{
						SetMenuItemInfo(hContextOverwrite, i, TRUE, &mii);
					}

					// 選択されたメニューのチェックマークをオンにする
					mii.fState = MFS_CHECKED;
					SetMenuItemInfo(hContextOverwrite, selectedId, FALSE, &mii);
				}

				// 選択された内容をグローバル変数に退避する
				cmOverwriteId[hWnd] = selectedId;

				switch (selectedId)
				{
				// 「作成日時で上書きする」
				case ID_OVERWRITE_CREATION:
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), FALSE);
					csTimeStr = _T(" ");
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), (LPCTSTR)csTimeStr);
					break;

				// 「更新（前回保存）日時で上書きする」
				case ID_OVERWRITE_WRITE:
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), FALSE);
					csTimeStr = _T(" ");
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), (LPCTSTR)csTimeStr);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), (LPCTSTR)csTimeStr);
					break;

				// いずれのメニューも選択されていない
				case 0:
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), TRUE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), TRUE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), FALSE);
					EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), FALSE);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), NULL);
					DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), NULL);
					break;
				}
			}
		}

		break;
	}

	return TRUE;
}

// 「適用」ボタンがクリックされた
BOOL OnApply(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	WIN32_FILE_ATTRIBUTE_DATA FileInfo = {0};
	DWORD dwFileAttributes;

	// 更新の必要性を確認する
	BOOL bControlChange;
	cmControlChange.Lookup(hWnd, bControlChange);
	if (!bControlChange)
	{
		return TRUE;
	}
	if ((BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_CREATION)) && (BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_WRITE)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ACCESS)) && (BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_PRINT)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_EDIT)) && (BST_INDETERMINATE == IsDlgButtonChecked(hWnd, IDC_CHECK_READONLY)) &&
		(BST_INDETERMINATE == IsDlgButtonChecked(hWnd, IDC_CHECK_HIDDEN)) && (BST_INDETERMINATE == IsDlgButtonChecked(hWnd, IDC_CHECK_ARCHIVE)) &&
		(BST_INDETERMINATE == IsDlgButtonChecked(hWnd, IDC_CHECK_INDEX)) && (BST_INDETERMINATE == IsDlgButtonChecked(hWnd, IDC_CHECK_SYSTEM)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ZONE)))
	{
		return TRUE;
	}

	// プロパティハンドラのCLSID を取得
	CAtlMap<CString, CLSID> cmOfficeBin;
	CAtlMap<CString, CLSID> cmOfficeXML;
	QueryHandlerClassId(cmOfficeBin, cmOfficeXML);

	// タイムスタンプなどを更新する
	std::vector<CString> svFullPath;
	svFullPath = cmFullPath[hWnd];
	for (UINT i = 0; i < svFullPath.size(); i++)
	{
		// アイテムの属性を取得する
		if (0 != GetFileAttributesEx(svFullPath[i], GetFileExInfoStandard, &FileInfo))
		{
			// 属性を退避する
			dwFileAttributes = FileInfo.dwFileAttributes;

			// 「作成日時で上書きする」「更新（前回保存）日時で上書きする」が選択されているか確認する
			UINT nOverwriteId;
			cmOverwriteId.Lookup(hWnd, nOverwriteId);
			switch (nOverwriteId)
			{
				// 「作成日時で上書きする」
			case ID_OVERWRITE_CREATION:
				FileInfo.ftLastWriteTime = FileInfo.ftCreationTime;
				FileInfo.ftLastAccessTime = FileInfo.ftCreationTime;
				break;

				// 「更新（前回保存）日時で上書きする」
			case ID_OVERWRITE_WRITE:
				FileInfo.ftCreationTime = FileInfo.ftLastWriteTime;
				FileInfo.ftLastAccessTime = FileInfo.ftLastWriteTime;
				break;

				// いずれのメニューも選択されていない
			case 0:
				GetDTPControl(hWnd, IDC_DTP_CREATION_DATE, IDC_DTP_CREATION_TIME, FileInfo.ftCreationTime);
				GetDTPControl(hWnd, IDC_DTP_WRITE_DATE, IDC_DTP_WRITE_TIME, FileInfo.ftLastWriteTime);
				GetDTPControl(hWnd, IDC_DTP_ACCESS_DATE, IDC_DTP_ACCESS_TIME, FileInfo.ftLastAccessTime);
				break;
			}

			if (FILE_ATTRIBUTE_DIRECTORY & FileInfo.dwFileAttributes)
			{
				// フォルダ

				if ((BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_FOLDER)) || (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_FILE)) ||
					(BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ZONE)))
				{
					// 再帰処理
					RecursiveOnApply(hWnd, svFullPath[i], cmOfficeBin, cmOfficeXML);
				}

				// フォルダのタイムスタンプや属性を適用
				OnApplyItem(hWnd, svFullPath[i], FileInfo, dwFileAttributes);
			}
			else
			{
				// ファイル

				// 「読み取り専用」属性の解除
				RemoveReadOnly(hWnd, svFullPath[i], &dwFileAttributes);

				// 「ゾーン情報を削除する」の状態を確認する
				if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ZONE))
				{
					// ゾーン情報の削除
					DeleteZoneId(hWnd, svFullPath[i]);
				}

				// 拡張子を取得する
				CString csExt;
				csExt = CPath(svFullPath[i]).GetExtension().MakeLower();

				// Office Binary 形式
				CLSID classId;
				if (cmOfficeBin.Lookup(csExt, classId))
				{
					// プロパティセットを適用（Office Binary 形式）
					SetIPropStgOfficeBinary(hWnd, svFullPath[i], csExt, FileInfo);
				}

				// Office OpenXML 形式
				if (cmOfficeXML.Lookup(csExt, classId))
				{
					// プロパティセットを適用（Office OpenXML 形式）
					SetIPropStgOfficeOpenXML(hWnd, svFullPath[i], csExt, FileInfo, cmOfficeBin, cmOfficeXML);
				}

				// ファイルのタイムスタンプや属性を適用
				OnApplyItem(hWnd, svFullPath[i], FileInfo, dwFileAttributes);
			}
		}
		else
		{
			// 属性の取得に失敗した
			CStringW csMessage;
			GetErrorMessage(&csMessage);
			ShowMessageAndAppendLog(hWnd, _T("OnApply-GetAttribute"), csMessage, svFullPath[i]);
		}
	}

	return TRUE;
}

// 「日時を同じにする」がクリックされた
void OnChangedCheckSame(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	SYSTEMTIME st = {0};
	MENUITEMINFO mii;
	BOOL bEnable = NULL;

	switch (Button_GetCheck((HWND)lParam))
	{
	// オン
	case BST_CHECKED:
		bEnable = FALSE;

		// 「タイムスタンプ」グループボックスのチェックボックスをオンにする
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_CREATION), BST_CHECKED);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_WRITE), BST_CHECKED);
		Button_SetCheck(GetDlgItem(hWnd, IDC_CHECK_ACCESS), BST_CHECKED);

		// 「作成日時」を「更新（前回保存）日時」および「アクセス日時」にコピーする
		DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), &st);
		DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), GDT_VALID, &st);
		DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), GDT_VALID, &st);
		DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), &st);
		DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), GDT_VALID, &st);
		DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), GDT_VALID, &st);

		// 「適用」ボタンを有効にする
		PropSheet_Changed(GetParent(hWnd), hWnd);
		cmControlChange[hWnd] = TRUE;
		break;

	// オフ
	case BST_UNCHECKED:
		bEnable = TRUE;

		// DateTimePicker コントロールを編集可能とする
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), NULL);
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), NULL);
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), NULL);
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), NULL);
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), NULL);
		DateTime_SetFormat(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), NULL);

		// コンテキストメニューのチェックマークをオフにする
		mii = {0};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_STATE;
		mii.fState = MFS_UNCHECKED;

		HMENU hContextOverwrite;
		cmContextOverwrite.Lookup(hWnd, hContextOverwrite);
		for (UINT i = 0; i < (UINT)GetMenuItemCount(hContextOverwrite); i++)
		{
			SetMenuItemInfo(hContextOverwrite, i, TRUE, &mii);
		}

		break;
	}

	// 「タイムスタンプ」グループボックス配下のチェックボックスの状態を変更する
	EnableWindow(GetDlgItem(hWnd, IDC_CHECK_CREATION), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_CHECK_WRITE), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_CHECK_ACCESS), bEnable);

	// 「作成日時」, 「更新（前回保存）日時」および「アクセス日時」DateTimePicker コントロールの状態を変更する
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_DATE), True);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_DATE), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_ACCESS_DATE), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_CREATION_TIME), True);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_WRITE_TIME), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_ACCESS_TIME), bEnable);

	return;
}

// 「印刷日クリア」がクリックされた
void OnChangedCheckPrintClear(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	BOOL bEnable = NULL;

	// 「前回印刷日」をオフにする
	CheckDlgButton(hWnd, IDC_CHECK_PRINT, BST_UNCHECKED);
	OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_PRINT, BST_UNCHECKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_PRINT));

	switch (Button_GetCheck((HWND)lParam))
	{
	// オン
	case BST_CHECKED:

		// 「前回印刷日」を無効にする
		bEnable = FALSE;
		CheckDlgButton(hWnd, IDC_CHECK_PRINT, BST_CHECKED);
		break;

	// オフ
	case BST_UNCHECKED:

		// 「前回印刷日」を有効にする
		bEnable = TRUE;
		break;
	}

	// 「前回印刷日」の状態を変更する
	EnableWindow(GetDlgItem(hWnd, IDC_CHECK_PRINT), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_PRINT_DATE), FALSE);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_PRINT_TIME), FALSE);

	return;
}

// 「編集時間クリア」がクリックされた
void OnChangedCheckEditClear(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	BOOL bEnable = NULL;

	// 「総編集時間」をオフにする
	CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_UNCHECKED);
	OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_EDIT, BST_UNCHECKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_EDIT));

	switch (Button_GetCheck((HWND)lParam))
	{
	// オン
	case BST_CHECKED:

		// 「総編集時間」を無効にする
		bEnable = FALSE;
		CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_CHECKED);
		ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_HIDE);
		break;

	// オフ
	case BST_UNCHECKED:

		// 「総編集時間」を有効にする
		bEnable = TRUE;
		CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_UNCHECKED);
		break;
	}

	// 「総編集時間」の状態を変更する
	EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT), bEnable);
	EnableWindow(GetDlgItem(hWnd, IDC_GROUP_EDIT), FALSE);
	EnableWindow(GetDlgItem(hWnd, IDC_SPIN_EDIT), FALSE);
	EnableWindow(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), FALSE);

	return;
}

// アイテムのタイムスタンプや属性を適用
void OnApplyItem(HWND hWnd, CString csItemPath, WIN32_FILE_ATTRIBUTE_DATA& FileInfo, DWORD dwFileAttributes)
{
	//  「作成日時で上書きする」「更新（前回保存）日時で上書きする」が選択されているか確認する
	UINT nOverwriteId;
	cmOverwriteId.Lookup(hWnd, nOverwriteId);
	if (0 == nOverwriteId)
	{
		// 「作成日時」
		if (BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_CREATION))
		{
			FileInfo.ftCreationTime.dwHighDateTime = 0;
			FileInfo.ftCreationTime.dwLowDateTime = 0;
		}

		// 「更新（前回保存）日時」
		if (BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_WRITE))
		{
			FileInfo.ftLastWriteTime.dwHighDateTime = 0;
			FileInfo.ftLastWriteTime.dwLowDateTime = 0;
		}

		// 「アクセス日時」
		if (BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ACCESS))
		{
			FileInfo.ftLastAccessTime.dwHighDateTime = 0;
			FileInfo.ftLastAccessTime.dwLowDateTime = 0;
		}
	}

	// コントロールからアイテムの属性を取得
	GetAttributeControl(hWnd, IDC_CHECK_READONLY, FILE_ATTRIBUTE_READONLY, &FileInfo.dwFileAttributes);
	GetAttributeControl(hWnd, IDC_CHECK_HIDDEN, FILE_ATTRIBUTE_HIDDEN , &FileInfo.dwFileAttributes);
	GetAttributeControl(hWnd, IDC_CHECK_ARCHIVE, FILE_ATTRIBUTE_ARCHIVE, &FileInfo.dwFileAttributes);
	GetAttributeControl(hWnd, IDC_CHECK_INDEX, FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, &FileInfo.dwFileAttributes);
	GetAttributeControl(hWnd, IDC_CHECK_SYSTEM, FILE_ATTRIBUTE_SYSTEM, &FileInfo.dwFileAttributes);

	// いずれの属性も設定しない場合, FILE_ATTRIBUTE_NORMAL 属性とする
	if ((BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_READONLY)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_HIDDEN)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ARCHIVE)) &&
		(BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_INDEX)) &&
		(BST_UNCHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_SYSTEM)) &&
		!(FILE_ATTRIBUTE_OFFLINE & FileInfo.dwFileAttributes) &&
		!(FILE_ATTRIBUTE_TEMPORARY & FileInfo.dwFileAttributes))
	{
		FileInfo.dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
	}

	// タイムスタンプの適用が必要か確認する
	if ((BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_CREATION)) ||
		(BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_WRITE)) ||
		(BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ACCESS)))
	{
		DWORD dwFlagsAndAttributes;
		if (FILE_ATTRIBUTE_DIRECTORY & FileInfo.dwFileAttributes)
		{
			// フォルダ
			dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL | FILE_FLAG_BACKUP_SEMANTICS;
		}
		else
		{
			// ファイル
			dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL;
		}

		// アイテムをオープンする
		HANDLE hFile;
		hFile = CreateFile(csItemPath, FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ, NULL, OPEN_EXISTING, dwFlagsAndAttributes, NULL);
		if (INVALID_HANDLE_VALUE != hFile)
		{
			// タイムスタンプを適用する
			if (0 == SetFileTime(hFile, &FileInfo.ftCreationTime, &FileInfo.ftLastAccessTime, &FileInfo.ftLastWriteTime))
			{
				// タイムスタンプの適用に失敗した
				CStringW csMessage;
				GetErrorMessage(&csMessage);
				ShowMessageAndAppendLog(hWnd, _T("OnApplyItem-SetTimeStamp"), csMessage, csItemPath);
			}

			// アイテムをクローズする
			CloseHandle(hFile);
		}
		else
		{
			//　アイテムのオープンに失敗した
			CStringW csMessage;
			GetErrorMessage(&csMessage);
			ShowMessageAndAppendLog(hWnd, _T("OnApplyItem-FileOpen"), csMessage, csItemPath);
		}
	}

	// 属性の適用が必要か確認する
	if (dwFileAttributes != FileInfo.dwFileAttributes)
	{
		// 属性を適用する
		if (0 == SetFileAttributes((LPCTSTR)csItemPath, FileInfo.dwFileAttributes))
		{
			// 属性の適用に失敗した
			CStringW csMessage;
			GetErrorMessage(&csMessage);
			ShowMessageAndAppendLog(hWnd, _T("OnApplyItem-SetAttribute"), csMessage, csItemPath);
		}
	}

	return;
}

// ダイアログボックス破棄
BOOL OnDestroyDialog(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	// グローバル変数を破棄する
	cmFullPath.RemoveKey(hWnd);
	cmOverwriteId.RemoveKey(hWnd);
	cmControlChange.RemoveKey(hWnd);
	cmMessageShow.RemoveKey(hWnd);

	// アイコンハンドルを廃棄する
	HICON hIcon;
	cmIconHandle.Lookup(hWnd, hIcon);
	if (NULL != hIcon)
	{
		DestroyIcon(hIcon);
		cmIconHandle.RemoveKey(hWnd);
	}

	// メニューハンドルを廃棄する
	HMENU hContextSame;
	cmContextSame.Lookup(hWnd, hContextSame);
	if (NULL != hContextSame)
	{
		DestroyMenu(hContextSame);
		cmContextSame.RemoveKey(hWnd);
	}
	HMENU hContextSetTime;
	cmContextSetTime.Lookup(hWnd, hContextSetTime);
	if (NULL != hContextSetTime)
	{
		DestroyMenu(hContextSetTime);
		cmContextSetTime.RemoveKey(hWnd);
	}
	HMENU hContextOverwrite;
	cmContextOverwrite.Lookup(hWnd, hContextOverwrite);
	if (NULL != hContextOverwrite)
	{
		DestroyMenu(hContextOverwrite);
		cmContextOverwrite.RemoveKey(hWnd);
	}

	// フォントハンドルを破棄する
	// HFONT hFont;
	// cmFontHandle.Lookup(hWnd, hFont);
	// if (NULL != hFont)
	// {
	//     DeleteObject(hFont);
	// 	   cmFontHandle.RemoveKey(hWnd);
	// }

	return 0;
}

// 指定したフォルダ配下のアイテム数を確認
void RecursiveFindFile(HWND hWnd, CString csFolderPath, UINT &nFolderCount, UINT &nFileCount)
{
	WIN32_FIND_DATA FindData;
	HANDLE hFind = NULL;

	hFind = FindFirstFile(csFolderPath + _T("\\*"), &FindData);
	if ((INVALID_HANDLE_VALUE == hFind) || ((HANDLE)ERROR_FILE_NOT_FOUND == hFind))
	{
		// 選択されたフォルダが削除された
		CStringW csMessage;
		GetErrorMessage(&csMessage);
		ShowMessageAndAppendLog(hWnd, _T("RecursiveFindFile"), csMessage, csFolderPath);
		return;
	}

	do
	{
		// カレントや親フォルダは対象外とする
		if (0 == _wcsicmp(FindData.cFileName, _T(".")) || (0 == _wcsicmp(FindData.cFileName, _T(".."))))
		{
			continue;
		}

		if (FILE_ATTRIBUTE_DIRECTORY & FindData.dwFileAttributes)
		{
			// 再帰処理
			RecursiveFindFile(hWnd, csFolderPath + _T("\\") + FindData.cFileName, nFolderCount, nFileCount);
			nFolderCount++;
		}
		else
		{
			nFileCount++;
		}
	} while (FindNextFile(hFind, &FindData));
	FindClose(hFind);

	return;
}

// 指定したフォルダ配下のアイテムにタイムスタンプや属性を適用
void RecursiveOnApply(HWND hWnd, CString csFolderPath, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML)
{
	WIN32_FIND_DATA FindData;
	HANDLE hFind = NULL;
	WIN32_FILE_ATTRIBUTE_DATA FileInfo = {0};
	DWORD dwFileAttributes;

	hFind = FindFirstFile(csFolderPath + _T("\\*"), &FindData);
	if (INVALID_HANDLE_VALUE == hFind)
	{
		// 選択されたフォルダが削除された
		CStringW csMessage;
		GetErrorMessage(&csMessage);
		ShowMessageAndAppendLog(hWnd, _T("RecursiveOnApply-FindFirst"), csMessage, csFolderPath);
		return;
	}

	do
	{
		// カレントや親フォルダは対象外とする
		if (0 == _wcsicmp(FindData.cFileName, _T(".")) || (0 == _wcsicmp(FindData.cFileName, _T(".."))))
		{
			continue;
		}

		// アイテムの属性を取得する
		CString csItemPath = csFolderPath + _T("\\") + FindData.cFileName;
		if (0 != GetFileAttributesEx(csItemPath, GetFileExInfoStandard, &FileInfo))
		{
			// 属性を退避する
			dwFileAttributes = FileInfo.dwFileAttributes;

			// 「作成日時で上書きする」「更新（前回保存）日時で上書きする」が選択されているか確認する
			UINT nOverwriteId;
			cmOverwriteId.Lookup(hWnd, nOverwriteId);
			switch (nOverwriteId)
			{
				// 「作成日時で上書きする」
			case ID_OVERWRITE_CREATION:
				FileInfo.ftLastWriteTime = FileInfo.ftCreationTime;
				FileInfo.ftLastAccessTime = FileInfo.ftCreationTime;
				break;

				// 「更新（前回保存）日時で上書きする」
			case ID_OVERWRITE_WRITE:
				FileInfo.ftCreationTime = FileInfo.ftLastWriteTime;
				FileInfo.ftLastAccessTime = FileInfo.ftLastWriteTime;
				break;

				// いずれのメニューも選択されていない
			case 0:
				GetDTPControl(hWnd, IDC_DTP_CREATION_DATE, IDC_DTP_CREATION_TIME, FileInfo.ftCreationTime);
				GetDTPControl(hWnd, IDC_DTP_WRITE_DATE, IDC_DTP_WRITE_TIME, FileInfo.ftLastWriteTime);
				GetDTPControl(hWnd, IDC_DTP_ACCESS_DATE, IDC_DTP_ACCESS_TIME, FileInfo.ftLastAccessTime);
				break;
			}

			if (FILE_ATTRIBUTE_DIRECTORY & FindData.dwFileAttributes)
			{
				// フォルダ

				// 再帰処理
				RecursiveOnApply(hWnd, csItemPath, cmOfficeBin, cmOfficeXML);

				// 「サブフォルダを更新する」の状態を確認する
				if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_FOLDER))
				{
					// フォルダにタイムスタンプや属性を適用
					OnApplyItem(hWnd, csItemPath, FileInfo, dwFileAttributes);
				}
			}
			else
			{
				// ファイル

				// 「ファイルを更新する」「ゾーン情報を削除する」の状態を確認する
				if ((BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_FILE)) || (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ZONE)))
				{

					// 「読み取り専用」属性の解除
					RemoveReadOnly(hWnd, csItemPath, &dwFileAttributes);

					// 「ゾーン情報を削除する」の状態を確認する
					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_ZONE))
					{
						// ゾーン情報の削除
						DeleteZoneId(hWnd, csItemPath);
					}

					// 「ファイルを更新する」の状態を確認する
					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_FILE))
					{
						// 拡張子を取得する
						CString csExt;
						csExt = CPath(csItemPath).GetExtension().MakeLower();

						// Office Binary 形式
						CLSID classId;
						if (cmOfficeBin.Lookup(csExt, classId))
						{
							// プロパティセットを適用（Office Binary 形式）
							SetIPropStgOfficeBinary(hWnd, csItemPath, csExt, FileInfo);
						}

						// Office OpenXML 形式
						if (cmOfficeXML.Lookup(csExt, classId))
						{
							// プロパティセットを適用（Office OpenXML 形式）
							SetIPropStgOfficeOpenXML(hWnd, csItemPath, csExt, FileInfo, cmOfficeBin, cmOfficeXML);
						}

						// ファイルにタイムスタンプや属性を適用
						OnApplyItem(hWnd, csItemPath, FileInfo, dwFileAttributes);
					}
					else
					{
						// 「読み取り専用」属性を解除している場合は, 復元する
						if (dwFileAttributes != FileInfo.dwFileAttributes)
						{
							// 属性を適用する
							if (0 == SetFileAttributes((LPCTSTR)csItemPath, FileInfo.dwFileAttributes))
							{
								// 属性の適用に失敗した
								CStringW csMessage;
								GetErrorMessage(&csMessage);
								ShowMessageAndAppendLog(hWnd, _T("RecursiveOnApply-SetAttribute"), csMessage, csItemPath);
							}
						}
					}
				}
			}
		}
	} while (FindNextFile(hFind, &FindData));
	FindClose(hFind);

	return;
}

// プロパティセットを適用（Office Binary 形式）
void SetIPropStgOfficeBinary(HWND hWnd, CString csFilePath, CString csExt, WIN32_FILE_ATTRIBUTE_DATA& FileInfo)
{
	HRESULT hr;

	// ルートストレージオブジェクトをオープンする
	IPropertySetStorage* ppObject = {0};
	DWORD grfMode = STGM_READWRITE | STGM_SHARE_DENY_WRITE | STGM_DIRECT_SWMR;
	hr = StgOpenStorageEx(csFilePath, grfMode, STGFMT_STORAGE, 0, NULL, 0, IID_IPropertySetStorage, reinterpret_cast<void**>(&ppObject));
	if (SUCCEEDED(hr))
	{
		// プロパティセットをオープンする
		IPropertyStorage* ppprstg;
		hr = ppObject->Open(FMTID_SummaryInformation, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, &ppprstg);
		if (SUCCEEDED(hr))
		{
			FILETIME ft = {0};

			PROPSPEC rgpspec[1] = {0};
			rgpspec[0].ulKind = PRSPEC_PROPID;

			// PROPVARIANT 構造体を初期化する
			PROPVARIANT rgpropvar[1] = {0};
			PropVariantInit(&rgpropvar[0]);

			// 「作成日時」がオンの場合, プロパティセットを更新する
			if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_CREATION))
			{
				rgpspec[0].propid = PIDSI_CREATE_DTM;
				rgpropvar[0].vt = VT_FILETIME;
				rgpropvar[0].filetime = FileInfo.ftCreationTime;
				hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
				if (FAILED(hr))
				{
					// プロパティセットの更新に失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-CREATE_DTM"), csMessage, csFilePath);
				}
			}

			// 「更新（前回保存）日時」がオンの場合, プロパティセットを更新する
			if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_WRITE))
			{
				rgpspec[0].propid = PIDSI_LASTSAVE_DTM;
				rgpropvar[0].vt = VT_FILETIME;
				rgpropvar[0].filetime = FileInfo.ftLastWriteTime;
				hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
				if (FAILED(hr))
				{
					// プロパティセットの更新に失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-LASTSAVE_DTM"), csMessage, csFilePath);
				}
			}

			// 「前回印刷日」がオンの場合, プロパティセットを更新する
			if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_PRINT))
			{
				rgpspec[0].propid = PIDSI_LASTPRINTED;

				if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_PRINT_CLEAR))
				{
					// 印刷日クリア
					rgpropvar[0].vt = VT_EMPTY;
					ft.dwHighDateTime = 0;
					ft.dwLowDateTime = 0;
				}
				else
				{
					rgpropvar[0].vt = VT_FILETIME;
					GetDTPControl(hWnd, IDC_DTP_PRINT_DATE, IDC_DTP_PRINT_TIME, ft);
				}

				rgpropvar[0].filetime = ft;
				hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
				if (FAILED(hr))
				{
					// プロパティセットの更新に失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-LASTPRINTED"), csMessage, csFilePath);
				}
			}

			// Excel は対象外
			if (!((_T(".xls") == csExt) || (_T(".xlt") == csExt)))
			{
				// 「総編集時間」がオンの場合, プロパティセットを更新する
				if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_EDIT))
				{
					rgpspec[0].propid = PIDSI_EDITTIME;

					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_EDIT_CLEAR))
					{
						// 編集時間クリア
						rgpropvar[0].vt = VT_EMPTY;
						ft.dwHighDateTime = 0;
						ft.dwLowDateTime = 0;
					}
					else
					{
						rgpropvar[0].vt = VT_FILETIME;
						GetEditControl(hWnd, ft);
					}

					rgpropvar[0].filetime = ft;
					hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
					if (FAILED(hr))
					{
						// プロパティセットの更新に失敗した
						CStringW csMessage;
						GetHRESULTValue(hr, &csMessage);
						ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-EDITTIME"), csMessage, csFilePath);
					}
				}
			}

			// PROPVARIANT 構造体を解放する
			PropVariantClear(&rgpropvar[0]);

			// プロパティセットをクローズする
			ppprstg->Release();
		}
		else
		{
			// プロパティセットのオープンに失敗した
			CStringW csMessage;
			GetHRESULTValue(hr, &csMessage);
			ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-Open"), csMessage, csFilePath);
		}

		// ルートストレージオブジェクトをクローズする
		ppObject->Release();
	}
	else
	{
		// ルートストレージオブジェクトのオープンに失敗した
		CStringW csMessage;
		GetHRESULTValue(hr, &csMessage);
		ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeBinary-StgOpenStorageEx"), csMessage, csFilePath);
	}

	return;
}

// プロパティセットから日時を取得（Office Binary 形式）
void GetIPropStgOfficeBinary(HWND hWnd, CString csFilePath, CString csExt)
{
	HRESULT hr;

	// ルートストレージオブジェクトをオープンする
	IPropertySetStorage* ppObject = {0};
	DWORD grfMode = STGM_READ | STGM_SHARE_DENY_NONE | STGM_DIRECT_SWMR;
	hr = StgOpenStorageEx(csFilePath, grfMode, STGFMT_STORAGE, 0, NULL, 0, IID_IPropertySetStorage, reinterpret_cast<void**>(&ppObject));
	if (SUCCEEDED(hr))
	{
		// プロパティセットをオープンする
		IPropertyStorage* ppprstg;
		hr = ppObject->Open(FMTID_SummaryInformation, STGM_READ | STGM_SHARE_EXCLUSIVE, &ppprstg);
		if (SUCCEEDED(hr))
		{
			PROPSPEC rgpspec[1] = {0};
			rgpspec[0].ulKind = PRSPEC_PROPID;

			// PROPVARIANT 構造体を初期化する
			PROPVARIANT rgpropvar[1] = {0};
			PropVariantInit(&rgpropvar[0]);

			// 「前回印刷日」
			rgpspec[0].propid = PIDSI_LASTPRINTED;
			rgpropvar[0].vt = VT_FILETIME;
			hr = ppprstg->ReadMultiple(1, &rgpspec[0], &rgpropvar[0]);
			if (SUCCEEDED(hr))
			{
				if (VT_EMPTY != rgpropvar[0].vt)
				{
					// 「前回印刷日」をオンにする
					CheckDlgButton(hWnd, IDC_CHECK_PRINT, BST_CHECKED);
					OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_PRINT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_PRINT));

					// 日時をコントロールに設定する
					SetDTPControl(hWnd, IDC_DTP_PRINT_DATE, IDC_DTP_PRINT_TIME, rgpropvar[0].filetime);
				}
			}
			else
			{
				// プロパティセットの取得に失敗した
				CStringW csMessage;
				GetHRESULTValue(hr, &csMessage);
				ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeBinary-LASTPRINTED"), csMessage, csFilePath);
			}

			// Excel は対象外
			if (!((_T(".xls") == csExt) || (_T(".xlt") == csExt)))
			{

				// 「総編集時間」
				rgpspec[0].propid = PIDSI_EDITTIME;
				rgpropvar[0].vt = VT_FILETIME;
				hr = ppprstg->ReadMultiple(1, &rgpspec[0], &rgpropvar[0]);
				if (SUCCEEDED(hr))
				{
					if (VT_EMPTY != rgpropvar[0].vt)
					{
						// 「総編集時間」をオンにする
						CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_CHECKED);
						OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_EDIT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_EDIT));

						// 日時をコントロールに設定する
						SetEditControl(hWnd, rgpropvar[0].filetime);
					}
				}
				else
				{
					// プロパティセットの取得に失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeBinary-PIDSI_EDITTIME"), csMessage, csFilePath);
				}
			}
			else
			{
				// 「総編集時間」「編集時間クリア」を無効にする
				EnableWindow(GetDlgItem(hWnd, IDC_GROUP_EDIT), FALSE);
				ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_HIDE);
				EnableWindow(GetDlgItem(hWnd, IDC_SPIN_EDIT), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT), FALSE);
				EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT_CLEAR), FALSE);
			}

			// PROPVARIANT 構造体を解放する
			PropVariantClear(&rgpropvar[0]);

			// プロパティセットをクローズする
			ppprstg->Release();
		}
		else
		{
			// プロパティセットのオープンに失敗した
			CStringW csMessage;
			GetHRESULTValue(hr, &csMessage);
			ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeBinary-Open"), csMessage, csFilePath);
		}

		// ルートストレージオブジェクトをクローズする
		ppObject->Release();
	}
	else
	{
		// ルートストレージオブジェクトのオープンに失敗した
		CStringW csMessage;
		GetHRESULTValue(hr, &csMessage);
		ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeBinary-StgOpenStorageEx"), csMessage, csFilePath);
	}

	return;
}

// プロパティセットを適用（Office OpenXML 形式）
void SetIPropStgOfficeOpenXML(HWND hWnd, CString csFilePath, CString csExt, WIN32_FILE_ATTRIBUTE_DATA& FileInfo, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML)
{
	HRESULT hr;

	// COM クラスを取得する
	CLSID classId;
	cmOfficeXML.Lookup(csExt, classId);
	IPersistFile* ppv = {0};
	hr = CoCreateInstance(classId, NULL, CLSCTX_INPROC_SERVER, IID_IPersistFile, reinterpret_cast<void**>(&ppv));
	if (SUCCEEDED(hr))
	{
		// ファイルをオープンする
		hr = ppv->Load(csFilePath, STGM_READWRITE | STGM_SHARE_DENY_WRITE | STGM_DIRECT_SWMR);
		if (SUCCEEDED(hr))
		{
			// インターフェイスを取得する
			IPropertySetStorage* ppvObject = {0};
			hr = ppv->QueryInterface(IID_IPropertySetStorage, reinterpret_cast<void**>(&ppvObject));
			if (SUCCEEDED(hr))
			{
				// プロパティセットをオープンする
				IPropertyStorage* ppprstg;
				hr = ppvObject->Open(FMTID_SummaryInformation, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, &ppprstg);
				if (SUCCEEDED(hr))
				{
					FILETIME ft = {0};

					PROPSPEC rgpspec[1] = {0};
					rgpspec[0].ulKind = PRSPEC_PROPID;

					// PROPVARIANT 構造体を初期化する
					PROPVARIANT rgpropvar[1] = {0};
					PropVariantInit(&rgpropvar[0]);

					// 「作成日時」がオンの場合, プロパティセットを更新する
					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_CREATION))
					{
						rgpspec[0].propid = PIDSI_CREATE_DTM;
						rgpropvar[0].vt = VT_FILETIME;
						rgpropvar[0].filetime = FileInfo.ftCreationTime;
						hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
						if (FAILED(hr))
						{
							// プロパティセットの更新に失敗した
							CStringW csMessage;
							GetHRESULTValue(hr, &csMessage);
							ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-CREATE_DTM"), csMessage, csFilePath);
						}
					}

					// 「更新（前回保存）日時」がオンの場合, プロパティセットを更新する
					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_WRITE))
					{
						rgpspec[0].propid = PIDSI_LASTSAVE_DTM;
						rgpropvar[0].vt = VT_FILETIME;
						rgpropvar[0].filetime = FileInfo.ftLastWriteTime;
						hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
						if (FAILED(hr))
						{
							// プロパティセットの更新に失敗した
							CStringW csMessage;
							GetHRESULTValue(hr, &csMessage);
							ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-LASTSAVE_DTM"), csMessage, csFilePath);
						}
					}

					// 「前回印刷日」がオンの場合, プロパティセットを更新する
					if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_PRINT))
					{
						rgpspec[0].propid = PIDSI_LASTPRINTED;

						if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_PRINT_CLEAR))
						{
							rgpropvar[0].vt = VT_EMPTY;
							ft.dwHighDateTime = 0;
							ft.dwLowDateTime = 0;
						}
						else
						{
							rgpropvar[0].vt = VT_FILETIME;
							GetDTPControl(hWnd, IDC_DTP_PRINT_DATE, IDC_DTP_PRINT_TIME, ft);
						}

						rgpropvar[0].filetime = ft;
						hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
						if (FAILED(hr))
						{
							// プロパティセットの更新に失敗した
							CStringW csMessage;
							GetHRESULTValue(hr, &csMessage);
							ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-LASTPRINTED"), csMessage, csFilePath);
						}
					}

					// Excel は対象外
					if (!((_T(".xlsx") == csExt) || (_T(".xlsm") == csExt) || (_T(".xlsb") == csExt) || (_T(".xltx") == csExt) || (_T(".xltm") == csExt)))
					{
						// 「総編集時間」がオンの場合, プロパティセットを更新する
						if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_EDIT))
						{
							rgpspec[0].propid = PIDSI_EDITTIME;

							if (BST_CHECKED == IsDlgButtonChecked(hWnd, IDC_CHECK_EDIT_CLEAR))
							{
								rgpropvar[0].vt = VT_EMPTY;
								ft.dwHighDateTime = 0;
								ft.dwLowDateTime = 0;
							}
							else
							{
								rgpropvar[0].vt = VT_FILETIME;
								GetEditControl(hWnd, ft);
							}

							rgpropvar[0].filetime = ft;
							hr = ppprstg->WriteMultiple(1, &rgpspec[0], &rgpropvar[0], 0);
							if (FAILED(hr))
							{
								// プロパティセットの更新に失敗した
								CStringW csMessage;
								GetHRESULTValue(hr, &csMessage);
								ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-EDITTIME"), csMessage, csFilePath);
							}
						}
					}

					// PROPVARIANT 構造体を解放する
					PropVariantClear(&rgpropvar[0]);

					// プロパティセットをクローズする
					ppprstg->Release();
				}
				else
				{
					// プロパティセットのオープンに失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-Open"), csMessage, csFilePath);
				}

				// インターフェイスを解放する
				ppvObject->Release();
			}
			else
			{
				// インターフェイスの取得に失敗した
				CStringW csMessage;
				GetHRESULTValue(hr, &csMessage);
				ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-QueryInterface"), csMessage, csFilePath);
			}

			// ファイルをクローズする
			ppv->Release();
		}
		else
		{
			// ファイルのオープンに失敗した
			CStringW csMessage;
			GetHRESULTValue(hr, &csMessage);
			ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-Load"), csMessage, csFilePath);
		}
	}
	else
	{
		// COM クラスの取得に失敗した
		CStringW csMessage;
		GetHRESULTValue(hr, &csMessage);
		ShowMessageAndAppendLog(hWnd, _T("SetIPropStgOfficeOpenXML-CoCreateInstance"), csMessage, csFilePath);
	}

	return;
}

// プロパティセットから日時を取得（Office OpenXML 形式）
void GetIPropStgOfficeOpenXML(HWND hWnd, CString csFilePath, CString csExt, CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML)
{
	HRESULT hr;

	// COM クラスを取得する
	CLSID classId;
	cmOfficeXML.Lookup(csExt, classId);
	IPersistFile* ppv = {0};
	hr = CoCreateInstance(classId, NULL, CLSCTX_INPROC_SERVER, IID_IPersistFile, reinterpret_cast<void**>(&ppv));
	if (SUCCEEDED(hr))
	{
		// ファイルをオープンする
		hr = ppv->Load(csFilePath, STGM_READ | STGM_SHARE_DENY_NONE | STGM_DIRECT_SWMR);
		if (SUCCEEDED(hr))
		{
			// インターフェイスを取得する
			IPropertySetStorage* ppvObject = {0};
			hr = ppv->QueryInterface(IID_IPropertySetStorage, reinterpret_cast<void**>(&ppvObject));
			if (SUCCEEDED(hr))
			{
				// プロパティセットをオープンする
				IPropertyStorage* ppprstg;
				hr = ppvObject->Open(FMTID_SummaryInformation, STGM_READ | STGM_SHARE_EXCLUSIVE, &ppprstg);
				if (SUCCEEDED(hr))
				{
					PROPSPEC rgpspec[1] = {0};
					rgpspec[0].ulKind = PRSPEC_PROPID;

					// PROPVARIANT 構造体を初期化する
					PROPVARIANT rgpropvar[1] = {0};
					PropVariantInit(&rgpropvar[0]);

					// 「前回印刷日」
					rgpspec[0].propid = PIDSI_LASTPRINTED;
					rgpropvar[0].vt = VT_FILETIME;
					hr = ppprstg->ReadMultiple(1, &rgpspec[0], &rgpropvar[0]);
					if (SUCCEEDED(hr))
					{
						if (VT_EMPTY != rgpropvar[0].vt)
						{
							// 「前回印刷日」をオンにする
							CheckDlgButton(hWnd, IDC_CHECK_PRINT, BST_CHECKED);
							OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_PRINT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_PRINT));
							SetDTPControl(hWnd, IDC_DTP_PRINT_DATE, IDC_DTP_PRINT_TIME, rgpropvar[0].filetime);
						}
					}
					else
					{
						// プロパティセットの取得に失敗した
						CStringW csMessage;
						GetHRESULTValue(hr, &csMessage);
						ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-LASTPRINTED"), csMessage, csFilePath);
					}

					// Excel は対象外
					if (!((_T(".xlsx") == csExt) || (_T(".xlsm") == csExt) || (_T(".xlsb") == csExt) || (_T(".xltx") == csExt) || (_T(".xltm") == csExt)))
					{
						// 「総編集時間」
						rgpspec[0].propid = PIDSI_EDITTIME;
						rgpropvar[0].vt = VT_FILETIME;
						hr = ppprstg->ReadMultiple(1, &rgpspec[0], &rgpropvar[0]);
						if (SUCCEEDED(hr))
						{
							if (VT_EMPTY != rgpropvar[0].vt)
							{
								// 「総編集時間」をオンにする
								CheckDlgButton(hWnd, IDC_CHECK_EDIT, BST_CHECKED);
								OnLeftClicked(hWnd, MAKEWPARAM(IDC_CHECK_EDIT, BN_CLICKED), (LPARAM)GetDlgItem(hWnd, IDC_CHECK_EDIT));

								// 日時をコントロールに設定する
								SetEditControl(hWnd, rgpropvar[0].filetime);
							}
						}
						else
						{
							// プロパティセットの取得に失敗した
							CStringW csMessage;
							GetHRESULTValue(hr, &csMessage);
							ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-EDITTIME"), csMessage, csFilePath);
						}
					}
					else
					{
						// 「総編集時間」「編集時間クリア」を無効にする
						EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT), FALSE);
						EnableWindow(GetDlgItem(hWnd, IDC_GROUP_EDIT), FALSE);
						ShowWindow(GetDlgItem(hWnd, IDC_EDIT_EDIT), SW_HIDE);
						EnableWindow(GetDlgItem(hWnd, IDC_SPIN_EDIT), FALSE);
						EnableWindow(GetDlgItem(hWnd, IDC_CHECK_EDIT_CLEAR), FALSE);
					}

					// PROPVARIANT 構造体を解放する
					PropVariantClear(&rgpropvar[0]);

					// プロパティセットをクローズする
					ppprstg->Release();
				}
				else
				{
					// プロパティセットのオープンに失敗した
					CStringW csMessage;
					GetHRESULTValue(hr, &csMessage);
					ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-Open"), csMessage, csFilePath);
				}

				// インターフェイスを解放する
				ppvObject->Release();
			}
			else
			{
				// インターフェイスの取得に失敗した
				CStringW csMessage;
				GetHRESULTValue(hr, &csMessage);
				ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-QueryInterface"), csMessage, csFilePath);
			}

			// ファイルをクローズする
			ppv->Release();
		}
		else
		{
			// ファイルのオープンに失敗した
			CStringW csMessage;
			GetHRESULTValue(hr, &csMessage);
			ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-Load"), csMessage, csFilePath);
		}
	}
	else
	{
		// COM クラスの取得に失敗した
		CStringW csMessage;
		GetHRESULTValue(hr, &csMessage);
		ShowMessageAndAppendLog(hWnd, _T("GetIPropStgOfficeOpenXML-CoCreateInstance"), csMessage, csFilePath);
	}

	return;
}

// プロパティハンドラのCLSID を取得
void QueryHandlerClassId(CAtlMap<CString, CLSID>& cmOfficeBin, CAtlMap<CString, CLSID>& cmOfficeXML)
{
	// Office Binary 形式で対応できる拡張子
	cmOfficeBin[_T(".doc")] = CLSID_NULL;	// Word 文書
	cmOfficeBin[_T(".dot")] = CLSID_NULL;	// Word 97-2003 テンプレート
	cmOfficeBin[_T(".xls")] = CLSID_NULL;	// Excel 97-2003 ブック
	cmOfficeBin[_T(".xlt")] = CLSID_NULL;	// Excel 97-2003 テンプレート
	cmOfficeBin[_T(".ppt")] = CLSID_NULL;	// PowerPoint 97-2003 プレゼンテーション
	cmOfficeBin[_T(".pot")] = CLSID_NULL;	// PowerPoint 97-2003 テンプレート

	// プロパティハンドラのCLSID を取得する
	CString extension[] = {
		_T(".docx"),	// Word 文書
		_T(".docm"),	// Word マクロ有効文書
		_T(".dotx"),	// Word テンプレート
		_T(".dotm"),	// Word マクロ有効テンプレート
		_T(".xlsx"),	// Excel ブック
		_T(".xlsm"),	// Excel マクロ有効ブック
		_T(".xlsb"),	// Excel バイナリブック
		_T(".xltx"),	// Excel テンプレート
		_T(".xltm"),	// Excel マクロ有効テンプレート
		_T(".pptx"),	// PowerPoint プレゼンテーション
		_T(".pptm"),	// PowerPoint マクロ有効プレゼンテーション
		_T(".potx"),	// PowerPoint テンプレート
		_T(".potm")		// PowerPoint マクロ有効テンプレート
	};
	for (auto& clsId : extension) {
		CRegKey cReg;
		LONG lStat;
		CLSID guidValue;

		// Office 製品のインストール状況によりプロパティハンドラが存在しない可能性があるため
		// 拡張子ごとに確認する必要がある

		// レジストリキーの存在を確認する
		lStat = cReg.Open(HKEY_CLASSES_ROOT, clsId + _T("\\ShellEx\\PropertyHandler"), KEY_QUERY_VALUE);
		if (ERROR_SUCCESS == lStat)
		{
			// CLSID を取得する
			cReg.QueryGUIDValue(NULL, guidValue);
			cmOfficeXML[clsId] = guidValue;

			// cRegKey を解放する
			cReg.Close();
		}
	}
}

// 日時をDateTimePicker コントロールに設定
void SetDTPControl(HWND hWnd, UINT nIDDlgDate, UINT nIDDlgTime, FILETIME FileTime)
{
	FILETIME ftLocal;
	SYSTEMTIME st;

	// FileTime からSystemTime 構造体に変換する
	FileTimeToLocalFileTime(&FileTime, &ftLocal);
	FileTimeToSystemTime(&ftLocal, &st);

	// 日時をDateTimePicker コントロールに設定する
	DateTime_SetSystemtime(GetDlgItem(hWnd, nIDDlgDate), GDT_VALID, (LPARAM)&st);
	DateTime_SetSystemtime(GetDlgItem(hWnd, nIDDlgTime), GDT_VALID, (LPARAM)&st);

	return;
}

// DateTimePicker コントロールから日時を取得
void GetDTPControl(HWND hWnd, UINT nIDDlgDate, UINT nIDDlgTime, FILETIME& FileTime)
{
	// DateTimePicker コントロールから日時を取得する
	SYSTEMTIME st = {0}, stDate = {0}, stTime = {0};
	DateTime_GetSystemtime(GetDlgItem(hWnd, nIDDlgDate), (LPARAM)&stDate);
	DateTime_GetSystemtime(GetDlgItem(hWnd, nIDDlgTime), (LPARAM)&stTime);

	// 2 つのDateTimePicker コントロールから取得した日時をSystemTime 構造体にまとめる
	st.wYear = stDate.wYear;
	st.wMonth = stDate.wMonth;
	st.wDay = stDate.wDay;
	st.wHour = stTime.wHour;
	st.wMinute = stTime.wMinute;
	st.wSecond = stTime.wSecond;

	// SystemTime からFileTime 構造体に変換する
	FILETIME ftLocal;
	SystemTimeToFileTime(&st, &ftLocal);
	LocalFileTimeToFileTime(&ftLocal, &FileTime);

	return;
}

// 総編集時間をコントロールに設定
void SetEditControl(HWND hWnd, FILETIME FileTime)
{
	// 編集時間をコントロールに設定する
	SYSTEMTIME st;
	FileTimeToSystemTime(&FileTime, &st);
	DateTime_SetSystemtime(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), GDT_VALID, &st);

	// 24時間をこえる分を日数に換算する
	ULONGLONG lDuration;
	lDuration = (((ULONGLONG)FileTime.dwHighDateTime) << 32) + FileTime.dwLowDateTime;

	// PIDSI_EDITTIME はナノ秒単位
	// 経過日数 = 総編集時間 / 10000000 ナノ秒（= 1秒）/ 86400 秒（= 1日）
	UINT nDuration;
	nDuration = (UINT)floor(lDuration / 10000000.0 / 86400.0);

	// 編集日数をコントロールに設定する
	CString csString;
	csString.Format(_T("%d"), nDuration);
	SetWindowText(GetDlgItem(hWnd, IDC_EDIT_EDIT), csString);

	return;
}

// コントロールから総編集時間を取得
void GetEditControl(HWND hWnd, FILETIME& FileTime)
{
	// コントロールから編集時間を取得する
	SYSTEMTIME st = {0};
	DateTime_GetSystemtime(GetDlgItem(hWnd, IDC_DTP_EDIT_TIME), (LPARAM)&st);
	st.wYear = 1601;
	st.wMonth = 1;
	st.wDay = 1;
	SystemTimeToFileTime(&st, &FileTime);

	// コントロールから経過日を取得する
	CString csWndTitle;
	UINT nCch;
	nCch = GetWindowTextLength(GetDlgItem(hWnd, IDC_EDIT_EDIT)) + 1;
	csWndTitle.GetBuffer(nCch);
	GetWindowText(GetDlgItem(hWnd, IDC_EDIT_EDIT), (LPWSTR)(LPCWSTR)csWndTitle, nCch);
	csWndTitle.ReleaseBuffer();

	// PIDSI_EDITTIME はナノ秒単位
	// 経過時間 = 経過日数 * 86400 秒（= 1日）* 10000000 ナノ秒（= 1秒）
	ULONGLONG lDuration;
	lDuration = (((ULONGLONG)FileTime.dwHighDateTime) << 32) + FileTime.dwLowDateTime + (ULONGLONG)(_wtoi(csWndTitle) * 86400.0 * 10000000.0);
	
	// FILETIME 構造体に変換する
	FileTime.dwHighDateTime = ((lDuration >> 32) & 0xffffffff);
	FileTime.dwLowDateTime = (lDuration & 0xffffffff);

	return;
}

// アイテムの属性をコントロールに設定
void SetAttributeControl(HWND hWnd, UINT nIDDlgItem, DWORD dwFileAttribute, DWORD dwFileAttributes)
{
	if (nIDDlgItem != IDC_CHECK_INDEX)
	{
		// 「インデックス」以外
		if (dwFileAttribute & dwFileAttributes)
		{
			// オン
			CheckDlgButton(hWnd, nIDDlgItem, BST_CHECKED);
		}
		else
		{
			// オフ
			CheckDlgButton(hWnd, nIDDlgItem, BST_UNCHECKED);
		}
	}
	else
	{
		// 「インデックス」のみ
		if (dwFileAttribute & dwFileAttributes)
		{
			// オフ
			CheckDlgButton(hWnd, nIDDlgItem, BST_UNCHECKED);
		}
		else
		{
			// オン
			CheckDlgButton(hWnd, nIDDlgItem, BST_CHECKED);
		}
	}

	return;
}

// コントロールからアイテムの属性を取得
void GetAttributeControl(HWND hWnd, UINT nIDDlgItem, DWORD dwFileAttribute, DWORD* pdwFileAttributes)
{
	if (nIDDlgItem != IDC_CHECK_INDEX)
	{
		// 「インデックス」以外
		switch (Button_GetCheck(GetDlgItem(hWnd, nIDDlgItem)))
		{
			// オン
		case BST_CHECKED:
			*pdwFileAttributes = *pdwFileAttributes | dwFileAttribute;
			break;

			// オフ
		case BST_UNCHECKED:
			*pdwFileAttributes = *pdwFileAttributes & ~dwFileAttribute;
			break;
		}
	}
	else
	{
		// 「インデックス」のみ
		switch (Button_GetCheck(GetDlgItem(hWnd, nIDDlgItem)))
		{
			// オン
		case BST_CHECKED:
			*pdwFileAttributes = *pdwFileAttributes & ~dwFileAttribute;
			break;

			// オフ
		case BST_UNCHECKED:
			*pdwFileAttributes = *pdwFileAttributes | dwFileAttribute;
			break;
		}
	}

	return;
}

// 「読み取り専用」属性の解除
void RemoveReadOnly(HWND hWnd, CString csFilePath, DWORD* pdwFileAttributes)
{
	// ファイルの「読み取り専用」属性が有効な場合は, 解除する
	// （有効の場合, タイムスタンプ, 属性などの更新ができないため）
	if (FILE_ATTRIBUTE_READONLY & *pdwFileAttributes)
	{
		*pdwFileAttributes = *pdwFileAttributes & ~FILE_ATTRIBUTE_READONLY;
		if (0 == SetFileAttributes((LPCTSTR)csFilePath, *pdwFileAttributes))
		{
			// ファイルの「読み取り属性」解除に失敗した
			CStringW csMessage;
			GetErrorMessage(&csMessage);
			ShowMessageAndAppendLog(hWnd, _T("RemoveReadOnly"), csMessage, csFilePath);
		}
	}

	return;
}

// ゾーン情報の削除
void DeleteZoneId(HWND hWnd, CString csFilePath)
{
	// ゾーン情報の有無を確認する
	if (CPath(csFilePath + ADS_ZONE_ID).FileExists())
	{
		// ゾーン情報を削除する
		if (0 == DeleteFile(csFilePath + ADS_ZONE_ID))
		{
			// ゾーン情報の削除に失敗した
			CStringW csMessage;
			GetErrorMessage(&csMessage);
			ShowMessageAndAppendLog(hWnd, _T("DeleteZoneId"), csMessage, csFilePath);
		}
	}

	return;
}

// メッセージ表示およびログ出力
void ShowMessageAndAppendLog(HWND hWnd, CStringW csLocation, CStringW csMessage, CStringW csItemname)
{
	CString csLogFilename;
	CString csHWND;
	HANDLE hFile;
	DWORD NumberOfBytesWritten;

	// ログファイルのパス名を決定する
	csHWND.Format(_T("%p"), hWnd);
	GetEnvironmentVariable(_T("TEMP"), csLogFilename.GetBufferSetLength(NTFS_MAX_PATH + 1), NTFS_MAX_PATH + 1);
	csLogFilename.ReleaseBuffer();
	csLogFilename += _T("\\") + CString(MAKEINTRESOURCE(IDS_LOG_FILENAME));
	csLogFilename.Replace(_T("%HWND%"), csHWND);

	// ログファイルの存在を確認する
	if (CPath(csLogFilename).FileExists())
	{
		// 既存ファイルをオープンする
		hFile = CreateFile(csLogFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

		// 追記するため, ファイルポインタを最後に移動する
		SetFilePointer(hFile, 0, NULL, FILE_END);
	}
	else
	{
		// 新規にファイルを作成する
		hFile = CreateFile(csLogFilename, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);

		// UTF-16LE のBOM を書き込む
		WORD wBOM = 0xFEFF;
		WriteFile(hFile, &wBOM, sizeof(wBOM), &NumberOfBytesWritten, NULL);
	}

	// 現在時刻を書き込む
	SYSTEMTIME st;
	GetLocalTime(&st);
	CStringW csDateTime;
	csDateTime.Format(_T("%04d/%02d/%02d %02d:%02d:%02d  "), st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
	WriteFile(hFile, csDateTime, csDateTime.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// 発生個所を書き込む
	WriteFile(hFile, csLocation, csLocation.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// 空白を書き込む
	CStringW csSpace = _T(" ");
	WriteFile(hFile, csSpace, csSpace.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// メッセージを書き込む
	WriteFile(hFile, csMessage, csMessage.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// アイテム名を書き込む
	WriteFile(hFile, csItemname, csItemname.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// 改行コードを書き込む
	CStringW csCrLf = _T("\r\n");
	WriteFile(hFile, csCrLf, csCrLf.GetLength() * sizeof(TCHAR), &NumberOfBytesWritten, NULL);

	// ファイルをクローズする
	CloseHandle(hFile);

	// メッセージを表示する
	BOOL bMessageShow;
	cmMessageShow.Lookup(hWnd, bMessageShow);
	if (bMessageShow)
	{
		UINT nRetCode;
		UINT uType = MB_OKCANCEL | MB_ICONWARNING | MB_DEFBUTTON1;
		CString csText;
		csText = csMessage + _T("\r\n") + csItemname + _T("\r\n\r\n") + CString(MAKEINTRESOURCE(IDS_MESSAGE_COMPLEMENT));
		csText.Replace(_T("%HWND%"), csHWND);
		nRetCode = MessageBox(hWnd, csText, CString(MAKEINTRESOURCE(IDS_PROJNAME)), uType);

		// 「キャンセル」ボタンがクリックされたか確認する
		if (IDCANCEL == nRetCode)
		{
			// メッセージ表示を抑止する
			cmMessageShow[hWnd] = FALSE;
		}
	}

	return;
}

// エラーメッセージを取得
void GetErrorMessage(CStringW* csMessage)
{
	// エラーコードを取得する
	DWORD dwMessageId;
	dwMessageId = GetLastError();

	// エラーメッセージを取得する
	DWORD dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
	LPVOID lpBuffer = {0};
	FormatMessage(dwFlags, NULL, dwMessageId, NULL, (LPTSTR)&lpBuffer, 0, NULL);
	*csMessage = (LPCTSTR)lpBuffer;
	LocalFree(lpBuffer);

	// 前後の空白を取り除く
	*csMessage->Trim();

	return;
}

// HRESULT メッセージを取得
void GetHRESULTValue(HRESULT hr, CStringW* csMessage)
{
	switch (hr)
	{
	case E_ABORT:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_ABORT);
		break;
	case E_ACCESSDENIED:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_ACCESSDENIED);
		break;
	case E_FAIL:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_FAIL);
		break;
	case E_HANDLE:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_HANDLE);
		break;
	case E_INVALIDARG:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_INVALIDARG);
		break;
	case E_NOINTERFACE:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_NOINTERFACE);
		break;
	case E_NOTIMPL:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_NOTIMPL);
		break;
	case E_OUTOFMEMORY:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_OUTOFMEMORY);
		break;
	case E_POINTER:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_POINTER);
		break;
	case E_UNEXPECTED:
		*csMessage = (CStringW)MAKEINTRESOURCE(IDS_E_UNEXPECTED);
		break;
	default:
		csMessage->Format(_T("HRESULT = %08X  "), hr);
		break;
	}

	return;
}
