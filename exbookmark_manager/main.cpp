#include "pch.h"
#include "resource.h"

namespace apn::exbookmark
{
	//
	// メインダイアログのウィンドウハンドルです。
	//
	HWND main_dialog = {};

	//
	// COMサーバー(dll)があるフォルダです。
	//
	std::filesystem::path server_dir;

	//
	// COMサーバー(dll)のパスです。
	//
	std::filesystem::path server_path;

	//
	// 指定されたモジュールのファイルパスを返します。
	//
	std::filesystem::path get_module_file_name(HINSTANCE instance)
	{
		wchar_t module_file_name[MAX_PATH] = {};
		::GetModuleFileNameW(instance,
			module_file_name, std::size(module_file_name));
		return module_file_name;
	}

	//
	// COMサーバーをインストールします。
	//
	BOOL on_install()
	{
		auto param = std::format(LR"***("{}")***", server_path.wstring());

		::ShellExecuteW(main_dialog, L"runas",
			L"regsvr32.exe", param.c_str(), server_dir.c_str(), SW_SHOW);

		return TRUE;
	}

	//
	// COMサーバーをアンインストールします。
	//
	BOOL on_uninstall()
	{
		auto param = std::format(LR"***(/u "{}")***", server_path.wstring());

		::ShellExecuteW(main_dialog, L"runas",
			L"regsvr32.exe", param.c_str(), server_dir.c_str(), SW_SHOW);

		return TRUE;
	}

	//
	// ファイル選択ダイアログを表示します。
	//
	BOOL on_test_common_dialog()
	{
		std::wstring file_name(MAX_PATH, L'\0');

		// ファイル選択ダイアログ用の構造体を設定します。
		OPENFILENAMEW ofn = { sizeof(ofn) };
		ofn.hwndOwner = main_dialog;
		ofn.Flags = OFN_FILEMUSTEXIST;
		ofn.lpstrTitle = L"ファイルを選択";
		ofn.lpstrFile = file_name.data();
		ofn.nMaxFile = (DWORD)file_name.size();
		ofn.lpstrFilter = L"すべてのファイル (*.*)\0*.*\0";

		// ファイル選択ダイアログを表示します。
		auto result = ::GetOpenFileNameW(&ofn);

		return TRUE;
	}

	//
	// インプロセスでエクスプローラを表示します。
	//
	BOOL on_test_explorer()
	{
		// エクスプローラブラウザを作成します。
		IExplorerBrowserPtr browser;
		::CoCreateInstance(CLSID_ExplorerBrowser,
			nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&browser));

		// エクスプローラブラウザを初期化します。
		auto parent = ::GetDesktopWindow();
		auto rc = RECT { 100, 100, 1000, 1000 };
		auto hr = browser->Initialize(parent, &rc, nullptr);

		// エクスプローラブラウザをポップアップウィンドウ化します。
		IOleWindowPtr window = browser;
		auto hwnd = HWND {};
		window->GetWindow(&hwnd);
		::SetParent(hwnd, nullptr);
		auto style = (DWORD)::GetWindowLong(hwnd, GWL_STYLE);
		style &= ~WS_CHILD;
		style |= WS_POPUP | WS_CAPTION | WS_THICKFRAME |
			WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;
		::SetWindowLong(hwnd, GWL_STYLE, style);
#if 0
		// 最前面に表示する場合の処理です。
//		::SetWindowLongPtr(hwnd, GWLP_HWNDPARENT, (LONG_PTR)main_dialog); // メインダイアログより前に表示する場合
		::SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_SHOWWINDOW);
#else
		// 通常表示する場合の処理です。
		::SetWindowPos(hwnd, nullptr, 0, 0, 0, 0,
			SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_SHOWWINDOW);
#endif
		::SetForegroundWindow(hwnd);
		::SetActiveWindow(hwnd);
		::SetFocus(hwnd);

		// エクスプローラブラウザでデスクトップを表示します。
		{
			PIDLIST_ABSOLUTE desktop_idl = {};
			::SHGetKnownFolderIDList(
				FOLDERID_Desktop, 0, nullptr, &desktop_idl);
			browser->BrowseToIDList(desktop_idl, SBSP_ABSOLUTE);
			::ILFree(desktop_idl);
		}

		// メッセージループを促すためにタイマーを開始します。
		auto timer_id = ::SetTimer(nullptr, 0, 500, nullptr);

		// メッセージループを開始します。
		MSG msg = {};
		while (::GetMessageW(&msg, nullptr, 0, 0) > 0)
		{
			// ウィンドウが無効の場合はループを終了します。
			if (!::IsWindow(hwnd)) break;

			// メッセージをディスパッチします。
			::TranslateMessage(&msg);
			::DispatchMessageW(&msg);
		}

		// タイマーを終了します。
		::KillTimer(nullptr, timer_id);

		// エクスプローラブラウザを後始末します。
		browser->Destroy();

		return TRUE;
	}

	//
	// メインダイアログのダイアログプロシージャです。
	//
	INT_PTR CALLBACK main_dlg_proc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
	{
		switch (message)
		{
		case WM_INITDIALOG:
			{
				main_dialog = hwnd;

				auto path = get_module_file_name(nullptr);
				server_dir = path.parent_path();
				server_path = server_dir / L"exbookmark.dll";

				return TRUE;
			}
		case WM_COMMAND:
			{
				auto code = HIWORD(wParam);
				auto control_id = LOWORD(wParam);
				auto control = (HWND)lParam;

				switch (control_id)
				{
				case IDC_INSTALL: return on_install();
				case IDC_UNINSTALL: return on_uninstall();
				case IDC_TEST_COMMON_DIALOG: return on_test_common_dialog();
				case IDC_TEST_EXPLORER: return on_test_explorer();
				case IDOK:
				case IDCANCEL:
					{
						// アプリケーションを終了します。
						::PostQuitMessage(0);

						// ダイアログを終了します。
						::EndDialog(hwnd, control_id);

						return TRUE;
					}
				}

				break;
			}
		}

		return FALSE;
	}

	//
	// エントリポイントです。
	//
	EXTERN_C int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE prev_instance, LPWSTR cmd_line, int cmd_show)
	{
		auto hr = ::CoInitialize(nullptr);
		::DialogBoxW(instance, MAKEINTRESOURCE(IDD_MAIN), ::GetDesktopWindow(), main_dlg_proc);
		::CoUninitialize();

		return 0;
	}
}
