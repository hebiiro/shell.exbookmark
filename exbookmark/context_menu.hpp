#pragma once

namespace apn::exbookmark
{
	//
	// コンテキストメニューのCLSIDリテラル文字列です。
	//
	#define STR_CLSID_MyContextMenu "705A0E13-8955-4F7A-BA22-8B78E38318F7"

	//
	// このクラスはコンテキストメニューハンドラーです。
	//
	struct DECLSPEC_UUID(STR_CLSID_MyContextMenu)
		ContextMenu : Unknown<IObjectWithSite, IShellExtInit, IContextMenu>
	{
		//
		// サーバーにエントリします。
		//
		inline static const auto entry = Entry<ContextMenu> {
			L"exbookmark.context_menu",
			L"{" STR_CLSID_MyContextMenu L"}",
			L"exbookmark.context_menu",
			{
				L"*\\shellex\\ContextMenuHandlers\\exbookmark",
				L"Folder\\shellex\\ContextMenuHandlers\\exbookmark",
				L"Directory\\Background\\shellex\\ContextMenuHandlers\\exbookmark",
			},
		};

		//
		// 選択項目のファイルパスです。
		//
		std::filesystem::path selected_item_path;

		//
		// 現在開いているフォルダのファイルパスです。
		//
		std::filesystem::path current_folder_path;

		//
		// 操作対象のフォルダのファイルパスです。
		//
		std::filesystem::path target_folder_path;

		//
		// 操作対象のファイル名です。
		//
		std::wstring target_file_spec;

		//
		// タイマーの繰り返し回数です。
		//
		long timer_count = {};

		//
		// コンテキストメニューのサイトです。
		//
		IUnknownPtr site;

		//
		// シェルブラウザです。
		//
		IShellBrowserPtr shell_browser;

		//
		// フォルダツリーです。
		//
		FolderTree folder_tree;

		//
		// GDI+マネージャです。
		// GDI+の初期化と後始末を自動的に実行します。
		//
//		utils::GdiplusManager gdiplus_manager;

		//
		// ブラウザを初期化します。
		//
		HRESULT init_browser()
		{
			// サイトが無効の場合は何もしません。
			if (!site) return E_FAIL;

			// サイトからシェルブラウザを取得します。
			shell_browser = utils::get_shell_browser(site);

			// シェルブラウザを取得できなかった場合は何もしません。
			if (!shell_browser) return E_FAIL;

			return S_OK;
		}

		//
		// ブラウザを後始末します。
		//
		HRESULT exit_browser()
		{
			shell_browser = nullptr;

			return S_OK;
		}

		//
		// 対象ファイルを選択します。
		//
		HRESULT select_target_file()
		{
			// 操作対象のファイル名が存在しない場合は何もしません。
			if (target_file_spec.empty()) return S_FALSE;

			// シェルブラウザが無効の場合は何もしません。
			if (!shell_browser) return S_FALSE;

			// シェルビューを取得します。
			IShellViewPtr shell_view;
			auto hr = shell_browser->QueryActiveShellView(&shell_view);
			if (!shell_view) return hr;

			// フォルダビューを取得します。
			IFolderViewPtr folder_view = shell_view;
			if (!folder_view) return E_FAIL;

			// シェルフォルダを取得します。
			IShellFolderPtr shell_folder;
			hr = folder_view->GetFolder(IID_PPV_ARGS(&shell_folder));
			if (!shell_folder) return hr;

			// 相対的なIDリストを取得します。
			utils::idl::unique_ptr file_spec_idl;
			hr = shell_folder->ParseDisplayName(
				nullptr,
				nullptr,
				target_file_spec.data(),
				nullptr,
				utils::idl::out_param(file_spec_idl),
				nullptr);
			if (!file_spec_idl) return hr;

			// 操作対象のファイルを選択します。
			hr = shell_view->SelectItem(
				file_spec_idl.get(),
				SVSI_DESELECTOTHERS |
				SVSI_ENSUREVISIBLE |
				SVSI_FOCUSED |
				SVSI_KEYBOARDSELECT);

			return hr;
		}

		//
		// サイト(シェルブラウザ)で指定されたパスを表示します。
		//
		HRESULT navigate(const std::filesystem::path& path)
		{
			// パスがフォルダの場合は
			if (std::filesystem::is_directory(path))
			{
				// 操作対象のフォルダのファイルパスを取得します。
				target_folder_path = path;

				// 操作対象のファイル名は存在しないので消去します。
				target_file_spec.clear();
			}
			// パスがフォルダではない場合は
			else
			{
				// 操作対象のフォルダのファイルパスを取得します。
				target_folder_path = path.parent_path();

				// 操作対象のファイル名を取得します。
				target_file_spec = path.filename();
			}

			// 操作対象のフォルダのパスをIDリストに変換します。
			utils::idl::unique_ptr folder_idl(
				utils::parse_display_name(target_folder_path));
			if (!folder_idl) return E_FAIL;

			// ブラウザの初期化に失敗した場合は何もしません。
			auto hr = init_browser();
			if (FAILED(hr)) return hr;

			// ブラウザで操作対象のフォルダを表示します。
			hr = shell_browser->BrowseObject(
				folder_idl.get(), SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
			if (FAILED(hr)) return hr;

			// 対象ファイルが空の場合はここで処理を終了します。
			if (target_file_spec.empty()) return S_OK;

			// この時点ではまだファイル一覧が作成されていないので項目を選択できません。
			// なので、ウィンドウタイマーを使用して遅延実行します。

			// シェルブラウザのウィンドウハンドルを取得します。
			auto hwnd = HWND{}; shell_browser->GetWindow(&hwnd);
#ifdef _DEBUG
			auto class_name = utils::get_class_name(hwnd);
#endif
			// 参照カウントを増やします。
			AddRef();

			// タイマーの繰り返し回数をリセットします。
			timer_count = {};

			// タイマーを開始します。
			if (!::SetTimer(hwnd, (UINT_PTR)this, 200,
				[](HWND hwnd, UINT message, UINT_PTR timer_id, DWORD time)
			{
				// thisポインタを取得します。
				auto p = (ContextMenu*)timer_id;

				// タイマーの繰り返し回数が少ない場合は
				if (p->timer_count++ < 10)
				{
					// 項目を選択します。
					auto hr = p->select_target_file();

					// 項目の選択に失敗した場合は引き続きタイマーを繰り返します。
					if (FAILED(hr)) return;
				}

				// タイマーを終了します。
				::KillTimer(hwnd, timer_id);

				// 参照カウントを減らします。
				p->Release();
			}))
			{
				// タイマーの作成に失敗した場合は参照カウントを減らします。
				Release();

				return E_FAIL;
			}

			return hr;
		}

		//
		// Unknown::QueryInterfaceImpl()の実装です。
		//
		virtual void* QueryInterfaceImpl(REFIID iid) override
		{
			if (iid == __uuidof(IUnknown)) return static_cast<IContextMenu*>(this);
			else if (iid == __uuidof(IObjectWithSite)) return static_cast<IObjectWithSite*>(this);
			else if (iid == __uuidof(IShellExtInit)) return static_cast<IShellExtInit*>(this);
			else if (iid == __uuidof(IContextMenu)) return static_cast<IContextMenu*>(this);

			return nullptr;
		}

		//
		// IObjectWithSite::SetSite()の実装です。
		//
		IFACEMETHODIMP SetSite(IUnknown* site)
		{
			return this->site = site, S_OK;
		}
        
		//
		// IObjectWithSite::GetSite()の実装です。
		//
		IFACEMETHODIMP GetSite(REFIID iid, void** ppv)
		{
			return site.QueryInterface(iid, ppv);
		}

		//
		// IShellExtInit::Initialize()の実装です。
		//
		IFACEMETHODIMP Initialize(LPCITEMIDLIST folder_idl, IDataObject* data_object, HKEY prog_id)
		{
			// 一旦パスを消去します。
			selected_item_path.clear();
			current_folder_path.clear();

			// 現在選択されている項目が有効の場合は
			if (data_object)
			{
				FORMATETC fe = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
				STGMEDIUM stm = {};
				if (FAILED(data_object->GetData(&fe, &stm))) return E_INVALIDARG;

				if (auto drop = (HDROP)::GlobalLock(stm.hGlobal))
				{
					// 現在選択されている項目のパスを取得します。
					wchar_t path[MAX_PATH] = {};
					::DragQueryFileW(drop, 0, path, std::size(path));
					selected_item_path = path;

					::GlobalUnlock(stm.hGlobal);
				}

				::ReleaseStgMedium(&stm);
			}

			// 現在開いているフォルダのIDリストが有効の場合は
			if (folder_idl)
			{
				// 現在開いているフォルダのパスを取得します。
				current_folder_path = utils::get_file_path(folder_idl);
			}

			return S_OK;
		}

		//
		// IContextMenu::QueryContextMenu()の実装です。
		//
		IFACEMETHODIMP QueryContextMenu(HMENU menu, UINT index_menu, UINT first_command_id, UINT last_command_id, UINT flags)
		{
			// 同じメニューで複数回呼ばれる場合があるのでチェックします。
			static HMENU last_menu = {};
			if (flags & CMF_DEFAULTONLY || last_menu == menu) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
			last_menu = menu;

			// フォルダツリーメニューを構築します。
			auto c = folder_tree.build_menu(menu, index_menu, first_command_id, last_command_id);

			// 構築できた場合はセパレータも追加します。
			if (c) ::InsertMenuW(menu, index_menu++, MF_SEPARATOR | MF_BYPOSITION, 0, nullptr);

			// 追加した項目数を返します。
			return MAKE_HRESULT(SEVERITY_SUCCESS, 0, c);
		}

		//
		// IContextMenu::InvokeCommand()の実装です。
		//
		IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO _mici)
		{
			auto mici = (LPCMINVOKECOMMANDINFOEX)_mici;
			if (mici->cbSize < sizeof(*mici)) return E_FAIL;
			if (HIWORD(mici->lpVerb) != 0) return E_FAIL;
#ifdef _DEBUG
			auto class_name = utils::get_class_name(mici->hwnd);
#endif
			// 対象のパスを取得します。
			auto path = selected_item_path;

			// 対象のパスが無効の場合はフォルダのパスを対象にします。
			if (path.empty()) path = current_folder_path;

			// 対象のパスが無効の場合は何もしません。
			if (path.empty()) return E_FAIL;

			// コマンドIDのオフセットを取得します。
			auto command_id_offset = LOWORD(mici->lpVerb);

			// コマンドIDに対応するイテレータを取得します。
			auto it = folder_tree.nodes.find(command_id_offset);

			// イテレータが無効の場合は何もしません。
			if (it == folder_tree.nodes.end()) return E_FAIL;

			// ノードを取得します。
			const auto& node = it->second;

			if (node.type == node.c_type.c_add)
			{
				// 対象のパスを追加します。
				return folder_tree.add_path(node, path);
			}
			else if (node.type == node.c_type.c_remove)
			{
				// 対象のパスを削除します。
				return folder_tree.remove_path(node, path);
			}
			else
			{
				// 実行フラグが立っている場合は
				if (node.is_execute)
				{
					// パスを実行します。
					return ::ShellExecuteW(nullptr, nullptr,
						node.path.c_str(), nullptr, nullptr, SW_SHOWDEFAULT) ? S_OK : E_FAIL;
				}
				else
				{
					// サイトでパスを表示します。
					return navigate(node.path);
				}
			}
		}

		//
		// IContextMenu::GetCommandString()の実装です。
		//
		IFACEMETHODIMP GetCommandString(UINT_PTR command_id, UINT type, UINT* reserved, LPSTR name, UINT cch)
		{
			return E_NOTIMPL;
		}
	};
}
