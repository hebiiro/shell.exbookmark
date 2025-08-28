#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはフォルダツリーを管理します。
	//
	struct FolderTree
	{
		//
		// このクラスはメニュー構築用コンテキストです。
		//
		struct Context {
			HMENU menu;
			UINT index_menu;
			UINT first_command_id;
			UINT last_command_id;
			UINT command_id;
			std::filesystem::path path;
		};

		//
		// このクラスはノードです。
		//
		struct Node {
			//
			// このクラスはノードのタイプです。
			//
			inline static constexpr struct Type {
				inline static constexpr wchar_t c_folder[] = L"folder";
				inline static constexpr wchar_t c_file[] = L"file";
				inline static constexpr wchar_t c_separator[] = L"separator";
				inline static constexpr wchar_t c_add[] = L"add";
				inline static constexpr wchar_t c_remove[] = L"remove";
			} c_type;

			//
			// このノードに対応するxml要素です。
			//
			MSXML2::IXMLDOMElementPtr element;

			std::wstring label;
			std::wstring type;
			std::filesystem::path path;
			BOOL is_execute;
		};

		//
		// フォントツリーファイルのxmlドキュメントです。
		//
		MSXML2::IXMLDOMDocumentPtr document;

		//
		// ノードのコレクションです。
		//
		std::unordered_map<UINT, Node> nodes;

		//
		// ビットマップのコレクションです。
		//
		std::vector<utils::bitmap::unique_ptr> bitmaps;

		//
		// フォントツリーファイルのパスを返します。
		//
		inline static auto get_font_tree_file_path()
		{
			return utils::get_module_file_name(hive.instance).replace_extension() / L"folder_tree.xml";
		}

		//
		// 指定された要素の子ノード要素のリストを返します。
		//
		inline static MSXML2::IXMLDOMNodeListPtr get_node_elements(const MSXML2::IXMLDOMElementPtr& element)
		{
			return element->selectNodes(L"node");
		}

		//
		// メニューを構築します。
		//
		UINT build_menu(HMENU menu, UINT index_menu, UINT first_command_id, UINT last_command_id)
		{
			try
			{
				// 既存のノードを消去します。
				nodes.clear();

				// 既存のビットマップを消去します。
				bitmaps.clear();

				// フォントツリーファイルを読み込みます。
				document = utils::load_document(get_font_tree_file_path());

				// ドキュメントを読み込めなかった場合は何もしません。
				if (!document) return 0;

				// コンテキストを作成します。
				Context context = { menu, index_menu, first_command_id, last_command_id, first_command_id };

				// ドキュメントエレメントの子ノード要素を読み込みます。
				return read_nodes(get_node_elements(document->documentElement), context);
			}
			catch (_com_error& e)
			{
				(void)e;
			}

			return 0;
		}

		//
		// フォルダツリーファイルに指定されたパスを追加します。
		//
		HRESULT add_path(const Node& node, const std::wstring& path)
		{
			try
			{
				// ドキュメントが存在しない場合は何もしません。
				if (!document) return E_FAIL;

				// 要素を取得します。
				auto element = node.element;

				// 要素が存在しない場合は何もしません。
				if (!element) return E_FAIL;

				// 親要素を取得します。
				auto parent_element = element->parentNode;

				// 親要素が存在しない場合は何もしません。
				if (!parent_element) return E_FAIL;

				// 新しい要素を作成します。
				auto new_element = document->createElement(L"node");
				new_element->setAttribute(L"path", path.c_str());

				// 兄弟要素として先頭に追加します。
				parent_element->insertBefore(new_element,
					parent_element->firstChild.GetInterfacePtr());

				// フォントツリーファイルに保存します。
				return utils::save_document(document, get_font_tree_file_path());
			}
			catch (_com_error& e)
			{
				return e.Error();
			}

			return E_FAIL;
		}

		//
		// フォルダツリーファイルから指定されたパスを削除します。
		//
		HRESULT remove_path(const Node& node, const std::wstring& path)
		{
			try
			{
				// ドキュメントが存在しない場合は何もしません。
				if (!document) return E_FAIL;

				// 要素を取得します。
				auto element = node.element;

				// 要素が存在しない場合は何もしません。
				if (!element) return E_FAIL;

				// 親要素を取得します。
				auto parent_element = element->parentNode;

				// 親要素が存在しない場合は何もしません。
				if (!parent_element) return E_FAIL;

				// 全てのノードを走査します。
				for (const auto& pair : nodes)
				{
					const auto& target_node = pair.second;

					// パスが異なる場合は何もしません。
					if (target_node.path.compare(path)) continue;

					// 削除対象の要素を取得します。
					auto target_element = target_node.element;

					// 要素が存在しない場合は何もしません。
					if (!target_element) continue;

					// 親要素を取得します。
					auto target_parent_element = target_element->parentNode;

					// 親要素が存在しない場合は何もしません。
					if (!target_parent_element) continue;

					// 親要素が異なる場合は何もしません。
					if (target_parent_element != parent_element) continue;

					// 要素を削除します。
					target_parent_element->removeChild(target_element);

					// フォントツリーファイルに保存します。
					return utils::save_document(document, get_font_tree_file_path());
				}
			}
			catch (_com_error& e)
			{
				return e.Error();
			}

			return E_FAIL;
		}

		//
		// 子ノード要素を読み込みます。
		//
		UINT read_nodes(const MSXML2::IXMLDOMNodeListPtr& node_elements, Context& context)
		{
			// 作成したメニュー項目数です。
			auto result = UINT {};

			/// 子ノード要素の個数を取得します。
			auto c = node_elements->length;

			/// 子ノード要素を走査します。
			for (decltype(c) i = 0; i < c; i++)
			{
				// 子ノード要素を読み込みます。
				result += read_node(node_elements->item[i], context);
			}

			// 作成したメニュー項目数を返します。
			return result;
		}

		//
		// ノード要素を読み込みます。
		//
		UINT read_node(const MSXML2::IXMLDOMElementPtr& element, Context& context)
		{
			// 作成したメニュー項目数です。
			auto result = UINT {};

			/// ノードを作成します。
			auto node = Node {
				.element = element,
				.label = utils::get_string(element, L"label"),
				.type = utils::get_string(element, L"type"),
				.path = utils::expand_env_string(utils::get_string(element, L"path")),
				.is_execute = utils::has_attribute(element, L"execute"),
			};

			// ノードタイプがセパレータの場合は
			if (node.type == node.c_type.c_separator)
			{
				// このノード要素をセパレータとして追加します。
				MENUITEMINFO mii = { sizeof(mii) };
				mii.fMask = MIIM_FTYPE;
				mii.fType = MFT_SEPARATOR;
				::InsertMenuItemW(context.menu, context.index_menu++, TRUE, &mii);

				return result;
			}

			/// 残りのノード要素の属性を取得します。
			auto icon_path = utils::get_string(element, L"icon_path");
			auto icon_index = utils::get_string(element, L"icon_index");
			auto is_absolute = utils::has_attribute(element, L"absolute");

			// 絶対パス指定がない場合は
			if (!is_absolute)
			{
				// パスが相対パスの場合は
				if (node.path.is_relative())
				{
					// コンテキストのパスと結合します。
					node.path = context.path / node.path;
				}
			}

			// アイコンパスが空の場合はノードのパスを使用します。
			if (icon_path.empty()) icon_path = node.path;

			// ビットマップを取得します。
			auto bitmap = [&]() -> HBITMAP
			{
				// アイコンパスが空の場合はnullptrを返します。
				if (icon_path.empty()) return nullptr;

				// アイコンインデックスを数値に変換します。
				auto index = wcstol(icon_index.c_str(), nullptr, 0);

				// アイコンを取得します。
				utils::icon::unique_ptr icon(icon_index.empty() ?
					utils::get_shell_icon(icon_path) : utils::get_icon(icon_path, index));

				// アイコンをビットマップに変換して返します。
				return utils::to_bitmap(icon.get());
			} ();

			// ビットマップをコレクションに追加します。
			bitmaps.emplace_back(bitmap);

			/// 子ノード要素のリストを取得します。
			auto node_elements = get_node_elements(element);

			/// 子ノード要素が存在しない場合は
			if (node_elements->length == 0)
			{
				/// ラベルが空の場合は
				if (node.label.empty())
				{
					// パスをラベルにします。
					node.label = node.path;
				}
				/// ラベルが空ではない場合は
				else
				{
					/// パスが空ではない場合は
					if (!node.path.empty())
					{
						/// ラベルとパスを結合します。
						node.label += L"\t" + node.path.wstring();
					}
				}

				// コマンドIDのオフセットをキーにしてノードをコレクションに追加します。
				nodes[context.command_id - context.first_command_id] = node;

				// このノード要素をメニュー項目として追加します。
				MENUITEMINFO mii = { sizeof(mii) };
				mii.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_BITMAP | MIIM_ID;
				mii.fType = MFT_STRING;
				mii.dwTypeData = node.label.data();
				mii.hbmpItem = bitmap;
				mii.wID = context.command_id++;
				::InsertMenuItemW(context.menu, context.index_menu++, TRUE, &mii);

				// 作成したメニュー項目数を増やします。
				result++;
			}
			/// 子ノード要素が存在する場合は
			else
			{
				/// ラベルが空の場合はパスをラベルにします。
				if (node.label.empty()) node.label = node.path;

				// サブメニューを作成します。
				auto sub_menu = ::CreatePopupMenu();

				// このノード要素をサブメニューとして追加します。
				MENUITEMINFO mii = { sizeof(mii) };
				mii.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_BITMAP | MIIM_SUBMENU;
				mii.fType = MFT_STRING;
				mii.dwTypeData = node.label.data();
				mii.hbmpItem = bitmap;
				mii.hSubMenu = sub_menu;
				::InsertMenuItemW(context.menu, context.index_menu++, TRUE, &mii);

				// サブコンテキストを作成します。
				auto sub_context = Context {
					sub_menu,
					0,
					context.first_command_id,
					context.last_command_id,
					context.command_id,
					node.path,
				};

				// 子ノード要素を読み込みます。
				result += read_nodes(node_elements, sub_context);

				// コンテキストのコマンドIDを更新します。
				context.command_id = sub_context.command_id;
			}

			// 作成したメニュー項目数を返します。
			return result;
		}
	};
}
