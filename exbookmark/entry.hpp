#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはサーバーエントリの実装です。
	//
	template <typename T>
	struct Entry : EntryBase
	{
		//
		// COMオブジェクト名です。
		//
		std::wstring name_string;

		//
		// CLSID文字列です。
		//
		std::wstring clsid_string;

		//
		// ProgID文字列です。
		//
		std::wstring progid_string;

		//
		// 追加のレジストリキーです。
		//
		std::vector<std::wstring> addition_keys;

		//
		// CLSIDキー文字列です。
		//
		std::wstring clsid_key;

		//
		// ProgIDキー文字列です。
		//
		std::wstring progid_key;

		//
		// コンストラクタです。
		//
		Entry(
			const std::wstring& name_string,
			const std::wstring& clsid_string,
			const std::wstring& progid_string,
			const std::vector<std::wstring>& addition_keys)
			: name_string(name_string)
			, clsid_string(clsid_string)
			, progid_string(progid_string)
			, addition_keys(addition_keys)
			, clsid_key(L"CLSID\\" + clsid_string)
			, progid_key(progid_string)
		{
			server.entries.emplace_back(this);
		}

		//
		// 指定されたクラスオブジェクトを返します。
		//
		virtual HRESULT on_get_class_object(REFCLSID clsid, REFIID iid, void** ppv) override
		{
			if (clsid == __uuidof(T))
				return utils::create_interface<ClassFactory<T>>(iid, ppv);

			return CLASS_E_CLASSNOTAVAILABLE;
		}

		//
		// レジストリに登録します。
		//
		virtual HRESULT on_register_server() override
		{
			//
			// この関数はレジストリキーをセットします。
			//
			const auto set_registry_key = [](
				HKEY root,
				const std::wstring& sub_key,
				const std::wstring& value_name,
				const std::wstring& value) -> BOOL
			{
				// レジストリキーを作成します。
				auto key = HKEY {};
				auto result = ::RegCreateKeyExW(root, sub_key.c_str(), 0, nullptr,
					REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &key, nullptr);
				if (result != ERROR_SUCCESS) return FALSE;

				// レジストリ値をセットします。
				result = ::RegSetValueExW(key,value_name.empty() ? nullptr : value_name.c_str(),
					0, REG_SZ,
					(const BYTE*)value.data(),
					(DWORD)((value.size() + 1) * sizeof(wchar_t)));

				// レジストリキーを閉じます。
				::RegCloseKey(key);

				return (result == ERROR_SUCCESS);
			};

			// このDLLのファイルパスを取得します。
			auto module_file_name = utils::get_module_file_name(hive.instance);

			// CLSIDキーを作成します。
			if (!set_registry_key(HKEY_CLASSES_ROOT, clsid_key, L"", name_string)) return E_FAIL;
			if (!set_registry_key(HKEY_CLASSES_ROOT, clsid_key + L"\\InprocServer32", L"", module_file_name)) return E_FAIL;
			if (!set_registry_key(HKEY_CLASSES_ROOT, clsid_key + L"\\InprocServer32", L"ThreadingModel", L"Apartment")) return E_FAIL;

			// ProgIDキーを作成します。
			if (!set_registry_key(HKEY_CLASSES_ROOT, progid_key, L"", name_string)) return E_FAIL;
			if (!set_registry_key(HKEY_CLASSES_ROOT, progid_key + L"\\CLSID", L"", clsid_string)) return E_FAIL;

			// 追加のレジストリキーを作成します。
			for (const auto& addition_key : addition_keys)
				if(!set_registry_key(HKEY_CLASSES_ROOT, addition_key, L"", clsid_string)) return E_FAIL;

			return S_OK;
		}

		//
		// レジストリから削除します。
		//
		virtual HRESULT on_unregister_server() override
		{
			//
			// この関数はレジストリキーを削除します。
			//
			const auto delete_registry_key = [](
				HKEY root,
				const std::wstring& sub_key) -> BOOL
			{
				// レジストリキーを削除します。
				auto result = ::RegDeleteTreeW(root, sub_key.c_str());

				return (result == ERROR_SUCCESS);
			};

			auto result = FALSE;

			// CLSIDキーを削除します。
			result |= delete_registry_key(HKEY_CLASSES_ROOT, clsid_key);

			// ProgIDキーを削除します。
			result |= delete_registry_key(HKEY_CLASSES_ROOT, progid_key);

			// 追加のレジストリキーを削除します。
			for (const auto& addition_key : addition_keys)
				result |= delete_registry_key(HKEY_CLASSES_ROOT, addition_key);

			return result ? S_OK : E_FAIL;
		}
	};
}
