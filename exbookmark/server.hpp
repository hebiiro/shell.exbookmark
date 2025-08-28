#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはCOMサーバーです。
	//
	inline struct Server
	{
		//
		// エントリのコレクションです。
		//
		std::vector<EntryBase*> entries;

		//
		// このDLLをアンロードできる場合はS_OKを返します。
		//
		HRESULT can_unload_now()
		{
			// オブジェクトが存在しないときだけアンロードできます。
			return (hive.nb_objects == 0) ? S_OK : S_FALSE;
		}

		//
		// 指定されたクラスオブジェクトを返します。
		//
		HRESULT get_class_object(REFCLSID clsid, REFIID iid, void** ppv)
		{
			for (auto& entry : entries)
			{
				auto hr = entry->on_get_class_object(clsid, iid, ppv);
				if (SUCCEEDED(hr)) return hr;
			}

			return CLASS_E_CLASSNOTAVAILABLE;
		}

		//
		// このDLLをレジストリに登録します。
		//
		HRESULT register_server()
		{
			for (auto& entry : entries)
			{
				auto hr = entry->on_register_server();
				if (FAILED(hr)) return hr;
			}

			return S_OK;
		}

		//
		// このDLLをレジストリから削除します。
		//
		HRESULT unregister_server()
		{
			auto ret_value = S_OK;

			for (auto& entry : entries)
			{
				auto hr = entry->on_unregister_server();
				if (FAILED(hr)) ret_value = hr;
			}

			return ret_value;
		}
	} server;
}
