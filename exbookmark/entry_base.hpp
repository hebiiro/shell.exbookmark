#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはサーバーエントリです。
	//
	struct EntryBase {
		virtual ~EntryBase() {}
		virtual HRESULT on_get_class_object(REFCLSID clsid, REFIID iid, void** ppv) = 0;
		virtual HRESULT on_register_server() = 0;
		virtual HRESULT on_unregister_server() = 0;
	};
}
