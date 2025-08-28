#include "pch.h"
#include "hive.hpp"
#include "utils.hpp"
#include "unknown.hpp"
#include "class_factory.hpp"
#include "entry_base.hpp"
#include "server.hpp"
#include "entry.hpp"
#include "folder_tree.hpp"
#include "context_menu.hpp"

namespace apn::exbookmark
{
	//
	// DLLエントリポイントです。
	//
	EXTERN_C BOOL APIENTRY DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved)
	{
		switch (reason)
		{
		case DLL_PROCESS_ATTACH:
			{
				::DisableThreadLibraryCalls(hive.instance = instance);

				break;
			}
		case DLL_PROCESS_DETACH:
			{
				break;
			}
		}

		return TRUE;
	}

	//
	// このDLLをアンロードできる場合はS_OKを返します。
	//
	STDAPI DllCanUnloadNow()
	{
		return server.can_unload_now();
	}

	//
	// 指定されたクラスオブジェクトを返します。
	//
	STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void** ppv)
	{
		return server.get_class_object(clsid, iid, ppv);
	}

	//
	// このDLLをレジストリに登録します。
	//
	STDAPI DllRegisterServer()
	{
		return server.register_server();
	}

	//
	// このDLLをレジストリから削除します。
	//
	STDAPI DllUnregisterServer()
	{
		return server.unregister_server();
	}
}
