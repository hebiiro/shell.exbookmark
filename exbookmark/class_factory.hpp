#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはクラスファクトリです。
	//
	template <typename T>
	struct ClassFactory : Unknown<IClassFactory>
	{
		//
		// Unknown::QueryInterfaceImpl()の実装です。
		//
		virtual void* QueryInterfaceImpl(REFIID iid) override
		{
			if (iid == __uuidof(IUnknown)) return static_cast<IClassFactory*>(this);
			else if (iid == __uuidof(IClassFactory)) return static_cast<IClassFactory*>(this);

			return nullptr;
		}

		IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID iid, void** ppv)
		{
			if (outer) return CLASS_E_NOAGGREGATION;

			return utils::create_interface<T>(iid, ppv);
		}

		IFACEMETHODIMP LockServer(BOOL lock)
		{
			return S_OK;
		}
	};
}
