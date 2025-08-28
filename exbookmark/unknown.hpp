#pragma once

namespace apn::exbookmark
{
	//
	// このクラスはIUnknownの実装です。
	//
	template <typename... Bases>
	struct DECLSPEC_NOVTABLE Unknown : Bases...
	{
		virtual void* QueryInterfaceImpl(REFIID iid) = 0;

		//
		// 参照カウントです。
		//
		ULONG nb_refs = 1;

		//
		// コンストラクタです。
		//
		Unknown()
		{
			::InterlockedIncrement(&hive.nb_objects);
		}

		//
		// デストラクタです。
		//
		virtual ~Unknown()
		{
			::InterlockedDecrement(&hive.nb_objects);
		}

		//
		// IUnknown::QueryInterface()の実装です。
		//
		IFACEMETHODIMP QueryInterface(REFIID iid, void** ppv)
		{
			// インターフェイスが見つかった場合は
			if (*ppv = QueryInterfaceImpl(iid))
				return AddRef(), S_OK; // 参照カウントを増やします。
			else
				return E_NOINTERFACE;
		}

		//
		// IUnknown::AddRef()の実装です。
		//
		IFACEMETHODIMP_(ULONG) AddRef()
		{
			return ::InterlockedIncrement(&nb_refs);
		}

		//
		// IUnknown::Release()の実装です。
		//
		IFACEMETHODIMP_(ULONG) Release()
		{
			if (auto c = ::InterlockedDecrement(&nb_refs))
				return c;
			else
				return delete this, 0;
		}
	};
}
