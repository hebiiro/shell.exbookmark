#pragma once

namespace apn::exbookmark::utils
{
	//
	// このクラスはwin32のハンドルを管理するユーティリティです。
	//
	template <typename T, auto deleter>
	struct handle_utils
	{
		using type = std::remove_pointer<T>::type;
		using unique_ptr = std::unique_ptr<type, decltype(deleter)>;

		struct shared_ptr : std::shared_ptr<type>
		{
			shared_ptr() {}
			shared_ptr(T x) : std::shared_ptr<type>(x, deleter) {}
			auto reset(T x) { return __super::reset(x, deleter); }
		};

		using weak_ptr = std::weak_ptr<type>;

		struct out_param
		{
			unique_ptr& ptr; T value;
			out_param(unique_ptr& ptr) : ptr(ptr), value() {}
			~out_param() { ptr.reset(value); }
			operator T*() { return &value; }
		};
	};

	using dc = handle_utils<HDC, [](HDC p) { ::DeleteDC(p); }>;
	using bitmap = handle_utils<HBITMAP, [](HBITMAP p) { ::DeleteObject(p); }>;
	using icon = handle_utils<HICON, [](HICON p) { ::DestroyIcon(p); }>;
	using theme = handle_utils<HTHEME, [](HTHEME p) { ::CloseThemeData(p); }>;
	using idl = handle_utils<LPITEMIDLIST, [](LPVOID p) { ::CoTaskMemFree(p); }>;

	namespace gdi
	{
		//
		// このクラスはデバイスコンテキストとGDIオブジェクトをバインドします。
		//
		struct selector
		{
			HDC dc; HGDIOBJ obj, old_obj;
			selector(HDC dc, HGDIOBJ obj) : dc(dc), obj(obj), old_obj(::SelectObject(dc, obj)) {}
			~selector() { ::SelectObject(dc, old_obj); }
		};
	}

	//
	// このクラスはGDI+を管理します。
	//
	struct GdiplusManager {
		Gdiplus::GdiplusStartupInput si;
		Gdiplus::GdiplusStartupOutput so;
		ULONG_PTR token;
		ULONG_PTR hook_token;
		GdiplusManager() {
			si.SuppressBackgroundThread = TRUE;
			Gdiplus::GdiplusStartup(&token, &si, &so);
			so.NotificationHook(&hook_token);
		}
		~GdiplusManager() {
			so.NotificationUnhook(hook_token);
			Gdiplus::GdiplusShutdown(token);
		}
	};

	//
	// 最初のインターフェイスを作成して返します。
	//
	template <typename T>
	HRESULT create_interface(REFIID iid, void** ppv)
	{
		// インスタンスを作成します。
		auto p = new T();

		// 最初のインターフェイスを取得します。
		auto hr = p->QueryInterface(iid, ppv);

		// インターフェイスが取得できなかった場合は削除してから返します。
		return (hr == S_OK) ? (hr) : (delete p, hr);
	}

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
	// 指定されたウィンドウのクラス名を返します。
	//
	std::wstring get_class_name(HWND hwnd)
	{
		wchar_t class_name[MAX_PATH] = {};
		::GetClassNameW(hwnd, class_name, std::size(class_name));
		return class_name;
	}

	//
	// PIDLを文字列パスに変換して返します。
	//
	std::wstring get_file_path(PCIDLIST_ABSOLUTE idl)
	{
		wchar_t buffer[MAX_PATH] = {};
		::SHGetPathFromIDListW(idl, buffer);
		return buffer;
	}

	//
	// 文字列パスをPIDLに変換して返します。
	//
	PIDLIST_ABSOLUTE parse_display_name(const std::wstring& path)
	{
		auto idl = PIDLIST_ABSOLUTE {};
		::SHParseDisplayName(path.c_str(), nullptr, &idl, 0, nullptr);
		return idl;
	}

	//
	// シェルブラウザを返します。
	//
	IShellBrowserPtr get_shell_browser(IUnknown* site)
	{
		// サービスプロバイダーを取得します。
		if (IServiceProviderPtr service_provider = site)
		{
			const auto get_service = [&](const auto& sid)
			{
				IShellBrowserPtr shell_browser;
				service_provider->QueryService(sid, IID_PPV_ARGS(&shell_browser));
				return shell_browser;
			};

//			if (auto shell_browser = get_service(SID_SShellBrowser)) return shell_browser;
//			if (auto shell_browser = get_service(SID_SInPlaceBrowser)) return shell_browser;
			if (auto shell_browser = get_service(SID_STopLevelBrowser)) return shell_browser;
		}

		return {};
	}

	//
	// ビットマップの画像サイズを返します。
	//
	SIZE get_size(HBITMAP bitmap)
	{
		BITMAP bm = {};
		::GetObjectW(bitmap, sizeof(bm), &bm);
		return SIZE { bm.bmWidth, bm.bmHeight };
	}

	//
	// アイコンをビットマップに変換して返します。
	//
	HBITMAP to_bitmap(HICON icon, SIZE target_size)
	{
		// アイコン情報を取得します。
		ICONINFO ii = {};
		if (!::GetIconInfo(icon, &ii)) return {};

		// アイコンのビットマップを取得します。
		bitmap::unique_ptr color_bitmap(ii.hbmColor);
		bitmap::unique_ptr mask_bitmap(ii.hbmMask);
#if 0
		// カラービットマップを返します。
		return color_bitmap.release();
#else
		// メモリDCを作成します。
		dc::unique_ptr mem_dc(::CreateCompatibleDC(nullptr));
		if (!mem_dc) return {};

		// アイコンの画像サイズを取得します。
		auto src_size = get_size(color_bitmap.get());

		// 出力画像サイズを取得します。
		auto out_size = SIZE {
			(target_size.cx > 0) ? target_size.cx : src_size.cx,
			(target_size.cy > 0) ? target_size.cy : src_size.cy,
		};

		if (src_size.cx == out_size.cx && src_size.cy == out_size.cy)
		{
			// 出力用DIBを作成します。
			BITMAPINFO out_bmi = {};
			out_bmi.bmiHeader.biSize        = sizeof(out_bmi.bmiHeader);
			out_bmi.bmiHeader.biWidth       = out_size.cx;
			out_bmi.bmiHeader.biHeight      = -out_size.cy; // top-down
			out_bmi.bmiHeader.biPlanes      = 1;
			out_bmi.bmiHeader.biBitCount    = 32;
			out_bmi.bmiHeader.biCompression = BI_RGB;
			auto out_bits = LPVOID {};
			bitmap::unique_ptr dib(::CreateDIBSection(
				mem_dc.get(), &out_bmi, DIB_RGB_COLORS, &out_bits, nullptr, 0));

			// カラービットマップを出力用DIBにコピーします。
			::GetDIBits(mem_dc.get(), color_bitmap.get(),
				0, out_size.cy, out_bits, &out_bmi, DIB_RGB_COLORS);

			// 出力用DIBを返します。
			return dib.release();
		}

		// カラービットマップのピクセルデータを32bit DIB形式で取得します。
		BITMAPINFO color_bmi = {};
		color_bmi.bmiHeader.biSize        = sizeof(color_bmi.bmiHeader);
		color_bmi.bmiHeader.biWidth       = src_size.cx;
		color_bmi.bmiHeader.biHeight      = -src_size.cy; // top-down
		color_bmi.bmiHeader.biPlanes      = 1;
		color_bmi.bmiHeader.biBitCount    = 32;
		color_bmi.bmiHeader.biCompression = BI_RGB;
		std::vector<BYTE> color_buffer(src_size.cx * src_size.cy * 4);
		::GetDIBits(mem_dc.get(), color_bitmap.get(),
			0, src_size.cy, color_buffer.data(), &color_bmi, DIB_RGB_COLORS);
#if 0
		// カラービットマップがアルファを持っているか確認します。
		auto has_alpha = [&color_buffer]()
		{
			// ピクセルデータを走査します。
			for (size_t i = 0; i < color_buffer.size() / 4; i++)
			{
				// アルファ値が0ではない場合はTRUEを返します。
				if (color_buffer[i * 4 + 3]) return TRUE;
			}

			// アルファ値がすべて0だった場合はFALSEを返します。
			return FALSE;
		}();

		// カラービットマップがアルファを持っていない場合は
		if (!has_alpha)
		{
			BITMAPINFO mask_bmi = {};
			mask_bmi.bmiHeader.biSize        = sizeof(mask_bmi.bmiHeader);
			mask_bmi.bmiHeader.biWidth       = src_size.cx;
			mask_bmi.bmiHeader.biHeight      = -src_size.cy; // top-down
			mask_bmi.bmiHeader.biPlanes      = 1;
			mask_bmi.bmiHeader.biBitCount    = 32;
			mask_bmi.bmiHeader.biCompression = BI_RGB;
			std::vector<BYTE> mask_buffer(src_size.cx * src_size.cy * 4);
			auto result = ::GetDIBits(mem_dc.get(), mask_bitmap.get(),
				0, src_size.cy, mask_buffer.data(), &mask_bmi, DIB_RGB_COLORS);

			// ピクセルデータを走査します。
			for (size_t i = 0; i < mask_buffer.size() / 4; i++)
			{
				// マスクのRGB値をカラーのアルファに変換します。
				color_buffer[i * 4 + 3] = mask_buffer[i * 4 + 0] ? 0 : 255;
			}
		}
#endif
		// 出力用DIBを作成します。
		BITMAPINFO out_bmi = {};
		out_bmi.bmiHeader.biSize        = sizeof(out_bmi.bmiHeader);
		out_bmi.bmiHeader.biWidth       = out_size.cx;
		out_bmi.bmiHeader.biHeight      = -out_size.cy; // top-down
		out_bmi.bmiHeader.biPlanes      = 1;
		out_bmi.bmiHeader.biBitCount    = 32;
		out_bmi.bmiHeader.biCompression = BI_RGB;
		auto out_bits = LPVOID {};
		bitmap::unique_ptr dib(::CreateDIBSection(
			mem_dc.get(), &out_bmi, DIB_RGB_COLORS, &out_bits, nullptr, 0));
		if (!dib) return nullptr;

		{
			gdi::selector selector(mem_dc.get(), dib.get());

			::StretchDIBits(mem_dc.get(),
				0, 0, out_size.cx, out_size.cy,
				0, 0, src_size.cx, src_size.cy,
				color_buffer.data(), &color_bmi,
				DIB_RGB_COLORS, SRCCOPY);
		}

		return dib.release();
#endif
	}

	//
	// アイコンをビットマップに変換して返します。
	//
	HBITMAP to_bitmap(HICON icon)
	{
#if 0
		auto bitmap = HBITMAP {};
		Gdiplus::Bitmap(icon).GetHBITMAP(Gdiplus::Color(), &bitmap);
		return bitmap;
#else
		auto w = ::GetSystemMetrics(SM_CXSMICON);
		auto h = ::GetSystemMetrics(SM_CYSMICON);

		return to_bitmap(icon, { w, h });
#endif
	}

	//
	// 指定されたファイルのアイコンを取得します。
	//
	HICON get_icon(const std::wstring& path, int icon_index)
	{
		auto icon = HICON {};
		::ExtractIconExW(path.c_str(), icon_index, nullptr, &icon, 1);
		return icon;
	}

	//
	// 指定されたファイルパスのアイコンを取得します。
	//
	HICON get_shell_icon(const std::wstring& path)
	{
		// シェルからファイル情報を取得します。
		SHFILEINFO sfi = {};
		::SHGetFileInfoW(path.c_str(), 0,
			&sfi, sizeof(sfi), SHGFI_ICON | SHGFI_SMALLICON);

		// アイコンを返します。
		return sfi.hIcon;
	}

	//
	// XMLドキュメントを作成して返します。
	//
	MSXML2::IXMLDOMDocumentPtr load_document(const std::wstring& path)
	{
		// フォルダツリーファイルのパスを取得します。
		auto file_name = utils::get_module_file_name(hive.instance).replace_extension() / L"folder_tree.xml";

		// XMLドキュメントを作成します。
		MSXML2::IXMLDOMDocumentPtr document(__uuidof(MSXML2::DOMDocument));

		// XMLファイルを開きます。
		if (document->load(path.c_str()) == VARIANT_FALSE)
			return {};

		// XMLドキュメントを返します。
		return document;
	}

	//
	// XMLドキュメントを保存します。
	//
	HRESULT save_document(const MSXML2::IXMLDOMDocumentPtr& document,
		const std::wstring& path, const std::wstring& encoding = L"UTF-8")
	{
		// 出力ストリームを作成します。
		IStreamPtr stream;
		::SHCreateStreamOnFileW(path.c_str(), STGM_WRITE | STGM_SHARE_DENY_WRITE | STGM_CREATE | STGM_DIRECT, &stream);
		if (!stream) return E_FAIL;

		// XMLライターを作成します。
		MSXML2::IMXWriterPtr writer(__uuidof(MSXML2::MXXMLWriter));
		writer->output = stream.GetInterfacePtr();
		writer->indent = VARIANT_TRUE;
		writer->byteOrderMark = VARIANT_TRUE; // ※このようにしてもBOMは付与されません。
#if 1
		writer->omitXMLDeclaration = VARIANT_FALSE;
		writer->version = L"1.0";
		writer->encoding = encoding.c_str();
#endif
		// SAXリーダーを作成します。
		MSXML2::ISAXXMLReaderPtr reader(__uuidof(MSXML2::SAXXMLReader));
		reader->putContentHandler(MSXML2::ISAXContentHandlerPtr(writer));

		// XMLファイルに保存します。
		return reader->parse(document.GetInterfacePtr());
	}

	//
	// 指定された名前の属性が存在する場合はTRUEを返します。
	//
	BOOL has_attribute(const MSXML2::IXMLDOMElementPtr& element, LPCWSTR name)
	{
		try
		{
			return !!element->getAttributeNode(name);
		}
		catch (_com_error& e)
		{
			(void)e;
		}

		return {};
	}

	//
	// 指定された名前の属性を取得して文字列として返します。
	//
	std::wstring get_string(const MSXML2::IXMLDOMElementPtr& element, LPCWSTR name)
	{
		try
		{
			// 属性を取得します。
			auto var = element->getAttribute(name);
			if (var.vt == VT_NULL) return {};

			// 属性を文字列に変換します。
			_bstr_t value = (_bstr_t)var;
			if (!(BSTR)value) return {};

			// 文字列を返します。
			return (BSTR)value;
		}
		catch (_com_error& e)
		{
			(void)e;
		}

		return {};
	}

	//
	// 指定されたパスの環境変数を展開して返します。
	//
	std::wstring expand_env_string(const std::wstring& path)
	{
		auto buffer_size = ::ExpandEnvironmentStringsW(path.c_str(), nullptr, 0);
		std::wstring buffer(buffer_size + 1, L'\0');
		if (!::ExpandEnvironmentStringsW(path.c_str(), buffer.data(), (DWORD)buffer.size()))
			return path;
		buffer.resize(wcslen(buffer.data()));
		return buffer;
	}
}
