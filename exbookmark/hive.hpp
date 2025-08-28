#pragma once

namespace apn::exbookmark
{
	//
	// このクラスは他クラスから共通して使用される変数を保持します。
	//
	inline struct Hive
	{
		inline static constexpr auto c_name = L"exbookmark";
		inline static constexpr auto c_display_name = L"EXブックマーク";

		//
		// このモジュールのインスタンスハンドルです。
		//
		HINSTANCE instance = nullptr;

		//
		// 現存するオブジェクトの総数です。
		//
		ULONG nb_objects = 0;

		//
		// コンフィグのファイル名です。
		//
		std::wstring config_file_name;

		//
		// メッセージボックスを表示します。
		//
		int32_t message_box(const std::wstring& text, HWND hwnd = nullptr, int32_t type = MB_OK | MB_ICONWARNING) {
			return ::MessageBoxW(hwnd, text.c_str(), c_display_name, type);
		}
	} hive;
}
