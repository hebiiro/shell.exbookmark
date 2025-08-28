#pragma once

#include <string>
#include <format>
#include <filesystem>
#include <vector>
#include <unordered_map>

#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>
#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")
#include <oleacc.h>
#pragma comment(lib, "oleacc.lib")
#include <vsstyle.h>
#include <vssym32.h>
#include <uxtheme.h>
#pragma comment(lib, "uxtheme.lib")
#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")

#if 0
#import <msxml3.dll>
#else
#import <msxml6.dll>
namespace MSXML2 {
	using DOMDocument = DOMDocument60;
	using MXXMLWriter = MXXMLWriter60;
	using SAXXMLReader = SAXXMLReader60;
}
#endif

_COM_SMARTPTR_TYPEDEF(IFolderView, __uuidof(IFolderView));

#pragma comment(linker, "\"/manifestdependency:type='win32' \
	name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
	processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
