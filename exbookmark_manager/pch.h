#pragma once

#include <string>
#include <filesystem>
#include <format>

#include <windows.h>
#include <shlobj.h>
#include <shdispid.h>
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
#include <comdef.h>

_COM_SMARTPTR_TYPEDEF(IExplorerBrowser, __uuidof(IExplorerBrowser));

#pragma comment(linker, "\"/manifestdependency:type='win32' \
	name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
	processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
