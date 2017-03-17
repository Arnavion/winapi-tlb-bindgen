#pragma once

#include <Windows.h>

#include <cstdio>

#define TRY(expr) do { HRESULT hr = (expr); if (FAILED(hr)) { if (hr == TYPE_E_CANTLOADLIBRARY) { std::exit(1); } else { std::abort(); } } } while (false)
#define ASSERT(expr) do { if (!(expr)) { std::abort(); } } while (false)
#define UNREACHABLE do { std::abort(); } while (false)
