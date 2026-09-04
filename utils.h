#pragma once
#include "pch.h"

std::wstring convert_path_from_win_to_winslash(const std::wstring& win_path);
std::wstring convert_path_from_win_to_fileprotocal(const std::wstring& win_path);
std::wstring convert_path_from_win_to_winescaped(const std::wstring& win_path);
std::wstring convert_path_from_win_to_unix(const std::wstring& win_path);
std::wstring convert_path_from_win_to_name(const std::wstring& win_path);
std::wstring convert_path_from_win_to_wsl(const std::wstring& win_path);
std::wstring ConvertPaths(std::vector<std::wstring> paths, std::wstring(*converter)(const std::wstring&));
std::wstring TruncateMiddle(const std::wstring &s, size_t maxLen);
BOOL SetClipboardTextW(const WCHAR *text, SIZE_T cch);