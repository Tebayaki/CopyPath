#include "pch.h"
#include "utils.h"

// C:/dir/name
std::wstring convert_path_from_win_to_winslash(const std::wstring& win_path) {
    std::wstring winslash;
    for (WCHAR c : win_path) {
        if (c == L'\\') {
            winslash.push_back(L'/');
        } else {
            winslash.push_back(c);
        }
    }
    return winslash;
}

// file:///C:/dir/name
std::wstring convert_path_from_win_to_fileprotocal(const std::wstring& win_path) {
    std::wstring file_url = L"file:///";
    for (WCHAR c : win_path) {
        if (c == L'\\') {
            file_url.push_back(L'/');
        } else {
            file_url.push_back(c);
        }
    }
    DWORD size = static_cast<DWORD>(file_url.size() * 3 + 1);
    std::wstring escaped(size, L'\0');
    if (!InternetCanonicalizeUrlW(file_url.c_str(), &escaped[0], &size, NULL)) {
        escaped.resize(size);
        if (!InternetCanonicalizeUrlW(file_url.c_str(), &escaped[0], &size, NULL)) {
            return file_url;
        }
    }
    escaped.resize(size);
    return escaped;
}

// C:\\dir\\name
std::wstring convert_path_from_win_to_winescaped(const std::wstring& win_path) {
    std::wstring winescaped;
    for (WCHAR c : win_path) {
        if (c == L'\\') {
            winescaped.append(L"\\\\");
        } else {
            winescaped.push_back(c);
        }
    }
    return winescaped;
}

// /c/dir/name
std::wstring convert_path_from_win_to_unix(const std::wstring& win_path) {
    std::wstring unix_path;
    size_t len = win_path.length();
    size_t i2 = 0;
    if (len >= 2 && win_path[1] == L':') {
        unix_path.push_back(L'/');
        unix_path.push_back(win_path[0]);
        i2 = 2;
    }
    for (; i2 < len; i2++) {
        WCHAR c = win_path[i2];
        if (c == L'\\') unix_path.push_back(L'/');
        else unix_path.push_back(c);
    }
    return unix_path;
}

// name
std::wstring convert_path_from_win_to_name(const std::wstring& win_path) {
    size_t pos = win_path.find_last_of(L'\\');
    std::wstring name_part = (pos == std::wstring::npos) ? win_path : win_path.substr(pos + 1);
    return name_part;
}

// /mnt/c/dir/name
// Only converts local drive-letter paths; UNC not specially handled.
std::wstring convert_path_from_win_to_wsl(const std::wstring &win_path) {
    if (win_path.empty()) return win_path;

    if (win_path.size() >= 2 && win_path[1] == L':') {
        wchar_t drive = towlower(win_path[0]);
        std::wstring rest;
        size_t idx = 2;
        if (win_path.size() > 2 && (win_path[2] == L'\\' || win_path[2] == L'/')) idx = 3;
        for (size_t i = idx; i < win_path.size(); ++i) {
            wchar_t c = win_path[i];
            if (c == L'\\') rest.push_back(L'/');
            else if (c == L' ') { rest.push_back(L'\\'); rest.push_back(L' '); }
            else rest.push_back(c);
        }
        std::wstring out = L"/mnt/";
        out.push_back(drive);
        if (!rest.empty()) {
            if (rest.front() != L'/') out.push_back(L'/');
            out += rest;
        }
        return out;
    }

    // Fallback: replace backslashes with slashes and ensure leading slash
    std::wstring s = win_path;
    for (wchar_t &c : s) if (c == L'\\') c = L'/';
    if (!s.empty() && s[0] != L'/') s.insert(s.begin(), L'/');
    return s;
}

std::wstring ConvertPaths(std::vector<std::wstring> paths, std::wstring (*converter)(const std::wstring&)) {
    std::wstring res;
    size_t size = paths.size();
    for (size_t i = 0; i < size; i++) {
        if (converter != nullptr) {
            std::wstring converted = converter(paths[i]);
            res += converted;
        } else {
            res += paths[i];
        }
        res.push_back(L'\n');
    }
    res.pop_back();
    return res;
}

// Helper: truncate a wide string in the middle with ellipsis so total length <= maxLen
std::wstring TruncateMiddle(const std::wstring &s, size_t maxLen) {
    if (s.size() <= maxLen) return s;
    if (maxLen <= 1) return s.substr(0, maxLen);
    size_t keepLeft = (maxLen - 1) / 2;
    size_t keepRight = (maxLen - 1) - keepLeft;
    return s.substr(0, keepLeft) + L'…' + s.substr(s.size() - keepRight);
}

BOOL SetClipboardTextW(const WCHAR *text, SIZE_T cch) {
    if (!OpenClipboard(NULL)) {
        return FALSE;
    }
    if (!EmptyClipboard()) {
        CloseClipboard();
        return FALSE;
    }
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, sizeof(WCHAR) * cch);
    if (hGlobal == NULL) {
        CloseClipboard();
        return FALSE;
    }
    WCHAR *lpText = (WCHAR *)GlobalLock(hGlobal);
    if (lpText == NULL) {
        CloseClipboard();
        return FALSE;
    }
    wmemcpy_s(lpText, cch, text, cch);
    GlobalUnlock(hGlobal);
    SetClipboardData(CF_UNICODETEXT, hGlobal);
    CloseClipboard();
    return TRUE;
}