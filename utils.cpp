#include "pch.h"
#include "utils.h"

// Convert Windows path like "C:\dir\sub" to WSL path "/mnt/c/dir/sub".
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