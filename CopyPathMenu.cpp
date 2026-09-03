// CopyPathMenu.cpp: CCopyPathMenu 的实现

#include "pch.h"
#include "CopyPathMenu.h"

// CCopyPathMenu

CCopyPathMenu::CCopyPathMenu() {
    hIcon__ = NULL;
}

CCopyPathMenu::~CCopyPathMenu() {
    if (hIcon__) {
        DeleteObject(hIcon__);
    }
}

STDMETHODIMP CCopyPathMenu::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID) {
    // Right click on background of explorer
    if (pidlFolder != nullptr) {
        if (paths__.size() != 1) {
            paths__.resize(1);
        }
        if (SHGetPathFromIDListW(pidlFolder, paths__[0].buf)) {
            paths__[0].size = MAX_PATH;
            return S_OK;
        }
        return E_INVALIDARG;
    }

    // Right click on item
    FORMATETC fmt = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM stg = {TYMED_HGLOBAL};
    if (FAILED(pdtobj->GetData(&fmt, &stg))) {
        return E_INVALIDARG;
    }

    HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
    if (hDrop == NULL) {
        return E_INVALIDARG;
    }

    UINT filesCnt = DragQueryFileW(hDrop, -1, NULL, 0);
    if (filesCnt == 0) {
        GlobalUnlock(stg.hGlobal);
        ReleaseStgMedium(&stg);
        return E_INVALIDARG;
    }
    if (paths__.size() != filesCnt) {
        paths__.resize(filesCnt);
    }
    for (UINT i = 0; i < filesCnt; i++) {
        SIZE_T size = DragQueryFileW(hDrop, i, paths__[i].buf, MAX_PATH);
        if (size == 0) {
            i--;
            filesCnt--;
            paths__.pop_back();
        }
        else {
            paths__[i].size = size + 1;
        }
    }

    GlobalUnlock(stg.hGlobal);
    ReleaseStgMedium(&stg);
    return S_OK;
}

// Helper: truncate a wide string in the middle with ellipsis so total length <= maxLen
static std::wstring TruncateMiddle(const std::wstring &s, size_t maxLen) {
    if (s.size() <= maxLen) return s;
    if (maxLen <= 1) return s.substr(0, maxLen);
    size_t keepLeft = (maxLen - 1) / 2;
    size_t keepRight = (maxLen - 1) - keepLeft;
    return s.substr(0, keepLeft) + L'…' + s.substr(s.size() - keepRight);
}

STDMETHODIMP CCopyPathMenu::InvokeCommand(CMINVOKECOMMANDINFO *pici) {
    if (HIWORD(pici->lpVerb) != 0) {
        return E_INVALIDARG;
    }
    size_t size = paths__.size();
    switch (LOWORD(pici->lpVerb)) {
        case COPYPATH_MENUITEMID_WIN:
            if (size == 1) {
                SetClipboardTextW(paths__[0].buf, paths__[0].size);
            }
            else {
                buf__.clear();
                for (size_t i = 0; i < size; i++) {
                    size_t len = paths__[i].size - 1;
                    for (size_t i2 = 0; i2 < len; i2++) {
                        buf__.push_back(paths__[i].buf[i2]);
                    }
                    buf__.push_back(L'\n');
                }
                buf__.back() = L'\0';
                SetClipboardTextW(buf__.c_str(), buf__.size());
            }
            break;
        case COPYPATH_MENUITEMID_WINSLSH:
        case COPYPATH_MENUITEMID_FILEPROTOCAL:
            buf__.clear();
            for (size_t i = 0; i < size; i++) {
                if (LOWORD(pici->lpVerb) == COPYPATH_MENUITEMID_FILEPROTOCAL) {
                    buf__.append(L"file:///");
                }
                size_t len = paths__[i].size - 1;
                for (size_t i2 = 0; i2 < len; i2++) {
                    WCHAR c = paths__[i].buf[i2];
                    if (c == L'\\') {
                        buf__.push_back(L'/');
                    }
                    else {
                        buf__.push_back(c);
                    }
                }
                buf__.push_back(L'\n');
            }
            buf__.back() = L'\0';
            SetClipboardTextW(buf__.c_str(), buf__.size());
            break;
        case COPYPATH_MENUITEMID_WINESCAPE:
            buf__.clear();
            for (size_t i = 0; i < size; i++) {
                size_t len = paths__[i].size - 1;
                for (size_t i2 = 0; i2 < len; i2++) {
                    WCHAR c = paths__[i].buf[i2];
                    if (c == L'\\') {
                        buf__ += L"\\\\";
                    }
                    else {
                        buf__.push_back(c);
                    }
                }
                buf__.push_back(L'\n');
            }
            buf__.back() = L'\0';
            SetClipboardTextW(buf__.c_str(), buf__.size());
            break;
        case COPYPATH_MENUITEMID_UNIX:
            buf__.clear();
            for (size_t i = 0; i < size; i++) {
                buf__.push_back(L'/');
                size_t len = paths__[i].size - 1;
                size_t i2 = 0;
                if (len >= 2 && paths__[i].buf[1] == L':') {
                    buf__.push_back(paths__[i].buf[0]);
                    i2 = 2;
                }
                for (; i2 < len; i2++) {
                    WCHAR c = paths__[i].buf[i2];
                    if (c == L'\\') {
                        buf__.push_back(L'/');
                    }
                    else {
                        buf__.push_back(c);
                    }
                }
                buf__.push_back(L'\n');
            }
            buf__.back() = L'\0';
            SetClipboardTextW(buf__.c_str(), buf__.size());
            break;
        case COPYPATH_MENUITEMID_NAME:
            buf__.clear();
            for (size_t i = 0; i < size; i++) {
                size_t len = paths__[i].size - 1;
                if (len == 3 && paths__[i].buf[1] == L':' && paths__[i].buf[2] == L'\\') {
                    buf__.push_back(paths__[i].buf[0]);
                    buf__ += L":\\";
                }
                size_t i2 = len - 1;
                for (;; i2--) {
                    if (paths__[i].buf[i2] == L'\\') {
                        ++i2;
                        break;
                    }
                    if (i2 == 0) {
                        break;
                    }
                }
                for (; i2 < len; i2++) {
                    buf__.push_back(paths__[i].buf[i2]);
                }
                buf__.push_back(L'\n');
            }
            buf__.back() = L'\0';
            SetClipboardTextW(buf__.c_str(), buf__.size());
            break;
        default:
            return E_INVALIDARG;
    }
    return S_OK;
}

STDMETHODIMP CCopyPathMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
    if (uFlags & CMF_DEFAULTONLY) {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
    }

    HMENU hSubMenu = CreatePopupMenu();

    // Prepare display strings using the first selected path (if available)
    std::wstring firstPath = L"";
    if (!paths__.empty()) {
        firstPath = paths__[0].buf;
    }

    auto replace_char = [](const std::wstring &s, WCHAR from, WCHAR to) {
        std::wstring out = s;
        for (WCHAR &c : out) {
            if (c == from) c = to;
        }
        return out;
    };

    // Build variants (keep raw path separately and apply truncation to the displayed part)
    std::wstring rawPath = firstPath;
    std::wstring winslash = replace_char(rawPath, L'\\', L'/');

    // (&A) C:\DIR\NAME
    const std::wstring prefixA = L"(&A) ";
    std::wstring winLabel = prefixA + TruncateMiddle(rawPath, MAX_LABEL_LEN);

    // (&S) C:/DIR/NAME
    const std::wstring prefixS = L"(&S) ";
    std::wstring winslashLabel = prefixS + TruncateMiddle(winslash, MAX_LABEL_LEN);

    // (&D) file:///C:/DIR/NAME
    const std::wstring prefixD = L"(&D) ";
    std::wstring fileProtoLabel = prefixD + TruncateMiddle(L"file:///" + winslash, MAX_LABEL_LEN);

    // (&F) C:\\DIR\\NAME  (escaped backslashes)
    std::wstring winEscaped;
    for (WCHAR c : rawPath) {
        if (c == L'\\') {
            winEscaped.append(L"\\\\");
        } else {
            winEscaped.push_back(c);
        }
    }
    const std::wstring prefixF = L"(&F) ";
    std::wstring winEscapedLabel = prefixF + TruncateMiddle(winEscaped, MAX_LABEL_LEN);

    // (&G) /C/DIR/NAME
    std::wstring unixPath;
    size_t len = rawPath.length();
    size_t i2 = 0;
    if (len >= 2 && rawPath[1] == L':') {
        unixPath.push_back(L'/');
        unixPath.push_back(rawPath[0]);
        i2 = 2;
    }
    for (; i2 < len; i2++) {
        WCHAR c = rawPath[i2];
        if (c == L'\\') unixPath.push_back(L'/');
        else unixPath.push_back(c);
    }
    const std::wstring prefixG = L"(&G) ";
    std::wstring unixLabel = prefixG + TruncateMiddle(unixPath, MAX_LABEL_LEN);

    // (&Q) NAME
    size_t pos = rawPath.find_last_of(L'\\');
    std::wstring namePart = (pos == std::wstring::npos) ? rawPath : rawPath.substr(pos + 1);
    const std::wstring prefixQ = L"(&Q) ";
    std::wstring nameLabel = prefixQ + TruncateMiddle(namePart, MAX_LABEL_LEN);

    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WIN, winLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WINSLSH, winslashLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_FILEPROTOCAL, fileProtoLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WINESCAPE, winEscapedLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_UNIX, unixLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_NAME, nameLabel.c_str());

    const BYTE bits[] = {0xff, 0xff, 0xf8, 0x1f, 0xe2, 0x47, 0xef, 0xf7, 0xef, 0xf7, 0xe8, 0x77, 0xef, 0xf7, 0xe8, 0x17, 0xef, 0xf7, 0xe8, 0x17, 0xef, 0xf7, 0xe8, 0xf7, 0xef, 0xe7, 0xe0, 0x07, 0xff, 0xff, 0xff, 0xff};
    hIcon__ = CreateBitmap(16, 16, 1, 1, bits);

    MENUITEMINFOW info = {sizeof(MENUITEMINFOW)};
    info.fMask = MIIM_FTYPE | MIIM_BITMAP | MIIM_STRING | MIIM_SUBMENU;
    info.fType = MFT_STRING;
    info.dwTypeData = L"Copy Path(&Q)";
    info.hSubMenu = hSubMenu;
    info.hbmpItem = hIcon__;
    InsertMenuItemW(hmenu, indexMenu, TRUE, &info);

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, COPYPATH_MENUITEMID_NEXT);
}

STDMETHODIMP CCopyPathMenu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax) {
    if (idCmd >= COPYPATH_MENUITEMID_NEXT) {
        return E_INVALIDARG;
    }
    if (uType == GCS_HELPTEXTW) {
        wmemcpy_s((WCHAR *)pszName, cchMax, L"Copy Path in specific format", 29);
        return S_OK;
    }
    else if (uType == GCS_HELPTEXTA) {
        memcpy_s(pszName, cchMax, "Copy Path in specific format", 29);
        return S_OK;
    }
    return E_INVALIDARG;
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
