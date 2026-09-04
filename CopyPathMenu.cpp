// CopyPathMenu.cpp: CCopyPathMenu 的实现

#include "pch.h"
#include "CopyPathMenu.h"
#include "utils.h"

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

STDMETHODIMP CCopyPathMenu::InvokeCommand(CMINVOKECOMMANDINFO *pici) {
    if (HIWORD(pici->lpVerb) != 0) {
        return E_INVALIDARG;
    }
    size_t size = paths__.size();
    std::wstring paths;
    switch (LOWORD(pici->lpVerb)) {
        case COPYPATH_MENUITEMID_WIN:
            paths = ConvertPaths(nullptr);
            break;
        case COPYPATH_MENUITEMID_WINSLSH:
            paths = ConvertPaths(convert_path_from_win_to_winslash);
            break;
        case COPYPATH_MENUITEMID_FILEPROTOCAL:
            paths = ConvertPaths(convert_path_from_win_to_fileprotocal);
            break;
        case COPYPATH_MENUITEMID_WINESCAPE:
            paths = ConvertPaths(convert_path_from_win_to_winescaped);
            break;
        case COPYPATH_MENUITEMID_UNIX:
            paths = ConvertPaths(convert_path_from_win_to_unix);
            break;
        case COPYPATH_MENUITEMID_NAME:
            paths = ConvertPaths(convert_path_from_win_to_name);
            break;
        case COPYPATH_MENUITEMID_WSL:
            paths = ConvertPaths(convert_path_from_win_to_wsl);
            break;
        default:
            return E_INVALIDARG;
    }
    SetClipboardTextW(paths.c_str(), paths.size());
    return S_OK;
}

STDMETHODIMP CCopyPathMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
    if (uFlags & CMF_DEFAULTONLY) {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
    }

    HMENU hSubMenu = CreatePopupMenu();

    // Prepare display strings using the first selected path (if available)
    std::wstring rawPath = paths__.empty() ? L"" : paths__[0].buf;

    // (&A) C:\DIR\NAME
    const std::wstring prefixA = L"(&A) ";
    std::wstring winLabel = prefixA + TruncateMiddle(rawPath, MAX_LABEL_LEN);

    // (&S) C:/DIR/NAME
    const std::wstring prefixS = L"(&S) ";
    const std::wstring winslash = convert_path_from_win_to_winslash(rawPath);
    std::wstring winslashLabel = prefixS + TruncateMiddle(winslash, MAX_LABEL_LEN);

    // (&D) file:///C:/DIR/NAME
    const std::wstring prefixD = L"(&D) ";
    std::wstring fileProtoPath = convert_path_from_win_to_fileprotocal(rawPath);
    std::wstring fileProtoLabel = prefixD + TruncateMiddle(fileProtoPath, MAX_LABEL_LEN);

    // (&F) C:\\DIR\\NAME  (escaped backslashes)
    const std::wstring prefixF = L"(&F) ";
    std::wstring winEscaped = convert_path_from_win_to_winescaped(rawPath);
    std::wstring winEscapedLabel = prefixF + TruncateMiddle(winEscaped, MAX_LABEL_LEN);

    // (&G) /C/DIR/NAME
    const std::wstring prefixG = L"(&G) ";
	std::wstring unixPath = convert_path_from_win_to_unix(rawPath);
    std::wstring unixLabel = prefixG + TruncateMiddle(unixPath, MAX_LABEL_LEN);

    // (&Q) NAME
    const std::wstring prefixQ = L"(&Q) ";
	std::wstring namePart = convert_path_from_win_to_name(rawPath);
    std::wstring nameLabel = prefixQ + TruncateMiddle(namePart, MAX_LABEL_LEN);

    // (&W) /mnt/c/dir/name  (WSL style)
    const std::wstring prefixWSL = L"(&W) ";
    std::wstring wslPathDisplay = convert_path_from_win_to_wsl(rawPath);
    std::wstring wslLabel = prefixWSL + TruncateMiddle(wslPathDisplay, MAX_LABEL_LEN);

    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WIN, winLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WINSLSH, winslashLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_FILEPROTOCAL, fileProtoLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WINESCAPE, winEscapedLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_UNIX, unixLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_NAME, nameLabel.c_str());
    AppendMenuW(hSubMenu, MF_STRING, (UINT_PTR)idCmdFirst + COPYPATH_MENUITEMID_WSL, wslLabel.c_str());

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
