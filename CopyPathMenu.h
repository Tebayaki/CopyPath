﻿// CopyPathMenu.h: CCopyPathMenu 的声明

#pragma once
#include "resource.h" // 主符号

#include "CopyPath_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;

enum {
    COPYPATH_MENUITEMID_WIN,
    COPYPATH_MENUITEMID_WINSLSH,
    COPYPATH_MENUITEMID_FILEPROTOCAL,
    COPYPATH_MENUITEMID_WINESCAPE,
    COPYPATH_MENUITEMID_UNIX,
    COPYPATH_MENUITEMID_NAME,
    COPYPATH_MENUITEMID_NEXT
};

BOOL SetClipboardTextW(const WCHAR *text, SIZE_T cch);

// CCopyPathMenu

class ATL_NO_VTABLE CCopyPathMenu : public CComObjectRootEx<CComSingleThreadModel>,
                                    public CComCoClass<CCopyPathMenu, &CLSID_CopyPathMenu>,
                                    public IDispatchImpl<ICopyPathMenu, &IID_ICopyPathMenu, &LIBID_CopyPathLib, /*wMajor =*/1, /*wMinor =*/0>,
                                    public IShellExtInit,
                                    public IContextMenu {
    public:
    CCopyPathMenu();
    ~CCopyPathMenu();

    DECLARE_REGISTRY_RESOURCEID(106)

    BEGIN_COM_MAP(CCopyPathMenu)
    COM_INTERFACE_ENTRY(ICopyPathMenu)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(IShellExtInit)
    COM_INTERFACE_ENTRY(IContextMenu)
    END_COM_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct() {
        return S_OK;
    }

    void FinalRelease() {
    }

    public:
    // IShellExtInit
    STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID);

    // IContextMenu
    STDMETHODIMP InvokeCommand(CMINVOKECOMMANDINFO *pici);

    STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);

    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax);

    private:
    typedef struct {
        WCHAR buf[MAX_PATH];
        SIZE_T size;
    } FILEPATH;

    std::vector<FILEPATH> paths__;
    std::wstring buf__;
    HBITMAP hIcon__;
};

OBJECT_ENTRY_AUTO(__uuidof(CopyPathMenu), CCopyPathMenu)
