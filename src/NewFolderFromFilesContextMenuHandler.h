#pragma once
#include <ShlObj.h>
#include <vector>
#include <string>

extern UINT g_cObjCount;

class NewFolderFromFilesContextMenuHandler : public IShellExtInit, public IContextMenu
{
protected:
    LONG m_ObjRefCount;
    std::vector<std::wstring> m_selectedFiles;
    std::wstring m_parentFolder;
    ~NewFolderFromFilesContextMenuHandler();

public:
    NewFolderFromFilesContextMenuHandler();

    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject);

    // IShellExtInit
    HRESULT STDMETHODCALLTYPE Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hKeyProgID);

    // IContextMenu
    HRESULT STDMETHODCALLTYPE GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax);
    HRESULT STDMETHODCALLTYPE InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    HRESULT STDMETHODCALLTYPE QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);

private:
    std::wstring GetCommonPrefix();
    std::wstring GenerateUniqueFolderName(const std::wstring& baseName);
};
