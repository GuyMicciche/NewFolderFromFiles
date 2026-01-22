#pragma once
#include <ShlObj.h>
#include <vector>
#include <string>
#include <map>

extern UINT g_cObjCount;

enum class OrganizeMode
{
    Default = 0,
    ByDay,
    ByMonth,
    ByYear,
    ByMonthYear,
    ByFullDate,
    ByTypeVideo,
    ByTypePhoto,
    ByTypeAudio,
    ByTypeDocument,
    ByTypeOther,
    ByExtension,
    BySize,
    Flatten,
    Numbered,
    Alphabetical,
    COUNT
};

class NewFolderFromFilesContextMenuHandler : public IShellExtInit, public IContextMenu
{
protected:
    LONG m_ObjRefCount;
    std::vector<std::wstring> m_selectedFiles;
    std::wstring m_parentFolder;
    UINT m_idCmdFirst;
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
    std::wstring GenerateUniqueFolderPath(const std::wstring& parent, const std::wstring& baseName);
    HRESULT ExecuteOrganize(OrganizeMode mode);
    HRESULT OrganizeDefault();
    HRESULT OrganizeByDate(OrganizeMode mode);
    HRESULT OrganizeByType();
    HRESULT OrganizeByExtension();
    HRESULT OrganizeBySize();
    HRESULT OrganizeFlatten();
    HRESULT OrganizeNumbered();
    HRESULT OrganizeAlphabetical();
    
    std::wstring GetFileTypeCategory(const std::wstring& path);
    std::wstring GetFileDateFolder(const std::wstring& path, OrganizeMode mode);
    std::wstring GetFileSizeCategory(const std::wstring& path);
    std::wstring GetFileExtension(const std::wstring& path);
    void SelectFolderInExplorer(const std::wstring& folderPath);
    void SelectMultipleFoldersInExplorer(const std::vector<std::wstring>& folders);
};
