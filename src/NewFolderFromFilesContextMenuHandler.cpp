#include "NewFolderFromFilesContextMenuHandler.h"
#include <Shlwapi.h>
#include <strsafe.h>
#include <algorithm>
#include <shobjidl.h>
#include <exdisp.h>
#include <atlbase.h>

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")

NewFolderFromFilesContextMenuHandler::~NewFolderFromFilesContextMenuHandler()
{
    InterlockedDecrement(&g_cObjCount);
}

NewFolderFromFilesContextMenuHandler::NewFolderFromFilesContextMenuHandler() : m_ObjRefCount(1)
{
    InterlockedIncrement(&g_cObjCount);
}

ULONG STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::AddRef()
{
    return InterlockedIncrement(&m_ObjRefCount);
}

ULONG STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::Release()
{
    LONG ref = InterlockedDecrement(&m_ObjRefCount);
    if (ref == 0)
    {
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    *ppvObject = nullptr;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        *ppvObject = static_cast<IShellExtInit*>(this);
    }
    else if (IsEqualIID(riid, IID_IShellExtInit))
    {
        *ppvObject = static_cast<IShellExtInit*>(this);
    }
    else if (IsEqualIID(riid, IID_IContextMenu))
    {
        *ppvObject = static_cast<IContextMenu*>(this);
    }
    else
    {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::Initialize(
    PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hKeyProgID)
{
    m_selectedFiles.clear();
    m_parentFolder.clear();

    if (!pdtobj)
        return E_INVALIDARG;

    FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stg = { TYMED_HGLOBAL };

    if (FAILED(pdtobj->GetData(&fmt, &stg)))
        return E_INVALIDARG;

    HDROP hDrop = static_cast<HDROP>(GlobalLock(stg.hGlobal));
    if (!hDrop)
    {
        ReleaseStgMedium(&stg);
        return E_INVALIDARG;
    }

    UINT fileCount = DragQueryFileW(hDrop, 0xFFFFFFFF, nullptr, 0);
    if (fileCount == 0)
    {
        GlobalUnlock(stg.hGlobal);
        ReleaseStgMedium(&stg);
        return E_INVALIDARG;
    }

    wchar_t filePath[MAX_PATH];
    for (UINT i = 0; i < fileCount; i++)
    {
        if (DragQueryFileW(hDrop, i, filePath, MAX_PATH))
        {
            m_selectedFiles.push_back(filePath);
        }
    }

    // Get parent folder from first file
    if (!m_selectedFiles.empty())
    {
        wchar_t parentPath[MAX_PATH];
        StringCchCopyW(parentPath, MAX_PATH, m_selectedFiles[0].c_str());
        PathRemoveFileSpecW(parentPath);
        m_parentFolder = parentPath;
    }

    GlobalUnlock(stg.hGlobal);
    ReleaseStgMedium(&stg);

    return m_selectedFiles.empty() ? E_FAIL : S_OK;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::GetCommandString(
    UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax)
{
    if (idCmd != 0)
        return E_INVALIDARG;

    if (uFlags == GCS_HELPTEXTW)
    {
        StringCchCopyW(reinterpret_cast<LPWSTR>(pszName), cchMax, 
            L"Create a new folder containing the selected items");
        return S_OK;
    }
    else if (uFlags == GCS_VERBW)
    {
        StringCchCopyW(reinterpret_cast<LPWSTR>(pszName), cchMax, L"newfolderfromfiles");
        return S_OK;
    }

    return E_NOTIMPL;
}

std::wstring NewFolderFromFilesContextMenuHandler::GetCommonPrefix()
{
    if (m_selectedFiles.empty())
        return L"New Folder";

    // Get just the filenames without extensions
    std::vector<std::wstring> names;
    for (const auto& path : m_selectedFiles)
    {
        wchar_t name[MAX_PATH];
        StringCchCopyW(name, MAX_PATH, PathFindFileNameW(path.c_str()));
        PathRemoveExtensionW(name);
        names.push_back(name);
    }

    if (names.size() == 1)
        return names[0];

    // Find common prefix
    std::wstring prefix = names[0];
    for (size_t i = 1; i < names.size() && !prefix.empty(); i++)
    {
        size_t j = 0;
        while (j < prefix.length() && j < names[i].length() && 
               towlower(prefix[j]) == towlower(names[i][j]))
        {
            j++;
        }
        prefix = prefix.substr(0, j);
    }

    // Trim trailing spaces, underscores, dashes
    while (!prefix.empty())
    {
        wchar_t last = prefix.back();
        if (last == L' ' || last == L'_' || last == L'-' || last == L'.')
            prefix.pop_back();
        else
            break;
    }

    return prefix.empty() ? L"New Folder" : prefix;
}

std::wstring NewFolderFromFilesContextMenuHandler::GenerateUniqueFolderName(const std::wstring& baseName)
{
    std::wstring folderPath = m_parentFolder + L"\\" + baseName;
    
    if (!PathFileExistsW(folderPath.c_str()))
        return folderPath;

    for (int i = 2; i < 1000; i++)
    {
        wchar_t buffer[32];
        StringCchPrintfW(buffer, 32, L" (%d)", i);
        folderPath = m_parentFolder + L"\\" + baseName + buffer;
        
        if (!PathFileExistsW(folderPath.c_str()))
            return folderPath;
    }

    return m_parentFolder + L"\\New Folder";
}

// Find the current Explorer window's shell view
static HRESULT GetActiveShellView(const std::wstring& folderPath, IShellView** ppShellView)
{
    *ppShellView = nullptr;
    
    CComPtr<IShellWindows> pShellWindows;
    HRESULT hr = pShellWindows.CoCreateInstance(CLSID_ShellWindows);
    if (FAILED(hr)) return hr;

    long count = 0;
    pShellWindows->get_Count(&count);

    for (long i = 0; i < count; i++)
    {
        CComVariant vi(i);
        CComPtr<IDispatch> pDisp;
        if (FAILED(pShellWindows->Item(vi, &pDisp)) || !pDisp)
            continue;

        CComQIPtr<IWebBrowserApp> pWebBrowserApp(pDisp);
        if (!pWebBrowserApp)
            continue;

        CComBSTR bstrUrl;
        if (FAILED(pWebBrowserApp->get_LocationURL(&bstrUrl)) || !bstrUrl)
            continue;

        // Check if this window is showing our folder
        CComQIPtr<IServiceProvider> pServiceProvider(pWebBrowserApp);
        if (!pServiceProvider)
            continue;

        CComPtr<IShellBrowser> pShellBrowser;
        if (FAILED(pServiceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&pShellBrowser))))
            continue;

        CComPtr<IShellView> pShellView;
        if (SUCCEEDED(pShellBrowser->QueryActiveShellView(&pShellView)))
        {
            // Check if this is the right folder
            CComQIPtr<IFolderView> pFolderView(pShellView);
            if (pFolderView)
            {
                CComPtr<IPersistFolder2> pPersistFolder;
                if (SUCCEEDED(pFolderView->GetFolder(IID_PPV_ARGS(&pPersistFolder))))
                {
                    PIDLIST_ABSOLUTE pidlFolder = nullptr;
                    if (SUCCEEDED(pPersistFolder->GetCurFolder(&pidlFolder)))
                    {
                        wchar_t szPath[MAX_PATH];
                        if (SHGetPathFromIDListW(pidlFolder, szPath))
                        {
                            if (_wcsicmp(szPath, folderPath.c_str()) == 0)
                            {
                                *ppShellView = pShellView.Detach();
                                CoTaskMemFree(pidlFolder);
                                return S_OK;
                            }
                        }
                        CoTaskMemFree(pidlFolder);
                    }
                }
            }
        }
    }

    return E_FAIL;
}

HRESULT NewFolderFromFilesContextMenuHandler::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    // Check if called by command ID
    if (HIWORD(pici->lpVerb) != 0)
    {
        if (lstrcmpiA(pici->lpVerb, "newfolderfromfiles") != 0)
            return E_INVALIDARG;
    }
    else if (LOWORD(pici->lpVerb) != 0)
    {
        return E_INVALIDARG;
    }

    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    // Use IFileOperation for proper undo support (single undo for entire operation)
    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr))
        return hr;

    // Set operation flags - allow undo, no confirmation dialogs
    hr = pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);
    if (FAILED(hr))
        return hr;

    // Generate folder name and path
    std::wstring suggestedName = GetCommonPrefix();
    std::wstring folderPath = GenerateUniqueFolderName(suggestedName);
    std::wstring folderName = PathFindFileNameW(folderPath.c_str());

    // Get parent folder as shell item
    CComPtr<IShellItem> pParentItem;
    hr = SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
    if (FAILED(hr))
        return hr;

    // Create the new folder operation
    hr = pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, folderName.c_str(), nullptr, nullptr);
    if (FAILED(hr))
        return hr;

    // Execute folder creation first
    hr = pFileOp->PerformOperations();
    if (FAILED(hr))
        return hr;

    // Now create a new file operation for the moves (same undo group)
    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr))
        return hr;

    hr = pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);
    if (FAILED(hr))
        return hr;

    // Get destination folder as shell item
    CComPtr<IShellItem> pDestFolder;
    hr = SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder));
    if (FAILED(hr))
    {
        // Folder might not exist yet, wait briefly
        Sleep(50);
        hr = SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder));
        if (FAILED(hr))
            return hr;
    }

    // Add move operations for each selected file
    for (const auto& filePath : m_selectedFiles)
    {
        CComPtr<IShellItem> pItem;
        hr = SHCreateItemFromParsingName(filePath.c_str(), nullptr, IID_PPV_ARGS(&pItem));
        if (SUCCEEDED(hr))
        {
            pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    // Execute moves
    hr = pMoveOp->PerformOperations();
    if (FAILED(hr))
        return hr;

    // Select the new folder in the current Explorer window and enter rename mode
    CComPtr<IShellView> pShellView;
    if (SUCCEEDED(GetActiveShellView(m_parentFolder, &pShellView)))
    {
        // Get the PIDL for the new folder relative to parent
        PIDLIST_ABSOLUTE pidlFolder = ILCreateFromPathW(folderPath.c_str());
        if (pidlFolder)
        {
            PCUITEMID_CHILD pidlChild = ILFindLastID(pidlFolder);
            
            // Select the item
            pShellView->SelectItem(pidlChild, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_EDIT);
            
            ILFree(pidlFolder);
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::QueryContextMenu(
    HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (uFlags & CMF_DEFAULTONLY)
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    // Only show if we have files selected
    if (m_selectedFiles.empty())
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    // Insert menu item
    MENUITEMINFOW mii = {};
    mii.cbSize = sizeof(MENUITEMINFOW);
    mii.fMask = MIIM_ID | MIIM_STRING | MIIM_STATE;
    mii.wID = idCmdFirst;
    mii.fState = MFS_ENABLED;
    mii.dwTypeData = const_cast<LPWSTR>(L"New folder with selection");

    if (!InsertMenuItemW(hmenu, indexMenu, TRUE, &mii))
        return HRESULT_FROM_WIN32(GetLastError());

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
}
