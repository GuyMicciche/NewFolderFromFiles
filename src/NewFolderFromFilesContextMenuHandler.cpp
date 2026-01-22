#include "NewFolderFromFilesContextMenuHandler.h"
#include <Shlwapi.h>
#include <strsafe.h>
#include <algorithm>
#include <shobjidl.h>
#include <exdisp.h>
#include <atlbase.h>
#include <set>

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")

// Menu command IDs (relative to idCmdFirst)
#define CMD_DEFAULT         0
#define CMD_SUBMENU_DEFAULT 1
#define CMD_BY_DAY          2
#define CMD_BY_MONTH        3
#define CMD_BY_YEAR         4
#define CMD_BY_MONTHYEAR    5
#define CMD_BY_FULLDATE     6
#define CMD_BY_TYPE         7
#define CMD_BY_EXTENSION    8
#define CMD_BY_SIZE         9
#define CMD_FLATTEN         10
#define CMD_NUMBERED        11
#define CMD_ALPHABETICAL    12
#define CMD_COUNT           13

NewFolderFromFilesContextMenuHandler::~NewFolderFromFilesContextMenuHandler()
{
    InterlockedDecrement(&g_cObjCount);
}

NewFolderFromFilesContextMenuHandler::NewFolderFromFilesContextMenuHandler() : m_ObjRefCount(1), m_idCmdFirst(0)
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
        *ppvObject = static_cast<IShellExtInit*>(this);
    else if (IsEqualIID(riid, IID_IShellExtInit))
        *ppvObject = static_cast<IShellExtInit*>(this);
    else if (IsEqualIID(riid, IID_IContextMenu))
        *ppvObject = static_cast<IContextMenu*>(this);
    else
        return E_NOINTERFACE;

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
            m_selectedFiles.push_back(filePath);
    }

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

    std::wstring prefix = names[0];
    for (size_t i = 1; i < names.size() && !prefix.empty(); i++)
    {
        size_t j = 0;
        while (j < prefix.length() && j < names[i].length() &&
            towlower(prefix[j]) == towlower(names[i][j]))
            j++;
        prefix = prefix.substr(0, j);
    }

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
    return GenerateUniqueFolderPath(m_parentFolder, baseName);
}

std::wstring NewFolderFromFilesContextMenuHandler::GenerateUniqueFolderPath(const std::wstring& parent, const std::wstring& baseName)
{
    std::wstring folderPath = parent + L"\\" + baseName;

    if (!PathFileExistsW(folderPath.c_str()))
        return folderPath;

    for (int i = 2; i < 1000; i++)
    {
        wchar_t buffer[32];
        StringCchPrintfW(buffer, 32, L" (%d)", i);
        folderPath = parent + L"\\" + baseName + buffer;

        if (!PathFileExistsW(folderPath.c_str()))
            return folderPath;
    }

    return parent + L"\\New Folder";
}

// Find the current Explorer window's shell view
static HRESULT GetActiveShellView(const std::wstring& folderPath, IShellView** ppShellView, IShellBrowser** ppShellBrowser = nullptr)
{
    *ppShellView = nullptr;
    if (ppShellBrowser) *ppShellBrowser = nullptr;

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

        CComQIPtr<IServiceProvider> pServiceProvider(pWebBrowserApp);
        if (!pServiceProvider)
            continue;

        CComPtr<IShellBrowser> pShellBrowser;
        if (FAILED(pServiceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&pShellBrowser))))
            continue;

        CComPtr<IShellView> pShellView;
        if (SUCCEEDED(pShellBrowser->QueryActiveShellView(&pShellView)))
        {
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
                                if (ppShellBrowser) *ppShellBrowser = pShellBrowser.Detach();
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

void NewFolderFromFilesContextMenuHandler::SelectFolderInExplorer(const std::wstring& folderPath)
{
    CComPtr<IShellView> pShellView;
    if (SUCCEEDED(GetActiveShellView(m_parentFolder, &pShellView)))
    {
        PIDLIST_ABSOLUTE pidlFolder = ILCreateFromPathW(folderPath.c_str());
        if (pidlFolder)
        {
            PCUITEMID_CHILD pidlChild = ILFindLastID(pidlFolder);
            pShellView->SelectItem(pidlChild, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_EDIT);
            ILFree(pidlFolder);
        }
    }
}

void NewFolderFromFilesContextMenuHandler::SelectMultipleFoldersInExplorer(const std::vector<std::wstring>& folders)
{
    if (folders.empty()) return;

    CComPtr<IShellView> pShellView;
    if (SUCCEEDED(GetActiveShellView(m_parentFolder, &pShellView)))
    {
        bool first = true;
        for (const auto& folderPath : folders)
        {
            PIDLIST_ABSOLUTE pidlFolder = ILCreateFromPathW(folderPath.c_str());
            if (pidlFolder)
            {
                PCUITEMID_CHILD pidlChild = ILFindLastID(pidlFolder);
                UINT flags = SVSI_SELECT | SVSI_ENSUREVISIBLE;
                if (first)
                {
                    flags |= SVSI_DESELECTOTHERS | SVSI_FOCUSED;
                    first = false;
                }
                pShellView->SelectItem(pidlChild, flags);
                ILFree(pidlFolder);
            }
        }
    }
}

std::wstring NewFolderFromFilesContextMenuHandler::GetFileExtension(const std::wstring& path)
{
    const wchar_t* ext = PathFindExtensionW(path.c_str());
    if (ext && *ext == L'.')
    {
        std::wstring result(ext + 1);
        // Convert to uppercase for consistency
        for (auto& c : result) c = towupper(c);
        return result;
    }
    return L"No Extension";
}

std::wstring NewFolderFromFilesContextMenuHandler::GetFileTypeCategory(const std::wstring& path)
{
    std::wstring ext = GetFileExtension(path);
    for (auto& c : ext) c = towlower(c);

    // Video
    if (ext == L"mp4" || ext == L"avi" || ext == L"mkv" || ext == L"mov" ||
        ext == L"wmv" || ext == L"flv" || ext == L"webm" || ext == L"m4v" ||
        ext == L"mpg" || ext == L"mpeg" || ext == L"3gp")
        return L"Video";

    // Photo
    if (ext == L"jpg" || ext == L"jpeg" || ext == L"png" || ext == L"gif" ||
        ext == L"bmp" || ext == L"tiff" || ext == L"tif" || ext == L"webp" ||
        ext == L"ico" || ext == L"svg" || ext == L"raw" || ext == L"psd" ||
        ext == L"heic" || ext == L"heif")
        return L"Photo";

    // Audio
    if (ext == L"mp3" || ext == L"wav" || ext == L"flac" || ext == L"aac" ||
        ext == L"ogg" || ext == L"wma" || ext == L"m4a" || ext == L"aiff")
        return L"Audio";

    // Document
    if (ext == L"doc" || ext == L"docx" || ext == L"pdf" || ext == L"txt" ||
        ext == L"rtf" || ext == L"odt" || ext == L"xls" || ext == L"xlsx" ||
        ext == L"ppt" || ext == L"pptx" || ext == L"csv" || ext == L"md")
        return L"Document";

    return L"Other";
}

std::wstring NewFolderFromFilesContextMenuHandler::GetFileDateFolder(const std::wstring& path, OrganizeMode mode)
{
    WIN32_FILE_ATTRIBUTE_DATA fileInfo;
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &fileInfo))
        return L"Unknown Date";

    SYSTEMTIME st;
    FileTimeToSystemTime(&fileInfo.ftLastWriteTime, &st);

    wchar_t buffer[64];
    switch (mode)
    {
    case OrganizeMode::ByDay:
        StringCchPrintfW(buffer, 64, L"%02d", st.wDay);
        break;
    case OrganizeMode::ByMonth:
        StringCchPrintfW(buffer, 64, L"%02d", st.wMonth);
        break;
    case OrganizeMode::ByYear:
        StringCchPrintfW(buffer, 64, L"%04d", st.wYear);
        break;
    case OrganizeMode::ByMonthYear:
        StringCchPrintfW(buffer, 64, L"%04d-%02d", st.wYear, st.wMonth);
        break;
    case OrganizeMode::ByFullDate:
    default:
        StringCchPrintfW(buffer, 64, L"%04d-%02d-%02d", st.wYear, st.wMonth, st.wDay);
        break;
    }
    return buffer;
}

std::wstring NewFolderFromFilesContextMenuHandler::GetFileSizeCategory(const std::wstring& path)
{
    WIN32_FILE_ATTRIBUTE_DATA fileInfo;
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &fileInfo))
        return L"Unknown Size";

    ULONGLONG size = (static_cast<ULONGLONG>(fileInfo.nFileSizeHigh) << 32) | fileInfo.nFileSizeLow;

    if (size < 1024 * 1024)  // < 1 MB
        return L"Small (under 1 MB)";
    else if (size < 100 * 1024 * 1024)  // < 100 MB
        return L"Medium (1-100 MB)";
    else
        return L"Large (over 100 MB)";
}

HRESULT NewFolderFromFilesContextMenuHandler::ExecuteOrganize(OrganizeMode mode)
{
    switch (mode)
    {
    case OrganizeMode::Default:
        return OrganizeDefault();
    case OrganizeMode::ByDay:
    case OrganizeMode::ByMonth:
    case OrganizeMode::ByYear:
    case OrganizeMode::ByMonthYear:
    case OrganizeMode::ByFullDate:
        return OrganizeByDate(mode);
    case OrganizeMode::ByTypeVideo:
    case OrganizeMode::ByTypePhoto:
    case OrganizeMode::ByTypeAudio:
    case OrganizeMode::ByTypeDocument:
    case OrganizeMode::ByTypeOther:
        return OrganizeByType();
    case OrganizeMode::ByExtension:
        return OrganizeByExtension();
    case OrganizeMode::BySize:
        return OrganizeBySize();
    case OrganizeMode::Flatten:
        return OrganizeFlatten();
    case OrganizeMode::Numbered:
        return OrganizeNumbered();
    case OrganizeMode::Alphabetical:
        return OrganizeAlphabetical();
    default:
        return E_INVALIDARG;
    }
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeDefault()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::wstring suggestedName = GetCommonPrefix();
    std::wstring folderPath = GenerateUniqueFolderName(suggestedName);
    std::wstring folderName = PathFindFileNameW(folderPath.c_str());

    CComPtr<IShellItem> pParentItem;
    hr = SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
    if (FAILED(hr)) return hr;

    pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, folderName.c_str(), nullptr, nullptr);
    hr = pFileOp->PerformOperations();
    if (FAILED(hr)) return hr;

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    Sleep(50);
    CComPtr<IShellItem> pDestFolder;
    hr = SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder));
    if (FAILED(hr)) return hr;

    for (const auto& filePath : m_selectedFiles)
    {
        CComPtr<IShellItem> pItem;
        if (SUCCEEDED(SHCreateItemFromParsingName(filePath.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
            pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
    }

    hr = pMoveOp->PerformOperations();
    if (FAILED(hr)) return hr;

    SelectFolderInExplorer(folderPath);
    return S_OK;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeByDate(OrganizeMode mode)
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    std::map<std::wstring, std::vector<std::wstring>> groups;
    for (const auto& file : m_selectedFiles)
        groups[GetFileDateFolder(file, mode)].push_back(file);

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    for (auto& [dateName, files] : groups)
    {
        std::wstring folderPath = GenerateUniqueFolderPath(m_parentFolder, dateName);
        std::wstring folderName = PathFindFileNameW(folderPath.c_str());

        if (!PathFileExistsW(folderPath.c_str()))
        {
            CComPtr<IShellItem> pParentItem;
            SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
            pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, folderName.c_str(), nullptr, nullptr);
        }
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (auto& [dateName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + dateName;
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        for (const auto& file : files)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeByType()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    std::map<std::wstring, std::vector<std::wstring>> groups;
    for (const auto& file : m_selectedFiles)
        groups[GetFileTypeCategory(file)].push_back(file);

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    for (auto& [typeName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + typeName;
        if (!PathFileExistsW(folderPath.c_str()))
        {
            CComPtr<IShellItem> pParentItem;
            SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
            pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, typeName.c_str(), nullptr, nullptr);
        }
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (auto& [typeName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + typeName;
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        for (const auto& file : files)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeByExtension()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    std::map<std::wstring, std::vector<std::wstring>> groups;
    for (const auto& file : m_selectedFiles)
        groups[GetFileExtension(file)].push_back(file);

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    for (auto& [extName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + extName;
        if (!PathFileExistsW(folderPath.c_str()))
        {
            CComPtr<IShellItem> pParentItem;
            SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
            pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, extName.c_str(), nullptr, nullptr);
        }
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (auto& [extName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + extName;
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        for (const auto& file : files)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeBySize()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    std::map<std::wstring, std::vector<std::wstring>> groups;
    for (const auto& file : m_selectedFiles)
        groups[GetFileSizeCategory(file)].push_back(file);

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    for (auto& [sizeName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + sizeName;
        if (!PathFileExistsW(folderPath.c_str()))
        {
            CComPtr<IShellItem> pParentItem;
            SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
            pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, sizeName.c_str(), nullptr, nullptr);
        }
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (auto& [sizeName, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + sizeName;
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        for (const auto& file : files)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeFlatten()
{
    // Collect all files from selected folders recursively
    std::vector<std::wstring> allFiles;
    
    for (const auto& path : m_selectedFiles)
    {
        DWORD attrs = GetFileAttributesW(path.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES)
            continue;
            
        if (attrs & FILE_ATTRIBUTE_DIRECTORY)
        {
            // Recursively find all files
            std::wstring searchPath = path + L"\\*";
            WIN32_FIND_DATAW fd;
            HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
            if (hFind != INVALID_HANDLE_VALUE)
            {
                do
                {
                    if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0)
                        continue;
                    
                    std::wstring fullPath = path + L"\\" + fd.cFileName;
                    if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
                        allFiles.push_back(fullPath);
                } while (FindNextFileW(hFind, &fd));
                FindClose(hFind);
            }
        }
        else
        {
            allFiles.push_back(path);
        }
    }

    if (allFiles.empty())
        return S_OK;

    CComPtr<IFileOperation> pMoveOp;
    HRESULT hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    CComPtr<IShellItem> pDestFolder;
    hr = SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder));
    if (FAILED(hr)) return hr;

    for (const auto& file : allFiles)
    {
        CComPtr<IShellItem> pItem;
        if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
            pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
    }

    return pMoveOp->PerformOperations();
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeNumbered()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    int folderNum = 1;
    
    for (const auto& file : m_selectedFiles)
    {
        wchar_t folderName[32];
        StringCchPrintfW(folderName, 32, L"Folder %d", folderNum++);
        
        std::wstring folderPath = GenerateUniqueFolderPath(m_parentFolder, folderName);
        std::wstring actualName = PathFindFileNameW(folderPath.c_str());

        CComPtr<IShellItem> pParentItem;
        SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
        pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, actualName.c_str(), nullptr, nullptr);
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (size_t i = 0; i < m_selectedFiles.size() && i < createdFolders.size(); i++)
    {
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(createdFolders[i].c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        CComPtr<IShellItem> pItem;
        if (SUCCEEDED(SHCreateItemFromParsingName(m_selectedFiles[i].c_str(), nullptr, IID_PPV_ARGS(&pItem))))
            pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT NewFolderFromFilesContextMenuHandler::OrganizeAlphabetical()
{
    if (m_selectedFiles.empty() || m_parentFolder.empty())
        return E_FAIL;

    std::map<std::wstring, std::vector<std::wstring>> groups;
    for (const auto& file : m_selectedFiles)
    {
        std::wstring filename = PathFindFileNameW(file.c_str());
        wchar_t letter[2] = { towupper(filename[0]), 0 };
        if (!iswalpha(letter[0]))
            wcscpy_s(letter, L"#");
        groups[letter].push_back(file);
    }

    CComPtr<IFileOperation> pFileOp;
    HRESULT hr = pFileOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    std::vector<std::wstring> createdFolders;
    for (auto& [letter, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + letter;
        if (!PathFileExistsW(folderPath.c_str()))
        {
            CComPtr<IShellItem> pParentItem;
            SHCreateItemFromParsingName(m_parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem));
            pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, letter.c_str(), nullptr, nullptr);
        }
        createdFolders.push_back(folderPath);
    }

    pFileOp->PerformOperations();
    Sleep(50);

    CComPtr<IFileOperation> pMoveOp;
    hr = pMoveOp.CoCreateInstance(CLSID_FileOperation);
    if (FAILED(hr)) return hr;

    pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

    for (auto& [letter, files] : groups)
    {
        std::wstring folderPath = m_parentFolder + L"\\" + letter;
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            continue;

        for (const auto& file : files)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
        }
    }

    hr = pMoveOp->PerformOperations();
    SelectMultipleFoldersInExplorer(createdFolders);
    return hr;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    if (HIWORD(pici->lpVerb) != 0)
        return E_INVALIDARG;

    UINT cmd = LOWORD(pici->lpVerb);

    switch (cmd)
    {
    case CMD_DEFAULT:
    case CMD_SUBMENU_DEFAULT:
        return ExecuteOrganize(OrganizeMode::Default);
    case CMD_BY_DAY:
        return ExecuteOrganize(OrganizeMode::ByDay);
    case CMD_BY_MONTH:
        return ExecuteOrganize(OrganizeMode::ByMonth);
    case CMD_BY_YEAR:
        return ExecuteOrganize(OrganizeMode::ByYear);
    case CMD_BY_MONTHYEAR:
        return ExecuteOrganize(OrganizeMode::ByMonthYear);
    case CMD_BY_FULLDATE:
        return ExecuteOrganize(OrganizeMode::ByFullDate);
    case CMD_BY_TYPE:
        return ExecuteOrganize(OrganizeMode::ByTypeVideo);  // Triggers full type sort
    case CMD_BY_EXTENSION:
        return ExecuteOrganize(OrganizeMode::ByExtension);
    case CMD_BY_SIZE:
        return ExecuteOrganize(OrganizeMode::BySize);
    case CMD_FLATTEN:
        return ExecuteOrganize(OrganizeMode::Flatten);
    case CMD_NUMBERED:
        return ExecuteOrganize(OrganizeMode::Numbered);
    case CMD_ALPHABETICAL:
        return ExecuteOrganize(OrganizeMode::Alphabetical);
    default:
        return E_INVALIDARG;
    }
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesContextMenuHandler::QueryContextMenu(
    HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (uFlags & CMF_DEFAULTONLY)
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    if (m_selectedFiles.empty())
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    m_idCmdFirst = idCmdFirst;

    // Create submenu
    HMENU hSubMenu = CreatePopupMenu();
    
    // Submenu items
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_SUBMENU_DEFAULT, L"New folder with selection");
    AppendMenuW(hSubMenu, MF_SEPARATOR, 0, nullptr);
    
    // Date submenu
    HMENU hDateMenu = CreatePopupMenu();
    AppendMenuW(hDateMenu, MF_STRING, idCmdFirst + CMD_BY_DAY, L"Day");
    AppendMenuW(hDateMenu, MF_STRING, idCmdFirst + CMD_BY_MONTH, L"Month");
    AppendMenuW(hDateMenu, MF_STRING, idCmdFirst + CMD_BY_YEAR, L"Year");
    AppendMenuW(hDateMenu, MF_STRING, idCmdFirst + CMD_BY_MONTHYEAR, L"Month-Year");
    AppendMenuW(hDateMenu, MF_STRING, idCmdFirst + CMD_BY_FULLDATE, L"Full Date");
    AppendMenuW(hSubMenu, MF_POPUP, (UINT_PTR)hDateMenu, L"By Date");
    
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_BY_TYPE, L"By Type");
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_BY_EXTENSION, L"By Extension");
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_BY_SIZE, L"By Size");
    AppendMenuW(hSubMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_FLATTEN, L"Flatten");
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_NUMBERED, L"Numbered");
    AppendMenuW(hSubMenu, MF_STRING, idCmdFirst + CMD_ALPHABETICAL, L"Alphabetical");

    // Main menu item with submenu
    MENUITEMINFOW mii = {};
    mii.cbSize = sizeof(MENUITEMINFOW);
    mii.fMask = MIIM_ID | MIIM_STRING | MIIM_STATE | MIIM_SUBMENU;
    mii.wID = idCmdFirst + CMD_DEFAULT;
    mii.fState = MFS_ENABLED;
    mii.hSubMenu = hSubMenu;
    mii.dwTypeData = const_cast<LPWSTR>(L"New folder with selection");

    if (!InsertMenuItemW(hmenu, indexMenu, TRUE, &mii))
        return HRESULT_FROM_WIN32(GetLastError());

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, CMD_COUNT);
}
