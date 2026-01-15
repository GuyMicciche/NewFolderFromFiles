#include <Windows.h>
#include <ShlObj.h>
#include <shobjidl.h>
#include <exdisp.h>
#include <atlbase.h>
#include <vector>
#include <string>
#include <strsafe.h>
#include <Shlwapi.h>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")

#define HOTKEY_ID 1
#define WM_TRAYICON (WM_USER + 1)

NOTIFYICONDATAW g_nid = {};
HWND g_hwnd = nullptr;
bool g_enabled = true;

// Settings stored in registry
const wchar_t* REG_KEY = L"Software\\NewFolderFromFiles";
const wchar_t* REG_ENABLED = L"HotkeyEnabled";

void SaveSettings()
{
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
    {
        DWORD val = g_enabled ? 1 : 0;
        RegSetValueExW(hKey, REG_ENABLED, 0, REG_DWORD, (BYTE*)&val, sizeof(val));
        RegCloseKey(hKey);
    }
}

void LoadSettings()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        DWORD val = 1, size = sizeof(val);
        RegQueryValueExW(hKey, REG_ENABLED, nullptr, nullptr, (BYTE*)&val, &size);
        g_enabled = (val != 0);
        RegCloseKey(hKey);
    }
}

void UpdateHotkey()
{
    UnregisterHotKey(g_hwnd, HOTKEY_ID);
    if (g_enabled)
    {
        // Ctrl+Alt+N
        RegisterHotKey(g_hwnd, HOTKEY_ID, MOD_CONTROL | MOD_ALT, 'N');
    }
}

std::wstring GenerateUniqueFolderName(const std::wstring& parentFolder, const std::wstring& baseName)
{
    std::wstring folderPath = parentFolder + L"\\" + baseName;
    
    if (!PathFileExistsW(folderPath.c_str()))
        return folderPath;

    for (int i = 2; i < 1000; i++)
    {
        wchar_t buffer[32];
        StringCchPrintfW(buffer, 32, L" (%d)", i);
        folderPath = parentFolder + L"\\" + baseName + buffer;
        
        if (!PathFileExistsW(folderPath.c_str()))
            return folderPath;
    }

    return parentFolder + L"\\New Folder";
}

std::wstring GetCommonPrefix(const std::vector<std::wstring>& files)
{
    if (files.empty())
        return L"New Folder";

    std::vector<std::wstring> names;
    for (const auto& path : files)
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
        {
            j++;
        }
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

void NewFolderFromSelection()
{
    CoInitialize(nullptr);

    CComPtr<IShellWindows> pShellWindows;
    if (FAILED(pShellWindows.CoCreateInstance(CLSID_ShellWindows)))
    {
        CoUninitialize();
        return;
    }

    // Find the foreground Explorer window
    HWND hwndForeground = GetForegroundWindow();
    
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

        HWND hwnd;
        if (FAILED(pWebBrowserApp->get_HWND((SHANDLE_PTR*)&hwnd)))
            continue;

        // Check if this is the foreground window
        if (hwnd != hwndForeground && GetAncestor(hwndForeground, GA_ROOTOWNER) != hwnd)
            continue;

        CComQIPtr<IServiceProvider> pServiceProvider(pWebBrowserApp);
        if (!pServiceProvider)
            continue;

        CComPtr<IShellBrowser> pShellBrowser;
        if (FAILED(pServiceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&pShellBrowser))))
            continue;

        CComPtr<IShellView> pShellView;
        if (FAILED(pShellBrowser->QueryActiveShellView(&pShellView)))
            continue;

        // Get selected items
        CComPtr<IDataObject> pDataObject;
        if (FAILED(pShellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&pDataObject))))
            continue;

        FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM stg = {};
        if (FAILED(pDataObject->GetData(&fmt, &stg)))
            continue;

        HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
        if (!hDrop)
        {
            ReleaseStgMedium(&stg);
            continue;
        }

        UINT fileCount = DragQueryFileW(hDrop, 0xFFFFFFFF, nullptr, 0);
        if (fileCount == 0)
        {
            GlobalUnlock(stg.hGlobal);
            ReleaseStgMedium(&stg);
            continue;
        }

        std::vector<std::wstring> selectedFiles;
        std::wstring parentFolder;

        wchar_t filePath[MAX_PATH];
        for (UINT j = 0; j < fileCount; j++)
        {
            if (DragQueryFileW(hDrop, j, filePath, MAX_PATH))
            {
                selectedFiles.push_back(filePath);
            }
        }

        GlobalUnlock(stg.hGlobal);
        ReleaseStgMedium(&stg);

        if (selectedFiles.empty())
            continue;

        // Get parent folder
        wchar_t parentPath[MAX_PATH];
        StringCchCopyW(parentPath, MAX_PATH, selectedFiles[0].c_str());
        PathRemoveFileSpecW(parentPath);
        parentFolder = parentPath;

        // Create folder and move files using IFileOperation
        CComPtr<IFileOperation> pFileOp;
        if (FAILED(pFileOp.CoCreateInstance(CLSID_FileOperation)))
            break;

        pFileOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

        std::wstring folderName = GetCommonPrefix(selectedFiles);
        std::wstring folderPath = GenerateUniqueFolderName(parentFolder, folderName);
        folderName = PathFindFileNameW(folderPath.c_str());

        CComPtr<IShellItem> pParentItem;
        if (FAILED(SHCreateItemFromParsingName(parentFolder.c_str(), nullptr, IID_PPV_ARGS(&pParentItem))))
            break;

        pFileOp->NewItem(pParentItem, FILE_ATTRIBUTE_DIRECTORY, folderName.c_str(), nullptr, nullptr);
        pFileOp->PerformOperations();

        // Move files
        CComPtr<IFileOperation> pMoveOp;
        if (FAILED(pMoveOp.CoCreateInstance(CLSID_FileOperation)))
            break;

        pMoveOp->SetOperationFlags(FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOFX_ADDUNDORECORD);

        Sleep(50);
        CComPtr<IShellItem> pDestFolder;
        if (FAILED(SHCreateItemFromParsingName(folderPath.c_str(), nullptr, IID_PPV_ARGS(&pDestFolder))))
            break;

        for (const auto& file : selectedFiles)
        {
            CComPtr<IShellItem> pItem;
            if (SUCCEEDED(SHCreateItemFromParsingName(file.c_str(), nullptr, IID_PPV_ARGS(&pItem))))
            {
                pMoveOp->MoveItem(pItem, pDestFolder, nullptr, nullptr);
            }
        }

        pMoveOp->PerformOperations();

        // Select new folder and enter rename mode
        PIDLIST_ABSOLUTE pidlFolder = ILCreateFromPathW(folderPath.c_str());
        if (pidlFolder)
        {
            PCUITEMID_CHILD pidlChild = ILFindLastID(pidlFolder);
            pShellView->SelectItem(pidlChild, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_EDIT);
            ILFree(pidlFolder);
        }

        break;
    }

    CoUninitialize();
}

void ShowContextMenu(HWND hwnd, POINT pt)
{
    HMENU hMenu = CreatePopupMenu();
    
    AppendMenuW(hMenu, g_enabled ? MF_CHECKED : MF_UNCHECKED, 1, L"Enable Hotkey (Ctrl+Alt+N)");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(hMenu, MF_STRING, 2, L"Exit");

    SetForegroundWindow(hwnd);
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, nullptr);
    DestroyMenu(hMenu);

    switch (cmd)
    {
    case 1:
        g_enabled = !g_enabled;
        SaveSettings();
        UpdateHotkey();
        break;
    case 2:
        Shell_NotifyIconW(NIM_DELETE, &g_nid);
        PostQuitMessage(0);
        break;
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_HOTKEY:
        if (wParam == HOTKEY_ID)
        {
            NewFolderFromSelection();
        }
        return 0;

    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP || lParam == WM_LBUTTONUP)
        {
            POINT pt;
            GetCursorPos(&pt);
            ShowContextMenu(hwnd, pt);
        }
        return 0;

    case WM_DESTROY:
        Shell_NotifyIconW(NIM_DELETE, &g_nid);
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int)
{
    // Check for existing instance
    HANDLE hMutex = CreateMutexW(nullptr, TRUE, L"NewFolderFromFilesHotkey");
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        CloseHandle(hMutex);
        return 0;
    }

    LoadSettings();

    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"NewFolderFromFilesHotkey";
    RegisterClassW(&wc);

    g_hwnd = CreateWindowExW(0, wc.lpszClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, hInstance, nullptr);

    // Create tray icon
    g_nid.cbSize = sizeof(g_nid);
    g_nid.hWnd = g_hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    // Use folder icon from shell
    SHFILEINFOW sfi = {};
    SHGetFileInfoW(L"folder", FILE_ATTRIBUTE_DIRECTORY, &sfi, sizeof(sfi), 
        SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
    g_nid.hIcon = sfi.hIcon ? sfi.hIcon : LoadIconW(nullptr, MAKEINTRESOURCEW(32512));
    StringCchCopyW(g_nid.szTip, ARRAYSIZE(g_nid.szTip), L"New Folder From Files (Ctrl+Alt+N)");
    Shell_NotifyIconW(NIM_ADD, &g_nid);

    UpdateHotkey();

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    UnregisterHotKey(g_hwnd, HOTKEY_ID);
    if (g_nid.hIcon) DestroyIcon(g_nid.hIcon);
    CloseHandle(hMutex);
    return 0;
}