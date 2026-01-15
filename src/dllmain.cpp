#include <Windows.h>
#include <ShlObj.h>
#include <strsafe.h>
#include <string>
#include <new>
#include "NewFolderFromFilesGUID.h"
#include "NewFolderFromFilesClassFactory.h"

HINSTANCE g_hInstance = nullptr;
UINT g_cObjCount = 0;

static const wchar_t* EXTENSION_NAME = L"NewFolderFromFiles";

BOOL APIENTRY DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInstance = hInstance;
        DisableThreadLibraryCalls(hInstance);
        break;
    }
    return TRUE;
}

STDAPI DllCanUnloadNow()
{
    return g_cObjCount > 0 ? S_FALSE : S_OK;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppvOut)
{
    if (!ppvOut)
        return E_INVALIDARG;

    *ppvOut = nullptr;

    if (!IsEqualCLSID(rclsid, CLSID_NewFolderFromFilesShellExtension))
        return CLASS_E_CLASSNOTAVAILABLE;

    NewFolderFromFilesClassFactory* pFactory = new (std::nothrow) NewFolderFromFilesClassFactory();
    if (!pFactory)
        return E_OUTOFMEMORY;

    HRESULT hr = pFactory->QueryInterface(riid, ppvOut);
    pFactory->Release();

    return hr;
}

static std::wstring GetCLSIDString()
{
    wchar_t* str = nullptr;
    StringFromCLSID(CLSID_NewFolderFromFilesShellExtension, &str);
    std::wstring result(str);
    CoTaskMemFree(str);
    return result;
}

static std::wstring GetModulePath()
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(g_hInstance, path, MAX_PATH);
    return path;
}

static HRESULT CreateRegistryKey(HKEY hRoot, const wchar_t* subKey, HKEY* pKey)
{
    DWORD disp;
    LONG result = RegCreateKeyExW(hRoot, subKey, 0, nullptr, 
        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, pKey, &disp);
    return HRESULT_FROM_WIN32(result);
}

static HRESULT SetRegistryValue(HKEY hKey, const wchar_t* name, const wchar_t* value)
{
    DWORD size = (DWORD)((wcslen(value) + 1) * sizeof(wchar_t));
    LONG result = RegSetValueExW(hKey, name, 0, REG_SZ, (BYTE*)value, size);
    return HRESULT_FROM_WIN32(result);
}

STDAPI DllRegisterServer()
{
    HRESULT hr;
    HKEY hKey = nullptr;
    std::wstring clsid = GetCLSIDString();
    std::wstring modulePath = GetModulePath();

    // Register CLSID
    std::wstring clsidKey = L"Software\\Classes\\CLSID\\" + clsid;
    hr = CreateRegistryKey(HKEY_LOCAL_MACHINE, clsidKey.c_str(), &hKey);
    if (FAILED(hr)) return hr;
    SetRegistryValue(hKey, nullptr, EXTENSION_NAME);
    RegCloseKey(hKey);

    // InprocServer32
    std::wstring inprocKey = clsidKey + L"\\InprocServer32";
    hr = CreateRegistryKey(HKEY_LOCAL_MACHINE, inprocKey.c_str(), &hKey);
    if (FAILED(hr)) return hr;
    SetRegistryValue(hKey, nullptr, modulePath.c_str());
    SetRegistryValue(hKey, L"ThreadingModel", L"Apartment");
    RegCloseKey(hKey);

    // Register for all files (*)
    std::wstring handlerKey = L"Software\\Classes\\*\\shellex\\ContextMenuHandlers\\" + std::wstring(EXTENSION_NAME);
    hr = CreateRegistryKey(HKEY_LOCAL_MACHINE, handlerKey.c_str(), &hKey);
    if (FAILED(hr)) return hr;
    SetRegistryValue(hKey, nullptr, clsid.c_str());
    RegCloseKey(hKey);

    // Register for folders
    handlerKey = L"Software\\Classes\\Folder\\shellex\\ContextMenuHandlers\\" + std::wstring(EXTENSION_NAME);
    hr = CreateRegistryKey(HKEY_LOCAL_MACHINE, handlerKey.c_str(), &hKey);
    if (FAILED(hr)) return hr;
    SetRegistryValue(hKey, nullptr, clsid.c_str());
    RegCloseKey(hKey);

    // Add to approved extensions list
    std::wstring approvedKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, approvedKey.c_str(), 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        SetRegistryValue(hKey, clsid.c_str(), EXTENSION_NAME);
        RegCloseKey(hKey);
    }

    // Notify shell of change
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

    return S_OK;
}

STDAPI DllUnregisterServer()
{
    std::wstring clsid = GetCLSIDString();

    // Delete InprocServer32
    std::wstring key = L"Software\\Classes\\CLSID\\" + clsid + L"\\InprocServer32";
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, key.c_str());

    // Delete CLSID
    key = L"Software\\Classes\\CLSID\\" + clsid;
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, key.c_str());

    // Delete handler for files
    key = L"Software\\Classes\\*\\shellex\\ContextMenuHandlers\\" + std::wstring(EXTENSION_NAME);
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, key.c_str());

    // Delete handler for folders
    key = L"Software\\Classes\\Folder\\shellex\\ContextMenuHandlers\\" + std::wstring(EXTENSION_NAME);
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, key.c_str());

    // Remove from approved list
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, 
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved",
        0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
    {
        RegDeleteValueW(hKey, clsid.c_str());
        RegCloseKey(hKey);
    }

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

    return S_OK;
}