#pragma once
#include <Windows.h>

extern UINT g_cObjCount;

class NewFolderFromFilesClassFactory : public IClassFactory
{
protected:
    LONG m_ObjRefCount;
    ~NewFolderFromFilesClassFactory();

public:
    NewFolderFromFilesClassFactory();

    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject);

    // IClassFactory
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject);
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock);
};
