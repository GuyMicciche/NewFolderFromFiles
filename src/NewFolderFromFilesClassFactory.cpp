#include "NewFolderFromFilesClassFactory.h"
#include "NewFolderFromFilesContextMenuHandler.h"

NewFolderFromFilesClassFactory::~NewFolderFromFilesClassFactory()
{
    InterlockedDecrement(&g_cObjCount);
}

NewFolderFromFilesClassFactory::NewFolderFromFilesClassFactory() : m_ObjRefCount(1)
{
    InterlockedIncrement(&g_cObjCount);
}

ULONG STDMETHODCALLTYPE NewFolderFromFilesClassFactory::AddRef()
{
    return InterlockedIncrement(&m_ObjRefCount);
}

ULONG STDMETHODCALLTYPE NewFolderFromFilesClassFactory::Release()
{
    LONG ref = InterlockedDecrement(&m_ObjRefCount);
    if (ref == 0)
    {
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesClassFactory::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    *ppvObject = nullptr;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        *ppvObject = static_cast<IClassFactory*>(this);
    }
    else if (IsEqualIID(riid, IID_IClassFactory))
    {
        *ppvObject = static_cast<IClassFactory*>(this);
    }
    else
    {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesClassFactory::CreateInstance(
    IUnknown* pUnkOuter, REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_INVALIDARG;

    *ppvObject = nullptr;

    if (pUnkOuter != nullptr)
        return CLASS_E_NOAGGREGATION;

    NewFolderFromFilesContextMenuHandler* pHandler = new (std::nothrow) NewFolderFromFilesContextMenuHandler();
    if (!pHandler)
        return E_OUTOFMEMORY;

    HRESULT hr = pHandler->QueryInterface(riid, ppvObject);
    pHandler->Release();

    return hr;
}

HRESULT STDMETHODCALLTYPE NewFolderFromFilesClassFactory::LockServer(BOOL fLock)
{
    if (fLock)
        InterlockedIncrement(&g_cObjCount);
    else
        InterlockedDecrement(&g_cObjCount);

    return S_OK;
}
