#include "StdAfx.h"
#include "MyOPCDataCallback.h"
#include "resource.h"
#include "opcDlg.h"


MyOPCDataCallback::MyOPCDataCallback(void):m_cnRef(0)
{
}


MyOPCDataCallback::~MyOPCDataCallback(void)
{
}

HRESULT STDMETHODCALLTYPE MyOPCDataCallback::QueryInterface (REFIID riid, LPVOID *ppv)
{
    // Validate the pointer
    if (ppv == NULL)
        return E_POINTER;       // invalid pointer
    // Standard COM practice requires that we invalidate output arguments
    // if an error is encountered.  Let's assume an error for now and invalidate
    // ppInterface.  We will reset it to a valid interface pointer later if we
    // determine requested ID is valid:
    *ppv = NULL;
    if (riid == IID_IUnknown)
        *ppv = (IUnknown*) this;
    else if (riid == IID_IOPCDataCallback)
        *ppv = (IOPCDataCallback*) this;
    else
        return (E_NOINTERFACE); //unsupported interface

    // Success: increment the reference counter.
    AddRef ();
    return (S_OK);
}

ULONG STDMETHODCALLTYPE MyOPCDataCallback::AddRef (void)
{
    // Atomically increment the reference count and return
    // the value.
    return InterlockedIncrement((volatile LONG*)&m_cnRef);
}

ULONG STDMETHODCALLTYPE MyOPCDataCallback::Release (void)
{
    if (InterlockedDecrement((volatile LONG*)&m_cnRef) == 0)
    {
        delete this;
        return 0;
    }
    else
        return m_cnRef;
}
HRESULT STDMETHODCALLTYPE MyOPCDataCallback::OnDataChange(
    /* [in] */ DWORD dwTransid,
    /* [in] */ OPCHANDLE hGroup,
    /* [in] */ HRESULT hrMasterquality,
    /* [in] */ HRESULT hrMastererror,
    /* [in] */ DWORD dwCount,
    /* [size_is][in] */ OPCHANDLE *phClientItems,
    /* [size_is][in] */ VARIANT *pvValues,
    /* [size_is][in] */ WORD *pwQualities,
    /* [size_is][in] */ FILETIME *pftTimeStamps,
    /* [size_is][in] */ HRESULT *pErrors)
{
    for( DWORD index=0; index<dwCount; index++ )
    {
        if( SUCCEEDED(pErrors[index]) )
        {
            ((CopcDlg*)m_obj)->UpdateItem(phClientItems[index],&pvValues[index],pwQualities[index],pftTimeStamps[index]);
        }

    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MyOPCDataCallback::OnReadComplete(
    /* [in] */ DWORD dwTransid,
    /* [in] */ OPCHANDLE hGroup,
    /* [in] */ HRESULT hrMasterquality,
    /* [in] */ HRESULT hrMastererror,
    /* [in] */ DWORD dwCount,
    /* [size_is][in] */ OPCHANDLE *phClientItems,
    /* [size_is][in] */ VARIANT *pvValues,
    /* [size_is][in] */ WORD *pwQualities,
    /* [size_is][in] */ FILETIME *pftTimeStamps,
    /* [size_is][in] */ HRESULT *pErrors)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MyOPCDataCallback::OnWriteComplete(
    /* [in] */ DWORD dwTransid,
    /* [in] */ OPCHANDLE hGroup,
    /* [in] */ HRESULT hrMastererr,
    /* [in] */ DWORD dwCount,
    /* [size_is][in] */ OPCHANDLE *pClienthandles,
    /* [size_is][in] */ HRESULT *pErrors)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MyOPCDataCallback::OnCancelComplete(  /* [in] */ DWORD dwTransid,
        /* [in] */ OPCHANDLE hGroup)
{
    return S_OK;
}

void MyOPCDataCallback::Register(CObject* _obj)
{
    m_obj=_obj;
}