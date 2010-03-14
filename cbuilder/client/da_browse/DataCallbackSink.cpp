// DataCallbackSink.cpp : Implementation of CDataCallbackSink
#include "DataCallbackSink.h"

/////////////////////////////////////////////////////////////////////////////
// CDataCallbackSink

HRESULT STDMETHODCALLTYPE CDataCallbackSink::OnDataChange(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   if (FEvDataChange != NULL)
	   FEvDataChange(dwTransid, hGroup, hrMasterquality, hrMastererror,
                    dwCount, phClientItems, pvValues, pwQualities,
                    pftTimeStamps, pErrors);
   return S_OK;
}

HRESULT STDMETHODCALLTYPE CDataCallbackSink::OnReadComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   if (FEvReadComplete != NULL)
	   FEvReadComplete(dwTransid, hGroup, hrMasterquality, hrMastererror,
                      dwCount, phClientItems, pvValues, pwQualities,
                      pftTimeStamps, pErrors);
   return S_OK;
}

HRESULT STDMETHODCALLTYPE CDataCallbackSink::OnWriteComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   if (FEvWriteComplete != NULL)
	   FEvWriteComplete(dwTransid, hGroup, hrMastererr, dwCount,
                       pClienthandles, pErrors);
   return S_OK;
}

HRESULT STDMETHODCALLTYPE CDataCallbackSink::OnCancelComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup)
{
   if (FEvCancelComplete != NULL)
	   FEvCancelComplete(dwTransid, hGroup);
   return S_OK;
}








