// DataCallbackSink.h : Declaration of the CDataCallbackSink

#ifndef __DATACALLBACKSINK_H_
#define __DATACALLBACKSINK_H_

#include <Opcda.h>
#include <vcl.h>

#include "CustomSinks.h"

// It could be anything, even all numbers in 0:
// {F8FE7C40-F9C1-11d3-96DB-00902787286C}
DEFINE_GUID(CLSID_DataCallbackSink,
   0xf8fe7c40, 0xf9c1, 0x11d3, 0x96, 0xdb, 0x0, 0x90, 0x27, 0x87, 0x28, 0x6c);

// Event types:
typedef void __fastcall (__closure *TOnDataChangeEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

typedef void __fastcall (__closure *TOnReadCompleteEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

typedef void __fastcall (__closure *TOnWriteCompleteEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

typedef void __fastcall (__closure *TOnCancelCompleteEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup);

/////////////////////////////////////////////////////////////////////////////
// CDataCallbackSink
class ATL_NO_VTABLE CDataCallbackSink :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CDataCallbackSink, &CLSID_DataCallbackSink>,
   public IOPCDataCallback
{
// properties to create handlers in the form:
protected:
	TOnDataChangeEvent FEvDataChange;
	TOnReadCompleteEvent FEvReadComplete;
	TOnWriteCompleteEvent FEvWriteComplete;
	TOnCancelCompleteEvent FEvCancelComplete;

public:
	__property TOnDataChangeEvent EvDataChange = {read=FEvDataChange,
                                                 write=FEvDataChange};
	__property TOnReadCompleteEvent EvReadComplete = {read=FEvReadComplete,
                                                     write=FEvReadComplete};
	__property TOnWriteCompleteEvent EvWriteComplete = {read=FEvWriteComplete,
                                                       write=FEvWriteComplete};
	__property TOnCancelCompleteEvent EvCancelComplete = {read=FEvCancelComplete,
                                                         write=FEvCancelComplete};
public:
   CDataCallbackSink()
	{
	}

DECLARE_NOT_AGGREGATABLE(CDataCallbackSink)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CDataCallbackSink)
	COM_INTERFACE_ENTRY(IOPCDataCallback)
END_COM_MAP()

// IOPCDataCallback: the sink implements this interface to receive the
// events from the OPC group.
public:
   virtual HRESULT STDMETHODCALLTYPE OnDataChange(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

   virtual HRESULT STDMETHODCALLTYPE OnReadComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

   virtual HRESULT STDMETHODCALLTYPE OnWriteComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

   virtual HRESULT STDMETHODCALLTYPE OnCancelComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup);
};

// Instantiable sink declaration
typedef TCustomSink<CDataCallbackSink, IID_IOPCDataCallback>
   CCreatableDataCallbackSink;

#endif //__DATACALLBACKSINK_H_
