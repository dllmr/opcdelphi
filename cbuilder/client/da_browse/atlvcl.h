/////////////////////////////////////////////////////////////////////////////
// ATLVCL.H - Provides the connective tissue between
//            the ATL framework and VCL components.
//
// $Revision:   1.75.1.0.1.0  $
// $Date:   28 Jun 2001 09:12:00  $
//
// Copyright (c) 1998,2000 Borland International
/////////////////////////////////////////////////////////////////////////////

#ifndef __ATLVCL_H_
#define __ATLVCL_H_

#pragma option push -VF

// These are required due to RTL differences between VC++ and BCB
//
#define _ATL_NO_FORCE_LIBS
#define _ATL_NO_DEBUG_CRT

// Defines _ASSERTE et al.
//
#include <utilcls.h>

// Delta for remapping messages for VCL compatibility
//
#if !defined(OCM_BASE)
  #define OCM_BASE (int)(8192)
#endif

#if !defined(__ATLBASE_H)
  #include <atl\atlbase.h>
#endif

#include <atl\atlmod.h>

#include <system.hpp>
#include <axctrls.hpp>
#include <objbase.h>
#include <sysutils.hpp>
#include <cguid.h>
#include <dir.h>
#include <safearry.h>


// Externs
// Pointer to save Initialization Procedure when using VCL
// (CBuilder3 backward compatibility)
//
extern void* SaveInitProc;

// Prototype of routines implemented in ATLVCL.CPP
//
bool __fastcall AutomationTerminateProc();
void __fastcall SaveVCLComponentToStream(TComponent *instance, LPSTREAM pStream);
void __fastcall LoadVCLComponentFromStream(TComponent *instance, LPSTREAM pStream);
TWinControl*    CreateReflectorWindow(HWND parent, Controls::TControl* Control);

#if !defined(__ATLCOM_H__)
  #include <atl\atlcom.h>
#endif
#include   <shellapi.h>
#if !defined(__ATLCTL_H__)
  #include <atl\atlctl.h>
#endif

#include <vcl.h>
#include <databkr.hpp>
#include <atl\axform.h>

// Forward type Declaration
//
template <class T> class DELPHICLASS TWinControlAccess;

// Forward function declarations
// (CBuilder3 backward compatibility)
//
void __fastcall InitAtlServer(void);

// Helper routines used by IMPL file of VCL ActiveX Controls
//

///////////////////////////////////////////////////////////////////////////////
// FONT property handlers
///////////////////////////////////////////////////////////////////////////////
inline void
SetVclCtlProp(Graphics::TFont* font, IFontDisp* fontdisp)
{
  Axctrls::SetOleFont(font, fontdisp);
}

inline void
SetVclCtlProp(Graphics::TFont* font, IFontDispPtr fontdisp)
{
  SetVclCtlProp(font, static_cast<IFontDisp*>(fontdisp));
}

// PUTREF version of FONT setter
inline void
SetVclCtlProp(Graphics::TFont* font, IFontDisp** fontdisp)
{
  SetVclCtlProp(font, *fontdisp);
}

inline void
SetVclCtlProp(Graphics::TFont* font, IFontDispPtr* fontdisp)
{
  SetVclCtlProp(font, *fontdisp);
}

inline void
GetVclCtlProp(Graphics::TFont* font, IFontDisp** fontdisp)
{
  _di_IFontDisp _di_font;
  Axctrls::GetOleFont(font, _di_font);
  *fontdisp = _di_font;
  if (*fontdisp)
    (*fontdisp)->AddRef();
}

inline void
GetVclCtlProp(Graphics::TFont* font, IFontDispPtr* fontdisp)
{
  GetVclCtlProp(font, reinterpret_cast<IFontDisp**>(fontdisp));
}

///////////////////////////////////////////////////////////////////////////////
// Picture property handlers
///////////////////////////////////////////////////////////////////////////////
inline void
GetVclCtlProp(TPicture* pic, IPictureDisp** ppPictDisp)
{
  _di_IPictureDisp _di_picture;
  Axctrls::GetOlePicture(pic, _di_picture);
  *ppPictDisp = _di_picture;
  if (*ppPictDisp)
    (*ppPictDisp)->AddRef();
}

inline void
GetVclCtlProp(TPicture* pic, IPictureDispPtr* ppPictDisp)
{
  GetVclCtlProp(pic, reinterpret_cast<IPictureDisp**>(ppPictDisp));
}

inline void
SetVclCtlProp(TPicture* pic, IPictureDisp* pPictDisp)
{
  Axctrls::SetOlePicture(pic, pPictDisp);
}

inline void
SetVclCtlProp(TPicture* pic, IPictureDisp** ppPictDisp)
{
  SetVclCtlProp(pic, *ppPictDisp);
}

inline void
SetVclCtlProp(TPicture* pic, IPictureDispPtr pPictDisp)
{
  SetVclCtlProp(pic, static_cast<IPictureDisp*>(pPictDisp));
}

inline void
SetVclCtlProp(TPicture* pic, IPictureDispPtr* ppPictDisp)
{
  SetVclCtlProp(pic, *ppPictDisp);;
}


///////////////////////////////////////////////////////////////////////////////
// Strings property handlers
///////////////////////////////////////////////////////////////////////////////
inline void
GetVclCtlProp(TStrings* str, IStrings **ppIString)
{
  _di_IStrings _di_strings;
  Axctrls::GetOleStrings(str, _di_strings);
  *ppIString = _di_strings;
  if (*ppIString)
    (*ppIString)->AddRef();
}
inline void
GetVclCtlProp(TStrings* str, IStringsPtr *ppIString)
{
  GetVclCtlProp(str, reinterpret_cast<IStrings**>(ppIString));
}
inline void
SetVclCtlProp(TStrings* str, IStrings* pIString)
{
  Axctrls::SetOleStrings(str, pIString);
}
inline void
SetVclCtlProp(TStrings* str, IStrings** ppIString)
{
  SetVclCtlProp(str, *ppIString);
}

inline void
SetVclCtlProp(TStrings* str, IStringsPtr pIString)
{
  SetVclCtlProp(str, static_cast<IStrings*>(pIString));
}
inline void
SetVclCtlProp(TStrings* str, IStringsPtr* ppIString)
{
  SetVclCtlProp(str, *ppIString);
}

/* IAppServer support */
/* IAppServerImpl is essentially Delphi's TRemoteDataModule; it implements
   the IAppServer interface, and is used to publish datasets residing
   in a TDataModule. Method calls on the IAppServer interface are 
   reflected to renamed methods of the style IAppServer_MethodName
   in order to avoid ambiguity problems in multiply derived classes.  */

template <class DM, class T , class Intf, const IID* piid, const GUID* plibid>
class ATL_NO_VTABLE IAppServerImpl: public IDispatchImpl<Intf, piid, plibid>
{
private:
  TCOMCriticalSection m_CS;
  
public:
  DM* m_DataModule;            // The Core. 
  // Note: This data module _must_ descend from TCRemoteDataModule.


  IAppServerImpl()
  {
   m_DataModule = new DM(NULL);
  }

  ~IAppServerImpl() 
  {
    m_DataModule->Free();
  }

  TCustomProvider* GetProvider(const AnsiString ProviderName)
  {
    return m_DataModule->GetProvider(ProviderName);
    // assumes that DM derives from TCRDM.
  }

  void RegisterProvider(TCustomProvider* Provider)
  {
    m_DataModule->RegisterProvider(Provider);
  }

  void UnRegisterProvider(TCustomProvider* Provider)
  {
    m_DataModule->UnRegisterProvider(Provider);
  }

  // Reflector methods (forwarding shims)
  // These methods lock and forward the call to methods on the
  // exported providers. Exception handling is expected to be
  // conducted by the function calling the reflector method.

  HRESULT STDMETHODCALLTYPE IAppServer_AS_GetProviderNames(OleVariant &Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = m_DataModule->CRDMGetProviderNames();
    return S_OK;
    // assumes that DM derives from TCRemoteDataModule.
  }

  HRESULT STDMETHODCALLTYPE IAppServer_AS_ApplyUpdates(WideString ProviderName,
     System::OleVariant Delta, int MaxErrors, int& ErrorCount, OleVariant& OwnerData,
     OleVariant& Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = GetProvider(ProviderName)->ApplyUpdates(Delta, MaxErrors, ErrorCount, OwnerData);
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE IAppServer_AS_GetRecords(const WideString ProviderName,
     int Count, int& RecsOut, int Options, const WideString CommandText, OleVariant& Params,
     OleVariant& OwnerData, OleVariant& Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = GetProvider(ProviderName)->GetRecords(Count, RecsOut, Options, CommandText,
                                                   Params, OwnerData);
    return S_OK;
  }
    
  HRESULT STDMETHODCALLTYPE IAppServer_AS_DataRequest(WideString ProviderName,
     OleVariant Data, System::OleVariant& Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = GetProvider(ProviderName)->DataRequest(Data);
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE IAppServer_AS_RowRequest(const WideString ProviderName,
     OleVariant Row, int RequestType, OleVariant& OwnerData, OleVariant& Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = GetProvider(ProviderName)->RowRequest(Row, RequestType, OwnerData);
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE IAppServer_AS_GetParams(const WideString ProviderName,
     OleVariant& OwnerData, OleVariant& Result)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    Result = GetProvider(ProviderName)->GetParams(OwnerData);
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE IAppServer_AS_Execute(const WideString ProviderName,
     const WideString CommandText, OleVariant& Params, OleVariant& OwnerData)
  {
    TCOMCriticalSection::Lock::Lock(m_CS);
    GetProvider(ProviderName)->Execute(CommandText, Params, OwnerData);
    return S_OK;
  }

  // IAppServer implementation
  // provides exception handling and forwards the call to a reflector
  // method.

  HRESULT STDMETHODCALLTYPE AS_GetProviderNames(OleVariant& Result)
  {
    ATLTRACE(_T("IAppServer::AS_GetProviderNames\n"));
    HRESULT hres;
    try
    {
      hres = static_cast<T*>(this)->IAppServer_AS_GetProviderNames(Result);
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }

  HRESULT STDMETHODCALLTYPE AS_ApplyUpdates (const WideString ProviderName,
     const OleVariant Delta, int MaxErrors, int& ErrorCount, OleVariant& OwnerData,
     OleVariant& Result)
  {
    ATLTRACE(_T("IAppServer::AS_ApplyUpdates\n"));
    HRESULT hres;
    try 
    {
       hres = static_cast<T*>(this)->IAppServer_AS_ApplyUpdates(ProviderName, Delta,
                                                                MaxErrors, ErrorCount,
                                                                OwnerData, Result);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }

  HRESULT STDMETHODCALLTYPE AS_GetRecords(const WideString ProviderName, int Count,
     int &RecsOut, int Options, const WideString CommandText, OleVariant &Params,
     OleVariant &OwnerData, OleVariant &Result)
  {
    ATLTRACE(_T("IAppServer::AS_GetRecords\n"));
    HRESULT hres;
    try 
    {
       hres = static_cast<T*>(this)->IAppServer_AS_GetRecords(ProviderName, Count,
                                                              RecsOut, Options,
                                                              CommandText, Params,
                                                              OwnerData, Result);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }
  
  HRESULT STDMETHODCALLTYPE AS_DataRequest(const WideString ProviderName, OleVariant Data,
     OleVariant&  Result)
  {
    ATLTRACE(_T("IAppServer::AS_DataRequest\n"));
    HRESULT hres;
    try
    {
       hres = static_cast<T*>(this)->IAppServer_AS_DataRequest(ProviderName, Data, Result);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }

    
  HRESULT STDMETHODCALLTYPE AS_RowRequest(const WideString ProviderName,
     OleVariant Row, int RequestType, OleVariant& OwnerData, OleVariant& Result)
  {
    ATLTRACE(_T("IAppServer::AS_RowRequest\n"));
    HRESULT hres;
    try 
    {
       hres = static_cast<T*>(this)->IAppServer_AS_RowRequest(ProviderName, Row,
                                                              RequestType, OwnerData,
                                                              Result);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }

  HRESULT STDMETHODCALLTYPE AS_GetParams(const WideString ProviderName, OleVariant& OwnerData,
     OleVariant& Result)
  {
    ATLTRACE(_T("IAppServer::AS_GetParams\n"));
    HRESULT hres;
    try 
    {
       hres = static_cast<T*>(this)->IAppServer_AS_GetParams(ProviderName, OwnerData, Result);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }


  HRESULT STDMETHODCALLTYPE AS_Execute(const WideString ProviderName, const WideString CommandText,
     OleVariant& Params, OleVariant& OwnerData)
  {
    ATLTRACE(_T("IAppServer::AS_GetParams\n"));
    HRESULT hres;
    try 
    {
       hres = static_cast<T*>(this)->IAppServer_AS_Execute(ProviderName, CommandText,
                                                           Params, OwnerData);
    }
    catch (Exception& e)
    {
       return(static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }
  
   
/*  Get Provider(const ProviderName: string); TCustomProvider; virtual;
}

/* IDataBroker support */
// IDataBrokerImpl class implements IDataBroker interface, and is
// used to publish datasets residing in a TDataModule.

/* IDataBroker was the interface exposed by TRemoteDataModule in MIDAS 2.0.
   It is maintained here for backwards compatability, but has been deprecated. */
};

template <class DM, class T, class Intf, const IID* piid, const GUID* plibid>
class ATL_NO_VTABLE IDataBrokerImpl: public IDispatchImpl<Intf, piid, plibid>
{
public:
  DM* m_DataModule;             // The Data Module

  IDataBrokerImpl()
  {
    m_DataModule = new DM(NULL);
  }
 ~IDataBrokerImpl()
  {
    m_DataModule->Free();
  }

  HRESULT IDataBroker_GetProviderNames(System::OleVariant &Result)
  {
    unsigned int TICount;
    _di_ITypeInfo TI;
    TPtr<Classes::TStringList> ProvProps(new TStringList);

    HRESULT hres = GetTypeInfoCount(&TICount);
    if (hres != S_OK)
      return hres;

    if (TICount)
    {
      hres = GetTypeInfo(0, 0, &TI);
      if (hres != S_OK)
        return hres;
      Databkr::EnumIProviderProps(TI, ProvProps);
      ::VariantClear(&(VARIANT&)Result);
      Result = Databkr::VarArrayFromStrings(ProvProps);
    }
    return S_OK;
  }

  // IDataBroker
  //
  HRESULT STDMETHODCALLTYPE GetProviderNames(System::OleVariant &Result)
  {
    ATLTRACE(_T("IDataBroker::GetProviderNames\n"));
    HRESULT hres;
    try
    {
      hres = static_cast<T*>(this)->IDataBroker_GetProviderNames(Result);
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return hres;
  }
};

// ISimpleFrameSite support
//
template <class T>
class ATL_NO_VTABLE ISimpleFrameSiteImpl
{
public:
  // IUnknown
  STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObject) = 0;
  _ATL_DEBUG_ADDREF_RELEASE_IMPL(ISimpleFrameSite)

  // ISimpleFrameSite
  STDMETHOD(PreMessageFilter)(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp,
                              LRESULT* plResult, DWORD* pdwCookie)
  {
    ATLTRACE(_T("ISimpleFrameImpl::PreMessageFilter\n"));
    T* pT = static_cast<T*>(this);
    return pT->ISimpleFrameSite_PreMessageFilter(hWnd, msg, wp, lp, plResult, pdwCookie);
  }

  STDMETHOD(PostMessageFilter)(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp,
                               LRESULT* plResult, DWORD dwCookie)
  {
    ATLTRACE(_T("ISimpleFrameImpl::PostMessageFilter\n"));
    T* pT = static_cast<T*>(this);
    return pT->ISimpleFrameSite_PostMessageFilter(hWnd, msg, wp, lp, plResult, dwCookie);
  }
};

/* ActiveX control support */

// TVclComControl is an ATL ActiveX control which encapsulates a VCL control
// T  is the user's class which implements the OLE Control.
// TVCL is the VCL type which is exposed as an ActiveX Control.
//
template <class T, class TVCL>
class ATL_NO_VTABLE TVclComControl: public CComControlBase, public CMessageMap
{
public:
  TVclComControl(void): m_VclCtl(0), m_VclWinCtl(0), m_hWnd(0),
                        m_CtlWndProc(0), CComControlBase(m_hWnd)
  {
    m_bWindowOnly = TRUE;
  }

 ~TVclComControl(void)
  {
    if (m_VclCtl)
    {
      m_VclCtl->WindowProc = m_CtlWndProc;
      delete m_VclCtl;
    }
    if ((m_VclWinCtl) && ((TWinControl*)m_VclWinCtl != (TWinControl*)m_VclCtl))
      delete m_VclWinCtl;
  }

  // IOleObject support
  //
  HRESULT IOleObject_SetClientSite(IOleClientSite *pClientSite)
  {
    HRESULT hres = CComControlBase::IOleObject_SetClientSite(pClientSite);
    if (hres == S_OK)
    {
      if (m_spClientSite != NULL)
      {
        DWORD MiscFlags;
        hres = static_cast<T*>(this)->GetMiscStatus(DVASPECT_CONTENT, &MiscFlags);

        if (!SUCCEEDED(hres))
          return hres;

        // The following double checks that the registry entry indeed corresponds to
        // the value specified by our Control. (NOTE: IOleObject's implementation
        // looks up the value in the registry).
        //
        _ASSERTE(MiscFlags == static_cast<T*>(this)->_GetObjectMiscStatus());

        if (MiscFlags & OLEMISC_SIMPLEFRAME)
        {
          m_spClientSite->QueryInterface(IID_ISimpleFrameSite,
                                        (void**)&m_spSimpleFrameSite);
        }

        // Initialize the helper class used to get ambient properties from the container
        m_AmbientDriver.Bind(m_spClientSite);

        // Get ambient properties from the container and update the control
        static_cast<T*>(this)->OnAmbientPropertyChange(0);

        // Enable Ctl3D style
        m_VclWinCtl->Perform(CM_PARENTCTL3DCHANGED, 1, 1);
      }
      else
      {
        m_spSimpleFrameSite = NULL;
        m_AmbientDriver = NULL;
      }
    }
    else
    {
      m_spClientSite = NULL;
      m_AmbientDriver = NULL;
    }
    return hres;
  }

  // IPersistStreamInit support
  //

  // Note: IPersistStreamInit_Load and IPersistStreamInit_Save were in
  // CComControlBase in ATL 2.1 but moved to IPersistStreamInitImpl in
  // ATL 3.0. As a result, the methods need to be renamed in C++Builder 5.

  HRESULT IPersistStreamInit_SaveVCL(LPSTREAM pStm, BOOL /* fClearDirty */, ATL_PROPMAP_ENTRY* pMap)
  {
    try
    {
      ::SaveVCLComponentToStream((TVCL*)m_VclCtl, pStm);
      //!?? Clear fClearDirty flag
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }

  HRESULT IPersistStreamInit_LoadVCL(LPSTREAM pStm, ATL_PROPMAP_ENTRY* pMap)
  {
    try
    {
      ::LoadVCLComponentFromStream((TVCL*)m_VclCtl, pStm);
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }

  // ISimpleFrameSite support
  //
  HRESULT ISimpleFrameSite_PreMessageFilter(HWND hWnd, UINT msg, WPARAM wp,
                                            LPARAM lp, LRESULT* plResult, DWORD* pdwCookie)
  {
    if (m_spSimpleFrameSite)
      return m_spSimpleFrameSite->PreMessageFilter(hWnd, msg, wp, lp,
                                                   plResult, pdwCookie);
    else
      return S_OK;
  }

  HRESULT ISimpleFrameSite_PostMessageFilter(HWND hWnd, UINT msg, WPARAM wp,
                                             LPARAM lp, LRESULT* plResult, DWORD dwCookie)
  {
    if (m_spSimpleFrameSite)
      return m_spSimpleFrameSite->PostMessageFilter(hWnd, msg, wp, lp,
                                                    plResult, dwCookie);
    else
      return S_OK;
  }


  // IPropertyNotifySink Support methods
  //
  // NOTE: If your control class derives from IPropertyNotifySink, this method calls
  // CFirePropNotifyEvent::FireOnRequestEdit to notify all connected IPropertyNotifySink interfaces
  // that the specified control property is about to change. If your control class does not derive
  // from IPropertyNotifySink, this method returns S_OK.
  //
  HRESULT FireOnRequestEdit(DISPID dispID)
  {
    T* pT = static_cast<T*>(this);
    return T::__ATL_PROP_NOTIFY_EVENT_CLASS::FireOnRequestEdit(pT->GetUnknown(), dispID);
  }

  // If your control class derives from IPropertyNotifySink, this method calls
  // CFirePropNotifyEvent::FireOnChanged to notify all connected IPropertyNotifySink interfaces
  // that the specified control property has changed. If your control class does not derive from
  // IPropertyNotifySink, this method returns S_OK
  //
  HRESULT FireOnChanged(DISPID dispID)
  {
    T* pT = static_cast<T*>(this);
    return T::__ATL_PROP_NOTIFY_EVENT_CLASS::FireOnChanged(pT->GetUnknown(), dispID);
  }

  // Retrieves pointer to requested interface
  // NOTE: Limited to interfaces listed in COM map table
  //
  virtual HRESULT ControlQueryInterface(const IID& iid, void** ppv)
  {
    T* pT = static_cast<T*>(this);
    return pT->_InternalQueryInterface(iid, ppv);
  }

  // Creates underlying window of the control.
  // NOTE: You may override this method to do something other than create an window (for
  //   example, to create two windows, one of which becomes a toolbar for your control).
  //
  HWND CreateControlWindow(HWND hWndParent, RECT& rcPos)
  {
    T* pT = static_cast<T*>(this);
    return pT->Create(hWndParent, rcPos);
  }

  // Returns the handle of the control, if successful. Otherwise, returns NULL.
  //
  HWND Create(HWND hWndParent, RECT& rcPos);

  // Sets the Window's show state (See ::ShowWindow of WIN32 SDK for details on 'nCmdShow')
  //
  BOOL ShowWindow(int nCmdShow)
  {
    _ASSERTE(m_hWnd);
    return ::ShowWindow(m_hWnd, nCmdShow);
  }

  // Overridable Message Handler method of Control.
  // NOTE: You may intercept messages in your control via an ATL Message Map.
  //   The default implementation of this method is to dispatch messages
  //   via the Message Map. If the message was not handled via a message
  //   map, the message is dispatched to the underlying VCL handler.
  //
  virtual void ControlWndProc(Messages::TMessage& Message);

  // Procedure handling messages of this control
  //
  virtual void __fastcall WndProc(Messages::TMessage& Message);

  // Method which destroys and re-create the control's window.
  // NOTE: This may be necessary if you need to change the window's style bits.
  //
  void RecreateWnd();

  // Data members
  //
  HWND m_hWnd;                    // Underlying Window Handle of our Control

protected:
  virtual void InitializeControl();
  void         Initialize();

  TWinControlAccess<TWinControl>*   m_VclWinCtl;
  TWinControlAccess<TVCL>*          m_VclCtl;
  TAutoDriver<IDispatch>            m_AmbientDriver;

private:
  TWndMethod                        m_CtlWndProc;
  CComPtr<ISimpleFrameSite>         m_spSimpleFrameSite;
};


// TVclComControl::Initialize
//
template <class T, class TVCL> void
TVclComControl<T, TVCL>::Initialize()
{
  // Retrieve handle to reflector Window defined in AXCTRLS unit
  //
  HWND hwndParkingWindow = Axctrls::ParkingWindow();

  // Create VCL Object Wrapper for our Control
  //
  m_VclCtl = (TWinControlAccess<TVCL>*)(TVCL::CreateParentedControl(__classid(TVCL), hwndParkingWindow));

  // Test whether our control requires message reflection and create a reflector window if yes
  //
  if (m_VclCtl->ControlStyle.Contains(csReflector))
    (TWinControl*)m_VclWinCtl = ::CreateReflectorWindow(hwndParkingWindow, m_VclCtl);
  else
    (TWinControl*)m_VclWinCtl = (TWinControl*)m_VclCtl;

  // Update the Window Procedure variables
  //
  m_CtlWndProc = m_VclCtl->WindowProc;
  m_VclCtl->WindowProc = WndProc;

  // Invoke virtual allowing Control to perform additional initialization
  //
  InitializeControl();
  ATLTRACE(_T("VCL control created and initialized\n"));
}

// InitializeControl is typically overriden in the derived class.
// It's used to initialize Closures to catch VCL events and turn
// them into OLE events.
//
template <class T, class TVCL> void
TVclComControl<T, TVCL>::InitializeControl(void)
{}

// TVclComControl::ReCreateWnd
//
template <class T, class TVCL> void
TVclComControl<T, TVCL>::RecreateWnd()
{
  if (m_VclWinCtl->HandleAllocated())
  {
    RECT CtlBounds = m_VclWinCtl->BoundsRect;
    HWND PrevWnd = ::GetWindow(m_VclWinCtl->Handle, GW_HWNDPREV);

    // set ATL window handle to NULL to prevent InPlaceDeactiveate()
    // from destroying the VCL control window too soon
    //
    m_hWndCD = NULL;
    IOleInPlaceObject_InPlaceDeactivate();
    m_VclWinCtl->DoDestroyHandle();
    m_VclWinCtl->UpdateControlState();
    if (InPlaceActivate(m_bUIActive ? OLEIVERB_INPLACEACTIVATE : OLEIVERB_HIDE,
                        &CtlBounds) == S_OK)
      ::SetWindowPos(m_VclWinCtl->Handle, PrevWnd, 0, 0, 0, 0,
                     SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
  }
}

// TVclComControl::Create
//
template <class T, class TVCL> HWND
TVclComControl<T, TVCL>::Create(HWND hWndParent, RECT& rcPos)
{
  _ASSERTE(m_VclCtl);

  // Set the parent handle here because with a parking window of NULL.  The
  // window is not created until the parent is set.
  //
  m_VclWinCtl->ParentWindow = hWndParent;
  m_VclWinCtl->BoundsRect = rcPos;

  // For ActiveForms the window handle will be NULL at this point; so, force create.
  //
  m_VclWinCtl->HandleNeeded();

  // For ActiveForms again, the window will not be visible. So, show it.
  //
  m_VclWinCtl->Visible = true;

  // Update and return handle
  //
  m_hWnd = m_VclWinCtl->GetWindowHandle();
  return m_hWnd;
}

// TVclComControl::ControlWndProc
//
template <class T, class TVCL> void
TVclComControl<T, TVCL>::ControlWndProc(Messages::TMessage& Message)
{
  if ((Message.Msg >= OCM_BASE) && (Message.Msg < OCM_BASE + WM_USER))
    Message.Msg = Message.Msg + (CN_BASE - OCM_BASE);

  // Allow ATL message maps to intercept messages before they are dispatch to VCL handlers
  //
  if (!ProcessWindowMessage(m_VclCtl->GetWindowHandle(), Message.Msg,
                            Message.WParam, Message.LParam,
                            *((LRESULT*)(&Message.Result)), 0))
    m_CtlWndProc(Message);

  if ((Message.Msg >= CN_BASE) && (Message.Msg < CN_BASE + WM_USER))
    Message.Msg = Message.Msg - (CN_BASE - OCM_BASE);
}

// TvclComControl::WndProc
// Procedure handling messages of this control
//
template <class T, class TVCL> void __fastcall
TVclComControl<T, TVCL>::WndProc(Messages::TMessage& Message)
{
  DWORD dwCookie;
  HWND Handle = m_VclCtl->GetWindowHandle();
  bool FilterMessage = ((Message.Msg < CM_BASE) || (Message.Msg >= 0xC000)) &&
                       m_spSimpleFrameSite && m_bInPlaceActive;
  if (FilterMessage)
  {
    if (m_spSimpleFrameSite->PreMessageFilter(Handle,
                                              Message.Msg, Message.WParam,
                                              Message.LParam, (LRESULT*)&Message.Result,
                                              &dwCookie) == S_FALSE)
      return;
  }

  T* pT = static_cast<T*>(this);
  CComPtr<IOleControlSite> spSite;
  switch (Message.Msg)
  {
  case WM_SETFOCUS:
  case WM_KILLFOCUS:
    ControlWndProc(Message);
    m_spClientSite->QueryInterface(IID_IOleControlSite, (void**)&spSite);
    if (spSite)
      spSite->OnFocus(Message.Msg == WM_SETFOCUS);
    break;

  case CM_VISIBLECHANGED:
    if ((TWinControl*)m_VclCtl != (TWinControl*)m_VclWinCtl)
      m_VclWinCtl->Visible = m_VclCtl->Visible;
    if (!(m_VclWinCtl->Visible))
      IOleInPlaceObject_UIDeactivate();
    ControlWndProc(Message);
    break;

  case CM_RECREATEWND:
    if ((m_bInPlaceActive) && ((TWinControl*)m_VclCtl == (TWinControl*)m_VclWinCtl))
      RecreateWnd();
    else
    {
      ControlWndProc(Message);
      pT->SendOnViewChange(DVASPECT_CONTENT);
    }
    break;

  case CM_INVALIDATE:
  case WM_SETTEXT:
    ControlWndProc(Message);
    if (!m_bInPlaceActive)
      pT->SendOnViewChange(DVASPECT_CONTENT);
    break;

  case WM_NCHITTEST:
    ControlWndProc(Message);
    if (Message.Result == HTTRANSPARENT)
      Message.Result = HTCLIENT;
    break;

  default:
    ControlWndProc(Message);
  }
  if (FilterMessage)
    m_spSimpleFrameSite->PostMessageFilter(Handle, Message.Msg,
                         Message.WParam,
                         Message.LParam,
                         (LRESULT*)&Message.Result,
                         dwCookie);
}

// TWinControlAccess: Template which wraps the Window interface of an ActiveX Control
//
template <class T>
class TWinControlAccess: public T
{
  __fastcall virtual TWinControlAccess(Classes::TComponent* AOwner): T(AOwner) {}
public:
  __fastcall TWinControlAccess(HWND ParentWindow): T(ParentWindow) {}
  HWND GetWindowHandle(void) {return WindowHandle;}
  void DoDestroyHandle(void) {DestroyHandle();}
};

// TVclControlImpl
//
template <class T,                        // User class implementing Control
          class TVCL,                     // Underlying VCL type used in One-Step Conversion
          const CLSID* pclsid,            // Class ID of Control
          const IID* piid,                // Primary interface of Control
          const IID* peventsid,           // Event (outgoing) interface of Control
          const GUID* plibid>             // GUID of TypeLibrary
class ATL_NO_VTABLE TVclControlImpl:
      public CComObjectRootEx<CComObjectThreadModel>,
      public CComCoClass<T, pclsid>,
      public TVclComControl<T, TVCL>,
      public IProvideClassInfo2Impl<pclsid, peventsid, plibid>,
      public IPersistStorageImpl<T>,
      public IPersistStreamInitImpl<T>,
      public IQuickActivateImpl<T>,
      public IOleControlImpl<T>,
      public IOleObjectImpl<T>,
      public IOleInPlaceActiveObjectImpl<T>,
      public IViewObjectExImpl<T>,
      public IOleInPlaceObjectWindowlessImpl<T>,
      public IDataObjectImpl<T>,
      public ISpecifyPropertyPagesImpl<T>,
      public IConnectionPointContainerImpl<T>,
      public IPropertyNotifySinkCP<T, CComDynamicUnkArray>,
      public ISupportErrorInfo,
      public ISimpleFrameSiteImpl<T>
{
public:

  TVclControlImpl()
  {}

 ~TVclControlImpl()
  {}

  // Returns pointer to outer IUnknown
  //
  virtual IUnknown* GetControllingUnknown()
  {
    return (static_cast<T*>(this))->GetUnknown();
  }

  // Empty Message Map provides default implementation of 'ProcessWindowMessage'
  //
  BEGIN_MSG_MAP(TVclControlImpl)
  END_MSG_MAP()

  // Map of connection points supported
  //
  BEGIN_CONNECTION_POINT_MAP(T)
     CONNECTION_POINT_ENTRY(*peventsid)
     CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
  END_CONNECTION_POINT_MAP()

  // This macro declares a static routine which returns the default OLEMISC_xxxx
  // flags used for controls. The macro may be redefined in the derived class
  //  to specify other OLEMISC_xxxx values.
  //
  DECLARE_OLEMISC_FLAGS(dwDefaultControlMiscFlags);

  // Verbs supported by Control. Provides default implementation of _GetVerbs()
  // This table can be redefined in the user's class if the latter supports additional
  // Verbs (or does not want the default 'Properties' verb).
  //
  BEGIN_VERB_MAP()
    VERB_ENTRY(0, L"Properties")
  END_VERB_MAP()

  // IOleInPlaceObject
  //
  STDMETHOD(SetObjectRects)(LPCRECT prcPos,LPCRECT prcClip)
  {
    try
    {
      if (m_VclWinCtl)
        m_VclWinCtl->SetBounds(prcPos->left, prcPos->top,
                     prcPos->right - prcPos->left, prcPos->bottom - prcPos->top);
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }

  HRESULT FinalConstruct()
  {
    try
    {
      Initialize();
    }
    catch(Exception &e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }

  void FinalRelease()
  {}

  // IViewObjectEx
  //
  STDMETHOD(GetViewStatus)(DWORD* pdwStatus)
  {
    ATLTRACE(_T("IViewObjectExImpl::GetViewStatus\n"));
    *pdwStatus = VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE;
    return S_OK;
  }

  // ISupportErrorInfo
  //
  STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
  {
    ATLTRACE(_T("ISupportErrorInfo::InterfaceSupportsErrorInfo\n"));
    const _ATL_INTMAP_ENTRY *pentries = static_cast<T*>(this)->_GetEntries();

    while (pentries->piid)
    {
      if (InlineIsEqualGUID(*(pentries->piid),riid))
        return S_OK;
      pentries++;
    }
    return S_FALSE;
  }

  // IOleControl
  // The following implementation supports BACKCOLOR, FORECOLOR, and FONT ambient
  // properties.  Note that a VCL control will ignore the container's BACKCOLOR
  // property unless the control's ParentColor property is True. FORECOLOR and
  // FONT properties are ignored unless the control's ParentFont property is True.
  //
  STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
  {
    ATLTRACE(_T("IOleControl::OnAmbientPropertyChange\n"));
    ATLTRACE(_T(" -- DISPID = %d (%d)\n"), dispid);
    if (m_VclWinCtl == NULL || !m_AmbientDriver)
      return S_OK;
    if (dispid == DISPID_AMBIENT_BACKCOLOR || dispid == 0)
    {
      TAutoArgs<0> arg;
      HRESULT hres = m_AmbientDriver.OlePropertyGet(DISPID_AMBIENT_BACKCOLOR, arg);
      if (hres != S_OK)
        return hres;
      m_VclWinCtl->Perform(CM_PARENTCOLORCHANGED, 1, arg.GetRetVariant());
    }
    if (dispid == DISPID_AMBIENT_FORECOLOR ||
        dispid == DISPID_AMBIENT_FONT ||
        dispid == 0)
    {
      TAutoArgs<0> arg;
      TPtr<TFont> pFont = new TFont;
      HRESULT hres = m_AmbientDriver.OlePropertyGet(DISPID_AMBIENT_FORECOLOR, arg);
      if (hres != S_OK)
        return hres;
      pFont->Color = (TColor)(int)arg.GetRetVariant();
      hres = m_AmbientDriver.OlePropertyGet(DISPID_AMBIENT_FONT, arg);
      if (hres != S_OK)
        return hres;
      SetOleFont(pFont, (IFontDisp*)(IUnknown*)arg.GetRetVariant());
      m_VclWinCtl->Perform(CM_PARENTFONTCHANGED, 1, int(&pFont));
    }
    return S_OK;
  }

  // IAtlSubCtl
  //
  HRESULT OnDraw(ATL_DRAWINFO& di)
  {
    try
    {
      if (m_VclCtl)
        m_VclCtl->PaintTo(di.hdcDraw, di.prcBounds->left, di.prcBounds->top);
    }
    catch (Exception& e)
    {
      return (static_cast<T*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }
};

// TValidateLicense - Standard implementation of 'IsLicenseValid' that simply returns
//             TRUE. When Control Licensing support is enabled, the Wizard
//             automatically assigns your control a GUID as the Runtime License
//             string. However, your control still needs to validate whether the
//             it is licensed on the machine (before handing out the runtime
//             license string for example). This class provides a standard implementation
//             of the validation routine that simply returns TRUE.
//
class TValidateLicense
{
public:
  static BOOL IsGUIDInFile(const WCHAR *szGUID, const CHAR *szLicFileName)
  {
    LPCSTR pName = 0;

    // Find name of LicFileName
    //
    OFSTRUCT ofs;
    ZeroMemory(&ofs, sizeof(ofs));
    ofs.cBytes = sizeof(OFSTRUCT);

    // We'll look for the license file in the directory of the OCX..
    // Then in the current, app, Windows, System or PATH directories!!
    //
    CHAR szModule[_MAX_PATH];
    ::GetModuleFileNameA(_Module.GetModuleInstance(), szModule, sizeof(szModule));
    CHAR szDir[_MAX_DIR];
    CHAR szDrv[_MAX_DRIVE];
    fnsplit(szModule, szDrv, szDir, 0, 0);
    fnmerge(szModule, szDrv, szDir, 0, 0);
    ::lstrcatA(szModule, szLicFileName);
    if (::OpenFile(szModule, &ofs, OF_EXIST) != HFILE_ERROR ||
        ::OpenFile(szLicFileName, &ofs, OF_EXIST) != HFILE_ERROR)
    {
      pName = ofs.szPathName;
    }
    else
    {
      // Could not find file anywhere
      //
      ATLTRACE(_T("Could not find License File\n"));
      return false;
    }

    // Create a TStringList and load from the file
    //
    _ASSERTE(pName);
    TPtr<TStringList> pList = new TStringList;
    pList->LoadFromFile(pName);
    int index = pList->IndexOf(AnsiString(szGUID));
    return (index != -1);
  }
};

// TLicenseString
//  Template built around the runtime license string of a Control.
// 'VerifyLicenseKey' compares string passed in to the runtime license key
// 'GetLicenseKey'  returns the runtime license key
// 'IsLicenseValid'   delegates to T::IsLicenseValid allowing a separate class to
//               handle the license verification.
//
template <class T>
class TLicenseString
{
protected:
  static BOOL VerifyLicenseKey(BSTR str)
  {
    ATLTRACE(_T("TLicenseString::VerifyLicenseKey\n"));
    return !lstrcmpW(str, T::GetLicenseString());
  }

  static BOOL GetLicenseKey(DWORD /*dwReserved*/, BSTR *pStr)
  {
    ATLTRACE(_T("TLicenseString::GetLicenseKey\n"));
    *pStr = ::SysAllocString(T::GetLicenseString());
    return TRUE;
  }

  static BOOL IsLicenseValid()
  {
    ATLTRACE(_T("TLicenseString::IsLicenseValid\n"));
    return T::IsLicenseValid();
  }
};

#include <axctrls.hpp>
#include <vclhew.hpp>


// TVCLPropertyPage
//
template <class IMPLCLASS, const CLSID* P_COCLASS_CLSID, class VCLCLASS>
class ATL_NO_VTABLE TVCLPropertyPage: public CComObjectRootEx<CComSingleThreadModel>,
                                      public IUnknown,
                                      public CComCoClass<IMPLCLASS, P_COCLASS_CLSID>
{
public:
  TVCLPropertyPage(void): m_PPImpl(0), m_InnerUnk(0)
  {}
 ~TVCLPropertyPage()
  {}

  DECLARE_PROTECT_FINAL_CONSTRUCT()

  HRESULT FinalConstruct(void)
  {
    ATLTRACE(_T("TVCLPropertyPage::FinalConstruct\n"));
    try
    {
      // Chain to Base Class
      HRESULT hres = CComObjectRootEx<CComSingleThreadModel>::FinalConstruct();
      if (hres != S_OK)
        return hres;

      // Create PropertyPageImpl passing it us as Controlling IUnknown
      IUnknown* pUnk = (static_cast<IMPLCLASS*>(this))->GetUnknown();
      m_PPImpl = new TPropertyPageImplHack(_di_IUnknown(pUnk));

      // Get IUnknown of PropertyPageImpl
      // using GetInterface hack.
      m_PPImpl->GetInterface(IID_IUnknown, reinterpret_cast<void*>(&m_InnerUnk));

      // Create Underlying VCL class
      m_PPImpl->PropertyPage = new VCLCLASS((TComponent*)NULL);

      // Initialize PropertyPage
      m_PPImpl->InitPropertyPage();
    }
    catch (Exception& e)
    {
      return (static_cast<IMPLCLASS*>(this))->Error(e.Message.c_str());
    }
    return S_OK;
  }

  void FinalRelease(void)
  {
    ATLTRACE(_T("TVCLPropertyPage::FinalRelease\n"));
    try
    {
      if (m_PPImpl->PropertyPage)
        delete m_PPImpl->PropertyPage;
      if (m_PPImpl)
        delete m_PPImpl;
      CComObjectRootEx<CComSingleThreadModel>::FinalRelease();
    }
    catch (Exception& e)
    {
      // don't propagate exception
    }
  }

  // Data members
  //
  TPropertyPageImplHack*              m_PPImpl;
  IUnknown*                           m_InnerUnk;
};

// VCL_CONTROL_COM_INTERFACE_ENTRIES
//
//  This macro defines the entries of the required Interfaces for exposing a VCL
//  component as an ActiveX Control
//
#define  VCL_CONTROL_COM_INTERFACE_ENTRIES(intf)  \
  COM_INTERFACE_ENTRY_IMPL(IViewObjectEx) \
  COM_INTERFACE_ENTRY_IMPL_IID(IID_IViewObject2, IViewObjectEx) \
  COM_INTERFACE_ENTRY_IMPL_IID(IID_IViewObject, IViewObjectEx) \
  COM_INTERFACE_ENTRY_IMPL_IID(IID_IOleInPlaceObject, IOleInPlaceObjectWindowless) \
  COM_INTERFACE_ENTRY_IMPL_IID(IID_IOleWindow, IOleInPlaceObjectWindowless) \
  COM_INTERFACE_ENTRY_IMPL(IOleInPlaceActiveObject) \
  COM_INTERFACE_ENTRY_IMPL(IOleControl) \
  COM_INTERFACE_ENTRY_IMPL(IOleObject) \
  COM_INTERFACE_ENTRY_IMPL(IQuickActivate) \
  COM_INTERFACE_ENTRY_IMPL(IPersistStorage) \
  COM_INTERFACE_ENTRY_IMPL(IPersistStreamInit) \
  COM_INTERFACE_ENTRY_IMPL(ISpecifyPropertyPages) \
  COM_INTERFACE_ENTRY_IMPL(IDataObject) \
  COM_INTERFACE_ENTRY_IMPL(ISimpleFrameSite) \
  COM_INTERFACE_ENTRY(IProvideClassInfo) \
  COM_INTERFACE_ENTRY(IProvideClassInfo2) \
  COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer) \
  COM_INTERFACE_ENTRY(ISupportErrorInfo) \
  COM_INTERFACE_ENTRY(intf) \
  COM_INTERFACE_ENTRY2(IDispatch, intf)

// VCLCONTROL_IMPL

// This macro is used to encapsulate the various base classes an ActiveX VCL Control derives from.
//
#define VCLCONTROL_IMPL(cppClass, CoClass, VclClass, intf, EventID) \
   public TVclControlImpl<cppClass, VclClass, &CLSID_##CoClass, &IID_##intf, &EventID, LIBID_OF_##CoClass>,\
   public IDispatchImpl<intf, &IID_##intf, LIBID_OF_##CoClass>, \
   public TEvents_##CoClass<cppClass>

// COM_MAP entry for VCL-based Property Page
//
#define  PROPERTYPAGE_COM_INTERFACE_ENTRIES                         \
     COM_INTERFACE_ENTRY(IUnknown)                                  \
     COM_INTERFACE_ENTRY_AGGREGATE(IID_IPropertyPage, m_InnerUnk)   \
     COM_INTERFACE_ENTRY_AGGREGATE(IID_IPropertyPage2, m_InnerUnk)

// Base classes of RemoteDataModule
//
// WARNING: this has changed between BCB 4 and BCB 5 due to changes in MIDAS

#define REMOTEDATAMODULE_IMPL(cppClass, CoClass, VclClass, intf) \
  public CComObjectRootEx<CComObjectThreadModel>,                \
  public CComCoClass<cppClass, &CLSID_##CoClass>,                \
  public IAppServerImpl<VclClass, cppClass, intf, &IID_##intf, LIBID_OF_##CoClass>

// Base class of VCL-based Property Page
//
#define PROPERTYPAGE_IMPL(cppClass, CoClass, VclClass) \
   public TVCLPropertyPage<cppClass, &CLSID_##CoClass, VclClass>

#define DECLARE_REMOTEDATAMODULE_REGISTRY(progid)                        \
  UPDATE_REGISTRY_METHOD(                                                \
     TRemoteDataModuleRegistrar RDMRegistrar(GetObjectCLSID(), progid);  \
     hres = RDMRegistrar.UpdateRegistry(bRegister);)

#define DECLARE_ACTIVEXCONTROL_REGISTRY(progid, idbmp) \
  UPDATE_REGISTRY_METHOD( \
     TAxControlRegistrar AXCR(GetObjectCLSID(), progid, idbmp, _GetObjectMiscStatus(), _GetVerbs()); \
     hres = AXCR.UpdateRegistry(bRegister);)

#pragma option pop

#endif //__ATLVCL_H_

