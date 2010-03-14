//---------------------------------------------------------------------------
#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
//---------------------------------------------------------------------------
#include "DataCallbackSink.h"
#include <ComCtrls.hpp>

class TMainForm : public TForm
{
__published:	// Composants gérés par l'EDI
   TEdit *edServer;
   TButton *btnConnect;
   TEdit *edItem;
   TLabel *Label1;
   TLabel *Label2;
   TMemo *memLog;
   TLabel *Label3;
   TButton *btnRead;
   TButton *btnAsyncRead;
   TEdit *edValue;
   TLabel *Label4;
   TButton *btnWrite;
   TButton *btnAsyncWrite;
   TButton *btnRefresh;
   TButton *btnClearLog;
        TListBox *ListBox1;
        TButton *svrButton;
        TStatusBar *StatusBar1;
        TListBox *ListBox2;
   void __fastcall FormDestroy(TObject *Sender);
   void __fastcall btnConnectClick(TObject *Sender);
   void __fastcall btnWriteClick(TObject *Sender);
   void __fastcall btnAsyncWriteClick(TObject *Sender);
   void __fastcall btnReadClick(TObject *Sender);
   void __fastcall btnAsyncReadClick(TObject *Sender);
   void __fastcall btnRefreshClick(TObject *Sender);
   void __fastcall btnClearLogClick(TObject *Sender);
        void __fastcall svrButtonClick(TObject *Sender);
        void __fastcall ListBox1Click(TObject *Sender);
        void __fastcall ListBox2Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
public:		// Déclarations de l'utilisateur
   __fastcall TMainForm(TComponent* Owner);

// implementation:
private:	// Déclarations de l'utilisateur
   // Task allocator
   CComPtr<IMalloc> m_ptrMalloc;

   // If the component is active or not
   VARIANT_BOOL m_vbActive;

   // Server Prog ID
   CComBSTR m_bstrServer;
   // The OPC server
   CComPtr<IOPCServer> m_ptrServer;
   // BrowseServerAddressSpace
//   CComPtr<IUnknown> m_ptrBrowse;
   IOPCBrowseServerAddressSpace *m_ptrBrowse;
   // Item flat identifier
   CComBSTR m_bstrItem;
   // Item Handle
   DWORD m_hItem;

   // Handle of the OPC group
   DWORD	m_hGroup;
   // Group parameters (see OPC doc):
   DWORD	m_dwRate;
   float	m_fDeadBand;
   // The group itself, referenced by an IUnknown interface pointer:
   CComPtr<IUnknown> m_ptrGroup;
   // IOPCSyncIO interface of the group
   CComPtr<IOPCSyncIO> m_ptrSyncIO;
   // IOPCAsyncIO2 interface of the group
   CComPtr<IOPCAsyncIO2> m_ptrAsyncIO;
   // IAsyncIO2 parameters
   DWORD m_dwID; // Used as Transaction ID
   DWORD m_dwCancelID; // The cancel id for asynchronous operations

   // Event Sink from catching the OPC group events,
   // particularly IOPCDataCallback ones.
   CCreatableDataCallbackSink m_DataCallbackSink;

private:
   void __fastcall GetServer();
   void __fastcall GetItems();
   void __fastcall ShowError(HRESULT hr, LPCSTR pszError);   
   void __fastcall CleanupItem();
   void __fastcall Cleanup();
   void __fastcall ConnectToServer();
   void __fastcall ConnectToItem();

   void __fastcall EnableButtons(bool bEnabled);
   void __fastcall DoLog(LPCTSTR pszMsg);

protected:
   // Event handlers:
   void __fastcall OnDataChange(
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

   void __fastcall OnReadComplete(
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

   void __fastcall OnWriteComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

   void __fastcall OnCancelComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup);
};
//---------------------------------------------------------------------------
extern PACKAGE TMainForm *MainForm;
//---------------------------------------------------------------------------
#endif
