//---------------------------------------------------------------------------
#include <vcl.h>
#include <assert.h>
#pragma hdrstop

#include "Main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMainForm *MainForm;
#define MAX_KEYLEN 256

// Ole initializer/deinitializer
static struct SOLEINIT
{
   SOLEINIT()
   {
      CoInitialize(NULL);
   }
   ~SOLEINIT()
   {
      CoUninitialize();
   }
} oleinit__;

inline AnsiString Variant2Str(VARIANT& v)
{
   Variant var(v);
   return VarToStr(var);
}

#ifdef StrToInt
#undef StrToInt
#endif // StrToInt

//---------------------------------------------------------------------------
__fastcall TMainForm::TMainForm(TComponent* Owner)
   : TForm(Owner),
   m_vbActive(VARIANT_FALSE),
   m_hGroup(0),
   m_dwRate(100),
   m_fDeadBand(0.0f),
   m_hItem(0),
   m_dwID(0),
   m_dwCancelID(0)
{
   CoGetMalloc(MEMCTX_TASK, &m_ptrMalloc);
   assert(m_ptrMalloc != NULL);
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::CleanupItem()
{
   if (0 != m_hItem)
   {
      assert(m_ptrGroup != NULL);
      // Get the item management interface of the group
      CComPtr<IOPCItemMgt> ptrItMgm;
      OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCItemMgt,
                              reinterpret_cast<LPVOID*>(&ptrItMgm)));
      assert(ptrItMgm != NULL);

      HRESULT* phResult = NULL;
      OLECHECK(ptrItMgm->RemoveItems(1, &m_hItem, &phResult));

      // Check the item result for errors
      assert(phResult != NULL);
      HRESULT hr = phResult[0];
      m_ptrMalloc->Free(phResult);
      OLECHECK((HRESULT)hr);

      // Final cleanup
      m_hItem = 0;
   }
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::Cleanup()
{
   // Free the current item
   CleanupItem();

   if (0 != m_hGroup)
   {
      // Disconnect group events
      m_DataCallbackSink.Disconnect();
      // Release the group
      m_ptrSyncIO.Release();
      m_ptrAsyncIO.Release();
      m_ptrGroup.Release();
      // Remove the group from the server itself
      OLECHECK(m_ptrServer->RemoveGroup(m_hGroup, FALSE));
      m_hGroup = 0;
   }
   assert(m_ptrSyncIO == NULL);
   assert(m_ptrAsyncIO == NULL);
   assert(m_ptrGroup == NULL);

   // Release the OPC server
   m_ptrServer.Release();

   // Set Deactivate Flag
   m_vbActive = VARIANT_FALSE;
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::ConnectToServer()
{
   assert(!m_vbActive);
   assert(m_bstrServer.Length() != 0);

   // INITIALIZATION:
   assert(m_ptrMalloc != NULL);

   // Create the OPC Server:
   CLSID clsid;
   OLECHECK(CLSIDFromProgID(m_bstrServer, &clsid));
   OLECHECK(CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IOPCServer,
                                     reinterpret_cast<LPVOID*>(&m_ptrServer)));

   OLECHECK(m_ptrServer->QueryInterface(IID_IOPCBrowseServerAddressSpace,
                                        reinterpret_cast<LPVOID*>(&m_ptrBrowse)));
   // Create a group with a unique name:
   LPOLESTR pszGRPID = NULL;
   try
   {
      GUID guidGroupName;
      OLECHECK(CoCreateGuid(&guidGroupName));
      OLECHECK(StringFromCLSID(clsid, &pszGRPID));
      assert(pszGRPID != NULL);
      DWORD dwRevisedRate = 0;
      OLECHECK(m_ptrServer->AddGroup(pszGRPID, TRUE, m_dwRate, 0, 0, &m_fDeadBand,
			      0, &m_hGroup, &dwRevisedRate, IID_IUnknown,
   			   reinterpret_cast<LPUNKNOWN*>(&m_ptrGroup)));
      m_dwRate = dwRevisedRate;
      assert(m_ptrGroup != NULL);
      CoTaskMemFree(pszGRPID);
   }
   catch(...)
   {
      if (pszGRPID != NULL)
      {
         CoTaskMemFree(pszGRPID);
      }
      throw;
   }

   // Get the sync IO interface of the group:
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCSyncIO,
                           reinterpret_cast<LPVOID*>(&m_ptrSyncIO)));

   // Get the async IO interface of the group:
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCAsyncIO2,
                           reinterpret_cast<LPVOID*>(&m_ptrAsyncIO)));

   // Connect the event handlers to those of this form:
   m_DataCallbackSink.EvDataChange = OnDataChange;
   m_DataCallbackSink.EvReadComplete = OnReadComplete;
   m_DataCallbackSink.EvWriteComplete = OnWriteComplete;
   m_DataCallbackSink.EvCancelComplete = OnCancelComplete;

   // Connect the IOPCDataCallback sink to the group
   OLECHECK(m_DataCallbackSink.Connect(m_ptrGroup));
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::ConnectToItem()
{
   assert(!m_vbActive);
   assert(m_ptrGroup != NULL);
   assert(m_bstrItem.Length() != 0);

   // Get the item management interface of the group
   CComPtr<IOPCItemMgt> ptrItMgm;
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCItemMgt,
                           reinterpret_cast<LPVOID*>(&ptrItMgm)));

	OPCITEMDEF itemdef;
	HRESULT *phResult = NULL;
	OPCITEMRESULT *pItemState = NULL;

	// Define one item
	//
   USES_CONVERSION;
	itemdef.szItemID = m_bstrItem;
	itemdef.szAccessPath = T2OLE(_T(""));
	itemdef.bActive = TRUE;
	itemdef.hClient = reinterpret_cast<DWORD>(Handle);
	itemdef.dwBlobSize = 0;
	itemdef.pBlob = NULL;
	itemdef.vtRequestedDataType = 0;

	// Add then items and check the hresults
	//
	OLECHECK(ptrItMgm->AddItems(1, &itemdef, &pItemState, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);

   if (SUCCEEDED(phResult[0]))
   {
      // Store item server handle for future use
      assert(pItemState != NULL);
      m_hItem = pItemState[0].hServer;

      if (pItemState[0].pBlob != NULL)
      {
         m_ptrMalloc->Free(pItemState[0].pBlob);
	   }

	   // Free the returned results
	   //
      m_ptrMalloc->Free(phResult);
	   m_ptrMalloc->Free(pItemState);

      // We always must call the SetEnabled of AsyncIO2:
      assert(m_ptrAsyncIO != NULL);
      OLECHECK(m_ptrAsyncIO->SetEnable(TRUE));

      // If it gets here, the server is active
      m_vbActive = VARIANT_TRUE;
   }
   else
   {
      HRESULT hr = phResult[0];

	   // Free the returned results
	   //
      m_ptrMalloc->Free(phResult);
	   m_ptrMalloc->Free(pItemState);

      OLECHECK((HRESULT)hr);// Always the exception is raised here
   }
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::EnableButtons(bool bEnabled)
{
   btnRead->Enabled = bEnabled;
   btnAsyncRead->Enabled = bEnabled;
   btnWrite->Enabled = bEnabled;
   btnAsyncWrite->Enabled = bEnabled;
   btnRefresh->Enabled = bEnabled;
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::DoLog(LPCTSTR pszMsg)
{
   memLog->Lines->Add(pszMsg);
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::OnDataChange(
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
   DoLog(Format(_T("DataChange: Value = '%s'"),
               ARRAYOFCONST((Variant2Str(*pvValues)))).c_str());
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::OnReadComplete(
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
   DoLog(Format(_T("ReadComplete: Value = '%s'"),
               ARRAYOFCONST((Variant2Str(*pvValues)))).c_str());
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::OnWriteComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   DoLog(_T("Async Write completed!"));
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::OnCancelComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup)
{
}

//---------------------------------------------------------------------------

void __fastcall TMainForm::FormDestroy(TObject *Sender)
{
   Cleanup();
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::btnConnectClick(TObject *Sender)
{
   static const LPCTSTR szConnect = _T("&Connect");
   static const LPCTSTR szDisconnect = _T("Dis&connect");

   if (btnConnect->Tag == 0)
   {
      try
      {
         m_bstrServer.Empty();
         m_bstrServer = WideString(edServer->Text);

         ConnectToServer();
	 StatusBar1->SimpleText ="Connected to server: "+edServer->Text +" Select Item...";
         GetItems();

         btnConnect->Tag = 1;
         btnConnect->Caption = szDisconnect;

         EnableButtons(true);
      }
      catch(...)
      {
         Cleanup();
         throw;
      }
   }
   else
   {
      try
      {
         Cleanup();

         btnConnect->Tag = 0;
         btnConnect->Caption = szConnect;

	 StatusBar1->SimpleText ="Disonnected...";
         EnableButtons(false);
      }
      catch(...)
      {
         throw;
      }
   }
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnWriteClick(TObject *Sender)
{
   // Only integers, by now:
   CComVariant vValue = Sysutils::StrToInt(edValue->Text);

   assert(m_ptrMalloc != NULL);
   assert(m_ptrSyncIO != NULL);

   HRESULT *phResult = NULL;
   OLECHECK(m_ptrSyncIO->Write(1, &m_hItem, &vValue, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);
   HRESULT hr = phResult[0];
   m_ptrMalloc->Free(phResult);
   OLECHECK((HRESULT)hr);
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnAsyncWriteClick(TObject *Sender)
{
   // Only integers, by now:
   CComVariant vValue = Sysutils::StrToInt(edValue->Text);

   assert(m_ptrMalloc != NULL);
   assert(m_ptrAsyncIO != NULL);

   HRESULT *phResult = NULL;
   OLECHECK(m_ptrAsyncIO->Write(1, &m_hItem, &vValue, m_dwID,
                          &m_dwCancelID, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);
   HRESULT hr = phResult[0];
   m_ptrMalloc->Free(phResult);
   OLECHECK((HRESULT)hr);
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnReadClick(TObject *Sender)
{
   try
   {
      assert(m_ptrMalloc != NULL);
      assert(m_ptrSyncIO != NULL);

      HRESULT *phResult = NULL;
      OPCITEMSTATE* pItemState;
      OLECHECK(m_ptrSyncIO->Read(OPC_DS_DEVICE, 1, &m_hItem, &pItemState, &phResult));

      assert(phResult != NULL);
      if (SUCCEEDED(phResult[0]))
      {
         assert(pItemState != NULL);

         DoLog(Format(_T("Synchronous Read: Value = '%s'"),
                     ARRAYOFCONST((Variant2Str(pItemState[0].vDataValue)))).c_str());

         // Free results
         m_ptrMalloc->Free(pItemState);
         m_ptrMalloc->Free(phResult);
      }
      else
      {
         HRESULT hr = phResult[0];

	      // Free the returned results
	      //
         m_ptrMalloc->Free(phResult);
	      m_ptrMalloc->Free(pItemState);

         OLECHECK(hr);// Always the exception is raised here
      }
   }
   catch(...)
   {
   }
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnAsyncReadClick(TObject *Sender)
{
   // Only integers, by now:
   CComVariant vValue = Sysutils::StrToInt(edValue->Text);

   assert(m_ptrMalloc != NULL);
   assert(m_ptrAsyncIO != NULL);

   HRESULT *phResult = NULL;
   OLECHECK(m_ptrAsyncIO->Read(1, &m_hItem, m_dwID, &m_dwCancelID, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);
   HRESULT hr = phResult[0];
   m_ptrMalloc->Free(phResult);
   OLECHECK((HRESULT)hr);
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnRefreshClick(TObject *Sender)
{
   assert(m_ptrAsyncIO != NULL);
   OLECHECK(m_ptrAsyncIO->Refresh2(OPC_DS_DEVICE, m_dwID, &m_dwCancelID));
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::btnClearLogClick(TObject *Sender)
{
   memLog->Lines->Clear();
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::GetServer()
{
//this code segment is take from SST_Client sample
// Copyright © 1998-1999 SST, a division of Woodhead Canada Limited
// www.sstech.on.ca
//
// Created by Richard Illes
// May 21, 1998
	// browse registry for OPC 1.0A Servers
	HKEY hk = HKEY_CLASSES_ROOT;
	TCHAR szKey[MAX_KEYLEN];
	for(int nIndex = 0; ::RegEnumKey(hk, nIndex, szKey, MAX_KEYLEN) == ERROR_SUCCESS; nIndex++)
	{
		HKEY hProgID;
		TCHAR szDummy[MAX_KEYLEN];
		if(::RegOpenKey(hk, szKey, &hProgID) == ERROR_SUCCESS)
		{
			LONG lSize = MAX_KEYLEN;
			if(::RegQueryValue(hProgID, "OPC", szDummy, &lSize) == ERROR_SUCCESS)
			{
                        ListBox1->Items->Add(szKey);
			}
			::RegCloseKey(hProgID);
		}
	}

if (ListBox1->Items->Count > 0){
        svrButton->Visible = false;
        StatusBar1->SimpleText ="Select Server and Connect.";
        }
else
        StatusBar1->SimpleText ="No Server found.";

}
//---------------------------------------------------------------------------


void __fastcall TMainForm::svrButtonClick(TObject *Sender)
{
GetServer();
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::ListBox1Click(TObject *Sender)
{
edServer->Text=ListBox1->Items->Strings[ListBox1->ItemIndex];
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::GetItems()
{
	// loop until all items are added
	char sz2[200];
	TCHAR szBuffer[256];
	HRESULT hr = 0;
	int nTestItem = 0; // how many items there are
	IEnumString* pEnumString = NULL;
	int nCount = 0;
	USES_CONVERSION;
        ListBox2->Items->Clear();

   OLECHECK(m_ptrBrowse->BrowseOPCItemIDs(OPC_FLAT, L""/*NULL*/, VT_EMPTY, 0, &pEnumString));
	LPOLESTR pszName = NULL;
        ULONG count = 0;
        while((hr = pEnumString->Next(1, &pszName, &count)) == S_OK)
        {
            ListBox2->Items->Add(OLE2T(pszName));
            ::CoTaskMemFree(pszName);
        }
        pEnumString->Release();
}
//---------------------------------------------------------------------------
void __fastcall TMainForm::ListBox2Click(TObject *Sender)
{
   edItem->Text=ListBox2->Items->Strings[ListBox2->ItemIndex];
   m_bstrItem.Empty();
   m_bstrItem = WideString(edItem->Text);
   ConnectToItem();
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::FormShow(TObject *Sender)
{
   svrButtonClick(NULL);
}
//---------------------------------------------------------------------------

