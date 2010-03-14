//---------------------------------------------------------------------------
#include <vcl.h>
#include <stdlib.h>
#include <stdio.h>
#pragma hdrstop
#define STRICT 1

#define OPC_SHUTDOWN_MSG	WM_USER+1

#include "aeclient.h"
#include "opcda.h"
#include "opc_ae.h"
#include "WTClientAPI.h"
#include "Unit2.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
void CALLBACK AECallback (HANDLE hConnect, char *pSource, FILETIME timestamp, char *pMsg, DWORD severity);
void CALLBACK OPCErrorCallback (DWORD hResult, char *pMsg);
void CALLBACK OPCShutdownCallback (HANDLE hConnect);

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{//server list
   String Name;
   int	i, Numbr;
   char buf[150];
   if (Sender == BitBtn1) MachineName = "";
   else MachineName = Edit1->Text;

   Numbr = NumberOfOPC_AEServers(MachineName.c_str());
   for (i=0; i<Numbr; i++)
   {
      GetOPC_AEServerName (i, buf, 100);
      Name = String(buf);
      ListBox1->Items->Add(buf);
      StatusBar1->SimpleText = "Select Server and Connect";//MachineName + ":"+ServerName;
   }
}
//---------------------------------------------------------------------------
void CALLBACK AECallback (HANDLE hConnect, char *pSource, FILETIME timestamp, char *pMsg, DWORD severity)
{
//  Callback to receive a message from
//	the Alarms & Events Server
        Form1->NewAEMsg (pSource, timestamp, pMsg, severity);
}
//---------------------------------------------------------------------------

void CALLBACK OPCErrorCallback (DWORD hResult, char *pMsg)
{
//	OPCErrorCallback (hResult, pMsg);
}
//---------------------------------------------------------------------------

void CALLBACK OPCShutdownCallback (HANDLE hConnect)
{
	// Post Message to shutdown the client
	// and return immediately to server
	PostMessage (Form1->Handle, OPC_SHUTDOWN_MSG, 0, 0);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn3Click(TObject *Sender)
{ //disconnect from server
        Timer1->Enabled = false;

	DisconnectOPC_AE(ConnectionHandle);
	ConnectionHandle = INVALID_HANDLE_VALUE;
        StatusBar1->SimpleText = "Disconnected";
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ListBox1Click(TObject *Sender)
{
        ServerName = ListBox1->Items->Strings[ListBox1->ItemIndex];
        ServerName = ServerName.Trim();
        MachineName = MachineName.Trim();
	char buf1[100], buf2[100];
	DWORD	BufferTime, MaxSize;
        //empty listview
        ListView1->Items->Clear();
	strcpy (buf1, MachineName.c_str());
//	strcpy (buf2,"ICONICS.Simulator.1");
//	strcpy (buf2,"Softing.OPCToolboxDemo_ServerDA.1");
	strcpy (buf2, ServerName.c_str());
        StatusBar1->SimpleText = buf2;//MachineName + ":"+ServerName;
	ConnectionHandle = ConnectOPC_AE(buf1, buf2);
	EnableShutdownNotification (ConnectionHandle, OPCShutdownCallback);

	BufferTime = 0;
	MaxSize = 1;

	Create_AE_Subscription (ConnectionHandle, (HANDLE)1, &BufferTime, &MaxSize);
	EnableAECallback(AECallback);
        Timer1->Enabled = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn4Click(TObject *Sender)
{ //server status
	OPCEVENTSERVERSTATUS status;
	WCHAR VendorString[200];
        String m_svrname, m_state;

	status.szVendorInfo = VendorString;
	if (!GetAESvrStatus(ConnectionHandle, &status, 200))
		{
		ShowMessage("Server Status Failure");
		return;
		}

	// display the results
	m_svrname = MachineName;
	m_svrname += " ";
	m_svrname += ServerName;
//	stime = CTime::CTime(status.ftStartTime);
//	dlg.m_starttime = stime.Format ("%H:%M:%S %d %b %Y");
//	stime = CTime::CTime(status.ftCurrentTime);
//	dlg.m_currenttime = stime.Format ("%H:%M:%S %d %b %Y");
//	stime = CTime::CTime(status.ftLastUpdateTime);
//	dlg.m_updatetime = stime.Format ("%H:%M:%S %d %b %Y");
//	dlg.m_groupcount = status.dwGroupCount;
//	dlg.m_bandwidth = status.dwBandWidth;
//	dlg.m_majorversion = status.wMajorVersion;
//	dlg.m_minorversion = status.wMinorVersion;
//	dlg.m_buildnumber = status.wBuildNumber;
//	dlg.m_vendorinfo = status.szVendorInfo;

	switch (status.dwServerState)
		{
		case OPC_STATUS_RUNNING: m_state = "Running"; break;
		case OPC_STATUS_FAILED: m_state = "Failed"; break;
		case OPC_STATUS_NOCONFIG: m_state = "Not Configured"; break;
		case OPC_STATUS_SUSPENDED: m_state = "Suspended"; break;
		case OPC_STATUS_TEST: m_state = "Test"; break;
		}

StatusBar1->SimpleText = m_svrname+"-"+ m_state+"-"
        +status.wMajorVersion+"-"+status.wMinorVersion+"-"
        +status.wBuildNumber+"-"+status.szVendorInfo;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormCreate(TObject *Sender)
{
        AEMsgArray = new TList();
// WTclientCoInit() initializes COM as MULTITHREADED
	WTclientCoInit();

}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
	int	i;
        Timer1->Enabled = false;
        ListView1->Items->Clear();
	// remove any unprocessed event msgs
	// and disconnect if not already
//	for (i=0; i<AEMsgArray->Count; i++)
//		delete ((TAEMsg *)AEMsgArray->Items[i]);
	AEMsgArray->Clear();
	if (ConnectionHandle != INVALID_HANDLE_VALUE)
		DisconnectOPC_AE(ConnectionHandle);
	CoUninitialize();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::NewAEMsg (char *pSource, FILETIME timestamp, char *pMsg, DWORD severity)
{
//
// Received control from the server
// Copy data into local array and
// return to server as quickly as possible
	TAEMsg	*pAEMsg;

	pAEMsg = new TAEMsg();
	pAEMsg->Source = pSource;
	pAEMsg->TimeStamp = timestamp;
	pAEMsg->Message = pMsg;
	pAEMsg->Severity = severity;

	AEMsgArray->Add ((TObject *)pAEMsg);

}
//---------------------------------------------------------------------------


void __fastcall TForm1::Timer1Timer(TObject *Sender)
{
   TAEMsg	*pAEMsg;
   FILETIME	time;
   SYSTEMTIME stTimeStamp;
   FILETIME ftLocal;   LPSTR lpszString;   char	buf[500];
   int	i;

   // update the ListCtrl with any
   // event msgs that have shown up since
   // the last time.
   while (AEMsgArray->Count > 0)
   {
      pAEMsg = (TAEMsg *)AEMsgArray->Items[0];

      i = ListView1->Items->Count;
      TListItem *lvitem = ListView1->Items->Add();
      strncpy (buf, pAEMsg->Source.c_str(), 500);
      lvitem->Caption = buf;

//      time = pAEMsg->TimeStamp;
    // Convert the file time to local time.
    if (!FileTimeToLocalFileTime(&(pAEMsg->TimeStamp), &ftLocal))        return ;    // Convert the local file time from UTC to system time.    FileTimeToSystemTime(&ftLocal, &stTimeStamp);    // Build a string showing the date and time.    lpszString = (char *)malloc(30);    wsprintf(lpszString, "%02d/%02d/%d  %02d:%02d:%02d",        stTimeStamp.wDay, stTimeStamp.wMonth, stTimeStamp.wYear,        stTimeStamp.wHour, stTimeStamp.wMinute,stTimeStamp.wSecond);      strncpy (buf, lpszString, 500);
      lvitem->SubItems->Add(buf);

      strncpy (buf, pAEMsg->Message.c_str(), 500);
      lvitem->SubItems->Add(buf);

      sprintf (buf, "%d", pAEMsg->Severity);
      lvitem->SubItems->Add(buf);
      free(lpszString);
      delete (pAEMsg);
      AEMsgArray->Delete(0);
      }
      // limit the number of items in the ListCtrl
      // to about 500
//      ListView1->UpdateItems(index,index);
      while (ListView1->Items->Count > 500)
        ListView1->Items->Delete(0);

}
//---------------------------------------------------------------------------
void __fastcall TForm1::OnShutdownRequest(TMessage& Msg)
{
	// server has requested disconnect
	// shutdown the connection
	//OnOpcDisconnect();
        BitBtn3Click(NULL);
	ShowMessage("Server Requested Disconnect");
}
//---------------------------------------------------------------------------

