//---------------------------------------------------------------------------
#include <vcl.h>
#include <stdlib.h>
#include <stdio.h>
#pragma hdrstop
#define STRICT 1

#include "Unit1.h"
#include "opcda.h"
#include "opc_ae.h"
#include "WTClientAPI.h"
#include "Unit2.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
void CALLBACK OPCUpdateCallback (HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality);
void CALLBACK OPCErrorCallback (DWORD hResult, char *pMsg);
void CALLBACK OPCShutdownCallback (HANDLE hConnect);

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{
	String Name;
	int	i, Numbr;
	char buf[150];
MachineName = "";

	Numbr = NumberOfOPCServers(true, MachineName.c_str());
	for (i=0; i<Numbr; i++)
		{
		GetOPCServerName (i, buf, 150);
		Name = String(buf);
		ListBox1->Items->Add(buf);
       StatusBar1->SimpleText = buf;//MachineName + ":"+ServerName;
		}
}
//---------------------------------------------------------------------------
void CALLBACK OPCUpdateCallback (HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality)
{
	Form1->OPCDataUpdate (hGroup, hItem, pVar, timestamp, quality);
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
//	PostMessage (hFrame, OPC_SHUTDOWN_MSG, 0, 0);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{
	char	buf[80];

        unsigned long m_rate=1000;
        float m_deadband=0.0;
//	strncpy (buf, dlg.m_name, 80);
	strcpy (buf, Edit1->Text.c_str());

	hGroup = AddOPCGroup (ConnectionHandle, buf, &m_rate, &m_deadband);

	if (hGroup == INVALID_HANDLE_VALUE)
		{
		sprintf(buf, "Error from AddGroup () (%lx)\n", 0);
		MessageBox (NULL,buf,"Error", MB_OK);
		return;
		}
        else StatusBar1->SimpleText = "Add group " + Edit1->Text+ "success";
	// OPC Group has been created
}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn3Click(TObject *Sender)
{
	char	buf[100];
	String Name;
	int	i, Numbr;

	Numbr = NumberOfOPCItems (ConnectionHandle);
	for (i=0; i<Numbr; i++)
		{
		GetOPCItemName (ConnectionHandle, i, buf, 100);
		Name = buf;
		ListBox2->Items->Add(Name);
		}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ListBox1Click(TObject *Sender)
{
        ServerName = ListBox1->Items->Strings[ListBox1->ItemIndex];
        ServerName = ServerName.Trim();
        MachineName = MachineName.Trim();
	char buf1[100], buf2[100];

	strcpy (buf1, "");//MachineName.c_str());
//	strcpy (buf2,"ICONICS.Simulator.1");
//	strcpy (buf2,"Softing.OPCToolboxDemo_ServerDA.1");
	strcpy (buf2, ServerName.c_str());
        StatusBar1->SimpleText = buf2;//MachineName + ":"+ServerName;
	ConnectionHandle = ConnectOPC(buf1, buf2, false);
	EnableOPCNotification (ConnectionHandle, OPCUpdateCallback);
	EnableShutdownNotification (ConnectionHandle, OPCShutdownCallback);

}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn4Click(TObject *Sender)
{
	OPCSERVERSTATUS	status;
	WCHAR VendorString[200];
        String m_svrname, m_state;

	status.szVendorInfo = VendorString;
	if (!GetSvrStatus (ConnectionHandle, &status, 200))
		{
		MessageBox (NULL,"Error","Server Status Failure", MB_OK);
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
        ItemList = new TList();
// WTclientCoInit() initializes COM as MULTITHREADED
	WTclientCoInit();

}
//---------------------------------------------------------------------------


void __fastcall TForm1::BitBtn5Click(TObject *Sender)
{

  TItemDef	*pItem;
  int i;

  pItem = new TItemDef();

  pItem->Name = ListBox2->Items->Strings[ListBox2->ItemIndex];//ItemName;
  pItem->Path = "";//Access;
  pItem->Handle = AddOPCItem(ConnectionHandle, hGroup, pItem->Name.c_str());
  ItemList->Add(pItem);
  ListView1->Items->Clear();
  for (i=0; i< ItemList->Count; i++)
    {
    pItem = (TItemDef *)ItemList->Items[i];
    TListItem *item = ListView1->Items->Add();
    item->Caption = pItem->Name;
    item->SubItems->Add("?");
    item->SubItems->Add("?");
    item->SubItems->Add("?");
    }//end for

}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
	int	i;
	TItemDef *pItem;
	//
	// if user closed the document w/o removing the OPC Group
	// clean-up the item list and notify the server that
	// we are done.
	//
	for (i=0; i<ItemList->Count; i++)
		{
		pItem = (TItemDef *)ItemList->Items[i];
		// remove each item from the server
        RemoveOPCItem(ConnectionHandle, hGroup, pItem->Handle);
		// delete the local item object
//		delete (pItem);
		}
	ItemList->Clear();

	if (hGroup != INVALID_HANDLE_VALUE)
		RemoveOPCGroup(ConnectionHandle, hGroup);
	if (ConnectionHandle != INVALID_HANDLE_VALUE)
		DisconnectOPC(ConnectionHandle);

}
//---------------------------------------------------------------------------
void __fastcall TForm1::OPCDataUpdate(HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality)
{
 int i;
TItemDef *pItem;
	// find the item definition
	for (i=0; i< ItemList->Count; i++)
		{
		pItem = (TItemDef *)ItemList->Items[i];
		if (pItem->Handle == (HANDLE)hItem)
			{
			// update the data
			VariantClear(&(pItem->Value));
			VariantCopy (&(pItem->Value), pVar);
			pItem->TimeStamp = timestamp;
			pItem->Quality = quality;
                        UpdateDataView(i);
			}
		}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::UpdateDataView(int index)
{
// redraw the updated item list
SYSTEMTIME stTimeStamp;
FILETIME ftLocal;LPSTR lpszString;TItemDef *pItem;
pItem = (TItemDef *)ItemList->Items[index];

String tmpstr;
    // Convert the file time to local time.
    if (!FileTimeToLocalFileTime(&(pItem->TimeStamp), &ftLocal))        return ;    // Convert the local file time from UTC to system time.    FileTimeToSystemTime(&ftLocal, &stTimeStamp);    // Build a string showing the date and time.    lpszString = (char *)malloc(20);    wsprintf(lpszString, "%02d/%02d/%d  %02d:%02d",        stTimeStamp.wDay, stTimeStamp.wMonth, stTimeStamp.wYear,        stTimeStamp.wHour, stTimeStamp.wMinute);    switch (pItem->Quality)
	{
	case OPC_QUALITY_GOOD:
	        tmpstr += " (Quality Good)";
	break;
	case OPC_QUALITY_BAD:
	        tmpstr += " (Quality Bad)";
	break;
	default:
	        tmpstr += " (Quality Uncertain)";
        break;
	}
//StatusBar1->SimpleText = tmpstr;
TListItem *item = ListView1->Items->Item[index];
//item->Caption = pItem->Name;
item->SubItems->Strings[0] = VarToStr(pItem->Value);
item->SubItems->Strings[1] = lpszString;
item->SubItems->Strings[2] = tmpstr;
ListView1->UpdateItems(index,index);
free(lpszString);
}
//---------------------------------------------------------------------------

