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
{//tree browser
/*	char	buf[100];
	String Name;
	int	i, Numbr;

	Numbr = NumberOfOPCItems (ConnectionHandle);
	for (i=0; i<Numbr; i++)
		{
		GetOPCItemName (ConnectionHandle, i, buf, 100);
		Name = buf;
		ListBox2->Items->Add(Name);
		}
*/
        TTreeNode *Node1;// = new TTreeNode();
        TTreeNode *pTree;// = new TTreeNode();
        HTREEITEM mItemId;

	// Browse to Root
	BrowseTo(ConnectionHandle, "");
        pTree=TreeView1->Items->Add(NULL, "Root");
        mItemId = pTree->ItemId;
/*
	// Browse To Root
	BrowseTo (pView->ConnectionHandle, "");
*/	// resolve all nodes one level below the root
	// As tree is expanded, keep adding nodes from
	// the OPC Server list until complete item name
	// tree is built
	ResolveOneLevel (pTree);

	return;// TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
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
        BitBtn2Click(NULL);
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
  TTreeNode *hItem = TreeView1->Selected;

  pItem->Name = QualifiedName(hItem);//ItemName;
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
void __fastcall TForm1::ResolveOneLevel (TTreeNode *hItem)
{
	int	i, numbr;
	char	buf[200];
	String UnqualifiedName;
	WORD NameSpace;

	TTreeNode *hBranch;// = new TTreeNode();
	TStringList  *pLocalArray = new TStringList();

	// if the TreeCtrl has already been resolved
	// to this level, simply return;
	if (hItem->HasChildren)
		return;// (TRUE);

	if (!BrowseTo(ConnectionHandle, QualifiedName(hItem).c_str()))
		return;//(FALSE);	// cannot find current Browse Position

	// Get All leaf names
	numbr = BrowseItems(ConnectionHandle, OPC_LEAF);
	for (i=0; i<numbr; i++)
		{
		if (GetOPCItemName (ConnectionHandle, i, buf, 200))
			{
			UnqualifiedName = UnqualifyName(buf);
        StatusBar1->SimpleText = UnqualifiedName;
			strncpy (buf, UnqualifiedName.c_str(), 200);
//			hBranch = TreeView1->InsertItem(TVIF_TEXT, _T(buf), 0, 0, 0, 0, 0, hItem, NULL);
			hBranch = TreeView1->Items->AddChild(hItem, String(buf));
			}
		}

	GetNameSpace (ConnectionHandle, &NameSpace);
	// if namespace is flat,  server will not differentiate
	// between branch & leaf browse requests
	if (NameSpace == OPC_NS_FLAT)
		return;//(TRUE);

	// Get all Branch names
	numbr = BrowseItems(ConnectionHandle, OPC_BRANCH);
	// Read node names into local array
	for (i=0; i<numbr; i++)
		{
		if (GetOPCItemName (ConnectionHandle, i, buf, 200))
			pLocalArray->Add(buf);
		}

	for (i=0; i<pLocalArray->Count; i++)
		{
		UnqualifiedName = UnqualifyName(pLocalArray->Strings[i]);
		strncpy (buf, UnqualifiedName.c_str(), 200);
                hBranch = TreeView1->Items->AddChild(hItem, String(buf));
		}


	return;// (TRUE);
}
//---------------------------------------------------------------------------
//
// Derive fully qualified name from current
// Tree position.
//
String __fastcall TForm1::QualifiedName (TTreeNode *hItem)
{
	TStringList *NodeList = new TStringList();
	String QName;
	int	i, numbr;

	if (hItem == TreeView1)
		return ("");
TTreeNode *CurrentNode = hItem;
	while (1)
		{
		NodeList->Add(CurrentNode->Text);
		CurrentNode = CurrentNode->Parent;
		if ((CurrentNode == TreeView1->Items->GetFirstNode())||	// IsRoot?
                    (CurrentNode == NULL))
			break;
		}

	numbr = NodeList->Count;
	QName = NodeList->Strings[numbr-1];
	for (i=1; i<numbr; i++)
		{
		QName += SetWTclientQualifier(0);
		QName += NodeList->Strings[numbr-i-1];
		}
	return (QName);
}
//---------------------------------------------------------------------------
//
// Derive unqualified name from passed string
//
String __fastcall TForm1::UnqualifyName (String Name)
{
	int	indx;
	String UnqualifiedName;

	indx = Name.LastDelimiter(SetWTclientQualifier(0));
	if (indx == 0)
		UnqualifiedName = Name;
        else
		{
		++indx;
		UnqualifiedName = Name.SubString(indx,Name.Length());
		}
	return (UnqualifiedName);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::TreeView1Expanded(TObject *Sender, TTreeNode *Node)
{
//
// whenere the TreeCtrl expands, resolve all nodes
// below the current position
//
	TTreeNode *hChildItem;

	hChildItem = Node->getFirstChild();
//        StatusBar1->SimpleText = Node->Text +"->"+hChildItem->Text;
	while (hChildItem != NULL)
		{
		ResolveOneLevel(hChildItem);
            hChildItem = Node->GetNextChild(hChildItem);
		}

}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn6Click(TObject *Sender)
{
char *Qual = (char *)malloc(2);
Qual = Edit2->Text.c_str();
SetWTclientQualifier(Qual[0]);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn7Click(TObject *Sender)
{
//	CStatic		*pHeader;
//	LV_COLUMN		lvcolumn;
//	CListCtrl		*pList;
//	CRect	rect;
	int	i,nProps;
	DWORD	PropertyID;
	VARTYPE	vt;
	VARIANT	var;
	char	Description[200];

  TListItem *Item=ListView1->Selected;
  ListView2->Items->Clear();
	VariantInit(&var);
	nProps = NumberOfItemProperties (ConnectionHandle, Item->Caption.c_str());
	for (i=0; i<nProps; i++)
	   {
	   if (GetItemPropertyDescription (ConnectionHandle, i, &PropertyID, &vt, Description, 200))
	     if (ReadPropertyValue (ConnectionHandle, Item->Caption.c_str(), PropertyID, &var))
		{
		//AddPropertytoList (pList, Description, var);
                TListItem *item = ListView2->Items->Add();
                item->Caption = String(Description);
                item->SubItems->Add(VarToStr(var));
		VariantClear (&var);
		}
	   }

//	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//---------------------------------------------------------------------------

