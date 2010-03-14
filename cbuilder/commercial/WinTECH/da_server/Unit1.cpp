//---------------------------------------------------------------------------
#include <vcl.h>
#include <stdio.h>
//#include<fstream.h>
#pragma hdrstop

#include "Unit1.h"
#include <variant.h>
#include "opcda.h"
#include "opc_ae.h"
#include "OpcError.h"
#define STRICT 1
#include "WTOPCsvrAPI.h"
#include "Unit2.h"
#include "Unit3.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
static const GUID CLSID_Svr = 
{ 0x6764a030, 0x70c, 0x11d3, { 0x80, 0xd6, 0x0, 0x60, 0x97, 0x58, 0x58, 0xbe } };

__fastcall CTag::CTag(void) : TObject()
{
	VariantInit(&Value);
	Value.vt = VT_R4;
	Value.fltVal = 0.0;
}

__fastcall CTag::~CTag(void)
{
	VariantClear(&Value);
}

void CALLBACK _export UnknownTagProc(LPSTR Path, LPSTR Name);
void CALLBACK _export TagRemovedProc(HANDLE hTag, LPSTR Path, LPSTR Name);
void CALLBACK _export WriteNotifyProc(HANDLE Handle, VARIANT *pNewValue, DWORD *pDevError);
void CALLBACK _export DeviceReadProc(HANDLE Handle, VARIANT *pNewValue, WORD *pQuality, FILETIME *pTimestamp);
void CALLBACK _export DisconnectProc(DWORD NumbrActiveClients);
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
	String	FilePath;
        char TmpStr[80]="                      ";
        TListItem  *ListItem;
	int	i, j, nTags;
        unsigned char nBytes;
	CTag	*pTag;
	WCHAR	*pWCHAR;
	VARIANT	PropertyValue;
        int iVal=0;
        bool boolVal=false;
        float fltVal=0.0;
        DWORD dwVal=0;
        BSTR bstrVal=WideString("");

if (OpenDialog1->Execute())
    {
        TFileStream *archive;
        archive = new TFileStream(OpenDialog1->FileName, fmOpenRead);

	for (i=0; i<TagList->Count; i++)
		delete ((CTag *)TagList->Items[i]);
	TagList->Clear();
        ListView1->Items->Clear();
//	archive >> nTags;
//        archive->Read(TmpStr,160/*nBytes*/);
        archive->Read(&nTags,4);
	for (i=0; i<nTags; i++)
		{
		pTag = new (CTag);
//		archive >> TmpStr;
                strset(TmpStr,'\0');
                archive->Read(&nBytes,1);
                archive->Read(TmpStr,nBytes);
		pTag->Name = TmpStr;
//		archive >> TmpStr;
                strset(TmpStr,'\0');
                archive->Read(&nBytes,1);
                archive->Read(TmpStr,nBytes);
		pTag->Description = TmpStr;
//		archive >> TmpStr;
                strset(TmpStr,'\0');
                archive->Read(&nBytes,1);
                archive->Read(TmpStr,nBytes);
		pTag->Units = TmpStr;
                archive->Read(&iVal,2);
		pTag->Value.vt=iVal;
		switch (pTag->Value.vt)
			{
			case VT_I2:
                        archive->Read(&iVal,2);
                        pTag->Value.iVal = iVal;
                        break;
			case VT_BOOL:
                        archive->Read(&boolVal,2);
			pTag->Value.boolVal=boolVal;
                        break;
			case VT_R4:
                        archive->Read(&fltVal,4);
                        pTag->Value.fltVal=fltVal;
                        break;
			case VT_BSTR:
                        strset(TmpStr,'\0');
                        archive->Read(&nBytes,1);
                        archive->Read(TmpStr,nBytes);
                        pWCHAR = WideString(TmpStr);
                        pTag->Value.bstrVal = SysAllocString (pWCHAR);
                        break;
			}
		for (j=0; j<4; j++)
			{
                        fltVal = 0.0;
                        archive->Read(&fltVal, 4);
			pTag->alarms[j]=fltVal;
                        dwVal=0;
                        archive->Read(&dwVal, 4);
			pTag->severity[j]=dwVal;
                        boolVal=false;
                        archive->Read(&boolVal, 4);
			pTag->enabled[j]=boolVal;
			}

		pTag->Handle = CreateTag((pTag->Name).c_str(), pTag->Value, OPC_QUALITY_GOOD, TRUE);
		TagList->Add((TObject *)pTag);
                ListItem = ListView1->Items->Add();
                ListItem->Caption = pTag->Name;
                ListItem->SubItems->Add(VarToStr(pTag->Value));
//                ListItem->SubItems->Strings[0] = VarToStr(pTag->Value);

		VariantInit(&PropertyValue);
		PropertyValue.vt = VT_BSTR;
		pWCHAR = WideString (pTag->Units);
		PropertyValue.bstrVal = SysAllocString (pWCHAR);
		SetTagProperties(pTag->Handle, 100, "EU Units", PropertyValue);
		VariantClear(&PropertyValue);

		PropertyValue.vt = VT_BSTR;
		pWCHAR = WideString (pTag->Description);
		PropertyValue.bstrVal = SysAllocString (pWCHAR);
		SetTagProperties (pTag->Handle, 101, "Item Description", PropertyValue);
		VariantClear(&PropertyValue);

		SetItemLevelAlarm (pTag->Handle, ID_LOLO_LIMIT, pTag->alarms[0], pTag->severity[0], pTag->enabled[0]);
		SetItemLevelAlarm (pTag->Handle, ID_LO_LIMIT, pTag->alarms[1], pTag->severity[1], pTag->enabled[1]);
		SetItemLevelAlarm (pTag->Handle, ID_HI_LIMIT, pTag->alarms[2], pTag->severity[2], pTag->enabled[2]);
		SetItemLevelAlarm (pTag->Handle, ID_HIHI_LIMIT, pTag->alarms[3], pTag->severity[3], pTag->enabled[3]);

		}


	delete(archive);
//	ConfigFile.Close();

}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button2Click(TObject *Sender)
{

	String	FilePath;
        char TmpStr[80]="                      ";
        TListItem  *ListItem;
	int	i, j, nTags;
        unsigned char nBytes;
	CTag	*pTag;
	WCHAR	*pWCHAR;
	VARIANT	PropertyValue;
        int iVal=0;
        bool boolVal=false;
        float fltVal=0.0;
        DWORD dwVal=0;
        String bstrVal="";

if (SaveDialog1->Execute())
    {
        TFileStream *archive;
        archive = new TFileStream(SaveDialog1->FileName, fmCreate);

	nTags = TagList->Count;
//	archive << nTags;
        archive->Write(&nTags,4);
	for (i=0; i<nTags; i++)
		{
                pTag = (CTag *) TagList->Items[i];
//		archive << pTag->Name;
                nBytes = pTag->Name.Length();
                archive->Write(&nBytes,1);
                archive->Write(pTag->Name.c_str(),nBytes);
//		archive << pTag->Description;
                nBytes = pTag->Description.Length();
                archive->Write(&nBytes,1);
                archive->Write(pTag->Description.c_str(),nBytes);
//		archive << pTag->Units;
                nBytes = pTag->Units.Length();
                archive->Write(&nBytes,1);
                archive->Write(pTag->Units.c_str(),nBytes);
//		archive << pTag->Value.vt;
		iVal = pTag->Value.vt;
                archive->Write(&iVal,2);
		switch (pTag->Value.vt)
			{
			case VT_I2:
//				archive << pTag->Value.iVal;
                	iVal = pTag->Value.iVal;
                        archive->Write(&iVal,2);
                        break;
			case VT_BOOL:
//				archive << pTag->Value.boolVal;
                        boolVal = pTag->Value.boolVal;
                        archive->Write(&boolVal,2);
			break;
			case VT_R4:
//			archive << pTag->Value.fltVal;
			fltVal = pTag->Value.fltVal;
                        archive->Write(&fltVal,4);
			break;
			case VT_BSTR:
                        bstrVal= pTag->Value.bstrVal;
                        nBytes = bstrVal.Length();
                        archive->Write(&nBytes,1);
                        archive->Write(bstrVal.c_str(),nBytes);
			break;
			}
		for (j=0; j<4; j++)
			{
//			archive << pTag->alarms[j];
                        fltVal = pTag->alarms[j];
                        archive->Write(&fltVal, 4);
//			archive << pTag->severity[j];
			dwVal=pTag->severity[j];
                        archive->Write(&dwVal, 4);
//			archive << pTag->enabled[j];
                        boolVal=pTag->enabled[j];
                        archive->Write(&boolVal, 4);
			}
		}
	delete(archive);
   }     
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
        TagList = new TList();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button3Click(TObject *Sender)
{
//
// Would generally register the server based on command
// line options during the InitInstance procedure in the App.
//
// This sample uses a menu option to write the registry entries.
//
	String	SvrName, SvrDescr, HelpPath;
	int	i;

	// get path to executable by parsing the help path
	// contained in CWinApp
//	HelpPath = AfxGetApp()->m_pszHelpFilePath;
//	i = HelpPath.ReverseFind('\\');
//	HelpPath = HelpPath.Left(i+1);
//	HelpPath += "WtSvrTest.exe";
   HelpPath = ExtractFilePath(Application->ExeName);
	SvrName = "DLLTestSvr";
	SvrDescr = "OPC Server to Test WT Toolkit";
	::UpdateRegistry((BYTE *)&CLSID_Svr, SvrName.c_str(), SvrDescr.c_str(), HelpPath.c_str());
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button10Click(TObject *Sender)
{
	String	SvrName;

	SvrName = "DLLTestSvr";
	UnregisterServer ((BYTE *)&CLSID_Svr, SvrName.c_str());

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{
// Add Item
	CTag	*pTag;
        TListItem  *ListItem;
//	CTagDlg	dlg;
	int	i, j, ntags;
	VARIANT	PropertyValue;
	WCHAR	*pWCHAR;
//	CWTSvrTestApp	*pApp;
	String buf;

//	pApp = (CWTSvrTestApp *)AfxGetApp();

	for (i=0; i<4; i++)
		{
		TagForm->m_alarms[i] = 0.0;
		TagForm->m_severity[i] = 200;
		TagForm->m_enabled[i] = FALSE;
		}
	TagForm->m_alarms[2] = 100.0;
	TagForm->m_alarms[3] = 100.0;
	TagForm->m_value.vt = VT_R4;
	TagForm->m_value.fltVal = 0.0;
	TagForm->m_x100 = FALSE;

	if (TagForm->ShowModal() != IDOK)
		return;

	if (TagForm->m_x100)
		ntags = 100;
	else
		ntags = 1;

	for (j=0; j<ntags; j++)
		{
		pTag = new CTag;
		pTag->Name = TagForm->m_name;
		if (TagForm->m_x100)
			{
			buf = ".TAG" + String(j+1);
			pTag->Name += buf;
			}
		pTag->Description = TagForm->m_description;
		pTag->Units = TagForm->m_units;
		VariantCopy (&(pTag->Value), &TagForm->m_value);
		CoFileTimeNow(&pTag->Time);
		for (i=0; i<4; i++)
			{
			pTag->alarms[i] = TagForm->m_alarms[i];
			pTag->severity[i] = TagForm->m_severity[i];
			pTag->enabled[i] = TagForm->m_enabled[i];
			}

		pTag->Handle = CreateTag(pTag->Name.c_str(), pTag->Value, OPC_QUALITY_GOOD, TRUE);

		TagList->Add ((TObject *)pTag);
                buf = pTag->Name;
                ListItem = ListView1->Items->Add();
                ListItem->Caption = pTag->Name;
                ListItem->SubItems->Add(VarToStr(pTag->Value));
//                ListItem->SubItems->Strings[0] = VarToStr(pTag->Value);

		VariantInit(&PropertyValue);
		PropertyValue.vt = VT_BSTR;
		pWCHAR = WideString (pTag->Units);
		PropertyValue.bstrVal = SysAllocString (pWCHAR);
//		pApp->WSTRFree (pWCHAR);
		SetTagProperties (pTag->Handle, 100, "EU Units", PropertyValue);
		VariantClear(&PropertyValue);

		PropertyValue.vt = VT_BSTR;
		pWCHAR = WideString (pTag->Description);
		PropertyValue.bstrVal = SysAllocString (pWCHAR);
//		pApp->WSTRFree (pWCHAR);
		SetTagProperties (pTag->Handle, 101, "Item Description", PropertyValue);
		VariantClear(&PropertyValue);

		SetItemLevelAlarm (pTag->Handle, ID_LOLO_LIMIT, pTag->alarms[0], pTag->severity[0], pTag->enabled[0]);
		SetItemLevelAlarm (pTag->Handle, ID_LO_LIMIT, pTag->alarms[1], pTag->severity[1], pTag->enabled[1]);
		SetItemLevelAlarm (pTag->Handle, ID_HI_LIMIT, pTag->alarms[2], pTag->severity[2], pTag->enabled[2]);
		SetItemLevelAlarm (pTag->Handle, ID_HIHI_LIMIT, pTag->alarms[3], pTag->severity[3], pTag->enabled[3]);
		}

	VariantClear (&TagForm->m_value);
//	Invalidate(TRUE);

}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormShow(TObject *Sender)
{
	// CLSID_Svr identifies this Server
	// (Created with guidgen.exe)
	// minimum server refresh rate is 250 ms
	InitWTOPCsvr ((BYTE *)&CLSID_Svr,250);

///	EnableWriteNotification (&WriteNotifyProc, TRUE);
///	EnableDisconnectNotification (&DisconnectProc);
///	EnableDeviceRead (&DeviceReadProc);

	//
	// If you want to override the default A&E operation of WtOPCsvr.dll
	// define your own callback object and overload the functions of interest.
//	pCallback = new (CMyAECallback);
//	SetAEServerCallback (TRUE, pCallback);

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button5Click(TObject *Sender)
{
	CTag	*pTag;
        SelectedIndex = ListView1->Selected->Index;
        pTag = (CTag *) TagList->Items[SelectedIndex];
	if (pTag == NULL)
		return;

	TagList->Delete(SelectedIndex);
        ListView1->Items->Delete(SelectedIndex);
	RemoveTag (pTag->Handle);
	delete (pTag);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button6Click(TObject *Sender)
{
//Update Tag
	CTag	*pTag;
	VARIANT	PropertyValue;
	int	i;
	WCHAR	*pWCHAR;
        SelectedIndex = ListView1->Selected->Index;
        pTag = (CTag *) TagList->Items[SelectedIndex];
	if (pTag == NULL)
		return;

	TagForm->m_name = pTag->Name;
	TagForm->m_description = pTag->Description;
	TagForm->m_units = pTag->Units;
	VariantCopy (&TagForm->m_value, &(pTag->Value));
	for (i=0; i<4; i++)
		{
		TagForm->m_alarms[i] = pTag->alarms[i];
		TagForm->m_severity[i] = pTag->severity[i];
		TagForm->m_enabled[i] = pTag->enabled[i];
		}
	if (TagForm->ShowModal() != IDOK)
		return;

	VariantCopy (&(pTag->Value), &TagForm->m_value);
	pTag->Description = TagForm->m_description;
	pTag->Units = TagForm->m_units;
	CoFileTimeNow(&pTag->Time);
	for (i=0; i<4; i++)
		{
		pTag->alarms[i] = TagForm->m_alarms[i];
		pTag->severity[i] = TagForm->m_severity[i];
		pTag->enabled[i] = TagForm->m_enabled[i];
		}

	UpdateTag (pTag->Handle, pTag->Value, OPC_QUALITY_GOOD);
        TListItem *item = ListView1->Items->Item[SelectedIndex];
        item->SubItems->Strings[0] = VarToStr(pTag->Value);

	VariantInit(&PropertyValue);
	PropertyValue.vt = VT_BSTR;
	pWCHAR = WideString (pTag->Units);
	PropertyValue.bstrVal = SysAllocString (pWCHAR);
	SetTagProperties (pTag->Handle, 100, "EU Units", PropertyValue);
	VariantClear(&PropertyValue);

	PropertyValue.vt = VT_BSTR;
	pWCHAR = WideString (pTag->Description);
	PropertyValue.bstrVal = SysAllocString (pWCHAR);
	SetTagProperties (pTag->Handle, 101, "Item Description", PropertyValue);
	VariantClear(&PropertyValue);

	SetItemLevelAlarm (pTag->Handle, ID_LOLO_LIMIT, pTag->alarms[0], pTag->severity[0], pTag->enabled[0]);
	SetItemLevelAlarm (pTag->Handle, ID_LO_LIMIT, pTag->alarms[1], pTag->severity[1], pTag->enabled[1]);
	SetItemLevelAlarm (pTag->Handle, ID_HI_LIMIT, pTag->alarms[2], pTag->severity[2], pTag->enabled[2]);
	SetItemLevelAlarm (pTag->Handle, ID_HIHI_LIMIT, pTag->alarms[3], pTag->severity[3], pTag->enabled[3]);

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button7Click(TObject *Sender)
{
//Dynamic Tags
	EnableUnknownItemNotification (UnknownTagProc);
	EnableItemRemovalNotification (TagRemovedProc);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button8Click(TObject *Sender)
{
//Force Client Refresh
        RefreshAllClients();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button9Click(TObject *Sender)
{
// User AE Message
//	ONEVENTSTRUCT	ev;
	SYSTEMTIME	SystemTime;
	FILETIME	TimeStamp;

	GetSystemTime(&SystemTime);		// Get current UTC Time
	SystemTimeToFileTime(&SystemTime, &TimeStamp); // and store it

	AEMessForm->m_severity = 200;
	if (AEMessForm->ShowModal() == IDOK)
		{
/*		ev.wChangeMask = 0;
		ev.wNewState = 0;
		ev.szSource = pApp->WSTRFromCString("User Message");
		ev.ftTime = TimeStamp;
		ev.szMessage = pApp->WSTRFromCString(dlg.m_msg);
		ev.dwEventType = OPC_SIMPLE_EVENT;
		ev.dwEventCategory = 1;
		ev.dwSeverity = dlg.m_severity;
		ev.szConditionName = pApp->WSTRFromCString("User Defined Condition");
		ev.szSubconditionName = pApp->WSTRFromCString("User Defined Subcondition");
		ev.wQuality = OPC_QUALITY_GOOD;
		ev.bAckRequired = FALSE;
		ev.ftActiveTime = TimeStamp;
		ev.dwCookie = 1;
		ev.dwNumEventAttrs = 0;
		ev.pEventAttributes = NULL;
		ev.szActorID = pApp->WSTRFromCString("");

		UserAEMessageEx (ev);	*/
		UserAEMessage(AEMessForm->m_msg.c_str(), AEMessForm->m_severity);
		}

}
//---------------------------------------------------------------------------
void __fastcall TForm1::UnknownTagHandler(LPSTR Path, LPSTR Name)
{
	CTag	*pTag;
	char buf[200];

	pTag = new (CTag);
	strncpy (buf, Path, 200);
	pTag->Name = buf;
	if (pTag->Name.Length() > 0)
		pTag->Name += ".";
	strncpy (buf, Name, 200);
	pTag->Name += buf;
	CoFileTimeNow(&pTag->Time);

	pTag->Handle = CreateTag (pTag->Name.c_str(), pTag->Value, OPC_QUALITY_GOOD, TRUE);
	TagList->Add((TObject *)pTag);

}
//---------------------------------------------------------------------------
void __fastcall TForm1::TagRemovedHandler(HANDLE hTag, LPSTR Path, LPSTR Name)
{
	CTag	*pTag;
	int	i;

	for (i=0; i<TagList->Count; i++)
		{
                pTag = (CTag *) TagList->Items[i];
		if (pTag->Handle == hTag)
			{
			TagList->Delete(i);
			delete (pTag);
			RemoveTag (hTag);
			SelectedIndex = -1;
			}
		}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::NotificationHandler(HANDLE Handle, VARIANT *pNewValue, DWORD *pDeviceError)
{
	int	i;
	CTag	*pTag;

	for (i=0; i<TagList->Count; i++)
		{
                pTag = (CTag *) TagList->Items[i];
		if (pTag->Handle == Handle)
			{
			VariantCopy (&(pTag->Value), pNewValue);
			UpdateTag (pTag->Handle, pTag->Value, OPC_QUALITY_GOOD);
//			Invalidate(TRUE);
			*pDeviceError = S_OK;
			return;
			}
		}
	*pDeviceError = OPC_E_INVALIDHANDLE;
}

//---------------------------------------------------------------------------
void __fastcall TForm1::DeviceReadHandler(HANDLE Handle, VARIANT *pNewValue, WORD *pQuality, FILETIME *pTimestamp)
{
	int	i;
	CTag	*pTag;
	SYSTEMTIME	systime;

	for (i=0; i<TagList->Count; i++)
		{
                pTag = (CTag *) TagList->Items[i];
		if (pTag->Handle == Handle)
			{
			VariantCopy  (pNewValue, &(pTag->Value));
			*pQuality = OPC_QUALITY_GOOD;
			GetSystemTime (&systime);
			SystemTimeToFileTime (&systime, pTimestamp);
			return;
			}
		}
	*pQuality = OPC_QUALITY_BAD;

}
//---------------------------------------------------------------------------
void CALLBACK UnknownTagProc(LPSTR Path, LPSTR Name)
{
	Form1->UnknownTagHandler(Path, Name);
}
//---------------------------------------------------------------------------
void CALLBACK TagRemovedProc(HANDLE hTag, LPSTR Path, LPSTR Name)
{
	Form1->TagRemovedHandler(hTag, Path, Name);
}
//---------------------------------------------------------------------------
void CALLBACK WriteNotifyProc(HANDLE Handle, VARIANT *pNewValue, DWORD *pDevError)
{
	Form1->NotificationHandler(Handle, pNewValue, pDevError);

}
//---------------------------------------------------------------------------
void CALLBACK DeviceReadProc(HANDLE Handle, VARIANT *pNewValue, WORD *pQuality, FILETIME *pTimestamp)
{
	Form1->DeviceReadHandler(Handle, pNewValue, pQuality, pTimestamp);

}
//---------------------------------------------------------------------------
void CALLBACK DisconnectProc(DWORD NumbrActiveClients)
{
	if (NumbrActiveClients == 0)
		{
		// If you want to end the server application
		// when the last client disconnects,
		// be sure to return from the callback
		// to release the client before shutting down.
		}
}
//---------------------------------------------------------------------------

