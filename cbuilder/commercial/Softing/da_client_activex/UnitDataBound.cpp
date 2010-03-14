//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "UnitDataBound.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "GWXGAUGELib_OCX"
#pragma link "SoftingAxC_OCX"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormActivate(TObject *Sender)
{
OPCDataControl1->Connect();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{
OPCDataControl1->Update();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OPCDataControl1DataChanged(TObject *Sender,
      short bOkay, OPCDataItemsPtr ChangedItems)
{
IOPCDataItemsDisp items(ChangedItems);
//items.m_bAutoRelease = FALSE; // The ctor does not AddRef, so it should not
																// call Release in dtor either

long nCnt = items->get_Count();
//        Caption += String(nCnt);
String strText;

for ( long i = 0 ; i < nCnt ; i++ )
 {
	TVariant vIndex(i);
	IOPCDataItemDisp item = items->get_Item(vIndex);

	TVariant vValue = item->get_Value();
	vValue.ChangeType(VT_BSTR);

	TDate vDate = item->get_Time();

	short vQuality = item->get_Quality();
        long QualityBit = (vQuality & 0xC0) >> 6;
        String QualityAsString;
        switch(QualityBit)
        {
        case 0:
                QualityAsString = "(bad)";
        break;
        case 1:
                QualityAsString = "(uncertain)";
        break;
        case 2:
                QualityAsString = "(N/A)";
        break;
        case 3:
                QualityAsString = "(good)";
        break;
        }

	strText = item->get_ItemID();
        strText += ":-->" + VarToStr(vValue) + "-->" + vDate.DateString() + "-->"
        + QualityAsString;
        switch (i)
        {
        case 0:	Edit1->Text=strText;break;
        case 1:	Edit2->Text=strText;break;
        case 2:	Edit3->Text=strText;break;
        case 3:	Edit4->Text=strText;break;
        case 4:	Edit5->Text=strText;break;
        case 5:	Edit6->Text=strText;break;
        }
	}

}
//---------------------------------------------------------------------------
void __fastcall TForm1::OPCDataControl1Connect(TObject *Sender)
{
Caption = "Connected";        
}
//---------------------------------------------------------------------------
