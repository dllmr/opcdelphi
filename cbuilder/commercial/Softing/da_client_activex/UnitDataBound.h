//---------------------------------------------------------------------------
#ifndef UnitDataBoundH
#define UnitDataBoundH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "GWXGAUGELib_OCX.h"
#include "SoftingAxC_OCX.h"
#include <Buttons.hpp>
#include <OleCtrls.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TBitBtn *BitBtn1;
        TEdit *Edit1;
        TOPCDataControl *OPCDataControl1;
        TEdit *Edit2;
        TEdit *Edit3;
        TEdit *Edit4;
        TEdit *Edit5;
        TEdit *Edit6;
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall OPCDataControl1DataChanged(TObject *Sender,
          short bOkay, OPCDataItemsPtr ChangedItems);
        void __fastcall OPCDataControl1Connect(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
