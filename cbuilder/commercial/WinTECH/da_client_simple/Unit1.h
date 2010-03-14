//---------------------------------------------------------------------------
#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ComCtrls.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TBitBtn *BitBtn1;
        TListBox *ListBox1;
        TBitBtn *BitBtn2;
        TListBox *ListBox2;
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TStatusBar *StatusBar1;
        TEdit *Edit1;
        TListView *ListView1;
        TBitBtn *BitBtn5;
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall ListBox1Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall BitBtn5Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
private:	// User declarations

	String	MachineName, ServerName;
	HANDLE	ConnectionHandle;
        HANDLE	hGroup;
        TList *ItemList;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
        void __fastcall OPCDataUpdate(HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality);
        void __fastcall UpdateDataView(int index);

};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
