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
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TStatusBar *StatusBar1;
        TEdit *Edit1;
        TListView *ListView1;
        TBitBtn *BitBtn5;
        TTreeView *TreeView1;
        TBitBtn *BitBtn6;
        TEdit *Edit2;
        TListView *ListView2;
        TBitBtn *BitBtn7;
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall ListBox1Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall BitBtn5Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall TreeView1Expanded(TObject *Sender,
          TTreeNode *Node);
        void __fastcall BitBtn6Click(TObject *Sender);
        void __fastcall BitBtn7Click(TObject *Sender);
private:	// User declarations

	String	MachineName, ServerName;
	HANDLE	ConnectionHandle;
        HANDLE	hGroup;
        TList *ItemList;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
        void __fastcall OPCDataUpdate(HANDLE hGroup, HANDLE hItem, VARIANT *pVar, FILETIME timestamp, DWORD quality);
        void __fastcall UpdateDataView(int index);
        void __fastcall ResolveOneLevel (TTreeNode *hItem);
        String __fastcall QualifiedName (TTreeNode *hItem);
        String __fastcall UnqualifyName (String Name);

};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
