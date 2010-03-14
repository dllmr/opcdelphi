//---------------------------------------------------------------------------
#ifndef aeclientH
#define aeclientH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TAEMsg : public TObject
{
public:

	String	Source, Message;
	DWORD	Severity;
	FILETIME TimeStamp;
};

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
        TTimer *Timer1;
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall ListBox1Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall Timer1Timer(TObject *Sender);
protected:
        void __fastcall OnShutdownRequest(TMessage& Msg);

BEGIN_MESSAGE_MAP
  MESSAGE_HANDLER(OPC_SHUTDOWN_MSG, TMessage, OnShutdownRequest)
END_MESSAGE_MAP(TForm)

private:	// User declarations

	String	MachineName, ServerName;
	HANDLE	ConnectionHandle;
        HANDLE	hGroup;
        TList *AEMsgArray;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
        void __fastcall NewAEMsg(char *pSource, FILETIME timestamp, char *pMsg, DWORD severity);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
