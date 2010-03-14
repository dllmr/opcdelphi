//---------------------------------------------------------------------------
#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <Dialogs.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TButton *Button1;
        TButton *Button2;
        TButton *Button3;
        TButton *Button4;
        TButton *Button5;
        TButton *Button6;
        TButton *Button7;
        TButton *Button8;
        TButton *Button9;
        TButton *Button10;
        TStatusBar *StatusBar1;
        TOpenDialog *OpenDialog1;
        TSaveDialog *SaveDialog1;
        TListView *ListView1;
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button10Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall Button6Click(TObject *Sender);
        void __fastcall Button7Click(TObject *Sender);
        void __fastcall Button8Click(TObject *Sender);
        void __fastcall Button9Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
private:	// User declarations
        TList *TagList;
        int	SelectedIndex;
public:		// User declarations
        void __fastcall UnknownTagHandler(LPSTR Path, LPSTR Name);
        void __fastcall TagRemovedHandler(HANDLE hTag, LPSTR Path, LPSTR Name);
        void __fastcall NotificationHandler(HANDLE Handle, VARIANT *pNewValue, DWORD *pDeviceError);
        void __fastcall DeviceReadHandler(HANDLE Handle, VARIANT *pNewValue, WORD *pQuality, FILETIME *pTimestamp);
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
class	CTag : public TObject
{
public:
	__fastcall CTag(void);
	__fastcall ~CTag(void);

	HANDLE	Handle;
	FILETIME	Time;
	String Name, Description, Units;
	VARIANT	Value;

// lolo, lo, hi, & hihi alarms
	float	alarms[4];
	DWORD	severity[4];
	BOOL	enabled[4];

};

#endif
