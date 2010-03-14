//---------------------------------------------------------------------------
#ifndef Unit2H
#define Unit2H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TTagForm : public TForm
{
__published:	// IDE-managed Components
        TRadioGroup *RadioGroup1;
        TEdit *Edit1;
        TEdit *Edit2;
        TEdit *Edit3;
        TEdit *Edit4;
        TLabel *Label1;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TRadioButton *RadioButton1;
        TGroupBox *GroupBox1;
        TCheckBox *CheckBox1;
        TEdit *Edit5;
        TEdit *Edit6;
        TGroupBox *GroupBox2;
        TCheckBox *CheckBox2;
        TEdit *Edit7;
        TEdit *Edit8;
        TGroupBox *GroupBox3;
        TCheckBox *CheckBox3;
        TEdit *Edit9;
        TEdit *Edit10;
        TGroupBox *GroupBox4;
        TCheckBox *CheckBox4;
        TEdit *Edit11;
        TEdit *Edit12;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *Label7;
        TButton *Button1;
        TButton *Button2;
        void __fastcall RadioGroup1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall RadioButton1Click(TObject *Sender);
        void __fastcall FormHide(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TTagForm(TComponent* Owner);
        VARIANT	m_value;
	BOOL	m_x100;

	String	m_name;
	String	m_description;
	String	m_units;

	float	m_alarms[4];
	DWORD	m_severity[4];
	BOOL	m_enabled[4];

};
//---------------------------------------------------------------------------
extern PACKAGE TTagForm *TagForm;
//---------------------------------------------------------------------------
#endif
