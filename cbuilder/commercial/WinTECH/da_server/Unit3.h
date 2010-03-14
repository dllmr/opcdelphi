//---------------------------------------------------------------------------
#ifndef Unit3H
#define Unit3H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
//---------------------------------------------------------------------------
class TAEMessForm : public TForm
{
__published:	// IDE-managed Components
        TEdit *Edit1;
        TEdit *Edit2;
        TButton *Button1;
        TButton *Button2;
private:	// User declarations
public:		// User declarations
        __fastcall TAEMessForm(TComponent* Owner);
        String	m_msg;
	int	m_severity;

};
//---------------------------------------------------------------------------
extern PACKAGE TAEMessForm *AEMessForm;
//---------------------------------------------------------------------------
#endif
