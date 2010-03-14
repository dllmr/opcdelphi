//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Unit2.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TTagForm *TagForm;
//---------------------------------------------------------------------------
__fastcall TTagForm::TTagForm(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TTagForm::RadioGroup1Click(TObject *Sender)
{
 if (RadioGroup1->ItemIndex == 0)  m_value.vt = VT_R4;
 else if (RadioGroup1->ItemIndex == 1)  m_value.vt = VT_I2;
 else if (RadioGroup1->ItemIndex == 2)  m_value.vt = VT_BOOL;
 else if (RadioGroup1->ItemIndex == 3)  m_value.vt = VT_BSTR;


}
//---------------------------------------------------------------------------
void __fastcall TTagForm::FormShow(TObject *Sender)
{
	switch (m_value.vt)
		{
		case VT_I2:	RadioGroup1->ItemIndex = 1;
				break;
		case VT_BOOL: RadioGroup1->ItemIndex = 2;
				break;
		case VT_BSTR: RadioGroup1->ItemIndex = 3;
				break;
		default: RadioGroup1->ItemIndex = 0;
				break;
		}
        Edit1->Text=m_name;
        Edit3->Text=m_description;
        Edit4->Text=m_units;
        Edit2->Text = VarToStr(m_value);
        CheckBox1->Checked=m_enabled[3];
        CheckBox2->Checked=m_enabled[2];
        CheckBox3->Checked=m_enabled[1];
        CheckBox4->Checked=m_enabled[0];

        Edit5->Text=FloatToStr(m_alarms[3]);
        Edit7->Text=FloatToStr(m_alarms[2]);
        Edit9->Text=FloatToStr(m_alarms[1]);
        Edit11->Text=FloatToStr(m_alarms[0]);

        Edit6->Text=IntToStr(m_severity[3]);
        Edit8->Text=IntToStr(m_severity[2]);
        Edit10->Text=IntToStr(m_severity[1]);
        Edit12->Text=IntToStr(m_severity[0]);
}
//---------------------------------------------------------------------------
void __fastcall TTagForm::RadioButton1Click(TObject *Sender)
{
	m_x100 = !m_x100;
}
//---------------------------------------------------------------------------
void __fastcall TTagForm::FormHide(TObject *Sender)
{
	String	buf;
//	CEdit	*pEdit;
	int	tempi=0;
	float	tempf=0.0;
	WCHAR	*pWCHAR;
//	CWTSvrTestApp	*pApp;

//	pApp = (CWTSvrTestApp *)AfxGetApp();

//	pEdit = (CEdit *)GetDlgItem(IDC_VALUE);
        m_name=Edit1->Text;
        m_description=Edit3->Text;
        m_units=Edit4->Text;
	buf=Edit2->Text;
	switch (m_value.vt)
		{
		case VT_I2:
		     m_value.intVal = StrToInt(buf);
		break;
		case VT_BOOL: if (buf == "true")
				  m_value.boolVal = TRUE;
			      else
				  {
				  tempi = StrToInt(buf);
				  if (tempi > 0)
				    m_value.boolVal = TRUE;
				  else
				    m_value.boolVal = FALSE;
				  }
				break;
		case VT_BSTR: 	pWCHAR = WideString (buf);
				m_value.bstrVal = SysAllocString (pWCHAR);
//					pApp->WSTRFree (pWCHAR);
				break;
		default: tempf= StrToFloat(buf);
			 m_value.fltVal = tempf;
			 break;
		}
		m_enabled[3] = CheckBox1->Checked;
		m_enabled[2] = CheckBox2->Checked;
		m_enabled[1] = CheckBox3->Checked;
		m_enabled[0] = CheckBox4->Checked;

		m_alarms[3] = StrToFloat(Edit5->Text);
		m_alarms[2] = StrToFloat(Edit7->Text);
		m_alarms[1] = StrToFloat(Edit9->Text);
		m_alarms[0] = StrToFloat(Edit11->Text);

		m_severity[3] = StrToInt(Edit6->Text);
		m_severity[2] = StrToInt(Edit8->Text);
		m_severity[1] = StrToInt(Edit10->Text);
		m_severity[0] = StrToInt(Edit12->Text);

}
//---------------------------------------------------------------------------

