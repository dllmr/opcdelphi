//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Unit2.h"
#include "opcda.h"
#include "opc_ae.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
__fastcall TItemDef::TItemDef(void) : TObject()
{
	VariantInit(&Value);
	CoFileTimeNow(&TimeStamp);
	Quality = OPC_QUALITY_BAD;
}

__fastcall TItemDef::~TItemDef(void)
{
	VariantClear(&Value);
}
