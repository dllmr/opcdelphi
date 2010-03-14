//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("ProjectData.res");
USEFORM("UnitDataBound.cpp", Form1);
USEUNIT("..\..\Imports\SoftingAxC_TLB.cpp");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->Run();
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        return 0;
}
//---------------------------------------------------------------------------
