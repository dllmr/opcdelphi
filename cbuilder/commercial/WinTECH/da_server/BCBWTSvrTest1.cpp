//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("BCBWTSvrTest1.res");
USEFORM("Unit1.cpp", Form1);
USELIB("wtopcsvr.lib");
USEFORM("Unit2.cpp", TagForm);
USEFORM("Unit3.cpp", AEMessForm);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TTagForm), &TagForm);
                 Application->CreateForm(__classid(TAEMessForm), &AEMessForm);
                 Application->Run();
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        return 0;
}
//---------------------------------------------------------------------------
