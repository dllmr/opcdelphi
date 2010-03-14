//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("aeProject.res");
USEFORM("aeclient.cpp", Form1);
USELIB("..\clienttst\wtlient.lib");
USEUNIT("..\clienttst\Unit2.cpp");
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
