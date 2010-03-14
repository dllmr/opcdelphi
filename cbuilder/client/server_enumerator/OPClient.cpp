//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
USERES("OPClient.res");
USEUNIT("opc\Opcda_i.c");
USEUNIT("opc\Opccomn_i.c");
USEUNIT("DataCallbackSink.cpp");
USEUNIT("atlprj.cpp");
USEUNIT("opc\opcda_cats.c");
USEUNIT("opc\opcenum_clsid.c");
USEFORM("Main.cpp", MainForm);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
   try
   {
       Application->Initialize();
       Application->CreateForm(__classid(TMainForm), &MainForm);
                 Application->Run();
   }
   catch (Exception &exception)
   {
       Application->ShowException(&exception);
   }
   return 0;
}
//---------------------------------------------------------------------------
