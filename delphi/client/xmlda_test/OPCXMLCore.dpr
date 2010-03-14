program OPCXMLCore;

uses
  Forms,
  main in 'main.pas' {Form1},
  MSHTTPPost in 'MSHTTPPost.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
