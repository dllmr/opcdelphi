//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
program DataAccess20Demo;

uses
  Forms,
  SysUtils,
  Main in 'Main.pas' {OPCForm},
  uOPCDemo in 'uOPCDemo.pas',
  sOPC_TLB in 'sOPC_TLB.pas',
  uTimer in '..\Source\uTimer.pas',
  comcat in '..\Source\comcat.pas',
  uGlobals in '..\Source\uGlobals.pas',
  uLogging in '..\Source\uLogging.pas',
  uMessageDlg in '..\Source\uMessageDlg.pas',
  uOPC in '..\Source\uOPC.pas',
  uOPCAsyncIO2 in '..\Source\uOPCAsyncIO2.pas',
  uOPCBrowse in '..\Source\uOPCBrowse.pas',
  uOPCDataAccess in '..\Source\uOPCDataAccess.pas',
  uOPCGroup in '..\Source\uOPCGroup.pas',
  uOPCItem in '..\Source\uOPCItem.pas',
  uOPCItemEnumerator in '..\Source\uOPCItemEnumerator.pas',
  uOPCNode in '..\Source\uOPCNode.pas',
  uOPCStringEnumerator in '..\Source\uOPCStringEnumerator.pas',
  uRegister in '..\Source\uRegister.pas';

{$R *.TLB}
{$R *.RES}

const
  ServerName = 'sOPC.DemoDataAccessServer20';
  ServerDescription = 'sOPC Demo OPC Data Access 2.0 Server';

begin
  // init OPC
  {$ifdef Debug}
  OPC.Init(ServerName, ServerDescription, true, true);
  {$else}
  OPC.Init(ServerName, ServerDescription);
  {$endif}
  OPC.OnRead := OnRead;
  OPC.OnWrite := OnWrite;
  OPC.OnInitAddressSpace := OnInitAddressSpace;
  // init Application
  Application.Initialize;
  Application.Title := 'Demo Data Access Server v2.0';
  Application.CreateForm(TOPCForm, OPCForm);
  OPC.Start;
  Application.Run;
end.

