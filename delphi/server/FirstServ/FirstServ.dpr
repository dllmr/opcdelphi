program FirstServ;

uses
  Forms,
  Main in 'Main.pas' {Form1},
  ServIMPL in 'ServIMPL.pas',
  RegDeRegServer in 'RegDeRegServer.pas',
  comcat in 'comcat.pas',
  Enumstring in 'Enumstring.pas',
  GroupUnit in 'GroupUnit.pas',
  ItemPropIMPL in 'ItemPropIMPL.pas',
  Globals in 'Globals.pas',
  ItemsUnit in 'ItemsUnit.pas',
  ItemAttributesOPC in 'ItemAttributesOPC.pas',
  EnumItemAtt in 'EnumItemAtt.pas',
  AsyncUnit in 'AsyncUnit.pas',
  ShutDownRequest in 'ShutDownRequest.pas' {ShutDownDlg},
  OPCErrorStrings in 'OPCErrorStrings.pas',
  EnumUnknown in 'EnumUnknown.pas',
  FirstServ_TLB in 'FirstServ_TLB.pas',
  OPCRemovedGroupUnit in 'OPCRemovedGroupUnit.pas',
  SysUtils,
  OPC_AE in 'OPC_Defines\OPC_AE.pas',
  OPCCOMN in 'OPC_Defines\OPCCOMN.pas',
  OPCDA in 'OPC_Defines\OPCDA.pas',
  OPCerror in 'OPC_Defines\OPCerror.pas',
  OPCHDA in 'OPC_Defines\OPCHDA.pas',
  OPCSEC in 'OPC_Defines\OPCSEC.pas',
  OPCtypes in 'OPC_Defines\OPCtypes.pas',
  IOPCCommonUnit in 'IOPCCommonUnit.pas',
  DataPointsUnit in 'DataPointsUnit.pas';

{$R *.TLB}

{$R *.RES}

begin
 if FindCmdLineSwitch('regserver',['-', '/'],true) then
  RegisterTheServer(serverName)
 else if FindCmdLineSwitch('unregserver',['-', '/'],true) then
  UnRegisterTheServer(serverName);

  Application.Initialize;
  Application.Title := 'MRD';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
