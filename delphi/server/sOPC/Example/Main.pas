unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, AxCtrls, ExtCtrls, Mask;

type
  TOPCForm = class(TForm)
    laDateTime: TLabel;
    paServer: TGroupBox;
    buShutDown: TButton;
    paClients: TPanel;
    laClients: TLabel;
    laClientCount: TLabel;
    laGroupCount: TLabel;
    laGroup: TLabel;
    buClose: TButton;
    Timer: TTimer;
    paPath: TPanel;
    paIncrement: TPanel;
    cbAuto: TCheckBox;
    buIncrement: TButton;
    edIncrement: TEdit;
    buReset: TButton;
    laValue: TLabel;
    IncTimer: TTimer;

    procedure buCloseClick(Sender: TObject);
    procedure buShutDownClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure buIncrementClick(Sender: TObject);
    procedure buResetClick(Sender: TObject);
    procedure IncTimerTimer(Sender: TObject);
  end;

var
  OPCForm: TOPCForm;

implementation

{$R *.DFM}

uses
  uOPC, uOPCDemo;

procedure PatchINT3; 
var
  NOP : Byte; 
  NTDLL: THandle; 
  BytesWritten: DWORD; 
  Address: Pointer; 
begin 
  if Win32Platform <> VER_PLATFORM_WIN32_NT then Exit; 
  NTDLL := GetModuleHandle('NTDLL.DLL'); 
  if NTDLL = 0 then Exit; 
  Address := GetProcAddress(NTDLL, 'DbgBreakPoint'); 
  if Address = nil then Exit; 
  try 
    if Char(Address^) <> #$CC then Exit; 

    NOP := $90; 
    if WriteProcessMemory(GetCurrentProcess, Address, @NOP, 1, BytesWritten) and 
      (BytesWritten = 1) then 
      FlushInstructionCache(GetCurrentProcess, Address, 1); 
  except 
    //Do not panic if you see an EAccessViolation here, it is perfectly harmless! 
    on EAccessViolation do ; 
    else raise; 
  end; 
end; 

procedure TOPCForm.buCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TOPCForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  OPC.ShutDown;
end;

procedure TOPCForm.buShutDownClick(Sender: TObject);
begin
  OPC.ShutDown;
end;

procedure TOPCForm.TimerTimer(Sender: TObject);
begin
  laDateTime.Caption := DateTimeToStr(Now);
  laClientCount.Caption := IntToStr(OPC.GetServerCount);
  laGroupCount.Caption := IntToStr(OPC.GetGroupCount);
  edIncrement.Text := IntToStr(uOPCDemo.Values[102]);
end;

procedure TOPCForm.FormCreate(Sender: TObject);
begin
  {$ifdef Debug}
  PatchINT3;
  {$endif}
  paPath.Caption := GetCurrentDir;
end;

procedure TOPCForm.buIncrementClick(Sender: TObject);
var
  i: integer;
begin
  for i := Low(Values) to High(Values) do inc(Values[i]);
end;

procedure TOPCForm.buResetClick(Sender: TObject);
begin
  ResetValues;
end;

procedure TOPCForm.IncTimerTimer(Sender: TObject);
begin
  if cbAuto.Checked then IncrementValues;
end;

end.

