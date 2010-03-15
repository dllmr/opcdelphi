//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uMessageDlg;

//******************************************************************************
interface

uses
  Dialogs, Classes;

//******************************************************************************

// call Message Dialog within Thread
function sMessageDlg(const Msg: string; AType: TMsgDlgType;
  AButtons: TMsgDlgButtons; HelpCtx: Longint): Word;

// call VCL method within Thread
procedure sSynchronize(Method: TThreadMethod);

//******************************************************************************
implementation

//******************************************************************************
type
  // Thread for synchronize calls
  sSyncThread = class(TThread)
  public
    constructor Create;
    procedure Execute; override;
    procedure sSynchronize(Method: TThreadMethod);
  end;

  sMessage = class
  public
    Msg: string;
    AType: TMsgDlgType;
    AButtons: TMsgDlgButtons;
    HelpCtx: Longint;
    ResultCode: Word;
    procedure MessageDlgShow;
  end;

//******************************************************************************
var
  SyncThread: sSyncThread;

//******************************************************************************
constructor sSyncThread.Create;
begin
  inherited Create(True);              // True -> create suspended
end;

procedure sSyncThread.Execute;
begin
end;

procedure sSyncThread.sSynchronize(Method: TThreadMethod);
// Synchronize is protected!
begin
  Synchronize(Method);
end;

//******************************************************************************
procedure sMessage.MessageDlgShow;
begin
  ResultCode := MessageDlg(Msg, AType, AButtons, HelpCtx);
end;

function sMessageDlg(const Msg: string; AType: TMsgDlgType;
                    AButtons: TMsgDlgButtons; HelpCtx: Longint): Word;
var
  m: sMessage;
begin
  m          := sMessage.Create;
  m.Msg      := Msg;
  m.AType    := AType;
  m.AButtons := AButtons;
  m.HelpCtx  := HelpCtx;
  sSynchronize(m.MessageDlgShow);
  Result := m.ResultCode;
  m.Free;
end;

//******************************************************************************
procedure sSynchronize(Method: TThreadMethod);
begin
  if SyncThread = nil then SyncThread := sSyncThread.Create;
  SyncThread.sSynchronize(Method);
end;

//******************************************************************************
initialization
  SyncThread := nil;

finalization
  if SyncThread <> nil then begin
    SyncThread.Free;
    SyncThread := nil;
  end;

end.

