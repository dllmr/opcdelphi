unit Main;

interface

uses Windows,Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
     RegDeRegServer,AxCtrls,ExtCtrls,ShutDownRequest;

type
  TForm1 = class(TForm)
    PulseTimer: TTimer;
    DateTimeLbl: TLabel;
    GroupBox1: TGroupBox;
    RegSerBtn: TButton;
    UnRegBtn: TButton;
    SDButton: TButton;
    Panel1: TPanel;
    Label2: TLabel;
    ClientConLbl: TLabel;
    GrpCountLbl: TLabel;
    Label1: TLabel;
    CloseBtn: TButton;
    procedure RegSerBtnClick(Sender: TObject);
    procedure CloseBtnClick(Sender: TObject);
    procedure SDButtonClick(Sender: TObject);
    procedure PulseTimerTimer(Sender: TObject);
    procedure UnRegBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  public
   clientsConnected:integer;
   procedure UpdateGroupCount;
  end;

var
  Form1: TForm1;


implementation

{$R *.DFM}

uses ServImpl,OPCCOMN,Globals,ActiveX,DataPointsUnit;

procedure TForm1.RegSerBtnClick(Sender: TObject);
begin
 RegisterTheServer(serverName);
end;

procedure TForm1.CloseBtnClick(Sender: TObject);
begin
 Close;
end;

procedure TForm1.SDButtonClick(Sender: TObject);
var
 i:integer;
 Obj: Pointer;
begin
 ShutDownDlg:=nil;
 try
  Application.CreateForm(TShutDownDlg,ShutDownDlg);
  for i:= low(theServers) to high(theServers) do
   begin
    ShutDownDlg.RadioGroup1.controls[i].enabled:=(theServers[i] <> nil);
    if ShutDownDlg.RadioGroup1.controls[i].enabled then
     ShutDownDlg.RadioGroup1.itemIndex:=i;        //select the last one why...
   end;                                           //just so one is selected ;)
  if ShutDownDlg.ShowModal = mrOk then
   begin
    i:=ShutDownDlg.RadioGroup1.itemIndex;
    if Assigned(theServers[i]) then
     if theServers[i].ClientIUnknown <> nil then
      if Succeeded(theServers[i].ClientIUnknown.QueryInterface(IOPCShutdown,Obj)) then
       IOPCShutdown(Obj).ShutdownRequest('MRD IOPCShutdown request.')
      else
       ShowMessage('The client does not support IOPCShutdown.');
   end;
  ShutDownDlg.Release;
 finally
 end;
end;

procedure TForm1.PulseTimerTimer(Sender: TObject);
var
 i:integer;
 cTime:TDateTime;
begin
 PulseTimer.enabled:=false;
 try
  cTime:=Now;

 //0     complete time and date
 //1     complete date
 //2     day
 //3     month
 //4     year
 //5     complete time
 //6     hour
 //7     minute
 //8     second
 //9     millisecond
 //10    Test_Tag_1
 //11    Test_Tag_1 Inverted
 //12    Test_Tag_2
 //13    Test_Tag_2 Inverted


  TRealDataPoint(rDataPoints[11]).aWord:=(not TRealDataPoint(rDataPoints[10]).aWord) and $FFFF;
  TRealDataPoint(rDataPoints[13]).aWord:=(not TRealDataPoint(rDataPoints[12]).aWord) and $FFFF;

  DecodeDate(cTime,TRealDataPoint(rDataPoints[4]).aWord,
                   TRealDataPoint(rDataPoints[3]).aWord,
                   TRealDataPoint(rDataPoints[2]).aWord);

  DecodeTime(cTime,TRealDataPoint(rDataPoints[6]).aWord,
                   TRealDataPoint(rDataPoints[7]).aWord,
                   TRealDataPoint(rDataPoints[8]).aWord,
                   TRealDataPoint(rDataPoints[9]).aWord);

  TRealDataPoint(rDataPoints[0]).aString:=TimeToStr(cTime) + ' ' + DateToStr(cTime);
  DateTimeLbl.Caption:=TRealDataPoint(rDataPoints[0]).aString;

  TRealDataPoint(rDataPoints[1]).aString:=
        DateToStr(EncodeDate(TRealDataPoint(rDataPoints[4]).aWord,
                             TRealDataPoint(rDataPoints[3]).aWord,
                             TRealDataPoint(rDataPoints[2]).aWord));

  TRealDataPoint(rDataPoints[5]).aString:=
        TimeToStr(EncodeTime(TRealDataPoint(rDataPoints[6]).aWord,
                             TRealDataPoint(rDataPoints[7]).aWord,
                             TRealDataPoint(rDataPoints[8]).aWord,
                             TRealDataPoint(rDataPoints[9]).aWord));


  for i:= low(theServers) to high(theServers) do
   if Assigned(theServers[i]) then
    theServers[i].TimeSlice(cTime);

 finally
  PulseTimer.enabled:=true;
 end;
end;

procedure TForm1.UnRegBtnClick(Sender: TObject);
begin
 UnRegisterTheServer(serverName);
end;

procedure TForm1.UpdateGroupCount;
var
 i,g:integer;
begin
 if Application.Terminated then Exit;
 clientsConnected:=0;
 g:=0;
 for i:= low(theServers) to high(theServers) do
  if Assigned(theServers[i]) then
   begin
    clientsConnected:=succ(clientsConnected);
    if Assigned(theServers[i].grps) then
     g:=g + theServers[i].grps.count;
    if Assigned(theServers[i].pubGrps) then
     g:=g + theServers[i].pubGrps.count;
   end;

 Form1.ClientConLbl.caption:=IntToStr(clientsConnected);
 GrpCountLbl.caption:=IntToStr(g);
 SDButton.enabled:=(clientsConnected  > 0);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
 SetUpRealDataPoints;
 UpdateGroupCount;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 FreeRealDataPoints;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
 if (clientsConnected > 0) then
  CanClose:=(MessageDlg('Clients are connected. Are you sure you want to quit?',
                        mtConfirmation,[mbYes,mbNo],0) =  mrYes)
 else
  CanClose:=true;
end;

end.
