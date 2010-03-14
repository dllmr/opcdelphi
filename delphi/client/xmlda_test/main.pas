{************************************************************************
*  Copyright  2003  Dipl.-Ing. H.-D. Kassl
*  http://www.dopc.kassl.de
*
*  LICENSE AGREEMENT
*  free, also for commerical use
*
*  Liability disclaimer:
*  THIS SOFTWARE IS DISTRIBUTED "AS IS" AND WITHOUT WARRANTIES AS TO
*  PERFORMANCE  OF MERCHANTABILITY OR ANY OTHER WARRANTIES WHETHER EXPRESSED
*  OR IMPLIED. YOU USE IT AT YOUR OWN RISK. THE AUTHOR WILL NOT BE LIABLE FOR
*  DATA LOSS, DAMAGES,  LOSS OF PROFITS OR ANY OTHER KIND OF LOSS WHILE USING
*  OR MISUSING THIS SOFTWARE.
*
*  Short description
*  Very simple OPC XML DA Client for Delphi 5,6,7 (Delphi 4 not tested)
*
*  Description
*  This example program shows the core OPC XML DA version 1.0 methods.
*  For further information you can download the OPC XML DA spezification from:
*  http://www.opcfoundation.org.
*
*  To get access to your OPC COM based server you can download our
*  dOPC XGate server, which augments any OPC COM based server with
*  an OPC XML DA interface, from:
*  http://www.dopc.kassl.de
*
*  The example SOAP requests are based on dOPC XGate with underlying
*  COM based Matrikon Simulation OPC Server. You can download this
*  server for free from:
*  http://www.matrikon.com
*
*  However, this program also works with other OPC XML DA
*  version 1.0 compliance servers
*
*  Have fun with OPC XML DA :-))))
*
************************************************************************}
unit main;

interface

{..$DEFINE INDYPOST}  // if you have installed INDY, you can also use INDY
uses
  Windows, Classes, Sysutils,Controls, Forms, StdCtrls, ExtCtrls,
  SHDocVw,
  {$IFDEF INDYPOST}
    indypost,
  {$ELSE}
    MSHTTPPost,
  {$ENDIF}
  OleCtrls;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    eURL: TEdit;
    bGo: TButton;
    cXMLCommand: TComboBox;
    WebBrowser1: TWebBrowser;
    Splitter1: TSplitter;
    Panel2: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    Memo1: TMemo;
    procedure FormShow(Sender: TObject);
    procedure cXMLCommandChange(Sender: TObject);
    procedure bGoClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
  public
  end;

var
  Form1: TForm1;

implementation
{$R *.DFM}

const
  SHB1 = '<?xml version="1.0" encoding="UTF-8" ?>';
  SHB2 = '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
  SHB3 = '<SOAP-ENV:Body>';
  //... XML Command
  SHE1 = '</SOAP-ENV:Body>';
  SHE2 = '</SOAP-ENV:Envelope>';


procedure TForm1.FormShow(Sender: TObject);
begin
  cXMLCommand.ItemIndex := 0;  // to first list item
  cXMLCommandChange(self);     // show SOAP Request
end;

// shows SOAP requests in memo
procedure TForm1.cXMLCommandChange(Sender: TObject);
var
  Filename : string;
begin
  // load current SOAP action file (GetStatus.x, Browse.x, ...)
  FileName := ExtractFilePath(Application.ExeName)+cXMLCommand.Text+'.x';
  Memo1.Lines.LoadFromFile(FileName);
end;

// change a given token <Token> (e.g #time#) in string <Source> to
// the given string <ToStr> (e.g. 11:00).
// The function give the result string back
function ChangeToken(Source,Token,ToStr: string): string;
var
  Dest : string;
  TokenPos : integer;
begin
    Dest := Source;
    repeat
      TokenPos := pos(Token,Dest);
      if TokenPos <> 0 then
      begin
         delete(Dest,TokenPos,length(Token));
         insert(ToStr,Dest,Tokenpos);
      end;
    until Tokenpos = 0;
    result := Dest;
end;

// add <sec> seconds to a given datetime <dt>
// result: datetime with added seconds
function AddSec(Dt: TDateTime; Sec: extended): TDateTime;
begin
  result :=  DT + Sec / (24 * 60 * 60)
end;


function GetTimeZoneBias: Integer;
var
  TimeZoneInfo: TTimeZoneInformation;
begin
  case GetTimeZoneInformation(TimeZoneInfo) of
  TIME_ZONE_ID_STANDARD: Result := TimeZoneInfo.Bias + TimeZoneInfo.StandardBias;
  TIME_ZONE_ID_DAYLIGHT: Result := TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias;
  else
    Result := 0;
  end;
end;

// converts datetime to xml datetime
function dDateTimeToXMLTime(Value: TDateTime; ApplyLocalBias: Boolean = True): WideString;
const
  Neg: array[Boolean] of string=  ('+', '-');
var
  Bias: Integer;
begin
  Result := FormatDateTime('yyyy''-''mm''-''dd''T''hh'':''nn'':''ss''.''zzz', Value);
  Bias := GetTimeZoneBias;
  if (Bias <> 0) and ApplyLocalBias then
  begin
    Result := Format('%s%s%.2d:%.2d', [Result, Neg[Bias > 0],
                                       Abs(Bias) div 60,
                                       Abs(Bias) mod 60]);
  end else
    Result := Result + 'Z';
end;


// the main method, to execute SOAP request
procedure TForm1.bGoClick(Sender: TObject);
var
  SoapAction : string;
  Request    : TStringStream;
  Response   : TStringStream;
  SOAPText   : string;
  SList      : TStringlist;
  FileName   : String;
  OldC       : TCursor;
begin
  Oldc := Screen.Cursor;
  SoapAction := 'http://opcfoundation.org/webservices/XMLDA/1.0/'+cXMLCommand.Text;
  SOAPText := Memo1.Lines.text;
  // converts token in SOAP request
  SOAPText := ChangeToken(SOAPText,'#now+1#',dDateTimeToXMLTime(AddSec(now,1)));
  SOAPText := ChangeToken(SOAPText,'#now#',dDateTimeToXMLTime(now));
  SOAPText := ChangeToken(SOAPText,'#now-1#',dDateTimeToXMLTime(AddSec(now,-1)));
  SOAPText := ChangeToken(SOAPText,'#now+1m#',dDateTimeToXMLTime(AddSec(now,60)));
  Request  := TStringStream.Create(SHB1+SHB2+SHB3+SOAPText+SHE1+SHE2);
  Response := TStringStream.Create('');
  SList    := TStringlist.create;
  try
    Screen.Cursor := crHourGlass;
    PostData(eUrl.Text,'','','',SoapAction,Request,Response); // send request
    Response.Position := 0;
    SList.LoadFromStream(Response);  // copy response to stringlist
    FileName := ExtractFilePath(Application.ExeName)+'Response.xml';
    SList.SaveToFile(FileName);      // save response to file
    WebBrowser1.Navigate(WideString(FileName)); // show response in browser
  finally
    Request.Free;
    Response.Free;
    SList.Free;
    Screen.Cursor := OldC;
  end;
end;

// execute SOAP request is also possible with keyboard
procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F9 then
    bGoClick(self);
end;

end.
