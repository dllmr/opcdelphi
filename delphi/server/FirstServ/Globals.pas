unit Globals;

interface

uses Windows,Messages,SysUtils,Classes,Graphics,StdCtrls,Forms,Dialogs,Controls,
     ShellAPI,ActiveX,OPCDA,Variants;

const
 serverName = 'MRD.DA2.1';
  
type
 itemIDStrings = record
  trunk,branch,leaf:string[255];
 end;

type
  itemProps = record
    PropID: longword;
    tagname:string[64];
    dataType:integer;
  end;

const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';

 posItems: array[0..13] of itemProps =
  ((PropID: 5000; tagname: 'Complete';                        dataType:VT_BSTR),
   (PropID: 5001; tagname: 'Date.Complete';                   dataType:VT_BSTR),
   (PropID: 5002; tagname: 'Date.Parts.Day';                  dataType:VT_UI2),
   (PropID: 5003; tagname: 'Date.Parts.Month';                dataType:VT_UI2),
   (PropID: 5004; tagname: 'Date.Parts.Year';                 dataType:VT_UI2),
   (PropID: 5005; tagname: 'Time.Complete';                   dataType:VT_BSTR),
   (PropID: 5006; tagname: 'Time.Parts.Hour';                 dataType:VT_UI2),
   (PropID: 5007; tagname: 'Time.Parts.Min';                  dataType:VT_UI2),
   (PropID: 5008; tagname: 'Time.Parts.Seconds';              dataType:VT_UI2),
   (PropID: 5009; tagname: 'Time.Parts.Millseconds';          dataType:VT_UI2),
   (PropID: 5010; tagname: 'Test_Tag_1.Actual';               dataType:VT_UI2),
   (PropID: 5011; tagname: 'Test_Tag_1.Inverted';             dataType:VT_UI2),
   (PropID: 5012; tagname: 'Test_Tag_2.Actual';               dataType:VT_UI2),
   (PropID: 5013; tagname: 'Test_Tag_2.Inverted';             dataType:VT_UI2));


 io2Read        = 1;
 io2Write       = 2;
 io2Refresh     = 3;
 io2Change      = 4;

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
function ReturnPropIDFromTagname(const s1:string):longword;
function ReturnTagnameFromPropID(PropID:longword):string;
function CanPropIDBeWritten(i:longword):boolean;
function ReturnDataTypeFromPropID(i:longword):integer;
procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
function ConvertVariant(cv:OleVariant; reqDataType:TVarType; var q:word; var err:HRESULT):OleVariant;
function IsVariantTypeOK(vType:integer):boolean;

implementation

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
var
 tempS:string;
 finish:boolean;
 nextloc,strLength: integer;
begin
 {$R-}
 strLength := length(theString);
 finish := false;
 SetLength(tempS,strLength);
 result := tempS;
 nextloc := 1;
 while not finish do
  begin
   if (start < 256) and (theString[start] <> theChar) and
      (theString[start] <> chr(13)) and (start <= strLength) then
    begin
     tempS[nextloc] := theString[start];
     Inc(nextloc);
     Inc(start);
    end
   else
    begin
     SetLength(tempS,nextloc-1);      {this sets the length of the string}
     finish:=true;                    {exit the loop}
     result:=tempS;                   {return the value}
    end;
  end;
 {$R+}
end;

function ReturnPropIDFromTagname(const s1:string):longword;
var
 i:integer;
begin
 result:=0;
 for i:= low(posItems) to high(posItems) do
  if (posItems[i].tagname = s1) then
   begin
    result:=posItems[i].PropID;
    Exit;
   end;
end;

function ReturnTagnameFromPropID(PropID:longword):string;
var
 i:integer;
begin
 result:='';
 for i:= low(posItems) to high(posItems) do
  if posItems[i].PropID = PropID then
   begin
    result:=posItems[i].tagname;
    Exit;
   end;
end;

function CanPropIDBeWritten(i:longword):boolean;
begin
 i:= i - posItems[low(posItems)].PropID;
 result:=boolean(i in [10,12]);               //the test Test_Tag_X's
end;

function ReturnDataTypeFromPropID(i:longword):integer;
var
 x:longword;
begin
 x:= i - posItems[low(posItems)].PropID;
 if (x <= high(posItems)) then
  result:=posItems[x].dataType
 else
  result:=VT_UI2;
end;

procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
var
 sTime:TSystemTime;
begin
 DateTimeToSystemTime(cTime,sTime);
 SystemTimeToFileTime(sTime,OPCTime);
 LocalFileTimeToFileTime(OPCTime,OPCTime);
end;

function ConvertVariant(cv:OleVariant; reqDataType:TVarType; var q:word; var err:HRESULT):OleVariant;
begin
 try
//  VariantClear(result);
  err:=VariantChangeTypeEx(result,cv, $400, 0, reqDataType);
  if (err = DISP_E_OVERFLOW) then
   q:=OPC_QUALITY_BAD;
 except
  err:=DISP_E_TYPEMISMATCH;
 end;
end;

function IsVariantTypeOK(vType:integer):boolean;
begin
 result:=boolean(vType in [VT_EMPTY..VT_CLSID]);
// result:=boolean(vType in [varEmpty..$14]);
end;

end.
