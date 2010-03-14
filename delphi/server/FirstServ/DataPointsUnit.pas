unit DataPointsUnit;

interface

uses Windows,ActiveX,ComObj,SysUtils,Dialogs,Classes,Globals;

type
  TRealDataPoint = class
  public
   aWord:word;
   aString:string;
  end;

 procedure FreeRealDataPoints;
 procedure SetUpRealDataPoints;

var
 rDataPoints:TList;

implementation

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
 //23    Test_Tag_2 Inverted

procedure FreeRealDataPoints;
var
 i:integer;
begin
 for i:=0 to rDataPoints.count - 1 do
  if Assigned(rDataPoints[i]) then
   TRealDataPoint(rDataPoints[i]).Free;
end;

procedure SetUpRealDataPoints;
var
 i:integer;
begin
 for i:=0 to 13 do
  rDataPoints.Add(TRealDataPoint.Create);

 TRealDataPoint(rDataPoints[10]).aWord:=1;
 TRealDataPoint(rDataPoints[12]).aWord:=2;

end;

initialization
 rDataPoints:=TList.Create;

finalization
 if Assigned(rDataPoints) then
  rDataPoints.Free;

end.
