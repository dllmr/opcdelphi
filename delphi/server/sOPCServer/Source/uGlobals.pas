//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uGlobals;

interface

uses
  Windows, ActiveX, SysUtils, ComObj,
  uLogging,
  OPCDA;

//******************************************************************************
const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';

//******************************************************************************
const
  ThreadingModel: TThreadingModel = tmFree;

type
  sMemoryKind = (mkHResult, mkItemState, mkItemResult, mkServerStatus,
    mkDWORD, mkPOleStr, mkVarType, mkOleVariant, mkWord, mkFileTime);

//******************************************************************************
var
  Logging: TLogging;

//******************************************************************************
function ConvertToFileTime(DateTime: TDateTime): TFileTime;
// converts time format TDateTime in TFileTime

function TaskMemAlloc(dwCount: DWORD; mk: sMemoryKind; var aResult: HResult): pointer;
// allocates Task Memory

procedure TaskMemFree(Memory: pointer);
// releases Task Memory

procedure OPCLog(Text: string);
// logging function

procedure OPCLogException(Text: string; E: Exception);
// OPC Exception log

//******************************************************************************
implementation

//******************************************************************************
function ConvertToFileTime(DateTime: TDateTime): TFileTime;
var
  sTime: TSystemTime;
begin
  DateTimeToSystemTime(DateTime, sTime);
  SystemTimeToFileTime(sTime, Result);
  LocalFileTimeToFileTime(Result, Result);
end;

//******************************************************************************
// Allocate Task Memory
//******************************************************************************
function TaskMemAlloc(dwCount: DWORD; mk: sMemoryKind; var aResult: HResult): pointer;
// Result = nil -> no memory allocated
var
  Size: integer;
begin
  try
    Size := 0;
    case mk of
      mkHResult:         Size := dwCount * sizeof(HRESULT);
      mkItemState:       Size := dwCount * sizeof(OPCITEMSTATE);
      mkItemResult:      Size := dwCount * sizeof(OPCITEMRESULT);
      mkServerStatus:    Size := sizeof(OPCSERVERSTATUS);
      mkDWORD:           Size := dwCount * sizeof(DWORD);
      mkPOleStr:         Size := dwCount * sizeof(POleStr);
      mkVarType:         Size := dwCount * sizeof(TVarType);
      mkOleVariant:      Size := dwCount * sizeof(OleVariant);
      mkWord:            Size := dwCount * sizeof(word);
      mkFileTime:        Size := dwCount * sizeof(TFileTime);
    end;
    Result := CoTaskMemAlloc(Size);
    if Result = nil
      then aResult := E_OUTOFMEMORY
      else FillChar(Result^, Size, 0);
  except
    on E: Exception do begin
      Result := nil;
      OPCLogException('TaskMemAlloc', E);
    end;
  end;
end;

procedure TaskMemFree(Memory: pointer);
begin
  if Memory <> nil then CoTaskMemFree(Memory);
end;

procedure OPCLog(Text: string);
begin
  if (Logging <> nil) then Logging.WriteIntoFile(Text);
  // PeekAt(Text);
end;

procedure OPCLogException(Text: string; E: Exception);
begin
  if Logging = nil then exit;
  Logging.WriteIntoFile(Format('%s %p %s - %s', [DateTimeToStr(Now), ExceptAddr, Text, E.Message]));
  // PeekError(Format('%s - Exception - %p - %s', [Text, ExceptAddr, E.Message]));
end;

initialization
  Logging := nil;

finalization

end.

