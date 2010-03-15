//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCStringEnumerator;

interface

uses
  Classes, WinProcs, ComObj, ActiveX;

type
  sOPCStringEnumerator = class(TComObject, IEnumString)
  private
    nextIndex: integer;
    FList: TStringList;
    FFree: boolean;
  public
    constructor Create(const aList: TStringList; aFree: boolean = True);
    destructor Destroy; override;
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
  end;

implementation

uses
  ComServ, SysUtils,
  OPCDA, uGlobals;

constructor sOPCStringEnumerator.Create(const aList: TStringList; aFree: boolean);
begin
  inherited Create;
  FList := aList;
  FFree := aFree;
  nextIndex := 0;
end;

destructor sOPCStringEnumerator.Destroy;
begin
  if FFree then FList.Free;
  inherited Destroy;
end;

function sOPCStringEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
 i: integer;
begin
  OPCLog(Format('sOPCStringEnumerator.Next - %d', [celt]));
  try
    // Argumente prüfen
    if pceltFetched = nil then begin
      Result := E_INVALIDARG;
      exit;
    end;
    pceltFetched^ := 0;

    // celt muß größer als 1 sein!
    if celt < 1 then begin
      Result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
      exit;
    end;

    // sind noch Einträge in der Liste?
    if (nextIndex >= FList.Count) then begin
      Result := S_FALSE;
      exit;
    end;

    // Einträge verarbeiten
    Result := S_FALSE;
    i := 0;
    while (i < celt) and (nextIndex < FList.Count) do begin
      TPointerList(elt)[i] := StringToLPOLESTR(FList[nextIndex]);
      inc(i);
      inc(nextIndex);
    end;

    pceltFetched^ := i;
    if i = celt then Result := S_OK;

  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCStringEnumerator.Next - Exception', []), E);
    end;
  end;
end;

function sOPCStringEnumerator.Skip(celt: Longint): HResult;
begin
  OPCLog(Format('sOPCStringEnumerator.Skip - %d', [celt]));
  try
    if (nextIndex + celt) <= FList.Count then begin
      inc(nextIndex, celt);
      Result := S_OK;
    end else begin
      nextIndex := FList.Count;
      Result := S_FALSE;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCStringEnumerator.Skip - Exception', []), E);
    end;
  end;
end;

function sOPCStringEnumerator.Reset: HResult;
begin
  nextIndex := 0;
  Result := S_OK;
end;

function sOPCStringEnumerator.Clone(out enm: IEnumString): HResult;
begin
  OPCLog(Format('sOPCStringEnumerator.Clone', []));
  try
    enm := sOPCStringEnumerator.Create(FList);
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCStringEnumerator.Clone - Exception', []), E);
    end;
  end;
end;

initialization

TComObjectFactory.Create(
  ComServer,
  sOPCStringEnumerator,
  IEnumString,
  'sOPCStringEnumerator',
  'sOPCStringEnumerator Description',
  ciInternal,
  ThreadingModel);

end.

