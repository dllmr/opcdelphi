//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCItemEnumerator;

interface

uses
  Windows, Classes, WinProcs, ComObj, ActiveX,
  OPCDA, OpcError;

type
  sOPCItemEnumerator = class(TComObject, IEnumOPCItemAttributes)
  private
    nextIndex: integer;
    FOPCItemList: TList;

  public
    constructor Create(OPCItemList: TList);
    function Next(
      celt: cardinal;
      out ppItemArray: POPCITEMATTRIBUTESARRAY;
      out pceltFetched: cardinal): HResult; stdcall;
    function Skip(celt: cardinal): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult; stdcall;
  end;

implementation

uses
  ComServ, SysUtils,
  uGlobals, uOPCItem;

constructor sOPCItemEnumerator.Create(OPCItemList: TList);
begin
  inherited Create;
  FOPCItemList := OPCItemList;
  nextIndex := 0;
end;

function sOPCItemEnumerator.Next(
  celt: cardinal;
  out ppItemArray: POPCITEMATTRIBUTESARRAY;
  out pceltFetched: cardinal): HResult; stdcall;
var
 i: cardinal;
begin
  OPCLog(Format('sOPCItemEnumerator.Next - %d', [celt]));
  try
    pceltFetched := 0;

    // celt muß größer als 1 sein!
    if celt < 1 then begin
      Result := RPC_X_ENUM_VALUE_OUT_OF_RANGE;
      exit;
    end;

    Result := E_FAIL;
    ppItemArray := POPCITEMATTRIBUTESARRAY(CoTaskMemAlloc(celt * sizeof(OPCITEMATTRIBUTES)));
    FillChar(ppItemArray^, celt * sizeof(OPCITEMATTRIBUTES), 0);
    if ppItemArray = nil then begin
      Result := E_OUTOFMEMORY;
      exit;
    end;

    // sind noch Einträge in der Liste?
    if (nextIndex >= FOPCItemList.Count) then begin
      Result := S_FALSE;
      exit;
    end;

    // Einträge verarbeiten
    i := 0;
    while (i < celt) and (nextIndex < FOPCItemList.Count) do begin
      sOPCItem(FOPCItemList[nextIndex]).GetOPCItemAttributes(ppItemArray[i]);
      inc(i);
      inc(nextIndex);
    end;

    pceltFetched := i;
    if i = celt then Result := S_OK;

  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCItemEnumerator.Next - Exception', []), E);
    end;
  end;
end;

function sOPCItemEnumerator.Skip(celt: cardinal):HResult;
begin
  OPCLog(Format('sOPCItemEnumerator.Skip - %d', [celt]));
  try
    if (nextIndex + integer(celt)) <= FOPCItemList.Count then begin
      nextIndex := nextIndex + integer(celt);
      Result := S_OK;
    end else begin
      nextIndex := FOPCItemList.Count;
      Result := S_FALSE;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCItemEnumerator.Skip - Exception', []), E);
    end;
  end;
end;

function sOPCItemEnumerator.Reset: HResult;
begin
  nextIndex := 0;
  Result := S_OK;
end;

function sOPCItemEnumerator.Clone(out ppEnumItemAttributes: IEnumOPCItemAttributes): HResult;
begin
  OPCLog(Format('sOPCItemEnumerator.Clone', []));
  try
    try
      ppEnumItemAttributes := sOPCItemEnumerator.Create(FOPCItemList);
      Result := S_OK;
    except
      Result := E_UNEXPECTED;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCItemEnumerator.Clone - Exception', []), E);
    end;
  end;
end;

initialization

TComObjectFactory.Create(
  ComServer,
  sOPCItemEnumerator,
  IID_IEnumOPCItemAttributes,
  'sOPCItemEnumerator',
  'sOPCItemEnumerator Description',
  ciInternal,
  ThreadingModel);

end.

