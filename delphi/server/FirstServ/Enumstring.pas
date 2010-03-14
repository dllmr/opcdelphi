unit EnumString;

interface

uses Classes,WinProcs,ComObj,ActiveX;

type
  TOPCStringsEnumerator = class(TComObject, IEnumString)
  private
    nextIndex:Integer;
    strList:TStringList;
  public
    constructor Create(const inStrings: TStringList);
    destructor Destroy;override;
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
  end;


implementation

uses ComServ,OPCDA;

const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';  //is in ole2.pas


constructor TOPCStringsEnumerator.Create(const inStrings: TStringList);
var
 i:integer;
begin
 inherited Create;
 strList := TStringList.create;
 for i:=0 to inStrings.count-1 do
  strList.add(inStrings[i]);
 nextIndex := 0;
end;

destructor TOPCStringsEnumerator.destroy;
begin
 strList.free;
 inherited destroy;
end;

function TOPCStringsEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
 i: integer;
begin
 i:=0;
 if celt < 1 then
  begin
   Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
   Exit;
  end;
 if pceltFetched = nil then
  begin
   Result:=E_INVALIDARG;
   Exit;
  end;

 Result := S_FALSE;
 while (i < celt) do
  begin
   if (nextIndex < strList.Count) then
    begin
     TPointerList(elt)[i]:=StringToLPOLESTR(strList[nextIndex]);
     Inc(i);
     Inc(nextIndex);
    end
   else
    begin
     Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
     Break;
    end;
  end;

 pceltFetched^:=i;
 if (i = celt) then
  Result := S_OK;
end;

function TOPCStringsEnumerator.Skip(celt: Longint): HResult;
begin
 if (nextIndex + celt) <= strList.Count then
  begin
   nextIndex:=nextIndex + celt;
   result:=S_OK;
  end
 else
  begin
   nextIndex:=strList.Count;
   result:=S_FALSE;
  end;
end;

function TOPCStringsEnumerator.Reset: HResult;
begin
 nextIndex:=0;                        result:=S_OK;
end;

function TOPCStringsEnumerator.Clone(out enm: IEnumString): HResult;
begin
 try
  enm:=TOPCStringsEnumerator.Create(strList);
  result:=S_OK;
 except
  result:=E_UNEXPECTED;
 end;
end;

initialization
 TComObjectFactory.Create(ComServer,
                            TOPCStringsEnumerator,
                            IEnumString,
                            'TOPCStringsEnumerator',
                            'MRD',
                            ciMultiInstance,
                            tmApartment);

end.
