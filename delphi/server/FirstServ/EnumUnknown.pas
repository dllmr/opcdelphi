unit EnumUnknown;

interface

uses Classes,WinProcs,ComObj,ActiveX;

type
  TMRDUnknownEnumerator = class(TComObject, IEnumUnknown)
  private
    theList:TList;
    nextIndex:integer;
  public
    constructor Create(inList:TList);
    destructor Destroy;override;
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumUnknown): HResult; stdcall;
    procedure AddAnotherList(inList:TList);
  end;

implementation

uses ComServ,OPCDA;

constructor TMRDUnknownEnumerator.Create(inList:TList);
var
 i:integer;
begin
 Inherited Create;

 if not Assigned(inList) then
  Exit;

 theList:=TList.Create;
 if not Assigned(theList) then
  Exit;

 nextIndex:=0;

 for i:=0 to inList.count - 1 do
  theList.add(inList[i]);
end;

destructor TMRDUnknownEnumerator.destroy;
begin
 theList.Free;
 Inherited Destroy;
end;

function TMRDUnknownEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
 i: integer;
begin
 i:=0;
 if (celt < 1) then
  begin
   result:=S_FALSE;
//   Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
   Exit;
  end;

 if (pceltFetched = nil) then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 result:=S_FALSE;
 while (i < celt) do
  begin
   if (nextIndex < theList.Count) then
    begin
     TPointerList(elt)[i]:=theList[nextIndex];
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

function TMRDUnknownEnumerator.Skip(celt: Longint): HResult;
begin
 if (nextIndex + celt) <= theList.Count then
  begin
   nextIndex:=nextIndex + celt;
   result:=S_OK;
  end
 else
  begin
   nextIndex:=theList.Count;
   result:=S_FALSE;
  end;
end;

function TMRDUnknownEnumerator.Reset: HResult;
begin
 nextIndex:=0;
 result:=S_OK;
end;

function TMRDUnknownEnumerator.Clone(out enm: IEnumUnknown): HResult;
begin
 try
  enm:=TMRDUnknownEnumerator.Create(theList);
  result:=S_OK;
 except
  result:=E_UNEXPECTED;
 end;
end;

procedure TMRDUnknownEnumerator.AddAnotherList(inList:TList);
var
 i:integer;
begin
 if not Assigned(inList) then
  Exit;
 for i:=0 to inList.count - 1 do
  theList.add(inList[i]);
end;

initialization
 TComObjectFactory.Create(ComServer,
                          TMRDUnknownEnumerator,
                          IEnumUnknown,
                          'TMRDUnknownEnumerator',
                          'MRD',
                          ciMultiInstance,
                          tmApartment);
end.
