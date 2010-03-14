unit EnumItemAtt;

interface

uses Windows,Classes,WinProcs,ComObj,ActiveX,ItemAttributesOPC,OPCDA,OpcError,
     Globals;

type
  TOPCItemAttEnumerator = class(TComObject,IEnumOPCItemAttributes)
  private
    iaList:TList;
    nextIndex:cardinal;
  public
    constructor Create(inList: TList);
    destructor Destroy;override;
    procedure PopulateRecord(var theRec:OPCITEMATTRIBUTES;i:integer);virtual;
    function Next(celt: Cardinal; out ppItemArray:POPCITEMATTRIBUTESARRAY;
                  out pceltFetched:Cardinal):HResult;stdcall;
    function Skip(celt:Cardinal):HResult;stdcall;
    function Reset:HResult; stdcall;
    function Clone(out ppEnumItemAttributes:IEnumOPCItemAttributes):HResult;stdcall;
  end;


implementation

uses ComServ;

constructor TOPCItemAttEnumerator.Create(inList:TList);
var
 i:integer;
 aItemObj:TOPCItemAttributes;
begin
 inherited Create;
 nextIndex:=0;
 iaList:=TList.create;
 for i:=0 to inList.count-1 do
  begin
   aItemObj:=TOPCItemAttributes.Create;
   TOPCItemAttributes(inList[i]).CopyYourSelf(aItemObj);
   iaList.add(aItemObj);
  end;
end;

destructor TOPCItemAttEnumerator.Destroy;
var
 i:integer;
begin
 if Assigned(iaList) and (iaList.count > 0) then
  for i:=0 to iaList.count - 1 do
   TOPCItemAttributes(iaList[i]).Free;
 iaList.free;
 inherited destroy;
end;

procedure TOPCItemAttEnumerator.PopulateRecord(var theRec:OPCITEMATTRIBUTES;i:integer);
begin
 FillChar(theRec,sizeOf(theRec),#0);                    //no variant conversion errors
 with TOPCItemAttributes(iaList[i]) do
  begin
   theRec.szAccessPath:=StringToLPOLESTR(szAccessPath);
   theRec.szItemID:=StringToLPOLESTR(szItemID);
   theRec.bActive:=bActive;
   theRec.hClient:=hClient;
   theRec.hServer:=hServer;
   theRec.dwAccessRights:=dwAccessRights;
   theRec.dwBlobSize:=dwBlobSize;
   theRec.pBlob:=pBlob;
   theRec.vtRequestedDataType:=vtRequestedDataType;
   theRec.vtCanonicalDataType:=vtCanonicalDataType;
   theRec.dwEUType:=dwEUType;
   theRec.vEUInfo:=vEUInfo;
 end;
end;

function TOPCItemAttEnumerator.Next(celt: Cardinal; out ppItemArray:POPCITEMATTRIBUTESARRAY;
                  out pceltFetched:Cardinal):HResult;
var
 i:Cardinal;
begin
 i:=0;
 pceltFetched:=i;
 if (celt < 1) then
  begin
   result:=S_FALSE;
//   Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
   Exit;
  end;

 Result:=E_FAIL;
 ppItemArray:=POPCITEMATTRIBUTESARRAY(CoTaskMemAlloc(celt * sizeof(OPCITEMATTRIBUTES)));
 if (ppItemArray = nil) then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 while (i < celt) do
  begin
   if (nextIndex < cardinal(iaList.Count)) then
    begin
     PopulateRecord(ppItemArray[i],nextIndex);
     Inc(i);
     Inc(nextIndex);
    end
   else
    begin
     Result:=S_FALSE;
     Break;
    end;
  end;

 pceltFetched:=i;
 if (i = celt) then
  Result:=S_OK;
end;

function TOPCItemAttEnumerator.Skip(celt:Cardinal):HResult;
begin
 if (nextIndex + celt) <= cardinal(iaList.Count) then
  begin
   nextIndex:=nextIndex + celt;
   result:=S_OK;
  end
 else
  begin
   nextIndex:=iaList.Count;
   result:=S_FALSE;
  end;
end;

function TOPCItemAttEnumerator.Reset: HResult;
begin
 nextIndex:=0;
 result:=S_OK;
end;

function TOPCItemAttEnumerator.Clone(out ppEnumItemAttributes:IEnumOPCItemAttributes): HResult;
begin
 try
  ppEnumItemAttributes:=TOPCItemAttEnumerator.Create(iaList);
  result:=S_OK;
 except
  result:=E_UNEXPECTED;
 end;
end;

initialization
 TComObjectFactory.Create(ComServer,
                            TOPCItemAttEnumerator,
                            IID_IEnumOPCItemAttributes,
                            'TOPCItemAttEnumerator',
                            'MRD',
                            ciMultiInstance,
                            tmApartment);

end.
