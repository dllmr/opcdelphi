unit ItemAttributesOPC;

interface

uses Windows,Classes,Globals,SysUtils,OPCDA;

type
  TOPCItemAttributes = class
  public
    bActive:longbool;
    dwEUType:integer;
    pBlob:PByteArray;
    vEUInfo:OleVariant;
    szAccessPath,szItemID:string;
    vtRequestedDataType,vtCanonicalDataType:word;
    hClient,hServer,dwAccessRights,dwBlobSize:longword;
    constructor Create;
    destructor Destroy;override;
    procedure CopyYourSelf(dObj:TOPCItemAttributes);
  end;


implementation

constructor TOPCItemAttributes.Create;
begin
 pBlob:=nil;                          dwBlobSize:=0;
end;

destructor TOPCItemAttributes.destroy;
begin
end;

procedure TOPCItemAttributes.CopyYourSelf(dObj:TOPCItemAttributes);
begin
 dObj.bActive:=bActive;
 dObj.dwEUType:=dwEUType;
 dObj.pBlob:=pBlob;
 dObj.vEUInfo:=vEUInfo;
 dObj.szAccessPath:=szAccessPath;
 dObj.szItemID:=szItemID;
 dObj.vtRequestedDataType:=vtRequestedDataType;
 dObj.vtCanonicalDataType:=vtCanonicalDataType;
 dObj.hClient:= hClient;
 dObj.hServer:=hServer;
 dObj.dwAccessRights:=dwAccessRights;
 dObj.dwBlobSize:=dwBlobSize;
end;

end.
