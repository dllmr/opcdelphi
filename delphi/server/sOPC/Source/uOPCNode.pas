//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCNode;

interface

uses
  SysUtils, Classes, Windows, ActiveX,
  OPCDA;

type
  sOPCNode = class
  public
    NodeType:            integer;      // 0 = Branch, 1 = Leaf, 2 = Property
    Ident:               integer;      // node ID
    Parent:              integer;      // parent of this node
    Instance:            integer;      // how much clients are accessing this node

    // item fields
    stItemDataID:        string;       // name of the branch or leaf
    stItemID:            string;       // fully qualified ItemID
    slAddressPathList:   TStringList;
    vtCanonicalDataType: TVarType;
    dwAccessRights:      DWORD;
    dwBlobSize:          DWORD;
    pBlob:               PByteArray;
    dwEUType:            OPCEUTYPE;
    vEUInfo:             OleVariant;

    // property fields
    dwPropertyID:        DWORD;
    stDescription:       string;
    vtPropertyDataType:  TVarType;
    vPropertyData:       OleVariant;

    // value, quality and cache fields
    InitCurrent:         boolean;      // True -> CurrentValue is initialized
    CurrentValue:        variant;      // current node value
    Quality:             word;         // current node quality
    LastUpdate:          TDateTime;    // time value set by device read or write

    constructor Create;
  end;

implementation

constructor sOPCNode.Create;
begin
  NodeType := 0;
  Ident := 0;
  Parent := 0;
  Instance := 0;
  stItemDataID := '';
  stItemID := '';
  slAddressPathList := nil;
  vtCanonicalDataType := VT_EMPTY;
  dwAccessRights := 0;
  dwBlobSize := 0;
  pBlob := nil;
  dwEUType := 0;
  vEUInfo := VT_EMPTY;
  dwPropertyID := 0;
  stDescription := '';
  vtPropertyDataType := VT_EMPTY;
  vPropertyData := 0;
  InitCurrent := False;
  VarClear(CurrentValue);
  Quality := OPC_QUALITY_OUT_OF_SERVICE;
  LastUpdate := 0;
end;

end.

