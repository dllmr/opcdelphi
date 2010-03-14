//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCItem;

interface

uses
  Windows, ActiveX, ComObj, Axctrls, SysUtils, Dialogs, Classes,
  OPCtypes, OPCDA, uGlobals, uOPCNode;

type
  sOPCItem = class
  protected
    // OPCNode
    FOPCNode:             sOPCNode;
    // OPC Attributes
    stAccessPath:         string;
    bActive:              boolean;
    hClient:              OPCHANDLE;
    vtRequestedDataType:  TVarType;         // VT_EMPTY -> use native type
    FCallBackTime:        TDateTime;        // last OPCItem call back
    FQuality:             word;
    FTag:                 integer;

  public
    constructor Create(aOPCNode: sOPCNode);
    destructor Destroy; override;

    procedure Copy(dItem: sOPCItem);

    function Read(DataSource: integer): variant;
    procedure Write(Value: variant);

    procedure GetOPCItemState(DataSource: word; var ItemState: OPCITEMSTATE);
    procedure GetOPCItemAttributes(var oia: OPCITEMATTRIBUTES);

    function WriteAble: boolean;
    // True -> writing on this tag is allowed

    function ReadAble: boolean;
    // True -> reading on this tag is allowed

    function SetRequestedDataType(aRequestedDataType: TVarType): boolean; virtual;
    // True  -> RequestedDataType is set
    // False -> RequestedDataType is not supported

    property AccessPath: string read stAccessPath write stAccessPath;
    property Active: boolean read bActive write bActive;
    property ClientHandle: OPCHANDLE read hClient write hClient;
    property Quality: word read FQuality write FQuality;
    property RequestedDataType: TVarType read vtRequestedDataType;
    property OPCNode: sOPCNode read FOPCNode;
    property CallBackTime: TDateTime read FCallBackTime write FCallBackTime;
    property Tag: integer read FTag write FTag;
  end;

implementation

uses
  uOPC;

constructor sOPCItem.Create(aOPCNode: sOPCNode);
begin
  inherited Create;
  FOPCNode := aOPCNode;
  bActive := False;
  FQuality := FOPCNode.Quality;
  stAccessPath := '';
  hClient := 0;
  vtRequestedDataType := VT_EMPTY;
  FCallBackTime := 0;
  FTag := 0;
  inc(FOPCNode.Instance);
end;

destructor sOPCItem.Destroy;
begin
  dec(FOPCNode.Instance);
  inherited;
end;

procedure sOPCItem.Copy(dItem: sOPCItem);
begin
  dItem.FOPCNode := FOPCNode;
  dItem.stAccessPath := stAccessPath;
  dItem.bActive := bActive;
  dItem.hClient := hClient;
  dItem.vtRequestedDataType := vtRequestedDataType;
  dItem.FQuality := FQuality;
end;

function sOPCItem.Read(DataSource: integer): variant;
begin
  // initialize cache
  if not FOPCNode.InitCurrent then begin
    DataSource := OPC_DS_DEVICE;
    FOPCNode.InitCurrent := True;
  end;
  case DataSource of
    OPC_DS_CACHE: begin
      if bActive
        then FQuality := FOPCNode.Quality
        else FQuality := OPC_QUALITY_OUT_OF_SERVICE;
    end;
    OPC_DS_DEVICE: begin
      OPC.ItemRead(FOPCNode, stAccessPath, vtRequestedDataType);
      FQuality := FOPCNode.Quality;
    end;
  end;
  Result := FOPCNode.CurrentValue;
end;

procedure sOPCItem.Write(Value: variant);
// Writes are not affected by the ACTIVE state of the group or item.
begin
  OPC.ItemWrite(FOPCNode, Value, stAccessPath, vtRequestedDataType);
  FQuality := FOPCNode.Quality;
end;

procedure sOPCItem.GetOPCItemState(DataSource: word; var ItemState: OPCITEMSTATE);
begin
  if (DataSource = OPC_DS_DEVICE) or bActive then begin
    ItemState.hClient := hClient;
    ItemState.vDataValue := Read(DataSource);
    ItemState.wQuality := FQuality;
    ItemState.ftTimeStamp := ConvertToFileTime(FOPCNode.LastUpdate);
  end else begin
    ItemState.hClient := hClient;
    ItemState.wQuality := OPC_QUALITY_OUT_OF_SERVICE;
  end;
end;

procedure sOPCItem.GetOPCItemAttributes(var oia: OPCITEMATTRIBUTES);
begin
  oia.szAccessPath := StringToLPOLESTR(stAccessPath);
  oia.szItemID := StringToLPOLESTR(FOPCNode.stItemID);
  oia.bActive := bActive;
  oia.hClient := hClient;
  oia.hServer := OPCHANDLE(self);
  oia.dwAccessRights := FOPCNode.dwAccessRights;
  oia.dwBlobSize := FOPCNode.dwBlobSize;
  oia.pBlob := FOPCNode.pBlob;
  oia.vtRequestedDataType := vtRequestedDataType;
  oia.vtCanonicalDataType := FOPCNode.vtCanonicalDataType;
  oia.dwEUType := FOPCNode.dwEUType;
  oia.vEUInfo := FOPCNode.vEUInfo;
end;

function sOPCItem.WriteAble: boolean;
begin
  Result := (FOPCNode.dwAccessRights and OPC_WRITEABLE) = OPC_WRITEABLE;
end;

function sOPCItem.ReadAble: boolean;
begin
  Result := (FOPCNode.dwAccessRights and OPC_READABLE) = OPC_READABLE;
end;

function sOPCItem.SetRequestedDataType(aRequestedDataType: TVarType): boolean;
begin
  Result := True;
  vtRequestedDataType := aRequestedDataType;
  // +++ später Event erzeugen -> prüfen, ob Typ zulässig
end;

end.

