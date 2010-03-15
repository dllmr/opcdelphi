//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCBrowse;

interface

uses
  SysUtils, Windows, ActiveX, Classes, ComCtrls,
  OPCDA, uOPCNode;

type
  sOPCBrowse = class
  protected
    FParent: integer;

  public
    constructor Create;

    destructor Destroy; override;

    function BrowseUp: boolean;

    function BrowseDown(stBranch: string): boolean;

    function BrowseTo(stItemID: string): boolean;

    function BrowseOPCItems(Filter: OPCBROWSETYPE): TStringList;

    function GetItemID(stItemDataID: string): string;

    function GetPropertyCount(stItemID: string): integer;

    function GetProperty(stItemID: string; Index: integer): sOPCNode;

    function GetPropertyData(stItemID: string; dwPropertyID: DWORD): OleVariant;

  end;

implementation

uses
 uGlobals, uOPC;

constructor sOPCBrowse.Create;
begin
  FParent := 0;
end;

destructor sOPCBrowse.Destroy;
begin
  inherited;
end;

function sOPCBrowse.BrowseUp: boolean;
// True -> Browse up
// False -> could not Browse up, we are at the root
begin
  Result := False;
  if FParent = 0 then exit;
  Result := True;
  FParent := OPC.Nodes[OPC.GetIndexOfNode(FParent)].Parent;
end;

function sOPCBrowse.BrowseDown(stBranch: string): boolean;
// True -> Browse down, ausgehend von der aktuellen Position
// False -> could not Browse down, no branch
var
  i: integer;
begin
  Result := False;
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].Parent = FParent) and
      (OPC.Nodes[i].stItemDataID = stBranch) and
      (OPC.Nodes[i].NodeType = 0)
    then begin
      Result := True;
      FParent := OPC.Nodes[i].Ident;
      exit;
    end;
  end;
end;

function sOPCBrowse.BrowseTo(stItemID: string): boolean;
// falls stItemID = '' -> auf Root Positionieren
// False -> stItemID nicht gefunden
// True  -> Ok
var
  i: integer;
begin
  Result := False;
  if stItemID = '' then begin
    Result := True;
    FParent := 0;
    exit;
  end;
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].stItemID = stItemID) and (OPC.Nodes[i].NodeType = 0) then begin
      Result := True;
      FParent := OPC.Nodes[i].Ident;
      exit;
    end;
  end;
end;

function sOPCBrowse.BrowseOPCItems(Filter: OPCBROWSETYPE): TStringList;
// liefert die Einträge von der aktuellen Position entsprechend dem angegebenen
// Filter in einer Stringliste zurück
var
  i: integer;
begin
  Result := TStringList.Create;
  case Filter of
    OPC_BRANCH: begin
      for i := 0 to High(OPC.Nodes) do begin
        if (OPC.Nodes[i].Parent = FParent) and (OPC.Nodes[i].NodeType = 0) then begin
          Result.Add(OPC.Nodes[i].stItemDataID);
          OPCLog(Format('sOPCBrowse.BrowseOPCItems - %s', [OPC.Nodes[i].stItemDataID]));
        end;
      end;
    end;
    OPC_LEAF: begin
      for i := 0 to High(OPC.Nodes) do begin
        if (OPC.Nodes[i].Parent = FParent) and (OPC.Nodes[i].NodeType = 1) then begin
          Result.Add(OPC.Nodes[i].stItemDataID);
          OPCLog(Format('sOPCBrowse.BrowseOPCItems - %s', [OPC.Nodes[i].stItemDataID]));
        end;
      end;
    end;
    OPC_FLAT: ;
  end;
end;

function sOPCBrowse.GetItemID(stItemDataID: string): string;
// liefert für stItemDataID die ItemID
// z.B. 'SetValue' -> 'Parameter.Profile.SetValue'
var
  i: integer;
begin
  Result := '?';
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].Parent = FParent) and (OPC.Nodes[i].stItemDataID = stItemDataID) then begin
      Result := OPC.Nodes[i].stItemID;
    end;
  end;
end;

function sOPCBrowse.GetPropertyCount(stItemID: string): integer;
var
  ix, i: integer;
begin
  Result := 0;
  ix := OPC.GetIndexOfstItemID(stItemID);
  if ix = -1 then exit;
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].Parent = OPC.Nodes[ix].Ident) and (OPC.Nodes[i].NodeType = 2)then begin
      inc(Result);
    end;
  end;
end;

function sOPCBrowse.GetProperty(stItemID: string; Index: integer): sOPCNode;
var
  ix, i, k: integer;
begin
  Result := nil;
  ix := OPC.GetIndexOfstItemID(stItemID);
  if ix = -1 then exit;
  k := 0;
  // +++ später in einem Durchlauf
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].Parent = OPC.Nodes[ix].Ident) and (OPC.Nodes[i].NodeType = 2) then begin
      if k = Index then begin
        Result := OPC.Nodes[i];
        exit;
      end;
      inc(k);
    end;
  end;
end;

function sOPCBrowse.GetPropertyData(stItemID: string; dwPropertyID: DWORD): OleVariant;
var
  ix, i: integer;
begin
  VarClear(Result);
  ix := OPC.GetIndexOfstItemID(stItemID);
  if ix = -1 then exit;
  for i := 0 to High(OPC.Nodes) do begin
    if (OPC.Nodes[i].Parent = OPC.Nodes[ix].Ident) and
      (OPC.Nodes[i].NodeType = 2) and
      (OPC.Nodes[i].dwPropertyID = dwPropertyID)
    then begin
      Result := OPC.Nodes[i].vPropertyData;
    end;
  end;
end;

end.

