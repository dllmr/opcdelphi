//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPC;

interface

uses
  SysUtils, Classes, Windows, ActiveX, SyncObjs,
  OPCDA,
  uOPCNode;

type
  sOnConnectEvent = procedure(aConnecting: boolean);
  // event on connect

  sOnRead = procedure (OPCNode: sOPCNode; Path: string; DataType: TVarType);
  // tag read event
  // OPCNode  = tag class
  // Path     = access path
  // DataType = requested data type from client

  sOnWrite = procedure (OPCNode: sOPCNode; Value: variant; Path: string; DataType: TVarType);
  // tag write event
  // OPCNode  = tag class
  // Value    = value to write
  // Path     = access path
  // DataType = requested data type from client

  sOnSetDataType = function(OPCNode: sOPCNode; DataType: TVarType): boolean;
  // not used

  sOnInitAddressSpace = procedure;
  // init address space event

  { sOnRead and sOnWrite must set the fields CurrentValue, Quality and LastUpdate
    of the OPCNode object.
    If LastUpdate is set, the server automatically updates the client if callback
    is enabled.
    Path is the desired AccessPath and DataType is the requested data type.
    If DataType = VT_EMPTY then the native (canonical) data type must be used.
  }

  sNodeArray = array of sOPCNode;

  sOPC = class
  protected
    FCS: TCriticalSection;
    FDataAccessServers: TList;
    FServerName: string;
    FServerDescription: string;
    FAddressSpaceInit: boolean;
    FNodes: sNodeArray;
    FCanStart: boolean;

    FOnRead: sOnRead;
    FOnWrite: sOnWrite;
    FOnSetDataType: sOnSetDataType;
    FOnInitAddressSpace: sOnInitAddressSpace;

    FOnConnect: sOnConnectEvent;

    procedure CreateQualifiedName;
    procedure CheckRegister;
    // check register or unregister

  public
    //**************************************************************************
    // this public functions and propertys are used internal
    //**************************************************************************
    function GetOPCNode(stItemID: string; Ident: integer = -1): sOPCNode;

    procedure ItemWrite(OPCNode: sOPCNode; Value: variant; Path: string; DataType: TVarType);

    procedure ItemRead(OPCNode: sOPCNode; Path: string; DataType: TVarType);

    procedure InitAddressSpace;

    function GetIndexOfNode(Ident: integer): integer;

    function GetIndexOfstItemID(stItemID: string): integer;

    procedure UpAndDown(aUp: boolean);

    procedure AddDataAccessServer(Server: TObject);
    // add new Data Access Server

    procedure RemoveDataAccessServer(Server: TObject);
    // removes Data Access Server

    property CanStart: boolean read FCanStart;
    // OPC can be started

    property Nodes: sNodeArray read FNodes;
    // all defined nodes

  public
    //**************************************************************************
    // this public functions and propertys are used external
    //**************************************************************************
    constructor Create;

    destructor Destroy; override;

    procedure Init(
      ServerName: string;
      ServerDescription: string;
      SetExePath: boolean = True;
      LogFlag: boolean = False);
    // init OPC
    // ServerName is the name of the OPC server e.g. MyVendor.Servername.1
    // ServerDescription is the OPC server description
    // if SetExePath = True then the actual path is set to the path where the
    // exe file is loaded.
    // if LogFlag = True then OPCLog is activated (for debug sessions only)

    procedure Start;
    // OPC can be started

    procedure ShutDown;
    // shuts down the OPC Server

    procedure AddBranch(Parent, Ident: integer; stItemID: string);
    // add a branch to name space

    procedure AddLeaf(Parent, Ident: integer; stItemID: string;
      vtDataType: TVarType; dwAccessRights: DWORD);
    // add a leaf (tag) to name space

    procedure AddProperty(Parent: integer; dwPropertyID: DWORD; stItemID,
      stDescription: string; vtDataType: TVarType; vPropertyData: variant);
    // add a property to name the space

    function GetServerCount: integer;
    // returns number of connected DataAccess servers

    function GetGroupCount: integer;
    // returns number of connected groups (private and public)

    property ServerName: string read FServerName;
    // OPC server name

    property ServerDescription: string read FServerDescription;
    // OPC server description

    property OnRead: sOnRead read FOnRead write FOnRead;
    // client reads data

    property OnWrite: sOnWrite read FOnWrite write FOnWrite;
    // client writes data

    property OnSetDataType: sOnSetDataType read FOnSetDataType write FOnSetDataType;
    // client wants to change data type, True -> Ok DataType is changed

    property OnInitAddressSpace: sOnInitAddressSpace read FOnInitAddressSpace write FOnInitAddressSpace;
    // create name space

    property OnConnect: sOnConnectEvent read FOnConnect write FOnConnect;
    // on connect event
  end;

var
  OPC: sOPC;

implementation

uses
  ComServ, Forms,
  uLogging,
  uGlobals, uRegister, uOPCDataAccess;

procedure sOPC.CreateQualifiedName;
var
  i, j, Parent: integer;
begin
  for i := 0 to High(FNodes) do begin
    FNodes[i].stItemID := FNodes[i].stItemDataID;
    Parent := FNodes[i].Parent;
    while True do begin
      j := GetIndexOfNode(Parent);
      if j = -1 then break;
      FNodes[i].stItemID := FNodes[j].stItemDataID + '.' + FNodes[i].stItemID;
      Parent := FNodes[j].Parent;
    end;
  end;
end;

constructor sOPC.Create;
begin
  FCS := TCriticalSection.Create;
  FDataAccessServers := TList.Create;
  FAddressSpaceInit := False;
  FCanStart := False;
  FOnRead := nil;
  FOnWrite := nil;
  FOnSetDataType := nil;
  FOnInitAddressSpace := nil;
  FOnConnect := nil;
end;

destructor sOPC.Destroy;
var
  i: integer;
begin
  for i := 0 to High(FNodes) do FNodes[i].Free;
  if Assigned(Logging) then Logging.Free;
  FDataAccessServers.Free;
  FCS.Free;
  inherited;
end;

procedure sOPC.CheckRegister;
begin
  case ComServer.StartMode of
    smRegServer: RegisterOPCServer;
    smUnregServer: UnRegisterOPCServer;
  end;
end;

procedure sOPC.Init(
  ServerName: string;
  ServerDescription: string;
  SetExePath: boolean = True;
  LogFlag: boolean = False);
begin
  FServerName := ServerName;
  FServerDescription := ServerDescription;
  if SetExePath then SetCurrentDir(ExtractFileDir(Application.ExeName));
  if LogFlag and (not Assigned(Logging)) then begin
    Logging := TLogging.Create('OPCLog.txt');
    Logging.WriteWithTimeStamp(True);
    OPCLog('Start Log');
  end;
end;

procedure sOPC.Start;
begin
  FCanStart := True;
end;

procedure sOPC.ShutDown;
var
  i: integer;
  l: Tlist;
begin
  FOnConnect := nil;
  // FDataAccessServers list is copied, because ShutDown Destroys the
  // DataAccessServer!
  l := TList.Create;
  FCS.Enter;
  try
    for i := 0 to FDataAccessServers.Count - 1 do l.Add(FDataAccessServers.Items[i]);
  finally
    FCS.Leave;
  end;
  for i := 0 to l.Count - 1 do sOPCDataAccess(l.Items[i]).ShutDown;
  l.Free;
end;

procedure sOPC.AddBranch(Parent, Ident: integer; stItemID: string);
var
  OPCNode: sOPCNode;
begin
  OPCNode := sOPCNode.Create;
  OPCNode.NodeType := 0;
  OPCNode.Ident := Ident;
  OPCNode.Parent := Parent;
  OPCNode.stItemDataID := stItemID;
  SetLength(FNodes, High(FNodes) + 2);
  FNodes[High(FNodes)] := OPCNode;
end;

procedure sOPC.AddLeaf(Parent, Ident: integer; stItemID: string;
  vtDataType: TVarType; dwAccessRights: DWORD);
// slAddressPathList: TStringList; als Parameter kommt später +++
var
  OPCNode: sOPCNode;
begin
  OPCNode := sOPCNode.Create;
  OPCNode.NodeType := 1;
  OPCNode.Ident := Ident;
  OPCNode.Parent := Parent;
  OPCNode.stItemDataID := stItemID;
  OPCNode.vtCanonicalDataType := vtDataType;
  OPCNode.dwAccessRights := dwAccessRights;
  SetLength(FNodes, High(FNodes) + 2);
  FNodes[High(FNodes)] := OPCNode;
end;

procedure sOPC.AddProperty(Parent: integer; dwPropertyID: DWORD; stItemID,
  stDescription: string; vtDataType: TVarType; vPropertyData: variant);
var
  OPCNode: sOPCNode;
begin
  OPCNode := sOPCNode.Create;
  OPCNode.NodeType := 2;
  OPCNode.Ident := -1;
  OPCNode.Parent := Parent;
  OPCNode.dwPropertyID := dwPropertyID;
  OPCNode.stItemDataID := stItemID;
  OPCNode.stDescription := stDescription;
  OPCNode.vtPropertyDataType := vtDataType;
  OPCNode.vPropertyData := vPropertyData;
  SetLength(FNodes, High(FNodes) + 2);
  FNodes[High(FNodes)] := OPCNode;
end;

function sOPC.GetOPCNode(stItemID: string; Ident: integer): sOPCNode;
var
  i: integer;
begin
  Result := nil;
  if Ident = -1 then begin
    for i := 0 to High(FNodes) do begin
      if FNodes[i].stItemID = stItemID then begin
        Result := FNodes[i];
        exit;
      end;
    end;
  end else begin
    for i := 0 to High(FNodes) do begin
      if FNodes[i].Ident = Ident then begin
        Result := FNodes[i];
        exit;
      end;
    end;
  end;
end;

function sOPC.GetServerCount: integer;
begin
  Result := FDataAccessServers.Count;
end;

function sOPC.GetGroupCount: integer;
var
  i: integer;
begin
  Result := 0;
  FCS.Enter;
  try
    for i := 0 to FDataAccessServers.Count - 1 do begin
      // count private groups
      inc(Result, sOPCDataAccess(FDataAccessServers.Items[i]).GroupCount(False));
      // count public groups
      inc(Result, sOPCDataAccess(FDataAccessServers.Items[i]).GroupCount(True));
    end;
  finally
    FCS.Leave;
  end;
end;

procedure sOPC.ItemWrite(OPCNode: sOPCNode; Value: variant; Path: string; DataType: TVarType);
begin
  if Assigned(FOnWrite) then begin
    try
      FOnWrite(OPCNode, Value, Path, DataType);
    except
      on E: Exception do OPCLogException('sOPC.ItemWrite', E);
    end;
  end else begin
    OPCNode.CurrentValue := Value;
    OPCNode.LastUpdate := Now;
  end;
end;

procedure sOPC.UpAndDown(aUp: boolean);
begin
  if not assigned(FOnConnect) then exit;
  // event is created when first DataAccess is connected or last is disconnected
  if (FDataAccessServers.Count = 1) then FOnConnect(aUp);
end;

procedure sOPC.ItemRead(OPCNode: sOPCNode; Path: string; DataType: TVarType);
begin
  try
    if Assigned(FOnRead) then FOnRead(OPCNode, Path, DataType);
  except
    on E: Exception do begin
      OPCLogException('sOPC.ItemRead', E);
      VarClear(OPCNode.CurrentValue);
      OPCNode.LastUpdate := Now;
      OPCNode.Quality := OPC_QUALITY_OUT_OF_SERVICE;
    end;
  end;
end;

// +++ OnSetDataType kommt später!

procedure sOPC.InitAddressSpace;
begin
  if FAddressSpaceInit then exit;
  FAddressSpaceInit := True;
  try
    if Assigned(FOnInitAddressSpace) then begin
      FOnInitAddressSpace;
    end else begin
      AddLeaf(0, 1, 'No address space defined!', VT_I4, OPC_READABLE + OPC_WRITEABLE);
    end;
  except
    on E: Exception do begin
      OPCLogException('InitAddressSpace', E);
      AddLeaf(0, 1, 'No address space defined (Exception)!', VT_I4, OPC_READABLE + OPC_WRITEABLE);
    end;
  end;
  CreateQualifiedName;
end;

function sOPC.GetIndexOfNode(Ident: integer): integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to High(Nodes) do begin
    if (FNodes[i].NodeType <> 2) and (FNodes[i].Ident = Ident) then begin
      Result := i;
      exit;
    end;
  end;
end;

function sOPC.GetIndexOfstItemID(stItemID: string): integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to High(FNodes) do begin
    if (FNodes[i].NodeType <> 2) and (FNodes[i].stItemID = stItemID) then begin
      Result := i;
      exit;
    end;
  end;
end;

procedure sOPC.AddDataAccessServer(Server: TObject);
begin
  FCS.Enter;
  try
    FDataAccessServers.Add(Server);
  finally
    FCS.Leave;
  end;
end;

procedure sOPC.RemoveDataAccessServer(Server: TObject);
begin
  FCS.Enter;
  try
    FDataAccessServers.Remove(Server);
  finally
    FCS.Leave;
  end;
end;

initialization

  OPC := sOPC.Create;

finalization

  OPC.CheckRegister;
  OPC.Free;

end.

