//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCGroup;

interface

uses
  Windows, ActiveX, ComObj, SysUtils, Dialogs, Classes, Axctrls, Forms,
  SyncObjs, ExtCtrls,
  uTimer, OPCtypes, OPCDA, OpcError, sOPC_TLB, uGlobals, uOPCNode, uOPCItem;

type
  sOPCGroup = class(
    TTypedComObject,
    IOPCGroup,
    IOPCItemMgt,
    IOPCGroupStateMgt,
    IOPCPublicGroupStateMgt,
    IOPCSyncIO,
    IConnectionPointContainer,
    IOPCAsyncIO2)

  protected
    FCriticalSection: TCriticalSection;
    FName: string;
    FActive: boolean;
    FRequestedUpdateRate: DWORD;
    FClientGroup: OPCHANDLE;
    FTimeBias: longint;
    FPercentDeadband: single;
    FLCID: DWORD;
    FDataAccess: TObject;              // OPC Data Access
    FPublicGroupFlag: boolean;         // True -> group is public
    FOPCLock: boolean;                 // OPC function calls are locked

    FOPCItemList: TList;
    AsyncIOList: TList;
    CallBackEnabled: boolean;

    FConnectionPoints: TConnectionPoints;
    FConnectionPoint: TConnectionPoint;
    FConnectEvent: TConnectEvent;
    FClientIUnknown: IUnknown;
    FDataCallback: pointer;
    FGroupTimer: TTimer;
    FTimer: sTimer;
    FEvaluationTimer: sTimer;

    // IOPCItemMgt
    function AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
      out ppAddResults: POPCITEMRESULTARRAY; out ppErrors: PResultList): HResult; stdcall;

    function ValidateItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY; bBlobUpdate: BOOL;
      out ppValidationResults: POPCITEMRESULTARRAY; out ppErrors: PResultList): HResult; stdcall;

    function RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      out ppErrors: PResultList): HResult; stdcall;

    function SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY; bActive: BOOL;
      out ppErrors: PResultList): HResult; stdcall;

    function SetClientHandles(dwCount: DWORD; phServer: POPCHANDLEARRAY; phClient: POPCHANDLEARRAY;
      out ppErrors: PResultList): HResult; stdcall;

    function SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY;
      pRequestedDatatypes: PVarTypeList; out ppErrors: PResultList): HResult; stdcall;

    function CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    // IOPCGroupStateMgt
    function GetState(out pUpdateRate: DWORD; out pActive: BOOL; out ppName: POleStr;
      out pTimeBias: Longint; out pPercentDeadband: Single; out pLCID: TLCID;
      out phClientGroup: OPCHANDLE; out phServerGroup: OPCHANDLE): HResult; overload; stdcall;

    function SetState(pRequestedUpdateRate: PDWORD; out pRevisedUpdateRate: DWORD; pActive: PBOOL;
      pTimeBias: PLongint; pPercentDeadband: PSingle; pLCID: PLCID;
      phClientGroup: POPCHANDLE): HResult; stdcall;

    function SetName(szName: POleStr): HResult; stdcall;

    function CloneGroup(szName: POleStr; const riid: TIID; out ppUnk:IUnknown): HResult; stdcall;

    // IOPCPublicGroupStateMgt
    function GetState(out pPublic: BOOL): HResult; overload; stdcall;

    function MoveToPublic: HResult; stdcall;

    // IOPCSyncIO
    function Read(dwSource: OPCDATASOURCE; dwCount: DWORD; phServer: POPCHANDLEARRAY;
      out ppItemValues: POPCITEMSTATEARRAY; out ppErrors: PResultList): HResult; overload; stdcall;

    function Write(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
      out ppErrors: PResultList): HResult; overload; stdcall;

    // IOPCAsyncIO2
    function Read(dwCount: DWORD; phServer: POPCHANDLEARRAY; dwTransactionID: DWORD;
      out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; overload; stdcall;

    function Write(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
      dwTransactionID: DWORD; out pdwCancelID: DWORD; out ppErrors: PResultList):
      HResult; overload; stdcall;

    function Refresh2(dwSource: OPCDATASOURCE; dwTransactionID: DWORD;
      out pdwCancelID: DWORD): HResult; stdcall;

    function Cancel2(dwCancelID:DWORD): HResult; stdcall;

    function SetEnable(bEnable: BOOL):HResult; stdcall;

    function GetEnable(out pbEnable: BOOL): HResult; stdcall;

    // IConnectionPointContainer
    function EnumConnectionPoints(out Enum: IEnumConnectionPoints): HResult; stdcall;

    function FindConnectionPoint(const iid: TIID; out cp: IConnectionPoint): HResult; stdcall;

    procedure ConnectEvent(const Sink: IUnknown; Connecting: Boolean);

    procedure SetRequestedUpdateRate(UpdateRate: DWORD);

    procedure EnterCriticalSection;
    procedure LeaveCriticalSection;

    procedure Timer(Sender: TObject);

  public
    constructor Create(Server: TObject);

    destructor Destroy; override;

    procedure Initialize; override;

    procedure Init(szName: string; bActive: BOOL; dwRequestedUpdateRate: DWORD;
      hClientGroup: OPCHANDLE; pTimeBias: PLongint; pPercentDeadband: PSingle;
      dwLCID: DWORD); virtual;

    function GetOPCItem(hServer: OPCHANDLE): sOPCItem;

    property Name: string read FName;
    property ClientIUnknown: IUnknown read FClientIUnknown;
    property ClientGroup: OPCHANDLE read FClientGroup;
    property OPCItemList: TList read FOPCItemList;
    property RequestedUpdateRate: DWORD read FRequestedUpdateRate;
    property PublicGroup: boolean read FPublicGroupFlag write FPublicGroupFlag;
    property DataCallback: pointer read FDataCallback;
    property DataAccess: TObject read FDataAccess;
    property OPCLock: boolean read FOPCLock write FOPCLock;

  end;

implementation

uses
  ComServ,
  uOPC, uOPCAsyncIO2, uOPCItemEnumerator, uOPCDataAccess;

//******************************************************************************
// IOPCItemMgt
//******************************************************************************
function sOPCGroup.AddItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  out ppAddResults: POPCITEMRESULTARRAY; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
  OPCNode: sOPCNode;
  ItemDef: POPCITEMDEF;
begin
  OPCLog(Format('sOPCGroup.AddItems - %d', [dwCount]));
  try
    ppAddResults := TaskMemAlloc(dwCount, mkItemResult, Result);
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if (ppAddResults = nil) or (ppErrors = nil) then begin
      TaskMemFree(ppAddResults);
      TaskMemFree(ppErrors);
      exit;
    end;

    EnterCriticalSection;
    try
      Result := S_OK;
      for i := 0 to integer(dwCount) - 1 do begin
        ItemDef := @POPCITEMDEFARRAY(@pItemArray^)[i];

        OPCNode := OPC.GetOPCNode(ItemDef.szItemID);
        if OPCNode = nil then begin
          Result := S_FALSE;
          ppErrors[i] := OPC_E_UNKNOWNITEMID;
          continue;
        end;

        // create and initialize OPCItem
        OPCItem := sOPCItem.Create(OPCNode);
        OPCItem.AccessPath := ItemDef.szAccessPath;
        OPCItem.Active := ItemDef.bActive;
        OPCItem.ClientHandle := ItemDef.hClient;
        if OPCItem.SetRequestedDataType(ItemDef.vtRequestedDataType) then begin
          ppErrors[i] := S_OK;
        end else begin
          ppErrors[i] := OPC_E_BADTYPE;
          Result := S_FALSE;
        end;
        // start FTimer, when first OPCItem is added
        if FOPCItemList.Count = 0 then FTimer.Start;
        FOPCItemList.Add(OPCItem);

        // set Results
        ppAddResults[i].hServer := OPCHANDLE(OPCItem);
        ppAddResults[i].vtCanonicalDataType := OPCNode.vtCanonicalDataType;
        ppAddResults[i].dwAccessRights := OPCNode.dwAccessRights;
        ppAddResults[i].dwBlobSize := OPCNode.dwBlobSize;
        ppAddResults[i].pBlob := OPCNode.pBlob;
      end;
    finally
      LeaveCriticalSection;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.AddItems - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.ValidateItems(dwCount: DWORD; pItemArray: POPCITEMDEFARRAY;
  bBlobUpdate: BOOL; out ppValidationResults: POPCITEMRESULTARRAY;
  out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCNode: sOPCNode;
  ItemDef: POPCITEMDEF;
begin
  OPCLog(Format('sOPCGroup.ValidateItems - %d', [dwCount]));
  try
    ppValidationResults := TaskMemAlloc(dwCount, mkItemResult, Result);
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if (ppValidationResults = nil) or (ppErrors = nil) then begin
      TaskMemFree(ppValidationResults);
      TaskMemFree(ppErrors);
      exit;
    end;

    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      ItemDef := @POPCITEMDEFARRAY(@pItemArray^)[i];

      OPCNode := OPC.GetOPCNode(ItemDef.szItemID);
      if OPCNode = nil then begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_UNKNOWNITEMID;
        continue;
      end;

      ppValidationResults[i].vtCanonicalDataType := OPCNode.vtCanonicalDataType;
      ppValidationResults[i].dwAccessRights := OPCNode.dwAccessRights;
      if bBlobUpdate then begin
        ppValidationResults[i].dwBlobSize := OPCNode.dwBlobSize;
        ppValidationResults[i].pBlob := OPCNode.pBlob;
      end;
      ppErrors[i] := S_OK;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.ValidateItems - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.RemoveItems(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.RemoveItems - %d', [dwCount]));
  try
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;
    EnterCriticalSection;
    try
      Result := S_OK;
      for i := 0 to integer(dwCount) - 1 do begin
        OPCItem := GetOPCItem(POPCHANDLEARRAY(@phServer^)[i]);
        if OPCItem <> nil then begin
          FOPCItemList.Remove(OPCItem);
          // stop FTimer, when last OPCItem is removed
          if FOPCItemList.Count = 0 then FTimer.Stop;
          OPCItem.Free;
          ppErrors[i] := S_OK;
        end else begin
          Result := S_FALSE;
          ppErrors[i] := OPC_E_INVALIDHANDLE;
        end;
      end;
    finally
      LeaveCriticalSection;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.RemoveItems - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetActiveState(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  bActive: BOOL; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.SetActiveState - %d', [dwCount]));
  try
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;
    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      OPCItem := GetOPCItem(POPCHANDLEARRAY(@phServer^)[i]);
      if OPCItem <> nil then begin
        OPCItem.Active := bActive;
        ppErrors[i] := S_OK;
      end else begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.SetActiveState - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetClientHandles(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  phClient: POPCHANDLEARRAY; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.SetClientHandles - %d', [dwCount]));
  try
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;
    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      OPCItem := GetOPCItem(POPCHANDLEARRAY(@phServer^)[i]);
      if OPCItem <> nil then begin
        OPCItem.ClientHandle := POPCHANDLEARRAY(@phClient^)[i];
        ppErrors[i] := S_OK;
      end else begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.SetClientHandles - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetDatatypes(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pRequestedDatatypes: PVarTypeList; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.SetDatatypes - %d', [dwCount]));
  try
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;
    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      OPCItem := GetOPCItem(POPCHANDLEARRAY(@phServer^)[i]);
      if OPCItem <> nil then begin
        if OPCItem.SetRequestedDataType(PVarTypeList(@pRequestedDatatypes^)[i])
          then ppErrors[i] := S_OK
          else ppErrors[i] := OPC_E_BADTYPE;
      end else begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.SetDatatypes - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.CreateEnumerator', []));
  try
    Result := S_OK;
    if (FOPCItemList = nil) or (FOPCItemList.Count = 0) then begin
      Result := S_FALSE;
      exit;
    end;
    ppUnk := sOPCItemEnumerator.Create(FOPCItemList);
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.CreateEnumerator - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCGroupStateMgt
//******************************************************************************
function sOPCGroup.GetState(out pUpdateRate: DWORD; out pActive: BOOL;
  out ppName: POleStr; out pTimeBias: Longint; out pPercentDeadband: Single;
  out pLCID: TLCID; out phClientGroup: OPCHANDLE; out phServerGroup: OPCHANDLE)
  :HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.GetState', []));
  try
    pUpdateRate := FRequestedUpdateRate;
    pActive := FActive;
    ppName := StringToLPOLESTR(FName);
    pTimeBias := FTimeBias;
    pPercentDeadband := FPercentDeadband;
    pLCID := FLCID;
    phClientGroup := FClientGroup;
    phServerGroup := OPCHANDLE(self);
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.GetState - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetState(pRequestedUpdateRate: PDWORD; out pRevisedUpdateRate: DWORD;
  pActive: PBOOL; pTimeBias: PLongint; pPercentDeadband: PSingle; pLCID: PLCID;
  phClientGroup: POPCHANDLE): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.SetState', []));
  try
    if Assigned(pRequestedUpdateRate) then SetRequestedUpdateRate(pRequestedUpdateRate^);
    if Assigned(pActive)              then FActive := pActive^;
    if Assigned(pTimeBias)            then FTimeBias := pTimeBias^;
    if Assigned(pPercentDeadband)     then FPercentDeadband := pPercentDeadband^;
    if Assigned(pLCID)                then FLCID := pLCID^;
    if Assigned(phClientGroup)        then FClientGroup := phClientGroup^;
    // return the closest update rate the server is able to provide for this group
    if (@pRevisedUpdateRate <> nil)   then pRevisedUpdateRate := FRequestedUpdateRate;
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.SetState - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetName(szName: POleStr): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.SetName - %s', [string(szName)]));
  FName := szName;
  Result := S_OK;
end;

function sOPCGroup.CloneGroup(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult; stdcall;
var
  newName: string;
  i: integer;
  OPCGroup: sOPCGroup;
  coi, OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.CloneGroup - %s', [string(szName)]));
  try
    if not (IsEqualIID(riid, IID_IOPCItemMgt) or IsEqualIID(riid, IID_IUnknown))
    then begin
      Result := E_NOINTERFACE;
      exit;
    end;

    newName := szName;
    if not sOPCDataAccess(FDataAccess).GenerateUniqueGroupName(newName) then begin
      Result := OPC_E_DUPLICATENAME;
      exit;
    end;

    // create group
    OPCGroup := sOPCGroup.Create(FDataAccess);
    sOPCDataAccess(FDataAccess).AddOPCGroup(OPCGroup);
    ppUnk := OPCGroup;
    // clone fields
    OPCGroup.FName := newName;
    OPCGroup.FActive := False;
    OPCGroup.FRequestedUpdateRate := FRequestedUpdateRate;
    OPCGroup.FClientGroup := FClientGroup;
    OPCGroup.FTimeBias := FTimeBias;
    OPCGroup.FPercentDeadband := FPercentDeadband;
    OPCGroup.FLCID := FLCID;
    OPCGroup.FPublicGroupFlag := False;
    OPCGroup.CallBackEnabled := CallBackEnabled;
    // clone OPCItem's
    for i := 0 to FOPCItemList.Count - 1 do begin
      coi := sOPCItem(FOPCItemList[i]);
      OPCItem := sOPCItem.Create(coi.OPCNode);
      coi.Copy(OPCItem);
      OPCGroup.FOPCItemList.Add(OPCItem);
    end;

    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.CloneGroup - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCPublicGroupStateMgt
//******************************************************************************
function sOPCGroup.GetState(out pPublic: BOOL): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.GetState', []));
  try
    pPublic := FPublicGroupFlag;
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.GetState - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.MoveToPublic: HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.MoveToPublic', []));
  try
    Result := E_FAIL;
    if FPublicGroupFlag then exit;
    Result := sOPCDataAccess(FDataAccess).ChangeGroupToPublic(FName);
    if Result = S_OK then FPublicGroupFlag := True;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.MoveToPublic - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCSyncIO
//******************************************************************************
function sOPCGroup.Read(dwSource: OPCDATASOURCE; dwCount: DWORD; phServer: POPCHANDLEARRAY;
  out ppItemValues: POPCITEMSTATEARRAY; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.Read - %d - %d', [integer(dwSource), dwCount]));
  try
    ppItemValues := TaskMemAlloc(dwCount, mkItemState, Result);
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if (ppItemValues = nil) or (ppErrors = nil) then begin
      TaskMemFree(ppItemValues);
      TaskMemFree(ppErrors);
      exit;
    end;

    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      OPCItem := GetOPCItem(PDWORDARRAY(@phServer^)[i]);
      if OPCItem <> nil then begin
        if OPCItem.ReadAble then begin
          if (dwSource = OPC_DS_DEVICE) or FActive then begin
            OPCItem.GetOPCItemState(dwSource, ppItemValues[i]);
            sOPCDataAccess(FDataAccess).LastDataUpdateToClient := Now;
          end else begin
            ppItemValues[i].wQuality := OPC_QUALITY_OUT_OF_SERVICE;
          end;
          ppErrors[i] := S_OK
        end else begin
          ppItemValues[i].wQuality := OPC_QUALITY_BAD;
          ppErrors[i] := OPC_E_BADRIGHTS;
          Result := S_FALSE;
        end;
      end else begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Read - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.Write(dwCount: DWORD; phServer: POPCHANDLEARRAY;
  pItemValues: POleVariantArray; out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCItem: sOPCItem;
begin
  OPCLog(Format('sOPCGroup.Write - %d', [dwCount]));
  try
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;

    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      OPCItem := GetOPCItem(PDWORDARRAY(@phServer^)[i]);
      if OPCItem <> nil then begin
        if not OPCItem.WriteAble then begin
          ppErrors[i] := OPC_E_BADRIGHTS
        end else begin
          OPCItem.Write(POleVariantArray(@pItemValues^)[i]);
          ppErrors[i] := S_OK
        end;
      end else begin
        Result := S_FALSE;
        ppErrors[i] := OPC_E_INVALIDHANDLE;
      end;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Write - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCAsyncIO2
//******************************************************************************
function sOPCGroup.Read(dwCount: DWORD; phServer: POPCHANDLEARRAY; dwTransactionID: DWORD;
  out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; stdcall;
// Reads are from Device and are not affected by the ACTIVE state of the group
// or item.
var
 AsyncIO2: sOPCAsyncIO2;
begin
  OPCLog(Format('sOPCGroup.Read - %d', [dwCount]));
  try
    if FClientIUnknown = nil then begin
      Result := CONNECT_E_NOCONNECTION;
      exit;
    end;

    if (dwCount = 0) or (dwTransactionID = 0) then begin
      Result := E_INVALIDARG;
      exit;
    end;

    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;

    AsyncIO2 := sOPCAsyncIO2.Create(self, omRead, dwTransactionID, dwCount, OPC_DS_DEVICE);
    pdwCancelID := AsyncIO2.CancelID;
    AsyncIO2.Copy_phServer(phServer);
    AsyncIOList.Add(AsyncIO2);

    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Read - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.Write(dwCount: DWORD; phServer: POPCHANDLEARRAY; pItemValues: POleVariantArray;
  dwTransactionID: DWORD; out pdwCancelID: DWORD; out ppErrors: PResultList): HResult; stdcall;
var
 AsyncIO2: sOPCAsyncIO2;
begin
  OPCLog(Format('sOPCGroup.Write - %d', [dwCount]));
  try
    if FClientIUnknown = nil then begin
      Result := CONNECT_E_NOCONNECTION;
      exit;
    end;
    if (dwCount = 0) or (dwTransactionID = 0) then begin
      Result := E_INVALIDARG;
      exit;
    end;

    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if ppErrors = nil then exit;

    AsyncIO2 := sOPCAsyncIO2.Create(self, omWrite, dwTransactionID, dwCount, OPC_DS_DEVICE);
    pdwCancelID := AsyncIO2.CancelID;
    AsyncIO2.Copy_phServer(phServer);
    AsyncIO2.Copy_pItemValues(pItemValues);
    AsyncIOList.Add(AsyncIO2);

    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Write - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.Refresh2(dwSource: OPCDATASOURCE; dwTransactionID: DWORD;
  out pdwCancelID: DWORD): HResult; stdcall;
var
  i: integer;
  AsyncIO2: sOPCAsyncIO2;
  OneActive: boolean;
begin
  OPCLog(Format('sOPCGroup.Refresh2 - %d', [integer(dwSource)]));
  try
    if FClientIUnknown = nil then begin
      Result := CONNECT_E_NOCONNECTION;
      exit;
    end;
    if (dwTransactionID = 0) then begin
      Result := E_INVALIDARG;
      exit;
    end;

    OneActive := False;
    for i := 0 to FOPCItemList.Count - 1 do begin
      if sOPCItem(FOPCItemList[i]).Active then begin
        OneActive := True;
        break;
      end;
    end;
    if (not FActive) or (not OneActive) then begin
      Result := E_FAIL;
      exit;
    end;

    AsyncIO2 := sOPCAsyncIO2.Create(self, omRefresh, dwTransactionID, 0, dwSource);
    pdwCancelID := AsyncIO2.CancelID;
    AsyncIOList.Add(AsyncIO2);
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Refresh2 - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.Cancel2(dwCancelID: DWORD): HResult; stdcall;
var
  i: integer;
begin
  OPCLog(Format('sOPCGroup.Cancel2 - %d', [dwCancelID]));
  try
    Result := E_FAIL;
    if (AsyncIOList = nil) or (AsyncIOList.Count = 0) then exit;
    for i := 0 to AsyncIOList.Count - 1 do begin
      if sOPCAsyncIO2(AsyncIOList[i]).CancelID = dwCancelID then begin
        Result := S_OK;
        sOPCAsyncIO2(AsyncIOList[i]).CancelFlag := True;
        break;
      end
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.Cancel2 - Exception', []), E);
    end;
  end;
end;

function sOPCGroup.SetEnable(bEnable: BOOL): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.SetEnable - %d', [longint(bEnable)]));
  CallBackEnabled := bEnable;
  Result := S_OK;
end;

function sOPCGroup.GetEnable(out pbEnable: BOOL): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.GetEnable', []));
  try
    pbEnable := CallBackEnabled;
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.GetEnable - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IConnectionPointContainer
//******************************************************************************
function sOPCGroup.EnumConnectionPoints(out Enum: IEnumConnectionPoints): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.EnumConnectionPoints - not implemented', []));
  Result := E_NOTIMPL;
end;

function sOPCGroup.FindConnectionPoint(const iid: TIID; out cp: IConnectionPoint): HResult; stdcall;
begin
  OPCLog(Format('sOPCGroup.FindConnectionPoint', []));
  try
    if (addr(cp) = nil) then begin
      Result := E_INVALIDARG;
      exit;
    end;

    if IsEqualGuid(iid, IID_IOPCDataCallback) then begin
      cp := FConnectionPoint;
      Result := S_OK;
    end else begin
      Result := E_NOINTERFACE;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCGroup.FindConnectionPoint - Exception', []), E);
    end;
  end;
end;

procedure sOPCGroup.Initialize;
begin
  inherited Initialize;
  FConnectionPoints := TConnectionPoints.Create(self);
  FConnectionPoint := TConnectionPoint.Create(FConnectionPoints, IID_IOPCDataCallback,
    ckSingle, FConnectEvent);
end;

constructor sOPCGroup.Create(Server: TObject);
begin
  OPCLog('sOPCGroup.Create');
  FCriticalSection := TCriticalSection.Create;
  FName := '';
  FActive := False;
  FRequestedUpdateRate := 0;
  FClientGroup := 0;
  FTimeBias := 0;
  FPercentDeadband := 0;
  FLCID := 0;
  FDataAccess := Server;
  FPublicGroupFlag := False;
  FOPCLock := False;
  FOPCItemList := TList.Create;
  AsyncIOList := TList.Create;
  CallBackEnabled := True;
  FConnectionPoints := nil;
  FConnectionPoint := nil;
  FConnectEvent := ConnectEvent;       // inherited create calls initialize!
  FClientIUnknown := nil;
  FDataCallback := nil;
  FTimer := sTimer.Create(False);
  FEvaluationTimer := sTimer.Create;
  FGroupTimer := TTimer.Create(nil);
  FGroupTimer.Interval := 50;
  FGroupTimer.Enabled := false;
  FGroupTimer.OnTimer := Timer;
  inherited Create;
end;

destructor sOPCGroup.Destroy;
var
 i: integer;
begin
  OPCLog('sOPCGroup.Destroy');
  FGroupTimer.Free;
  FConnectionPoint.Free;
  FConnectionPoints.Free;
  for i := 0 to FOPCItemList.Count - 1 do sOPCItem(FOPCItemList[i]).Free;
  FOPCItemList.Free;
  for i := 0 to AsyncIOList.Count - 1 do sOPCAsyncIO2(AsyncIOList[i]).Free;
  AsyncIOList.Free;
  FEvaluationTimer.Free;
  FTimer.Free;
  FCriticalSection.Free;
end;

procedure sOPCGroup.Init(szName: string; bActive: BOOL; dwRequestedUpdateRate: DWORD;
  hClientGroup: OPCHANDLE; pTimeBias: PLongint; pPercentDeadband: PSingle;
  dwLCID: DWORD);
begin
  FName := szName;
  FActive := bActive;
  SetRequestedUpdateRate(dwRequestedUpdateRate);
  FClientGroup := hClientGroup;
  if Assigned(pTimeBias) then FTimeBias := pTimeBias^;
  if Assigned(pPercentDeadband) then FPercentDeadband := pPercentDeadband^;
  FLCID := dwLCID;
end;

procedure sOPCGroup.ConnectEvent(const Sink: IUnknown; Connecting: Boolean);
begin
  if Connecting then begin
    FClientIUnknown := Sink;
    if not Succeeded(FClientIUnknown.QueryInterface(IOPCDataCallback,
      FDataCallback)) then FDataCallback := nil;
  end;
end;

procedure sOPCGroup.SetRequestedUpdateRate(UpdateRate: DWORD);
begin
  // lowest value for the update rate is 100 ms
  if UpdateRate < 100 then UpdateRate := 100;
  FRequestedUpdateRate := UpdateRate;
  FGroupTimer.Enabled := true;
end;

function sOPCGroup.GetOPCItem(hServer: OPCHANDLE): sOPCItem;
begin
  result := sOPCItem(hServer);
end;

procedure sOPCGroup.Timer(Sender: TObject);
const
  inTimer: boolean = false;
var
  AsyncIO2: sOPCAsyncIO2;
begin
  if inTimer then exit;
  inTimer := true;
  try
    // asynchronous I/O's
    while (AsyncIOList.Count > 0) do begin
      try
        AsyncIO2 := sOPCAsyncIO2(AsyncIOList[0]);
        AsyncIO2.ProcessRequest(CallBackEnabled);
        AsyncIO2.Free;
      except
        on E: Exception do OPCLogException('sOPCGroup.Timer 1', E);
      end;
      AsyncIOList.Delete(0);
    end;

    // Timer Refresh
    if FTimer.isRunning and (FTimer.msTime > FRequestedUpdateRate) then begin
      FTimer.Start;
      try
        if (not FActive) or (FOPCItemList.Count = 0) then exit;
        AsyncIO2 := sOPCAsyncIO2.Create(self, omTimerRefresh, 0, 0, 0);
        AsyncIO2.ProcessRequest(CallBackEnabled);
        AsyncIO2.Free;
      except
        on E: Exception do OPCLogException('sOPCGroup.Timer 2', E);
      end;
    end;
  finally
    inTimer := false;
  end;
end;

procedure sOPCGroup.EnterCriticalSection;
begin
  while True do begin
    FCriticalSection.Enter;
    if FOPCLock then begin
      FCriticalSection.Leave;
      Sleep(20);
    end else begin
      break;
    end;
  end;
end;

procedure sOPCGroup.LeaveCriticalSection;
begin
  FCriticalSection.Leave;
end;

initialization

TTypedComObjectFactory.Create(
  ComServer,
  sOPCGroup,
  Class_OPCGroup,
  ciInternal,
  ThreadingModel);

end.

