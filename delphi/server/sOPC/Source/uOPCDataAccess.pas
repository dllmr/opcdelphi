//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCDataAccess;

interface

uses
  Windows, ComObj, ActiveX, Axctrls, SysUtils, Dialogs, Classes, StdVCL, SyncObjs,
  OPCtypes, OPCDA, OPCCOMN, OpcError, uGlobals, uOPCNode, uOPCGroup,
  uOPCStringEnumerator, sOPC_TLB, uOPC, uOPCBrowse;

type
  sOPCDataAccess = class(
    TAutoObject,
    IOPCDataAccess20,
    IOPCServer,
    IOPCCommon,
    IOPCServerPublicGroups,
    IOPCBrowseServerAddressSpace,
    IPersist,
    IPersistFile,
    IConnectionPointContainer,
    IOPCItemProperties)

  protected
    // IPersistFile
    function GetClassID(out classID: TCLSID): HResult; stdcall;

    function IsDirty: HResult; stdcall;

    function Load(pszFileName: POleStr; dwMode: Longint): HResult; stdcall;

    function Save(pszFileName: POleStr; fRemember: BOOL): HResult; stdcall;

    function SaveCompleted(pszFileName: POleStr): HResult; stdcall;

    function GetCurFile(out pszFileName: POleStr): HResult; stdcall;

    // IConnectionPointContainer
    function EnumConnectionPoints(out Enum: IEnumConnectionPoints): HResult; stdcall;

    function FindConnectionPoint(const iid: TIID; out cp: IConnectionPoint): HResult; stdcall;

    // IOPCServer
    function AddGroup(szName:POleStr; bActive: BOOL; dwRequestedUpdateRate: DWORD;
      hClientGroup: OPCHANDLE; pTimeBias: PLongint; pPercentDeadband: PSingle;
      dwLCID: DWORD; out phServerGroup: OPCHANDLE; out pRevisedUpdateRate: DWORD;
      const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;

    function GetErrorString(dwError: HResult; dwLocale: TLCID;
      out ppString: POleStr): HResult; overload; stdcall;

    function GetGroupByName(szName: POleStr; const riid: TIID;
      out ppUnk: IUnknown): HResult; stdcall;

    function GetStatus(out ppServerStatus: POPCSERVERSTATUS): HResult; stdcall;

    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;

    function CreateGroupEnumerator(dwScope: OPCENUMSCOPE; const riid: TIID;
      out ppUnk: IUnknown): HResult; stdcall;

    // IOPCCommon
    function SetLocaleID(dwLcid: TLCID): HResult; stdcall;

    function GetLocaleID(out pdwLcid: TLCID): HResult; stdcall;

    function QueryAvailableLocaleIDs(out pdwCount: UINT; out pdwLcid: PLCIDARRAY): HResult; stdcall;

    function GetErrorString(dwError: HResult; out ppString: POleStr): HResult; overload; stdcall;

    function SetClientName(szName: POleStr): HResult; stdcall;

    // IOPCServerPublicGroups
    function GetPublicGroupByName(szName: POleStr; const riid: TIID;
      out ppUnk: IUnknown): HResult; stdcall;

    function RemovePublicGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;

    // IOPCBrowseServerAddressSpace
    function QueryOrganization(out pNameSpaceType: OPCNAMESPACETYPE): HResult; stdcall;

    function ChangeBrowsePosition(dwBrowseDirection: OPCBROWSEDIRECTION;
      szString: POleStr): HResult; stdcall;

    function BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE; szFilterCriteria: POleStr;
      vtDataTypeFilter: TVarType; dwAccessRightsFilter: DWORD;
      out ppIEnumString: IEnumString): HResult; stdcall;

    function GetItemID(szItemDataID: POleStr; out szItemID: POleStr): HResult; stdcall;

    function BrowseAccessPaths(szItemID: POleStr; out ppIEnumString: IEnumString): HResult; stdcall;

    // IOPCItemProperties
    function QueryAvailableProperties(szItemID: POleStr; out pdwCount: DWORD;
      out ppPropertyIDs: PDWORDARRAY; out ppDescriptions: POleStrList;
      out ppvtDataTypes: PVarTypeList): HResult; stdcall;

    function GetItemProperties(szItemID: POleStr; dwCount: DWORD;
      pdwPropertyIDs: PDWORDARRAY; out ppvData: POleVariantArray;
      out ppErrors: PResultList): HResult; stdcall;

    function LookupItemIDs(szItemID: POleStr; dwCount: DWORD; pdwPropertyIDs: PDWORDARRAY;
      out ppszNewItemIDs: POleStrList; out ppErrors: PResultList): HResult; stdcall;

  protected
    FOPCBrowse: sOPCBrowse;
    FCS: TCriticalSection;
    FOPCGroup: TList;
    FLcid: TLCID;
    FClientName: string;
    FStartTime: TDateTime;
    FLastDataUpdateToClient: TDateTime;

    FConnectionPoints: TConnectionPoints;
    FShutdown: TConnectionPoint;
    FConnectEvent: TConnectEvent;
    FClientIUnknown: IUnknown;

    procedure ConnectEvent(const Sink: IUnknown; Connecting: Boolean);

    procedure RemoveGroupByHandle(hServerGroup: OPCHANDLE);
    // removes the group 'hServerGroup' out of the group list and free's the group

    function CreateGroupNameList(Mode: integer): TStringList; virtual;
    // create a list of group names, the calling method must release the list!
    // Mode: 0 = list of private groups, 1 = list of public groups
    //       2 = list of private and public groups

    function FindGroupByName(Name: string; var ix: integer): sOPCGroup; virtual;
    // returns Group if 'Name' is found in private or public list

  public
    procedure Initialize; override;

    destructor Destroy; override;

    function ShutDown: boolean;

    function GroupCount(PublicFlag: boolean): integer; virtual;

    function GenerateUniqueGroupName(var Name: string): boolean;
    // generates a unique group name ouf of 'Name'
    // True -> new name generated

    function ChangeGroupToPublic(Name: string): HResult;

    procedure AddOPCGroup(OPCGroup: sOPCGroup);

    property LastDataUpdateToClient: TDateTime read FLastDataUpdateToClient
      write FLastDataUpdateToClient;
  end;

implementation

uses
  ComServ,
  uRegister;

//******************************************************************************
// IPersistFile
//******************************************************************************
function sOPCDataAccess.GetClassID(out classID: TCLSID): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.GetClassID - not implemented');
  Result := S_FALSE;
end;

function sOPCDataAccess.IsDirty: HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.IsDirty - not implemented');
  Result := S_FALSE;
end;

function sOPCDataAccess.Load(pszFileName: POleStr; dwMode: Longint): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.Load - not implemented');
  Result := S_FALSE;
end;

function sOPCDataAccess.Save(pszFileName: POleStr; fRemember: BOOL): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.Save - not implemented');
  Result := S_FALSE;
end;

function sOPCDataAccess.SaveCompleted(pszFileName: POleStr): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.SaveCompleted - not implemented');
  Result := S_FALSE;
end;

function sOPCDataAccess.GetCurFile(out pszFileName: POleStr): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.GetCurFile - not implemented');
  Result := S_FALSE;
end;

//******************************************************************************
// IConnectionPointContainer
//******************************************************************************
function sOPCDataAccess.EnumConnectionPoints(out Enum: IEnumConnectionPoints):HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.EnumConnectionPoints - not implemented');
  Result := E_NOTIMPL;
end;

function sOPCDataAccess.FindConnectionPoint(const iid: TIID; out cp: IConnectionPoint): HResult; stdcall;
begin
  OPCLog('sOPCDataAccess.FindConnectionPoint');
  if (addr(cp) = nil) then begin
    Result := E_INVALIDARG;
    exit;
  end;
  if IsEqualGuid(iid, IID_IOPCShutdown) then begin
    cp := FShutdown;
    Result := S_OK;
  end else
    Result := E_NOINTERFACE;
end;

//******************************************************************************
// IOPCServerIMPL
//******************************************************************************
function sOPCDataAccess.AddGroup(szName: POleStr; bActive: BOOL;
  dwRequestedUpdateRate: DWORD; hClientGroup: OPCHANDLE; pTimeBias: PLongint;
  pPercentDeadband: PSingle; dwLCID: DWORD; out phServerGroup: OPCHANDLE;
  out pRevisedUpdateRate: DWORD; const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
var
  newName: string;
  OPCGroup: sOPCGroup;
begin
  OPCLog(Format('sOPCDataAccess.AddGroup - %s', [szName]));
  try
    newName := szName;
    if not GenerateUniqueGroupName(newName) then begin
      Result := OPC_E_DUPLICATENAME;
      exit;
    end;

    OPCGroup := sOPCGroup.Create(self);
    if OPCGroup = nil then begin
      Result := E_OUTOFMEMORY;
      phServerGroup := 0;
      exit;
    end;

    FCS.Enter;
    try
      FOPCGroup.Add(OPCGroup);
    finally
      FCS.Leave;
    end;
    OPCGroup.Init(newName, bActive, dwRequestedUpdateRate, hClientGroup, pTimeBias, pPercentDeadband, dwLCID);
    phServerGroup := OPCHANDLE(OPCGroup);
    pRevisedUpdateRate := OPCGroup.RequestedUpdateRate;

    ppUnk := OPCGroup;
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.AddGroup - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.GetErrorString(dwError: HResult; dwLocale: TLCID;
  out ppString: POleStr): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.GetErrorString 1 - %d', [longint(dwError)]));
  ppString := StringToLPOLESTR(Format('Error %d', [longint(dwError)]));
  Result := S_OK;
end;

function sOPCDataAccess.GetGroupByName(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult; stdcall;
var
  OPCGroup: sOPCGroup;
  ix: integer;
begin
  OPCLog(Format('sOPCDataAccess.GetGroupByName - %s', [string(szName)]));
  try
    OPCGroup := FindGroupByName(szName, ix);
    if (addr(ppUnk) = nil) or (OPCGroup = nil) then begin
      Result := E_INVALIDARG;
      exit;
    end;
    Result := IUnknown(OPCGroup).QueryInterface(riid, ppUnk);
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.GetGroupByName - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.GetStatus(out ppServerStatus: POPCSERVERSTATUS): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.GetStatus', []));
  try
    if (addr(ppServerStatus) = nil) then begin
      Result := E_INVALIDARG;
      exit;
    end;

    ppServerStatus := TaskMemAlloc(0, mkServerStatus, Result);
    if ppServerStatus = nil then exit;

    Result := S_OK;
    ppServerStatus.ftStartTime      := ConvertToFileTime(FStartTime);
    ppServerStatus.ftCurrentTime    := ConvertToFileTime(Now);
    ppServerStatus.ftLastUpdateTime := ConvertToFileTime(FLastDataUpdateToClient);
    ppServerStatus.dwServerState    := OPC_STATUS_RUNNING;
    ppServerStatus.dwGroupCount     := GroupCount(False) + GroupCount(True);
    ppServerStatus.dwBandWidth      := $FFFFFFFF;
    ppServerStatus.wMajorVersion    := 1;
    ppServerStatus.wMinorVersion    := 1;
    ppServerStatus.wBuildNumber     := 1;
    ppServerStatus.szVendorInfo     := 'sOPC created by Schmid IT-Management, http://www.schmid-itm.de/';

  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.GetStatus - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
// This function should not be called for Public Groups!
begin
  OPCLog(Format('sOPCDataAccess.RemoveGroup', []));
  FOPCGroup.Remove(pointer(hServerGroup));
  Result := S_OK;
end;

function sOPCDataAccess.CreateGroupEnumerator(dwScope: OPCENUMSCOPE;
  const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
var
  Mode: integer;
begin
  OPCLog(Format('sOPCDataAccess.CreateGroupEnumerator - %d', [longint(dwScope)]));
  try
    Result := S_OK;
    Mode := 0;
    case dwScope of
      OPC_ENUM_PRIVATE_CONNECTIONS, OPC_ENUM_PRIVATE: Mode := 0;
      OPC_ENUM_PUBLIC_CONNECTIONS, OPC_ENUM_PUBLIC:   Mode := 1;
      OPC_ENUM_ALL_CONNECTIONS, OPC_ENUM_ALL:         Mode := 2;
    end;
    ppUnk := sOPCStringEnumerator.Create(CreateGroupNameList(Mode));
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.CreateGroupEnumerator - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCCommonIMPL
//******************************************************************************
function sOPCDataAccess.SetLocaleID(dwLcid: TLCID): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.SetLocaleID', []));
  FLcid := dwLcid;
  Result := S_OK;
end;

function sOPCDataAccess.GetLocaleID(out pdwLcid: TLCID): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.GetLocaleID', []));
  pdwLcid := FLcid;
  Result := S_OK;
end;

function sOPCDataAccess.QueryAvailableLocaleIDs(out pdwCount: UINT;
  out pdwLcid: PLCIDARRAY): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.QueryAvailableLocaleIDs', []));
  Result := S_FALSE;
end;

function sOPCDataAccess.GetErrorString(dwError: HResult; out ppString: POleStr): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.GetErrorString 2 - %d', [longint(dwError)]));
  ppString := StringToLPOLESTR(Format('Data Access Error %d ', [longint(dwError)]));
  Result := S_OK;
end;

function sOPCDataAccess.SetClientName(szName: POleStr): HResult; stdcall;
begin
  if (addr(szName) = nil) then begin
    Result := E_INVALIDARG;
      OPCLog(Format('sOPCDataAccess.SetClientName - Invalid Argument', []));
    exit;
  end;
  OPCLog(Format('sOPCDataAccess.SetClientName - %s', [string(szName)]));
  FClientName := szName;
  Result := S_OK;
end;

//******************************************************************************
// IOPCServerPublicGroupsIMPL
//******************************************************************************
function sOPCDataAccess.GetPublicGroupByName(szName: POleStr; const riid: TIID;
  out ppUnk: IUnknown): HResult; stdcall;
var
  OPCGroup: sOPCGroup;
  ix: integer;
begin
  try
    OPCGroup := FindGroupByName(szName, ix);
    if (addr(ppUnk) = nil) or (OPCGroup = nil) then begin
      Result := E_INVALIDARG;
      OPCLog(Format('sOPCDataAccess.GetPublicGroupByName - Invalid Argument', []));
      exit;
    end;
    OPCLog(Format('sOPCDataAccess.GetPublicGroupByName - %s', [string(szName)]));
    Result := IUnknown(OPCGroup).QueryInterface(riid, ppUnk);
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.GetPublicGroupByName - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.RemovePublicGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.RemovePublicGroup', []));
  try
    RemoveGroupByHandle(hServerGroup);
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.RemovePublicGroup - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCBrowseServerAddressSpaceIMPL
//******************************************************************************
function sOPCDataAccess.QueryOrganization(out pNameSpaceType: OPCNAMESPACETYPE): HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.QueryOrganization', []));
  pNameSpaceType := OPC_NS_HIERARCHIAL;
  Result := S_OK;
end;

function sOPCDataAccess.ChangeBrowsePosition(dwBrowseDirection: OPCBROWSEDIRECTION;
  szString: POleStr): HResult; stdcall;
var
  St: string;
begin
  try
    if not Assigned(szString)
      then St := ''
      else St := szString;
    OPCLog(Format('sOPCDataAccess.ChangeBrowsePosition - %d - s', [longint(dwBrowseDirection), St]));
    Result := S_OK;
    case dwBrowseDirection of
      OPC_BROWSE_UP: if not FOPCBrowse.BrowseUp then Result := E_FAIL;
      OPC_BROWSE_DOWN: if not FOPCBrowse.BrowseDown(St) then Result := E_FAIL;
      OPC_BROWSE_TO: if not FOPCBrowse.BrowseTo(St) then Result := E_INVALIDARG;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.ChangeBrowsePosition - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.BrowseOPCItemIDs(dwBrowseFilterType: OPCBROWSETYPE;
  szFilterCriteria: POleStr; vtDataTypeFilter: TVarType; dwAccessRightsFilter: DWORD;
  out ppIEnumString: IEnumString): HResult; stdcall;
var
  List: TStringList;
begin
  OPCLog(Format('sOPCDataAccess.BrowseOPCItemIDs - %d', [longint(dwBrowseFilterType)]));
  try
    Result := S_OK;
    List := FOPCBrowse.BrowseOPCItems(dwBrowseFilterType);
    ppIEnumString := sOPCStringEnumerator.Create(List);
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.BrowseOPCItemIDs - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.GetItemID(szItemDataID: POleStr; out szItemID: POleStr):
  HResult; stdcall;
begin
  OPCLog(Format('sOPCDataAccess.GetItemID - %s', [string(szItemDataID)]));
  try
    szItemID := StringToLPOLESTR(FOPCBrowse.GetItemID(szItemDataID));
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.GetItemID - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.BrowseAccessPaths(szItemID: POleStr;
  out ppIEnumString: IEnumString): HResult; stdcall;
var
  OPCNode: sOPCNode;
begin
  OPCLog(Format('sOPCDataAccess.BrowseAccessPaths - %s', [string(szItemID)]));
  try
    Result := S_OK;
    OPCNode := OPC.GetOPCNode(szItemID);
    if OPCNode.slAddressPathList = nil
      then Result := S_False
      else ppIEnumString := sOPCStringEnumerator.Create(OPCNode.slAddressPathList, False);
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.BrowseAccessPaths - Exception', []), E);
    end;
  end;
end;

//******************************************************************************
// IOPCItemProperties
//******************************************************************************
function sOPCDataAccess.QueryAvailableProperties(
  szItemID: POleStr;
  out pdwCount: DWORD;
  out ppPropertyIDs: PDWORDARRAY;
  out ppDescriptions: POleStrList;
  out ppvtDataTypes: PVarTypeList): HResult; stdcall;
var
  i: integer;
  OPCNode: sOPCNode;
begin
  OPCLog(Format('sOPCDataAccess.QueryAvailableProperties - %s', [string(szItemID)]));
  try
    OPCNode := OPC.GetOPCNode(szItemID);
    if OPCNode = nil then begin
      Result := OPC_E_INVALIDITEMID;
      exit;
    end;

    pdwCount := FOPCBrowse.GetPropertyCount(szItemID);
    if pdwCount = 0 then exit;

    ppPropertyIDs := TaskMemAlloc(pdwCount, mkDWORD, Result);
    ppDescriptions := TaskMemAlloc(pdwCount, mkPOleStr, Result);
    ppvtDataTypes := TaskMemAlloc(pdwCount, mkVarType, Result);
    if (ppPropertyIDs = nil) or (ppDescriptions = nil) or (ppvtDataTypes = nil) then begin
      TaskMemFree(ppPropertyIDs);
      TaskMemFree(ppDescriptions);
      TaskMemFree(ppvtDataTypes);
      exit;
    end;

    Result := S_OK;
    for i := 0 to pdwCount - 1 do begin
      OPCNode := FOPCBrowse.GetProperty(szItemID, i);
      ppPropertyIDs[i] := OPCNode.dwPropertyID;
      ppDescriptions[i] := StringToLPOLESTR(OPCNode.stDescription);
      ppvtDataTypes[i] := OPCNode.vtPropertyDataType;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.QueryAvailableProperties - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.GetItemProperties(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppvData: POleVariantArray;
  out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCNode: sOPCNode;
begin
  OPCLog(Format('sOPCDataAccess.GetItemProperties - %s', [string(szItemID)]));
  try
    OPCNode := OPC.GetOPCNode(szItemID);
    if OPCNode = nil then begin
      Result := OPC_E_INVALIDITEMID;
      exit;
    end;

    ppvData := TaskMemAlloc(dwCount, mkOleVariant, Result);
    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    if (ppvData = nil) or (ppErrors = nil) then begin
      TaskMemFree(ppvData);
      TaskMemFree(ppErrors);
      exit;
    end;

    Result := S_OK;
    for i := 0 to integer(dwCount) - 1 do begin
      ppvData[i] := FOPCBrowse.GetPropertyData(szItemID, PDWORDARRAY(pdwPropertyIDs)[i]);
      ppErrors[i] := S_OK;
    end;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.GetItemProperties - Exception', []), E);
    end;
  end;
end;

function sOPCDataAccess.LookupItemIDs(szItemID: POleStr; dwCount: DWORD;
  pdwPropertyIDs: PDWORDARRAY; out ppszNewItemIDs: POleStrList;
  out ppErrors: PResultList): HResult; stdcall;
var
  i: integer;
  OPCNode: sOPCNode;
begin
  OPCLog(Format('sOPCDataAccess.LookupItemIDs - %s - %d', [string(szItemID), dwCount]));
  try
    OPCNode := OPC.GetOPCNode(szItemID);
    if (OPCNode = nil) then begin
      Result := OPC_E_INVALIDITEMID;
      exit;
    end;
    if dwCount = 0 then begin
      Result := E_INVALIDARG;
      exit;
    end;

    ppErrors := TaskMemAlloc(dwCount, mkHResult, Result);
    ppszNewItemIDs := TaskMemAlloc(dwCount, mkPOleStr, Result);
    if (ppErrors = nil) or (ppszNewItemIDs = nil) then begin
      TaskMemFree(ppErrors);
      TaskMemFree(ppszNewItemIDs);
      exit;
    end;

    for i := 0 to integer(dwCount) - 1 do begin
      ppszNewItemIDs[i] := nil; // StringToLPOLESTR('...');  // +++ später
      ppErrors[i] := S_OK;
    end;
    Result := S_OK;
  except
    on E: Exception do begin
      Result := E_UNEXPECTED;
      OPCLogException(Format('sOPCDataAccess.LookupItemIDs - Exception', []), E);
    end;
  end;
end;

procedure sOPCDataAccess.Initialize;
begin
  OPCLog('sOPCDataAccess.Initialize - ' + InttoStr(longint(self)));
  // wait until the application is running
  repeat
    Sleep(100);
  until OPC.CanStart;
  inherited Initialize;

  FStartTime := Now;
  FLastDataUpdateToClient := 0;

  FConnectEvent := ConnectEvent;
  FConnectionPoints := TConnectionPoints.Create(self);
  FShutdown := TConnectionPoint.Create(FConnectionPoints, IID_IOPCShutdown,
    ckSingle, FConnectEvent);

  FCS := TCriticalSection.Create;
  FOPCGroup := TList.Create;
  OPC.InitAddressSpace;
  FOPCBrowse := sOPCBrowse.Create;

  OPC.AddDataAccessServer(self);
  OPC.UpAndDown(true);
end;

procedure sOPCDataAccess.ConnectEvent(const Sink: IUnknown; Connecting: Boolean);
begin
  OPCLog(Format('sOPCDataAccess.ConnectEvent - %d', [longint(Connecting)]));
  if not Connecting then OPC.UpAndDown(Connecting);
  if Connecting then FClientIUnknown := Sink;
end;

destructor sOPCDataAccess.Destroy;
var
  i: integer;
begin
  OPCLog('sOPCDataAccess.Destroy - ' + InttoStr(longint(self)));
  for i := 0 to FOPCGroup.Count - 1 do sOPCGroup(FOPCGroup.Items[i]).Free;
  FShutdown.Free;
  FConnectionPoints.Free;
  FOPCGroup.Free;
  FCS.Free;
  FOPCBrowse.Free;
  OPC.RemoveDataAccessServer(self);
  inherited;
end;

function sOPCDataAccess.ShutDown: boolean;
var
  Obj: pointer;
begin
  Result := False;
  if (FClientIUnknown <> nil) then begin
    if Succeeded(FClientIUnknown.QueryInterface(IOPCShutdown, Obj)) then begin
      if Assigned(Obj) then IOPCShutdown(Obj).ShutdownRequest('Terminated by User.');
      Result := True;
    end;
  end;
end;

procedure sOPCDataAccess.RemoveGroupByHandle(hServerGroup: OPCHANDLE);
begin
  FCS.Enter;
  try
    FOPCGroup.Delete(hServerGroup);
  finally
    FCS.Leave;
  end;
end;

function sOPCDataAccess.GroupCount(PublicFlag: boolean): integer;
var
  i: integer;
  OPCGroup: sOPCGroup;
begin
  Result := 0;
  FCS.Enter;
  try
    for i := 0 to FOPCGroup.Count - 1 do begin
      OPCGroup := sOPCGroup(FOPCGroup.Items[i]);
      if not (PublicFlag xor OPCGroup.PublicGroup) then inc(Result);
    end;
  finally
    FCS.Leave;
  end;
end;

function sOPCDataAccess.CreateGroupNameList(Mode: integer): TStringList;
var
  i: integer;
  OPCGroup: sOPCGroup;
begin
  Result := TStringList.Create;
  FCS.Enter;
  try
    for i := 0 to FOPCGroup.Count - 1 do begin
      OPCGroup := sOPCGroup(FOPCGroup.Items[i]);
      if (Mode = 2) or
        not ((Mode = 0) xor (not OPCGroup.PublicGroup) or
        not ((Mode = 1) xor OPCGroup.PublicGroup))
      then begin
        Result.Add(OPCGroup.Name);
      end;
    end;
  finally
    FCS.Leave;
  end;
end;

function sOPCDataAccess.FindGroupByName(Name: string; var ix: integer): sOPCGroup;
var
  i: integer;
  OPCGroup: sOPCGroup;
begin
  ix := -1;
  Result := nil;
  FCS.Enter;
  try
    for i := 0 to FOPCGroup.Count - 1 do begin
      OPCGroup := sOPCGroup(FOPCGroup.Items[i]);
      if Name = OPCGroup.Name then begin
        Result := OPCGroup;
        ix := i;
        exit;
      end;
    end;
  finally
    FCS.Leave;
  end;
end;

function sOPCDataAccess.GenerateUniqueGroupName(var Name: string): boolean;
var
  i, ix: integer;
  newName: string;
begin
  i := 1;
  Result := False;
  newName := Name;
  while (FindGroupByName(Name, ix) <> nil) and (i < 9999) do begin
    newName := Name + '_' + IntToStr(i);
    if (FindGroupByName(Name, ix) = nil) then break;
    inc(i);
  end;
  if i >= 9999 then exit;
  Result := True;
  Name := newName;
end;

function sOPCDataAccess.ChangeGroupToPublic(Name: string): HResult;
var
  OPCGroup: sOPCGroup;
  ix: integer;
begin
  Result := E_FAIL;
  OPCGroup := FindGroupByName(Name, ix);
  if OPCGroup = nil then exit;
  OPCGroup.PublicGroup := True;
  Result := S_OK;
end;

procedure sOPCDataAccess.AddOPCGroup(OPCGroup: sOPCGroup);
begin
  FCS.Enter;
  try
    FOPCGroup.Add(OPCGroup);
  finally
    FCS.Leave;
  end;
end;

initialization

TOPCAutoObjectFactory.Create(
  ComServer,
  sOPCDataAccess,
  CLASS_OPCDataAccess20,
  ciMultiInstance,
  ThreadingModel);

end.

