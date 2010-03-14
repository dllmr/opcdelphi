//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uOPCAsyncIO2;

interface

uses
  Windows, ActiveX, ComObj, Axctrls, SysUtils, Dialogs, Classes,
  OPCtypes, OPCDA, OpcError, uGlobals, uOPCGroup, uOPCItem;

type
  sOPCMode = (omRead, omWrite, omRefresh, omTimerRefresh);

  TFileTimeArray = array[0 .. 65535] of TFileTime;
  PTFileTimeArray = ^TFileTimeArray;

  sOPCAsyncIO2 = class
  private
    mClientID: PDWORDARRAY;
    mVariant: POleVariantArray;
    mQuality: PWordArray;
    mTime: PTFileTimeARRAY;
    mError: PResultList;

    OPCGroup: sOPCGroup;
    OPCMode: sOPCMode;
    dwTransactionID: DWORD;
    dwCount: DWORD;
    DataSource: integer;

    FCancelFlag: boolean;
    pdwCancelID: DWORD;

    hServer: array of OPCHANDLE;
    ItemValues: array of OleVariant;

    procedure Read(bEnable: boolean); virtual;
    procedure Write(bEnable: boolean); virtual;
    procedure Refresh(bEnable: boolean); virtual;

  public
    constructor Create(
      aOPCGroup: sOPCGroup;
      aOPCMode: sOPCMode;
      adwTransactionID,
      adwCount: DWORD;
      aDataSource: integer);

    function Init: boolean;
    // Result: False -> initialisation error

    procedure Done;

    procedure ProcessRequest(bEnable: boolean); virtual;

    procedure Copy_phServer(aphServer: POPCHANDLEARRAY);
    procedure Copy_pItemValues(apItemValues: POleVariantArray);

    property CancelFlag: boolean write FCancelFlag;
    property CancelID: DWORD read pdwCancelID write pdwCancelID;
  end;

implementation

uses
  uOPCDataAccess;

constructor sOPCAsyncIO2.Create(
  aOPCGroup: sOPCGroup;
  aOPCMode: sOPCMode;
  adwTransactionID,
  adwCount: DWORD;
  aDataSource: integer);
begin
  OPCGroup := aOPCGroup;
  OPCMode := aOPCMode;
  dwTransactionID := adwTransactionID;
  dwCount := adwCount;
  DataSource := aDataSource;
  pdwCancelID := DWORD(self);
  FCancelFlag := False;
end;

function sOPCAsyncIO2.Init: boolean;
var
  Dummy: HResult;
begin
  Result := False;

  mClientID := nil;
  mVariant := nil;
  mError := nil;
  mQuality := nil;
  mTime := nil;

  try
    mClientID := TaskMemAlloc(dwCount, mkDWORD, Dummy);
    mError    := TaskMemAlloc(dwCount, mkHResult, Dummy);
    if OPCMode <> omWrite then begin
      mVariant := TaskMemAlloc(dwCount, mkOleVariant, Dummy);
      mQuality := TaskMemAlloc(dwCount, mkWord, Dummy);
      mTime    := TaskMemAlloc(dwCount, mkFileTime, Dummy);
    end;
  except
    on E: Exception do begin
      OPCLogException('sOPCAsyncIO2.Init', E);
      Done;
      exit;
    end;
  end;

  if OPCMode = omWrite then begin
    if (mClientID = nil) or (mError = nil) then begin
      Done;
      exit;
    end;
  end else begin
    if (mQuality = nil)  or
       (mTime = nil)     or
       (mClientID = nil) or
       (mVariant = nil)  or
       (mError = nil)
    then begin
      Done;
      exit;
    end;
  end;

  Result := True;
end;

procedure sOPCAsyncIO2.Done;
var
  i: integer;
begin
  if mClientID <> nil then CoTaskMemFree(mClientID);
  if mVariant  <> nil then begin
    for i := 0 to dwCount - 1 do VarClear(mVariant[i]);
    CoTaskMemFree(mVariant);
  end;
  if mError    <> nil then CoTaskMemFree(mError);
  if mQuality  <> nil then CoTaskMemFree(mQuality);
  if mTime     <> nil then CoTaskMemFree(mTime);
end;

procedure sOPCAsyncIO2.Read(bEnable: boolean);
var
  i, c: integer;
  ReadResult: HResult;
  OPCItem: sOPCItem;
begin
  if not Init then exit;
  ReadResult := S_OK;
  try
    c := 0;
    for i := 0 to dwCount - 1 do begin
      OPCItem := OPCGroup.GetOPCItem(hServer[i]);
      if (OPCItem = nil) then begin
        mQuality[i] := OPC_QUALITY_BAD;
        mError[i] := OPC_E_INVALIDHANDLE;
        ReadResult := S_FALSE;
      end else if not OPCItem.Active then begin
        mQuality[i] := OPC_QUALITY_OUT_OF_SERVICE;
        mError[i] := E_FAIL;
        ReadResult := S_FALSE;
      end else if not OPCItem.ReadAble then begin
        mQuality[i] := OPC_QUALITY_BAD;
        mError[i] := OPC_E_BADRIGHTS;
        ReadResult := S_FALSE;
      end else begin
        inc(c);
        mTime[i] := ConvertToFileTime(OPCItem.OPCNode.LastUpdate);
        mClientID[i] := OPCItem.ClientHandle;
        mVariant[i] := OPCItem.Read(DataSource);
        mQuality[i] := OPCItem.Quality;
        mError[i] := S_OK;
      end;
    end;
    if bEnable and (OPCGroup.DataCallback <> nil) then begin
      sOPCDataAccess(OPCGroup.DataAccess).LastDataUpdateToClient := Now;
      if c > 0 then IOPCDataCallback(OPCGroup.DataCallback).OnReadComplete(
        dwTransactionID,
        OPCGroup.ClientGroup,
        OPC_QUALITY_GOOD,
        ReadResult,
        dwCount,
        @mClientID^,
        @mVariant^,
        @mQuality^,
        @mTime^,
        @mError^);
    end;
  except
    on E: Exception do OPCLogException('sOPCAsyncIO2.Read', E);
  end;
  Done;
end;

procedure sOPCAsyncIO2.Write(bEnable: boolean);
var
  i, c: integer;
  WriteResult: HResult;
  OPCItem: sOPCItem;
begin
  if not Init then exit;
  WriteResult := S_OK;
  try
    c := 0;
    for i := 0 to dwCount - 1 do begin
      OPCItem := OPCGroup.GetOPCItem(hServer[i]);
      if (OPCItem = nil) then begin
        mError[i] := OPC_E_INVALIDHANDLE;
        WriteResult := S_FALSE;
      end else if not OPCItem.Active then begin
        mError[i] := E_FAIL;
        WriteResult := S_FALSE;
      end else if not OPCItem.WriteAble then begin
        mError[i] := OPC_E_BADRIGHTS;
        WriteResult := S_FALSE;
      end else begin
        inc(c);
        mClientID[i] := OPCItem.ClientHandle;
        OPCItem.Write(ItemValues[i]);
        mError[i] := S_OK;
      end;
    end;
    if bEnable and (OPCGroup.DataCallback <> nil) then begin
      sOPCDataAccess(OPCGroup.DataAccess).LastDataUpdateToClient := Now;
      if c > 0 then IOPCDataCallback(OPCGroup.DataCallback).OnWriteComplete(dwTransactionID,
        OPCGroup.ClientGroup,
        WriteResult,
        dwCount,
        @mClientID^,
        @mError^);
    end;
  except
    on E: Exception do OPCLogException('sOPCAsyncIO2.Write', E);
  end;
  Done;
end;

procedure sOPCAsyncIO2.Refresh(bEnable: boolean);
var
  i, x: integer;
  RefreshResult: HResult;
  OPCItem: sOPCItem;
begin
  // set Active and ReadAble OPCItems to 1
  for i := 0 to OPCGroup.OPCItemList.Count - 1 do begin
    OPCItem := sOPCItem(OPCGroup.OPCItemList[i]);
    if OPCItem.Active and OPCItem.ReadAble
      then OPCItem.Tag := 1
      else OPCItem.Tag := 0;
  end;

  // update active OPCItem's on Timer or on Refresh Device
  if (OPCMode = omTimerRefresh) or (DataSource = OPC_DS_DEVICE) then begin
    for i := 0 to OPCGroup.OPCItemList.Count - 1 do begin
      OPCItem := sOPCItem(OPCGroup.OPCItemList[i]);
      if OPCItem.Tag = 1 then OPCItem.Read(OPC_DS_DEVICE);
    end;
  end;

  try
    OPCGroup.OPCLock := True;
    dwCount := 0;
    if (OPCMode = omTimerRefresh) then begin
      // count Items where the value has changed
      for i := 0 to OPCGroup.OPCItemList.Count - 1 do begin
        OPCItem := sOPCItem(OPCGroup.OPCItemList[i]);
        if OPCItem.Tag = 1 then begin
          if (OPCItem.CallBackTime <> OPCItem.OPCNode.LastUpdate) then begin
            OPCItem.Tag := 2;
            inc(dwCount);
          end;
        end;
      end;
    end else begin
      // count active Items
      for i := 0 to OPCGroup.OPCItemList.Count - 1 do begin
        OPCItem := sOPCItem(OPCGroup.OPCItemList[i]);
        if OPCItem.Tag = 1 then begin
          OPCItem.Tag := 2;
          inc(dwCount);
        end;
      end;
    end;

    if dwCount = 0 then exit;
    if not Init then exit;
    RefreshResult := S_OK;

    x := 0;
    for i := 0 to OPCGroup.OPCItemList.Count - 1 do begin
      OPCItem := sOPCItem(OPCGroup.OPCItemList[i]);
      if OPCItem.Tag = 2 then begin
        if (OPCMode = omTimerRefresh)
          then OPCItem.CallBackTime := OPCItem.OPCNode.LastUpdate;
        mVariant[x] := OPCItem.Read(OPC_DS_CACHE);
        mQuality[x] := OPCItem.Quality;
        mTime[x] := ConvertToFileTime(OPCItem.OPCNode.LastUpdate);
        mClientID[x] := OPCItem.ClientHandle;
        mError[x] := S_OK;
        inc(x);
      end;
    end;
  finally
    OPCGroup.OPCLock := False;
  end;

  if bEnable and (OPCGroup.DataCallback <> nil) then begin
    sOPCDataAccess(OPCGroup.DataAccess).LastDataUpdateToClient := Now;
    try
      IOPCDataCallback(OPCGroup.DataCallback).OnDataChange(dwTransactionID,
        OPCGroup.ClientGroup,
        OPC_QUALITY_GOOD,
        RefreshResult,
        dwCount,
        @mClientID^,
        @mVariant^,
        @mQuality^,
        @mTime^,
        @mError^);
    except
      on E: Exception do OPCLogException('sOPCAsyncIO2.Refresh', E);
    end;
  end;
  Done;
end;

procedure sOPCAsyncIO2.ProcessRequest(bEnable: boolean);
begin
  if FCancelFlag then exit;
  case OPCMode of
    omRead:         Read(bEnable);
    omWrite:        Write(bEnable);
    omRefresh,
    omTimerRefresh: Refresh(bEnable);
  end;
end;

procedure sOPCAsyncIO2.Copy_phServer(aphServer: POPCHANDLEARRAY);
var
  i: integer;
begin
  SetLength(hServer, dwCount);
  for i := 0 to dwCount - 1 do hServer[i] := aphServer[i];
end;

procedure sOPCAsyncIO2.Copy_pItemValues(apItemValues: POleVariantArray);
var
  i: integer;
begin
  SetLength(ItemValues, dwCount);
  for i := 0 to dwCount - 1 do ItemValues[i] := apItemValues[i];
  // +++ prüfen
end;

end.

