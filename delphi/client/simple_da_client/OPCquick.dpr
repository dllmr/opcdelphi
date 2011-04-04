program OPCquick;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}
{$IFDEF VER170}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$ENDIF}

uses
{$IF CompilerVersion >= 14}
  Variants,
{$IFEND}
  Windows, Forms, ComObj, ActiveX, SysUtils, OPCtypes, OPCDA, OPCutils;

{$R *.RES}

const
  OneSecond = 1 / (24 * 60 * 60);
  // these are for use with the MatrikonOPC sample server
  ServerProgID = 'Matrikon.OPC.Simulation';
  Item0Name = 'Triangle Waves.Real8';
  Item1Name = 'Bucket Brigade.Real4';
  RPC_C_AUTHN_LEVEL_NONE = 1;
  RPC_C_IMP_LEVEL_IMPERSONATE = 3;
  EOAC_NONE = 0;

type
  // class to receive IDataObject data change advises
  TOPCAdviseSink = class(TInterfacedObject, IAdviseSink)
  public
    procedure OnDataChange(const formatetc: TFormatEtc;
                            const stgmed: TStgMedium); stdcall;
    procedure OnViewChange(dwAspect: Longint; lindex: Longint); stdcall;
    procedure OnRename(const mk: IMoniker); stdcall;
    procedure OnSave; stdcall;
    procedure OnClose; stdcall;
  end;

  // class to receive IConnectionPointContainer data change callbacks
  TOPCDataCallback = class(TInterfacedObject, IOPCDataCallback)
  public
    function OnDataChange(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMasterquality: HResult; hrMastererror: HResult; dwCount: DWORD;
      phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
      pwQualities: PWordArray; pftTimeStamps: PFileTimeArray;
      pErrors: PResultList): HResult; stdcall;
    function OnReadComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMasterquality: HResult; hrMastererror: HResult; dwCount: DWORD;
      phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
      pwQualities: PWordArray; pftTimeStamps: PFileTimeArray;
      pErrors: PResultList): HResult; stdcall;
    function OnWriteComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
      hrMastererr: HResult; dwCount: DWORD; pClienthandles: POPCHANDLEARRAY;
      pErrors: PResultList): HResult; stdcall;
    function OnCancelComplete(dwTransid: DWORD; hGroup: OPCHANDLE):
      HResult; stdcall;
  end;

var
  ServerIf: IOPCServer;
  GroupIf: IOPCItemMgt;
  GroupHandle: OPCHANDLE;
  Item0Handle: OPCHANDLE;
  Item1Handle: OPCHANDLE;
  ItemType: TVarType;
  ItemValue: string;
  ItemQuality: Word;
  HR: HResult;
  AdviseSink: IAdviseSink;
  AsyncConnection: Longint;
  OPCDataCallback: IOPCDataCallback;
  StartTime: TDateTime;

// TOPCAdviseSink methods

// OPC standard says this is the only method we need to fill in
procedure TOPCAdviseSink.OnDataChange(const formatetc: TFormatEtc;
                                      const stgmed: TStgMedium);
var
  PG: POPCGROUPHEADER;
  PI1: POPCITEMHEADER1ARRAY;
  PI2: POPCITEMHEADER2ARRAY;
  PV: POleVariant;
  I: Integer;
  PStr: PWideChar;
  NewValue: string;
  WithTime: Boolean;
  ClientHandle: OPCHANDLE;
  Quality: Word;
begin
  // the rest of this method assumes that the item header array uses
  // OPCITEMHEADER1 or OPCITEMHEADER2 records,
  // so check this first to be defensive
  if (formatetc.cfFormat <> OPCSTMFORMATDATA) and
      (formatetc.cfFormat <> OPCSTMFORMATDATATIME) then Exit;
  // does the data stream provide timestamps with each value?
  WithTime := formatetc.cfFormat = OPCSTMFORMATDATATIME;

  PG := GlobalLock(stgmed.hGlobal);
  if PG <> nil then
  begin
    // we will only use one of these two values, according to whether
    // WithTime is set:
    PI1 := Pointer(PAnsiChar(PG) + SizeOf(OPCGROUPHEADER));
    PI2 := Pointer(PI1);
    if Succeeded(PG.hrStatus) then
    begin
      for I := 0 to PG.dwItemCount - 1 do
      begin
        if WithTime then
        begin
          PV := POleVariant(PAnsiChar(PG) + PI1[I].dwValueOffset);
          ClientHandle := PI1[I].hClient;
          Quality := (PI1[I].wQuality and OPC_QUALITY_MASK);
        end
        else begin
          PV := POleVariant(PAnsiChar(PG) + PI2[I].dwValueOffset);
          ClientHandle := PI2[I].hClient;
          Quality := (PI2[I].wQuality and OPC_QUALITY_MASK);
        end;
        if Quality = OPC_QUALITY_GOOD then
        begin
          // this test assumes we're not dealing with array data
          if TVarData(PV^).VType <> VT_BSTR then
          begin
            NewValue := VarToStr(PV^);
          end
          else begin
            // for BSTR data, the BSTR image follows immediately in the data
            // stream after the variant union;  the BSTR begins with a DWORD
            // character count, which we skip over as the BSTR is also
            // NULL-terminated
            PStr := PWideChar(PAnsiChar(PV) + SizeOf(OleVariant) + 4);
            NewValue := WideString(PStr);
          end;
          if WithTime then
          begin
            Writeln('New value for item ', ClientHandle, ' advised:  ',
                    NewValue, ' (with timestamp)');
          end
          else begin
            Writeln('New value for item ', ClientHandle, ' advised:  ',
                    NewValue);
          end;
        end
        else begin
          Writeln('Advise received for item ', ClientHandle,
                  ' , but quality not good');
        end;
      end;
    end;
    GlobalUnlock(stgmed.hGlobal);
  end;
end;

procedure TOPCAdviseSink.OnViewChange(dwAspect: Longint; lindex: Longint);
begin
end;

procedure TOPCAdviseSink.OnRename(const mk: IMoniker);
begin
end;

procedure TOPCAdviseSink.OnSave;
begin
end;

procedure TOPCAdviseSink.OnClose;
begin
end;

// TOPCDataCallback methods

function TOPCDataCallback.OnDataChange(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMasterquality: HResult; hrMastererror: HResult; dwCount: DWORD;
  phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
  pwQualities: PWordArray; pftTimeStamps: PFileTimeArray;
  pErrors: PResultList): HResult;
var
  ClientItems: POPCHANDLEARRAY;
  Values: POleVariantArray;
  Qualities: PWORDARRAY;
  I: Integer;
  NewValue: string;
begin
  Result := S_OK;
  ClientItems := POPCHANDLEARRAY(phClientItems);
  Values := POleVariantArray(pvValues);
  Qualities := PWORDARRAY(pwQualities);
  for I := 0 to dwCount - 1 do
  begin
    if Qualities[I] = OPC_QUALITY_GOOD then
    begin
      NewValue := VarToStr(Values[I]);
      Writeln('New callback for item ', ClientItems[I], ' received, value:  ',
              NewValue);
    end
    else begin
      Writeln('Callback received for item ', ClientItems[I],
              ' , but quality not good');
    end;
  end;
end;

function TOPCDataCallback.OnReadComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMasterquality: HResult; hrMastererror: HResult; dwCount: DWORD;
  phClientItems: POPCHANDLEARRAY; pvValues: POleVariantArray;
  pwQualities: PWordArray; pftTimeStamps: PFileTimeArray;
  pErrors: PResultList): HResult;
begin
  Result := OnDataChange(dwTransid, hGroup, hrMasterquality, hrMastererror,
    dwCount, phClientItems, pvValues, pwQualities, pftTimeStamps, pErrors);
end;

function TOPCDataCallback.OnWriteComplete(dwTransid: DWORD; hGroup: OPCHANDLE;
  hrMastererr: HResult; dwCount: DWORD; pClienthandles: POPCHANDLEARRAY;
  pErrors: PResultList): HResult;
begin
  // we don't use this facility
  Result := S_OK;
end;

function TOPCDataCallback.OnCancelComplete(dwTransid: DWORD;
  hGroup: OPCHANDLE): HResult;
begin
  // we don't use this facility
  Result := S_OK;
end;

// main program code

begin
  // among other things, this call makes sure that COM is initialized
  Application.Initialize;
  Writeln('========================================================');
  Writeln('Simple OPC client program, by Mike Dillamore, 1998-2009.');
  Writeln('Tested for compatibility with Delphi 3-7 and 2005-2010. ');
  Writeln('Requires OPC Simulation server from MatrikonOPC,        ');
  Writeln('but easily modified for use with other servers.         ');
  Writeln('========================================================');
  Writeln;

  // this is for DCOM:
  // without this, callbacks from the server may get blocked, depending on
  // DCOM configuration settings
  HR := CoInitializeSecurity(
    nil,                    // points to security descriptor
    -1,                     // count of entries in asAuthSvc
    nil,                    // array of names to register
    nil,                    // reserved for future use
    RPC_C_AUTHN_LEVEL_NONE, // the default authentication level for proxies
    RPC_C_IMP_LEVEL_IMPERSONATE,// the default impersonation level for proxies
    nil,                    // used only on Windows 2000
    EOAC_NONE,              // additional client or server-side capabilities
    nil                     // reserved for future use
    );
  if Failed(HR) then
  begin
    Writeln('Failed to initialize DCOM security');
  end;

  try
    // we will use the custom OPC interfaces, and OPCProxy.dll will handle
    // marshaling for us automatically (if registered)
    ServerIf := CreateComObject(ProgIDToClassID(ServerProgID)) as IOPCServer;
  except
    ServerIf := nil;
  end;
  if ServerIf <> nil then
  begin
    Writeln('Connected to OPC server');
  end
  else begin
    Writeln('Unable to connect to OPC server');
    Exit;
  end;

  // now add a group
  HR := ServerAddGroup(ServerIf, 'MyGroup', True, 500, 0, GroupIf, GroupHandle);
  if Succeeded(HR) then
  begin
    Writeln('Added group to server');
  end
  else begin
    Writeln('Unable to add group to server');
    Exit;
  end;

  // now add an item to the group
  HR := GroupAddItem(GroupIf, Item0Name, 0, VT_EMPTY, Item0Handle,
    ItemType);
  if Succeeded(HR) then
  begin
    Writeln('Added item 0 to group');
  end
  else begin
    Writeln('Unable to add item 0 to group');
    ServerIf.RemoveGroup(GroupHandle, False);
    Exit;
  end;
  // now add a second item to the group
  HR := GroupAddItem(GroupIf, Item1Name, 1, VT_EMPTY, Item1Handle,
    ItemType);
  if Succeeded(HR) then
  begin
    Writeln('Added item 1 to group');
  end
  else begin
    Writeln('Unable to add item 1 to group');
    ServerIf.RemoveGroup(GroupHandle, False);
    Exit;
  end;

  // set up an IDataObject advise callback for the group
  AdviseSink := TOPCAdviseSink.Create;
  HR := GroupAdviseTime(GroupIf, AdviseSink, AsyncConnection);
  if Failed(HR) then
  begin
    Writeln('Failed to set up IDataObject advise callback');
  end
  else begin
    Writeln('IDataObject advise callback established');
    // continue waiting for callbacks for 10 seconds
    StartTime := Now;
    while (Now - StartTime) < (10 * OneSecond) do
    begin
      Application.ProcessMessages;
    end;
    // end the IDataObject advise callback
    GroupUnadvise(GroupIf, AsyncConnection);
  end;

  // now set up an IConnectionPointContainer data callback for the group
  OPCDataCallback := TOPCDataCallback.Create;
  HR := GroupAdvise2(GroupIf, OPCDataCallback, AsyncConnection);
  if Failed(HR) then
  begin
    Writeln('Failed to set up IConnectionPointContainer advise callback');
  end
  else begin
    Writeln('IConnectionPointContainer data callback established');
    // continue waiting for callbacks for 10 seconds
    StartTime := Now;
    while (Now - StartTime) < (10 * OneSecond) do
    begin
      Application.ProcessMessages;
    end;
    // end the IConnectionPointContainer data callback
    GroupUnadvise2(GroupIf, AsyncConnection);
  end;

  // now try to read the item value synchronously
  HR := ReadOPCGroupItemValue(GroupIf, Item0Handle, ItemValue, ItemQuality);
  if Succeeded(HR) then
  begin
    if (ItemQuality and OPC_QUALITY_MASK) = OPC_QUALITY_GOOD then
    begin
      Writeln('Item 0 value read synchronously as ', ItemValue);
    end
    else begin
      Writeln('Item 0 value was read synchronously, but quality was not good');
    end;
  end
  else begin
    Writeln('Failed to read item 0 value synchronously');
  end;

  // finally write the value just read from item 0 into item 1
  // Note: WriteOPCGroupItemValues may be used to write multiple item values
  HR := WriteOPCGroupItemValue(GroupIf, Item1Handle, ItemValue);
  if Succeeded(HR) then
  begin
    Writeln('Item 1 value written synchronously');
  end
  else begin
    Writeln('Failed to write item 1 value synchronously');
  end;

  // wait for 1 second
  StartTime := Now;
  while (Now - StartTime) < OneSecond do
  begin
    Application.ProcessMessages;
  end;

  // and try to read it back
  HR := ReadOPCGroupItemValue(GroupIf, Item1Handle, ItemValue, ItemQuality);
  if Succeeded(HR) then
  begin
    if (ItemQuality and OPC_QUALITY_MASK) = OPC_QUALITY_GOOD then
    begin
      Writeln('Item 1 value read synchronously as ', ItemValue);
    end
    else begin
      Writeln('Item 1 value was read synchronously, but quality was not good');
    end;
  end
  else begin
    Writeln('Failed to read item 0 value synchronously');
  end;

  // now cleanup
  HR := ServerIf.RemoveGroup(GroupHandle, False);
  if Succeeded(HR) then
  begin
    Writeln('Removed group');
  end
  else begin
    Writeln('Unable to remove group');
  end;

  // Delphi runtime library will release all interfaces automatically when
  // variables go out of scope
end.
