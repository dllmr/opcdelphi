unit ServIMPL;

interface

uses Windows,ComObj,ActiveX,Axctrls,FirstServ_TLB,OPCDA,SysUtils,Dialogs,Classes,
     OPCCOMN,StdVCL,enumstring,ItemPropIMPL,Globals,OpcError,OPCErrorStrings,
     EnumUnknown,OPCTypes,IOPCCommonUnit;

type
  TDA2 = class(TAutoObject,IDA2,IOPCServer,IOPCCommon,IOPCBrowseServerAddressSpace,
               IConnectionPointContainer,IOPCItemProperties)
  private
   FIOPCCommon:TIOPCCommon;                     //IOPCCommon
   FIOPCItemProperties:TOPCItemProp;
   FIConnectionPoints:TConnectionPoints;
  protected
   property iFIOPCCommon:TIOPCCommon read FIOPCCommon write FIOPCCommon
                                                      implements IOPCCommon;
   property iFIConnectionPoints:TConnectionPoints read FIConnectionPoints
                          write FIConnectionPoints implements IConnectionPointContainer;
//IOPCServer begin
    function AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                      hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                      dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                    out pRevisedUpdateRate:DWORD;
                                    const riid: TIID;
                                    out ppUnk:IUnknown):HResult;stdcall;
    function GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult;overload; stdcall;
    function GetGroupByName(szName:POleStr; const riid: TIID; out ppUnk:IUnknown):HResult; stdcall;
    function GetStatus(out ppServerStatus:POPCSERVERSTATUS): HResult; stdcall;
    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
    function CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
//IOPCServer end

//IOPCBrowseServerAddressSpace begin
   function QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
   function ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                                 szString:POleStr):HResult;stdcall;
   function BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE; szFilterCriteria:POleStr;
                             vtDataTypeFilter:TVarType; dwAccessRightsFilter:DWORD;
                             out ppIEnumString:IEnumString):HResult;stdcall;
   function GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
   function BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
//IOPCBrowseServerAddressSpace end

  public
   grps,pubGrps,removedGroups:TList;
   localID:longword;
   clientName,errString:string;
   srvStarted,lastClientUpdate:TDateTime;
   FOnSDConnect: TConnectEvent;
   ClientIUnknown:IUnknown;
   property iFIOPCItemProperties:TOPCItemProp read FIOPCItemProperties
                                              write FIOPCItemProperties
                                              implements IOPCItemProperties;

   procedure CreateGroups;
   procedure Initialize; override;
   procedure ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);

   destructor Destroy;override;
   function GetNewGroupNumber:longword;
   function GetNewItemNumber:longword;
   function FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
   function FindIndexFromGroupNumberInRemovedGroup(groupNum:longword):integer;
   procedure GroupRemovingSelf(wGrp:TList;gNum:integer);
   function GetGroupCount(gList:TList):integer;
   function CreateGrpNameList(gList:TList):TStringList;
   function IsGroupNamePresent(gList:TList;theName:string):integer;
   function IsNameUsedInAnyGroup(theName:string):boolean;
   function IsThisGroupPublic(aList:TList):boolean;
   procedure TimeSlice(cTime:TDateTime);
   function CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
  end;

var
 theServers:array [0..10] of TDA2;

implementation

uses ComServ,Main,GroupUnit,OPCRemovedGroupUnit,Masks;

{$INCLUDE IOPCServerIMPL}
{$INCLUDE IOPCBrowseServerAddressSpaceIMPL}

function GetNextFreeServerSpot:integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if not Assigned(theServers[i]) then
   begin
    result:=i;
    Exit;
   end;
end;

function FindServerInArray(which:TDA2):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if Assigned(theServers[i]) then
   if (theServers[i] = which) then
    begin
     result:=i;
     Exit;
    end;
end;

function ReturnServerCount:integer;
var
 i:integer;
begin
 result:=0;
 for i:= low(theServers) to high(theServers) do
  if Assigned(theServers[i]) then
   Inc(result);
end;

procedure KillServers;
var
 i:integer;
begin
 for i:= high(theServers) downTo low(theServers) do
  begin
   CoDisconnectObject(TDA2(theServers[i]) as IUnknown,0);
   TDA2(theServers[i]).Free;
  end;
 FreeAndNil(theServers);
end;

procedure TDA2.CreateGroups;
begin
 if (grps <> nil) then
  Exit;
 grps:=TList.Create;
 grps.Capacity:=255;
 pubGrps:=TList.Create;
 pubGrps.Capacity:=grps.Capacity;

 removedGroups:=TList.Create;
end;

procedure TDA2.Initialize;
var
 i:integer;
begin
 i:=GetNextFreeServerSpot;
 if (i = -1) then
  Exit;
 inherited Initialize;
 srvStarted:=Now;
 lastClientUpdate:=0;
 localID:=LOCALE_SYSTEM_DEFAULT;

 FIOPCCommon:=TIOPCCommon.Create;
 FIConnectionPoints:=TConnectionPoints.Create(self);
 FIOPCItemProperties:=TOPCItemProp.Create;

 FOnSDConnect:=ShutdownOnConnect;
 FIConnectionPoints.CreateConnectionPoint(IID_IOPCShutdown,ckSingle,FOnSDConnect);

 CreateGroups;

 //hook into Main program here    may have multiple servers
 theServers[i]:=self;
 Form1.UpdateGroupCount;
end;

procedure TDA2.ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);
begin
 if connecting then
  ClientIUnknown:=Sink
 else
  ClientIUnknown:=nil
end;

destructor TDA2.Destroy;
var
 i:integer;
begin
 if Assigned(grps) then
  begin
   for i:= 0 to grps.count - 1 do
    TOPCGroup(grps.Items[i]).Free;
   grps.Free;
  end;

 if Assigned(pubGrps) then
  begin
   for i:= 0 to pubGrps.count - 1 do
    TOPCGroup(pubGrps.Items[i]).Free;
   pubGrps.Free;
  end;

 if Assigned(removedGroups) then
  begin
   for i:= removedGroups.count - 1 downTo 0 do
    if Assigned(removedGroups.Items[i]) then
     TOPCGroupRemoved(removedGroups.Items[i]).Free;
   FreeAndNil(removedGroups);
  end;

 i:=FindServerInArray(self);
 if (i <> -1) then
  theServers[i]:=nil;                //the client has let us be free ;)

 if Assigned(FIOPCCommon) then
  FreeAndNil(FIOPCCommon);
 if Assigned(FIConnectionPoints) then
  FIConnectionPoints.Free;
 if Assigned(FIOPCItemProperties) then
  FIOPCItemProperties.Free;



 Form1.UpdateGroupCount;
 Inherited;
end;

function TDA2.GetNewGroupNumber:longword;
const
 grpIndex:longword = 1;             //Assignable Typed Constants gota lovem
begin
 Inc(grpIndex);         //get us a new reference number
 result:=grpIndex;
end;

function TDA2.GetNewItemNumber:longword;
const
 itemIndex:longword = 1;
begin
 Inc(itemIndex);
 result:=itemIndex;
end;

function TDA2.FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to wGrp.count - 1 do
  if (TOPCGroup(wGrp[i]).serverHandle = gNum) then
   begin
    result:=i;
    Break;
   end;
end;

function TDA2.FindIndexFromGroupNumberInRemovedGroup(groupNum:longword):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to removedGroups.count -1 do
  if (TOPCGroupRemoved(removedGroups[i]).oldServerHandle = groupNum) then
   begin
    result:=i;
    Break;
   end;
end;

procedure TDA2.GroupRemovingSelf(wGrp:TList; gNum:integer);
var
 i:integer;
 oldGroup:TOPCGroupRemoved;
begin
 i:=FindIndexViaGrpNumber(wGrp,gNum);
 if (i <> -1) then
  begin
   with TOPCGroup(wGrp[i]) do
    oldGroup:=TOPCGroupRemoved.Create(tagName,serverHandle,wGrp);
   removedGroups.Add(oldGroup);
   wGrp.Delete(i);
   Form1.UpdateGroupCount;
  end;
end;

function TDA2.GetGroupCount(gList:TList):integer;
begin
 result:=0;
 if (gList = nil) then
  Exit;
 result:=gList.count;
end;

function TDA2.CreateGrpNameList(gList:TList):TStringList;
var
 i:integer;
begin
 result:=nil;
 if (gList = nil) then
  Exit;
 result:=TStringList.Create;
 for i:= 0 to gList.count - 1 do
  result.Add(TOPCGroup(gList.Items[i]).tagName);

 if (result.count = 0) then
  begin
   result.Free;
   result:=nil;
  end;
end;

function TDA2.IsGroupNamePresent(gList:TList; theName:string):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to gList.count - 1 do
  if (theName = TOPCGroup(gList.Items[i]).tagName) then
   begin
    result:=i;
    Break;
   end;
end;

function TDA2.IsNameUsedInAnyGroup(theName:string):boolean;
var
 i:integer;
begin
 result:=false;
 i:=IsGroupNamePresent(grps,theName);
 if (i <> -1) then
  begin
   result:=true;
   Exit;
  end;

 i:=IsGroupNamePresent(pubGrps,theName);
 if (i <> -1) then
  result:=true;
end;

function TDA2.IsThisGroupPublic(aList:TList):boolean;
begin
 result:=boolean((pubGrps <> nil)    and
                 (pubGrps.count = 0) and
                 (aList <> nil)      and
                 (pubGrps = aList));
end;

procedure TDA2.TimeSlice(cTime:TDateTime);
var
 i:integer;
begin
 lastClientUpdate:=cTime;

 if Assigned(grps) and (grps.count > 0) then
  begin
   i:=0;
   while i <= (grps.count - 1) do
    begin
     if Assigned(grps[i]) then
      TOPCGroup(grps[i]).TimeSlice(cTime);
     Inc(i);
    end;
  end;

 if Assigned(pubGrps) and (pubGrps.count > 0) then
  begin
   i:=0;
   while i <= (pubGrps.count - 1) do
    begin
     if Assigned(pubGrps[i]) then
      TOPCGroup(pubGrps[i]).TimeSlice(cTime);
     Inc(i);
    end;
  end;
end;

function TDA2.CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
var
 sGrp,dGrp:TOPCGroup;
begin
 sGrp:=TOPCGroup(aGrp);
 dGrp:=TOPCGroup.Create(self,grps);
 if (dGrp = nil) then
  begin
   result:=nil;
   res:=E_OUTOFMEMORY;
   Exit;
  end;

 grps.Add(dGrp);
 dGrp.tagName:=szName;
 sGrp.CloneYourSelf(dGrp);
 result:=dGrp;
 res:=S_OK;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TDA2, Class_DA2,
                            ciMultiInstance, tmApartment);

 //if an OPC client(s) is connected and the user has selected to quit after
 //the warning in the FormCloseQuery then do not let the system ask again in the
 //AutomationTerminateProc procedure in the VCL.
 ComServer.UIInteractive:=false;
finalization
 //if an OPC client is connected and this is a forced kill then if CoUnintialize
 //is not called here the OLE dll will generate an error when it is called after
 //we have killed the servers.
 CoUninitialize;
 KillServers;
end.
