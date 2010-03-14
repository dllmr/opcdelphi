function TDA2.QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
begin
 pNameSpaceType:=OPC_NS_FLAT;
 result:=S_OK;
end;

function TDA2.ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                              szString:POleStr):HResult;stdcall;
begin
 result:=E_FAIL
end;

function TDA2.BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE;
                               szFilterCriteria:POleStr;
                               vtDataTypeFilter:TVarType;
                               dwAccessRightsFilter:DWORD;
                               out ppIEnumString:IEnumString):HResult;stdcall;
var
 i:integer;
 tList:TStringList;

 function FilterOutAccess(dwAccessRightsFilter:DWORD; propID:integer):boolean;
 begin
  result:=false;
  if (dwAccessRightsFilter <> 0) then
   begin
    if CanPropIDBeWritten(propID) then
     result:=not (dwAccessRightsFilter in [OPC_WRITEABLE,3])
    else
     result:=not (dwAccessRightsFilter in [OPC_READABLE]);
   end;
 end;

 function FilterOutDataType(vtDataTypeFilter:TVarType; propID:integer):boolean;
 var
  vType:word;
 begin
  result:=false;
  if (vtDataTypeFilter <> VT_EMPTY) then
   begin
    vType:=ReturnDataTypeFromPropID(propID);
    result:=(vType <> vtDataTypeFilter);
   end;
 end;

begin
 result:=S_OK;
 tList:=nil;
 try
  if (dwAccessRightsFilter <> 0) then
   if not (dwAccessRightsFilter in [OPC_READABLE,OPC_WRITEABLE,3]) then
    begin
     result:=E_INVALIDARG;
     Exit;
    end;

//  if not (dwBrowseFilterType in [OPC_BRANCH,OPC_LEAF,OPC_FLAT]) then
//   begin
//    result:=E_INVALIDARG;
//    Exit;
//   end;

  tList:=TStringList.Create;
  if (tList = nil) then
   begin
    result:=E_OUTOFMEMORY;
    Exit;
   end;

  for i:= low(posItems) to high(posItems) - 1 do
   begin
    if FilterOutAccess(dwAccessRightsFilter,posItems[i].PropID) then
     Continue;

    if FilterOutDataType(vtDataTypeFilter,posItems[i].PropID) then
     Continue;

    tList.Add(posItems[i].tagname);
   end;

  if (length(szFilterCriteria) > 0) and Assigned(tList) then
   begin
    for i:= tList.count -1 downTo 0 do
     if not MatchesMask(tList[i], szFilterCriteria) then
      tList.Delete(i);

     if (tList.count = 0) then
      result:=S_FALSE;
   end;

  ppIEnumString:=TOPCStringsEnumerator.Create(tList);
 finally
  tList.Free;
 end;
end;

function TDA2.GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
var
 propID:integer;
begin
 result:=S_OK;
 if length(szItemDataID) = 0 then
  szItemID:=StringToLPOLESTR(szItemDataID)
 else
  begin
   propID:=ReturnPropIDFromTagname(szItemDataID);
   if (propID = 0) then
    result:=E_INVALIDARG              //do not know what to do
   else
    szItemID:=StringToLPOLESTR(szItemDataID);
  end;
end;

function TDA2.BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
begin
 result:=E_NOTIMPL;
end;
