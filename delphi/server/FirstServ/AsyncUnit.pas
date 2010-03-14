unit AsyncUnit;

interface

uses Windows,ActiveX,ComObj,FirstServ_TLB,SysUtils,Dialogs,Classes,ServIMPL,OPCDA,
     axctrls,Globals,GroupUnit,ItemsUnit,OpcError,OPCTypes;

type
  TAsyncIO2 = class
  public
   grp:TOPCGroup;
   source:integer;
   isCancelled:boolean;
   ppServer:PDWORDARRAY;
   ppValues:POleVariantArray;
   kind,clientTransID,cancelID,itemCount:longword;
   aStream:TMemoryStream;
   constructor Create(aGrp:TOPCGroup;ioKind,cID,count:longword;dwSource:integer);
   destructor Destroy;override;
   procedure DoCancelCallBack;
   procedure HandleRead(cTime:TDateTime);
   procedure HandleWrite(cTime:TDateTime);
   procedure HandleRefresh(cTime:TDateTime);
   procedure HandleChange(aStream:TMemoryStream; cTime:TDateTime);
   procedure HandleThisRequest(cTime:TDateTime);
  end;

implementation

uses Variants;

var
 Obj:Pointer;

constructor TAsyncIO2.Create(aGrp:TOPCGroup;ioKind,cID,count:longword;dwSource:integer);
begin
 grp:=aGrp;
 kind:=ioKind;
 clientTransID:=cID;
 itemCount:=count;
 ppServer:=nil;
 isCancelled:=false;
 source:=dwSource;
 cancelID:=grp.GenerateAsyncCancelID;
end;

destructor TAsyncIO2.Destroy;
begin
 grp:=nil;
 if (ppServer <> nil) then
  CoTaskMemFree(ppServer);
 if (ppValues <> nil) then
  CoTaskMemFree(ppValues);
end;

procedure TAsyncIO2.DoCancelCallBack;
begin
 if not Succeeded(TOPCGroup(grp).ClientIUnknown.QueryInterface(IID_IOPCDataCallback,Obj)) then
  Exit;
 IOPCDataCallback(Obj).OnCancelComplete(clientTransID,TOPCGroup(grp).clientHandle);
end;

procedure TAsyncIO2.HandleRead(cTime:TDateTime);
var
 i:longword;
 aItem:TOPCItem;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;
 ppTimeArray:PFileTimeArray;
 masterResult,masterQuality:HRESULT;
begin

 ppClientItems:=nil;
 pVariants:=nil;
 ppErrors:=nil;
 ppQualityArray:=nil;
 ppTimeArray:=nil;

 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount * sizeof(longword)));
  if (ppClientItems = nil) then
   Exit;

  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount * sizeof(OleVariant)));
  if (pVariants = nil) then
   Exit;

  FillChar(pVariants[0],itemCount * sizeOf(OleVariant),#0);   //no variant conversion errors

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));
  if (ppErrors = nil) then
   Exit;

  ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(word)));
  if (ppQualityArray = nil) then
   Exit;

  ppTimeArray:=PFileTimeArray(CoTaskMemAlloc(itemCount*sizeof(TFileTime)));
  if (ppTimeArray = nil) then
   Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;
  masterQuality:=S_OK;
  for i:= 0 to itemCount - 1 do
   begin
    ppTimeArray[i]:=aFileTime;
    pVariants[i]:=VT_EMPTY;
    ppClientItems[i]:=0;

    aItem:=TOPCItem(TOPCGroup(grp).clItems[ppServer[i]]);

//check access rights here

    aItem.CallBackRead(ppClientItems[i],pVariants[i],ppQualityArray[i],
                       ppErrors[i],true);

    if not aItem.bActive then
     ppQualityArray[i]:=OPC_QUALITY_OUT_OF_SERVICE;
   end;

  if isCancelled then
   Exit;

  IOPCDataCallback(Obj).OnReadComplete(clientTransID,
                                       grp.clientHandle,
                                       masterQuality,
                                       masterResult,
                                       itemCount,
                                       @ppClientItems^,
                                       @pVariants^,
                                       @ppQualityArray^,
                                       @ppTimeArray^,
                                       @ppErrors^);

 finally
  if (ppClientItems <> nil) then
   CoTaskMemFree(ppClientItems);
  if (pVariants <> nil) then
   CoTaskMemFree(pVariants);
  if (ppErrors <> nil) then
   CoTaskMemFree(ppErrors);
  if (ppQualityArray <> nil) then
   CoTaskMemFree(ppQualityArray);
  if (ppTimeArray <> nil) then
   CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleWrite(cTime:TDateTime);
var
 aItem:TOPCItem;
 vType:integer;
 ppErrors:PResultList;
 i,masterResult:longword;
 ppClientItems:PDWORDARRAY;

 procedure HandleException(x:integer);
 begin
  ppErrors[x]:=OPC_E_BADTYPE;
  masterResult:=S_FALSE;
 end;

begin
 ppClientItems:=nil;
 ppErrors:=nil;

 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount * sizeof(longword)));
  if (ppClientItems = nil) then
   Exit;

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount * sizeof(HRESULT)));
  if (ppErrors = nil) then
  Exit;

  masterResult:=S_OK;
  for i:= 0 to itemCount - 1 do
   begin
    if (ppServer[i] > longword((grp.clItems.count - 1))) then
     begin
      ppErrors[i]:=OPC_E_INVALIDITEMID;
      masterResult:=S_FALSE;
      Continue;
     end;

    aItem:=TOPCItem(grp.clItems[ppServer[i]]);
    ppClientItems[i]:=aItem.GetClientHandle;
    if not aItem.isWriteAble then
     begin
      ppErrors[i]:=OPC_E_BADRIGHTS;
      masterResult:=S_FALSE;
      Continue;
     end;

    vType:=VarType(ppValues[i]);
    if (vType = varEmpty) or not IsVariantTypeOK(vType) then
     begin
      ppErrors[i]:=OPC_E_BADTYPE;
      masterResult:=S_FALSE;
     end
    else
     begin
      try
       aItem.WriteItemValue(ppValues[i],ppErrors[i]);
      except
       HandleException(i);
      end;
     end;
   end;

  if isCancelled then
   Exit;

  IOPCDataCallback(Obj).OnWriteComplete(clientTransID,
                                        grp.clientHandle,
                                        masterResult,
                                        itemCount,
                                        @ppClientItems^,
                                        @ppErrors^);
 finally
  if ppClientItems <> nil then
   CoTaskMemFree(ppClientItems);
  if ppErrors <> nil then
   CoTaskMemFree(ppErrors);
 end;
end;

procedure TAsyncIO2.HandleRefresh(cTime:TDateTime);
var
 x:integer;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 i,masterResult,masterQuality:longword;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;
 ppTimeArray:PFileTimeArray;
begin
 ppClientItems:=nil;
 pVariants:=nil;
 ppErrors:=nil;
 ppQualityArray:=nil;
 ppTimeArray:=nil;

 try
  for i:= 0 to grp.clItems.count - 1 do
   if TOPCItem(grp.clItems[i]).GetActiveState then
    Inc(itemCount);

  if (itemCount = 0) then
   Exit;

  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount * sizeof(longword)));
  if (ppClientItems = nil) then
   Exit;

  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount * sizeof(OleVariant)));
  if (pVariants = nil) then
   Exit;

  FillChar(pVariants[0],itemCount * sizeOf(OleVariant),#0);   //no variant conversion errors

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount * sizeof(HRESULT)));
  if (ppErrors = nil) then
   Exit;

  ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount * sizeof(word)));
  if (ppQualityArray = nil) then
   Exit;

  ppTimeArray:=PFileTimeArray(CoTaskMemAlloc(itemCount * sizeof(TFileTime)));
  if (ppTimeArray = nil) then
   Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;
  masterQuality:=S_OK;

  x:=0;
  for i:= 0 to grp.clItems.count - 1 do
   if TOPCItem(grp.clItems[i]).GetActiveState then
    begin
     ppTimeArray[x]:=aFileTime;
     ppClientItems[x]:=TOPCItem(grp.clItems[i]).GetClientHandle;
     pVariants[x]:=TOPCItem(grp.clItems[i]).ReturnCurrentValue(source,
                            ppQualityArray[x],ppErrors[x]);

     if (ppQualityArray[x] <> OPC_QUALITY_GOOD) then
      masterQuality:=S_FALSE;
     Inc(x);
    end
   else
    begin
     ppQualityArray[x]:=OPC_QUALITY_OUT_OF_SERVICE;
     ppErrors[x]:=S_FALSE;
     masterQuality:=S_FALSE;
    end;

  if isCancelled then
   Exit;

  IOPCDataCallback(Obj).OnDataChange(clientTransID,
                                     grp.clientHandle,
                                     masterQuality,
                                     masterResult,
                                     itemCount,
                                     @ppClientItems^,
                                     @pVariants^,
                                     @ppQualityArray^,
                                     @ppTimeArray^,
                                     @ppErrors^);

 finally
  if (ppClientItems <> nil) then
   CoTaskMemFree(ppClientItems);
  if (pVariants <> nil) then
   CoTaskMemFree(pVariants);
  if (ppErrors <> nil) then
   CoTaskMemFree(ppErrors);
  if (ppQualityArray <> nil) then
   CoTaskMemFree(ppQualityArray);
  if (ppTimeArray <> nil) then
   CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleChange(aStream:TMemoryStream; cTime:TDateTime);
var
 x,k:integer;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 i,masterResult:longword;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;
 ppTimeArray:PFileTimeARRAY;
begin
 ppClientItems:=nil;
 pVariants:=nil;
 ppErrors:=nil;
 ppQualityArray:=nil;
 ppTimeArray:=nil;

 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(longword)));
  if (ppClientItems = nil) then
   Exit;

  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount * sizeof(OleVariant)));
  if (pVariants = nil) then
   Exit;

  FillChar(pVariants[0],itemCount * sizeOf(OleVariant),#0);   //no variant conversion errors

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));
  if (ppErrors = nil) then
   Exit;

  ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(word)));
  if (ppQualityArray = nil) then
   Exit;

  ppTimeArray:=PFileTimeARRAY(CoTaskMemAlloc(itemCount*sizeof(TFileTime)));
  if (ppTimeArray = nil) then
   Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;
  x:=0;

  aStream.Seek(0,soFromBeginning);
  for i:= 0 to itemCount - 1 do
   begin
    aStream.Read(k,sizeOf(k));
    if (TOPCGroup(grp).clItems.count >= k) then
     if Assigned(TOPCItem(TOPCGroup(grp).clItems[k])) then
      begin
       ppTimeArray[x]:=aFileTime;
       ppClientItems[x]:=TOPCItem(TOPCGroup(grp).clItems[k]).GetClientHandle;
       with TOPCItem(TOPCGroup(grp).clItems[k]) do
        CallBackRead(ppClientItems[x],pVariants[x], ppQualityArray[x],ppErrors[x],
                     false);
       Inc(x);
      end;
    end;

   if isCancelled then
    Exit;

   IOPCDataCallback(Obj).OnDataChange(clientTransID,
                                      TOPCGroup(grp).clientHandle,
                                      OPC_QUALITY_GOOD,
                                      masterResult,
                                      itemCount,
                                      @ppClientItems^,
                                      @pVariants^,
                                      @ppQualityArray^,
                                      @ppTimeArray^,
                                      @ppErrors^);

 finally
  if (ppClientItems <> nil) then
   CoTaskMemFree(ppClientItems);
  if (pVariants <> nil) then
   CoTaskMemFree(pVariants);
  if (ppErrors <> nil) then
   CoTaskMemFree(ppErrors);
  if (ppQualityArray <> nil) then
   CoTaskMemFree(ppQualityArray);
  if (ppTimeArray <> nil) then
   CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleThisRequest(cTime:TDateTime);
begin
 if isCancelled then
  Exit;

 if not Assigned(TOPCGroup(grp).ClientIUnknown) then
  Exit;

 if not Succeeded(TOPCGroup(grp).ClientIUnknown.QueryInterface(IID_IOPCDataCallback,Obj)) then
  Exit;

 case kind of
  io2Read:         HandleRead(cTime);
  io2Write:        HandleWrite(cTime);
  io2Refresh:      HandleRefresh(cTime);
  io2Change:       HandleChange(aStream, cTime)
 end;
end;

end.

