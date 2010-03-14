unit ItemsUnit;

interface

uses Windows,ActiveX,ComObj,FirstServ_TLB,SysUtils,Dialogs,Classes,ServIMPL,OPCDA,
     axctrls,Globals,ItemAttributesOPC;

type
  TOPCItem = class
  public
   servObj:TDA2;                         //the owner
   quality,oldQuality:word;
   strID:string;
   pBlob:PByteArray;
   bActive,isWriteAble,hasChanged:longbool;

//i normally have a class of item with a single type for all the types the server
//supports. Since I only have two types I just made one class type.
   currentValueWord,oldValueWord:word;
   currentValueString,oldValueString:string;

   serverItemNum,clientNum,itemIndex:longword;
   vtReqDataType,canonicalDataType:TVarType;
   constructor Create;
   function ChangedCheck:boolean;
   destructor Destroy;override;
   procedure CopyYourSelf(dItem:TOPCItem);
   procedure SetActiveState(state:longbool);
   function GetActiveState:longbool;
   function GetClientHandle:longword;
   function OnChangedCheck:boolean;
   procedure SetClientHandle(h:longword);
   procedure ResolveQuality(source:word);
   function ReturnCurrentValue(source:integer;  var q:word; var err:HRESULT):OleVariant;
   procedure SyncRead(source:word; var pStateRec:OPCITEMSTATE;
                                   var err:HRESULT);
   procedure WriteItemValue(v:oleVariant; var err:HRESULT);
   procedure FillInOPCItemObject(aObj:TOPCItemAttributes);
   procedure SetReqDataType(aType:TVarType);
   procedure CallBackRead(var cHandle:longword; var cValue:OleVariant;
                          var q:word; var err:HRESULT;
                          updateOldStorage:boolean);virtual;
   procedure SetOldValue;
  end;

implementation

uses DataPointsUnit;

constructor TOPCItem.Create;
begin
 Inherited;
 quality:=OPC_QUALITY_BAD;
 oldQuality:=not quality;
 isWriteAble:=false;
end;

function TOPCItem.ChangedCheck:boolean;
begin
 case canonicalDataType of
  VT_BSTR:
   begin
    currentValueString:=TRealDataPoint(rDataPoints[itemIndex]).aString;
    result:=(currentValueString <> oldValueString) or (quality <> oldQuality);
    oldValueString:=currentValueString;
   end;
  VT_UI2:
   begin
    currentValueWord:=TRealDataPoint(rDataPoints[itemIndex]).aWord;
    result:=(currentValueWord <> oldValueWord) or (quality <> oldQuality);
    oldValueWord:=currentValueWord;
   end;
  else
   result:=false;

 end;

 if result then                 //it is cleared when checked
  hasChanged:=true;

 oldQuality:=quality;
end;

destructor TOPCItem.Destroy;
begin
 Inherited;
end;

procedure TOPCItem.CopyYourSelf(dItem:TOPCItem);
begin
 dItem.servObj:=servObj;
 dItem.quality:=quality;
 dItem.strID:=strID;
 dItem.pBlob:=pBlob;
 dItem.bActive:= bActive;
 dItem.isWriteAble:=isWriteAble;

 dItem.currentValueString:=currentValueString;
 dItem.currentValueWord:=currentValueWord;
 dItem.oldValueWord:=oldValueWord;
 dItem.oldValueString:=oldValueString;

 dItem.serverItemNum:=serverItemNum;
 dItem.clientNum:=clientNum;
 dItem.itemIndex:=itemIndex;
 dItem.vtReqDataType:=vtReqDataType;
 dItem.canonicalDataType:=canonicalDataType;
end;

procedure TOPCItem.SetActiveState(state:longbool);
begin
 bActive:=state;
 if not bActive then
  quality:=OPC_QUALITY_OUT_OF_SERVICE
 else
  quality:=OPC_QUALITY_GOOD;
end;

function TOPCItem.GetActiveState:longbool;
begin
 result:=bActive;
end;

function TOPCItem.OnChangedCheck:boolean;
begin
//this is called for the for updates from groups.
 result:=hasChanged;
 hasChanged:=false;
end;

function TOPCItem.GetClientHandle:longword;
begin
 result:=clientNum;
end;

procedure TOPCItem.SetClientHandle(h:longword);
begin
 clientNum:=h;
end;

procedure TOPCItem.ResolveQuality(source:word);
begin
 if (source = OPC_DS_CACHE) then
  begin
   if bActive then                //in service so is it good
    quality:=OPC_QUALITY_GOOD
   else
    quality:=OPC_QUALITY_OUT_OF_SERVICE;
  end
 else                             //device
  quality:=OPC_QUALITY_GOOD;
end;

function TOPCItem.ReturnCurrentValue(source:integer; var q:word; var err:HRESULT):OleVariant;
begin
 ResolveQuality(source);
 q:=quality;
 err:=0;
 if (canonicalDataType = VT_BSTR) then
  result:=currentValueString
 else
  result:=currentValueWord;

 if (vtReqDataType <> canonicalDataType) then
  result:=ConvertVariant(result,vtReqDataType,q,err)
 else
  result:=result;
end;

procedure TOPCItem.SyncRead(source:word; var pStateRec:OPCITEMSTATE;
                                         var err:HRESULT);
begin
 ChangedCheck;                          //for a sync read force a read from the point
 pStateRec.vDataValue:=ReturnCurrentValue(source,pStateRec.wQuality,err);
 pStateRec.hClient:=GetClientHandle;
 DataTimeToOPCTime(servObj.lastClientUpdate,pStateRec.ftTimeStamp);
end;

procedure TOPCItem.WriteItemValue(v:oleVariant; var err:HRESULT);
begin
 err:=S_OK;
 if (canonicalDataType = VT_BSTR) then
  TRealDataPoint(rDataPoints[itemIndex]).aString:=v
 else
  begin
   if (v < low(word)) or (v > high(word)) then
    err:=DISP_E_OVERFLOW
   else
    TRealDataPoint(rDataPoints[itemIndex]).aWord:=v and $FFFF;
  end;
end;

procedure TOPCItem.FillInOPCItemObject(aObj:TOPCItemAttributes);
begin
 if (aObj = nil) then
  Exit;
 aObj.szAccessPath:='';
 aObj.szItemID:=strID;
 aObj.bActive:=bActive;
 aObj.hClient:=clientNum;
 aObj.hServer:=serverItemNum;
 if isWriteAble then
  aObj.dwAccessRights:=OPC_READABLE or OPC_WRITEABLE
 else
  aObj.dwAccessRights:=OPC_READABLE;

 aObj.vtRequestedDataType:=vtReqDataType;
 aObj.vtCanonicalDataType:=canonicalDataType;
 aObj.dwEUType:=OPC_NOENUM;
 aObj.vEUInfo:=VT_EMPTY;
end;

procedure TOPCItem.SetReqDataType(aType:TVarType);
begin
 if (aType = VT_EMPTY) then
  vtReqDataType:=canonicalDataType
 else
  vtReqDataType:=aType;
end;

procedure TOPCItem.CallBackRead(var cHandle:longword; var cValue:OleVariant; var q:word;
                                var err:HRESULT;
                                updateOldStorage:boolean);
begin
 err:=0;
 cHandle:=GetClientHandle;
 ResolveQuality(OPC_DS_CACHE);
 cValue:=ReturnCurrentValue(0,q,err);
 if updateOldStorage then
  begin
   oldQuality:=quality;
   oldValueWord:=currentValueWord;
   oldValueString:=currentValueString;
  end;
end;

procedure TOPCItem.SetOldValue;
begin
 if (canonicalDataType = VT_BSTR) then
  oldValueString:=''
 else
  oldValueWord:=1;
end;

end.
