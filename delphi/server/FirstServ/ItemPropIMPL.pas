unit ItemPropIMPL;

interface

uses Windows,ActiveX,ComObj,FirstServ_TLB,SysUtils,Dialogs,Classes,OPCDA,axctrls,
     OpcError,OPCTypes;

type
  TOPCItemProp = class
  public
   function QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                          out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                          out ppvtDataTypes:PVarTypeList):HResult;stdcall;
   function GetItemProperties(szItemID:POleStr;
                              dwCount:DWORD;
                              pdwPropertyIDs:PDWORDARRAY;
                              out ppvData:POleVariantArray;
                              out ppErrors:PResultList):HResult;stdcall;
   function LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                           out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
end;

implementation

uses ServIMPL,GLobals,Main,Variants,DataPointsUnit;


function TOPCItemProp.QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                       out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                       out ppvtDataTypes:PVarTypeList):HResult;stdcall;
var
 memErr:boolean;
 propID,x:longword;
begin
 if (length(szItemID) = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 propID:=ReturnPropIDFromTagname(szItemID);
 if (propID = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 pdwCount:=7;
 memErr:=false;
 ppPropertyIDs:=PDWORDARRAY(CoTaskMemAlloc(pdwCount * sizeof(DWORD)));
 if (ppPropertyIDs = nil) then
  memErr:=true;
 if not memErr then
  ppDescriptions:=POleStrList(CoTaskMemAlloc(pdwCount * sizeof(POleStr)));
 if (ppDescriptions = nil) then
  memErr:=true;
 if not memErr then
  ppvtDataTypes:=PVarTypeList(CoTaskMemAlloc(pdwCount * sizeof(TVarType)));
 if (ppvtDataTypes = nil) then
  memErr:=true;

 if memErr then
  begin
   if (ppPropertyIDs <> nil) then
    CoTaskMemFree(ppPropertyIDs);
   if (ppDescriptions <> nil) then
    CoTaskMemFree(ppDescriptions);
   if (ppvtDataTypes <> nil) then
    CoTaskMemFree(ppvtDataTypes);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

//the required 6
 x:=0;
 ppPropertyIDs[x]:=1;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_DATATYPE);
 ppvtDataTypes[x]:=VT_I2;
 Inc(x);

 ppPropertyIDs[x]:=2;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_VALUE);
 ppvtDataTypes[x]:=VT_VARIANT;
 Inc(x);

 ppPropertyIDs[x]:=3;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_QUALITY);
 ppvtDataTypes[x]:=VT_I2;
 Inc(x);

 ppPropertyIDs[x]:=4;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_TIMESTAMP);
 ppvtDataTypes[x]:=VT_DATE;
 Inc(x);

 ppPropertyIDs[x]:=5;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_ACCESS_RIGHTS);
 ppvtDataTypes[x]:=VT_I4;
 Inc(x);

 ppPropertyIDs[x]:=6;
 ppDescriptions[x]:=StringToLPOLESTR(OPC_PROPERTY_DESC_SCAN_RATE);
 ppvtDataTypes[x]:=VT_R4;
 Inc(x);

 ppPropertyIDs[x]:=propID;
 ppDescriptions[x]:=StringToLPOLESTR(ReturnTagnameFromPropID(propID));
 ppvtDataTypes[x]:=ReturnDataTypeFromPropID(propID);
 result:=S_OK;
end;

function TOPCItemProp.GetItemProperties(szItemID:POleStr;
                                        dwCount:DWORD;
                                        pdwPropertyIDs:PDWORDARRAY;
                                        out ppvData:POleVariantArray;
                                        out ppErrors:PResultList):HResult;stdcall;
var
 i,x:integer;
 memErr:boolean;
 propID:longword;
begin

 if (length(szItemID) = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 if (dwCount = 0) then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 propID:=ReturnPropIDFromTagname(szItemID);
 if (propID = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 memErr:=false;
 ppvData:=POleVariantArray(CoTaskMemAlloc(dwCount * sizeof(OleVariant)));
 if (ppvData = nil) then
  memErr:=true;

 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount * sizeof(HRESULT)));
 if (ppErrors = nil) then
  memErr:=true;

 if memErr then
  begin
   if (ppvData <> nil) then
    CoTaskMemFree(ppvData);
   if (ppErrors <> nil) then
    CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 FillChar(ppvData[0], dwCount * sizeOf(OleVariant),#0);

 result:=S_OK;
 for i:= 0 to dwCount - 1 do
  begin
//the first six are predefined by the spec
   case pdwPropertyIDs[i] of
    1:  ppvData[i]:=VarAsType(ReturnDataTypeFromPropID(propID),VT_I2);
    2:
     begin
      x:=(propID - posItems[low(posItems)].PropID);
      if (ReturnDataTypeFromPropID(propID) = VT_BSTR) then
       ppvData[i]:=TRealDataPoint(rDataPoints[x]).aString
      else
       ppvData[i]:=TRealDataPoint(rDataPoints[x]).aWord;
     end;
    3:  ppvData[i]:=VarAsType(OPC_QUALITY_GOOD,VT_I2);
    4:  ppvData[i]:=VarAsType(Now,VT_DATE);
    5:
     begin
      if CanPropIDBeWritten(propID) then
       ppvData[i]:=VarAsType(OPC_READABLE or OPC_WRITEABLE or 3,VT_I4)
      else
       ppvData[i]:=VarAsType(OPC_READABLE,VT_I4);
     end;
    6:  ppvData[i]:=VarAsType(1000,VT_R4);
    5000..5013:
     begin
      x:=(pdwPropertyIDs[i] - posItems[low(posItems)].PropID);
      if (ReturnDataTypeFromPropID(propID) = VT_BSTR) then
       ppvData[i]:=TRealDataPoint(rDataPoints[x]).aString
      else
       ppvData[i]:=TRealDataPoint(rDataPoints[x]).aWord;
     end;
    else
     begin
      ppErrors[i]:=OPC_E_INVALID_PID;
      result:=S_FALSE;
      Continue;
     end;
   end;
   ppErrors[i]:=S_OK;
  end;

end;

function TOPCItemProp.LookupItemIDs(szItemID:POleStr; dwCount:DWORD;
                                    pdwPropertyIDs:PDWORDARRAY;
                                    out ppszNewItemIDs:POleStrList;
                                    out ppErrors:PResultList): HResult; stdcall;
var
 i:integer;
 propID:longword;
 memErr:boolean;
begin
 if (length(szItemID) = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 if (dwCount < 1) then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 propID:=ReturnPropIDFromTagname(szItemID);
 if (propID = 0) then
  begin
   result:=OPC_E_INVALIDITEMID;
   Exit;
  end;

 memErr:=false;
 ppszNewItemIDs:=POleStrList(CoTaskMemAlloc(dwCount*sizeof(POleStr)));
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if (ppErrors = nil) then
  memErr:=true;

 if memErr then
  begin
   if (ppszNewItemIDs <> nil) then
    CoTaskMemFree(ppszNewItemIDs);
   if (ppErrors <> nil) then
    CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;

 for i:= 0 to dwCount - 1 do
  begin
   if (pdwPropertyIDs[i] < 7) then              //do not return the first 6 --compliance
    begin
     ppszNewItemIDs[i]:=StringToLPOLESTR('');
     ppErrors[i]:=OPC_E_INVALID_PID;
     result:=S_FALSE;
     Continue;
    end;

   if (pdwPropertyIDs[i] <> propID) then       //we only have the one
    begin
     ppszNewItemIDs[i]:=StringToLPOLESTR('');           //this device does not have the id
     ppErrors[i]:=OPC_E_INVALID_PID;
     result:=S_FALSE;
     Continue;
    end;


   ppszNewItemIDs[i]:=StringToLPOLESTR(szItemID);
   ppErrors[i]:=S_OK;
  end;

end;

end.
