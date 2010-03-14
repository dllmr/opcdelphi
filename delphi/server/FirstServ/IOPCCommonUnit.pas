unit IOPCCommonUnit;

interface

uses Windows,ComObj,ActiveX,Axctrls,SysUtils,Dialogs,Classes,StdVCL,OPCCOMN;

//*******************************************************************TIOPCCommon
type
 TIOPCCommon = class
 public
  localID:longword;
  clientName:string;
  constructor Create;
  function SetLocaleID(dwLcid:TLCID):HResult;stdcall;
  function GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
  function QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
  function GetErrorString(dwError:HResult; out ppString:POleStr):HResult;overload;stdcall;
  function SetClientName(szName:POleStr):HResult;stdcall;
end;

implementation

uses OPCErrorStrings;

constructor TIOPCCommon.Create;
begin
 clientName:='Client has not set the name';
 localID:=LOCALE_SYSTEM_DEFAULT;
end;

function TIOPCCommon.SetLocaleID(dwLcid:TLCID):HResult;stdcall;
begin
 if (dwLcid = LOCALE_SYSTEM_DEFAULT) or (dwLcid = LOCALE_USER_DEFAULT) then
  begin
   localID:=dwLcid;
   result:=S_OK;
  end
 else
  result:=E_INVALIDARG;
end;

function TIOPCCommon.GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
begin
 pdwLcid:=localID;
 result:=S_OK;
end;

function TIOPCCommon.QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
begin
 pdwCount:=2;
 pdwLcid:=PLCIDARRAY(CoTaskMemAlloc(pdwCount * sizeof(LCID)));
 if (pdwLcid = nil) then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;
 pdwLcid[0]:=LOCALE_SYSTEM_DEFAULT;
 pdwLcid[1]:=LOCALE_USER_DEFAULT;
 result:=S_OK;
end;

function TIOPCCommon.GetErrorString(dwError:HResult; out ppString:POleStr):HResult;stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TIOPCCommon.SetClientName(szName:POleStr):HResult;stdcall;
begin
 if (addr(szName) = nil) then
  begin
   Result:=E_INVALIDARG;
   Exit;
  end;
 clientName:=szName;
 result:=S_OK;
end;


end.
