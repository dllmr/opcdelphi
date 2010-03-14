{
Unit to wrap COMCAT.DLL functions. See MSDN for documentation on the use of Component Categories.
Copyright 1999, Dan Miser (dmiser@execpc.com).

All use is free subject to the following terms:
  1) The entire file, including comments, must remain intact.
  2) The author reserves the right to exclude specific companies from using this file,
     solely based on the author's discretion.
}
unit ComCat;

interface

uses
  Windows, Classes, AXCtrls, ActiveX;

// ICatRegister helper functions
function CreateComponentCategory(catid : TIID; catDescription : PWideChar) : HRESULT;
function UnCreateComponentCategory(catid : TGUID; catDescription : PWideChar) : HRESULT;
function RegisterCLSIDInCategory(clsid : TGUID; catid : TGUID) : HRESULT;
function UnRegisterCLSIDInCategory(clsid : TGUID; catid : TGUID) : HRESULT;

//ICatInformation helper functions
function GetCategoryList(Strings: TStrings): HRESULT;
function GetCategoryCLSIDList(catid: TGUID; Strings: TStrings): HRESULT;
function GetCategoryProgIDList(catid: TGUID; Strings: TStrings): HRESULT;

implementation

uses
  ComObj;

// ICatRegister
function CreateComponentCategory(catid : TGUID; catDescription : PWideChar) : HRESULT;
const
  MAX_LEN = 127;
var
  pcr : ICatRegister;
  catinfo : PCATEGORYINFO;
  len : integer;
  s : string;
begin
  Result:=CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatRegister, pcr);
  OleCheck(Result);

  // Make sure the HKCR\Component Categories\{..catid...} key is registered
  New(catinfo);
  try
    catinfo^.catid := catid;
    catinfo^.lcid := GetUserDefaultLCID; // or use $0409; (english)

    // Make sure the provided description is not too long.
    // Only copy the first 127 characters if it is
    s:=WideCharToString(catDescription);
    len := length(s);
    if len>MAX_LEN then
      SetLength(S, MAX_LEN);
    StringToWideChar(s, catinfo^.szDescription, len+1); // Need room for NULL character too

    Result := pcr.RegisterCategories(1, catinfo);
  finally
    Dispose(catinfo);
  end;
end;

//MRD
function UnCreateComponentCategory(catid : TGUID; catDescription : PWideChar) : HRESULT;
const
  MAX_LEN = 127;
var
  pcr : ICatRegister;
  catinfo : PCATEGORYINFO;
  len : integer;
  s : string;
begin
  Result:=CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatRegister, pcr);
  OleCheck(Result);

  // Make sure the HKCR\Component Categories\{..catid...} key is registered
  New(catinfo);
  try
    catinfo^.catid := catid;
    catinfo^.lcid := GetUserDefaultLCID; // or use $0409; (english)

    // Make sure the provided description is not too long.
    // Only copy the first 127 characters if it is
    s:=WideCharToString(catDescription);
    len := length(s);
    if len>MAX_LEN then
      SetLength(S, MAX_LEN);
    StringToWideChar(s, catinfo^.szDescription, len+1); // Need room for NULL character too

    Result := pcr.UnRegisterCategories(1, catinfo);
  finally
    Dispose(catinfo);
  end;
end;
//MRD

function RegisterCLSIDInCategory(clsid : TGUID; catid : TGUID) : HRESULT;
var
  pcr : ICatRegister;
  rgcatid : array[0..1] of TGUID;
begin
  Result := CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatRegister, pcr);
  OleCheck(Result);

  // Register this category as being "implemented" by the class.
  rgcatid[0] := catid;
  Result := pcr.RegisterClassImplCategories(clsid, 1, @rgcatid);
end;

function UnRegisterCLSIDInCategory(clsid : TGUID; catid : TGUID) : HRESULT;
var
  pcr : ICatRegister;
  rgcatid : array[0..1] of TGUID;
begin
  Result := CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatRegister, pcr);
  OleCheck(Result);

  // Unregister this category as being "implemented" by the class.
  rgcatid[0] := catid;
  Result := pcr.UnRegisterClassImplCategories(clsid, 1, @rgcatid);
end;

// ICatInformation
function GetCategoryList(Strings: TStrings): HRESULT;
var
  pci: ICatInformation;
  peci: IEnumCATEGORYINFO;
//  Fetched: integer;
  Fetched: UINT;      // MRD
  catinfo: TCATEGORYINFO;
begin
  Result:=CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatInformation, pci);
  OleCheck(Result);

  pci.EnumCategories(GetUserDefaultLCID, peci);
  while peci.Next(1, catinfo, Fetched) = S_OK do
    Strings.Add(WideCharToString(catinfo.szDescription));
end;

function GetCategoryCLSIDList(catid: TGUID; Strings: TStrings): HRESULT;
var
  pci: ICatInformation;
  peg: IEnumGUID;
//  Fetched: integer;
  Fetched: UINT;          //MRD
  guid: TGUID;
begin
  Result:=CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatInformation, pci);
  OleCheck(Result);

  // We could expand this to accept a list of GUIDs:
  // both implemented and required
//  pci.EnumClassesOfCategories(1, @catid, -1, nil, peg);
  pci.EnumClassesOfCategories(1, @catid, UINT(-1), nil, peg);           //MRD
  while peg.Next(1, guid, Fetched) = S_OK do
    Strings.Add(GUIDToString(guid));
end;
//    function EnumClassesOfCategories(cImplemented: UINT; rgcatidImpl: Pointer; cRequired: UINT; rgcatidReq: Pointer; out ppenumClsid: IEnumGUID): HResult; stdcall;

function GetCategoryProgIDList(catid: TGUID; Strings: TStrings): HRESULT;
var
  pci: ICatInformation;
  peg: IEnumGUID;
//  Fetched: integer;
  Fetched: UINT;             //MRD
  guid: TGUID;
begin
  Result:=CoCreateInstance(CLSID_StdComponentCategoryMgr, nil,
    CLSCTX_INPROC_SERVER, ICatInformation, pci);
  OleCheck(Result);

  // We could expand this to accept a list of GUIDs:
  // both implemented and required
//  pci.EnumClassesOfCategories(1, @catid, -1, nil, peg);
  pci.EnumClassesOfCategories(1, @catid, UINT(-1), nil, peg);   //MRD
  while peg.Next(1, guid, Fetched) = S_OK do
    Strings.Add(ClassIDToProgID(guid));
end;

end.
