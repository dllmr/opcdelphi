unit RegDeRegServer;

interface

uses Windows,Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
     FirstServ_TLB,Registry,OPCDA,ActiveX;

procedure RegisterTheServer(const name:string);
procedure UnRegisterTheServer(const name:string);

implementation

uses ComObj,ComCat;

procedure RegisterTheServer(const name:string);
var
 hr:HRESULT;
 aCLSIDString:string;
begin
 aCLSIDString:=GUIDToString(CLASS_DA2);
                                                //key, value name, value
//1. HKEY_CLASSES_ROOT\Vendor.Drivername.Version = A Description of your server
 CreateRegKey(name,'','MRD Data Access Server Version 2.0',HKEY_CLASSES_ROOT);

//2. HKEY_CLASSES_ROOT\Vendor.Drivername.Version\CLSID = {Your Server’s unique CLSID}
 CreateRegKey(name + '\CLSID','',aCLSIDString,HKEY_CLASSES_ROOT);

//3. HKEY_CLASSES_ROOT\Vendor.Drivername.Version\OPC
 CreateRegKey(name + '\OPC','','',HKEY_CLASSES_ROOT);

//4. HKEY_CLASSES_ROOT\Vendor.Drivername.Version\OPC\Vendor vendor name
 CreateRegKey(name + '\OPC\Vendor','','@2008 Everest Software',HKEY_CLASSES_ROOT);

//5. HKEY_CLASSES_ROOT\CLSID\{Your Server’s unique CLSID} = A Description of your server
 CreateRegKey('CLSID\' + aCLSIDString,'','MRD OPC Data Access',HKEY_CLASSES_ROOT);

//6. HKEY_CLASSES_ROOT\CLSID\{Your Server’s unique CLSID}\ProgID = Vendor.Drivername.Version
 CreateRegKey('CLSID\' + aCLSIDString + '\ProgID','',name,HKEY_CLASSES_ROOT);

//make the category and register. Must be present for 2.0
 try
  CoInitialize(nil);
  hr:=CreateComponentCategory(CATID_OPCDAServer20,'OPC Data Access Server Version 2.05');
  if Failed(hr) then
   ;

  hr:=RegisterCLSIDInCategory(CLASS_DA2,CATID_OPCDAServer20);
  if Failed(hr) then
   ;

 finally
  CoUninitialize;
 end;
end;

procedure UnRegisterTheServer(const name:string);
var
 hr:HRESULT;
 aCLSIDString:string;
begin
 aCLSIDString:=GUIDToString(CLASS_DA2);

//delete sub keys first
 DeleteRegKey('CLSID\' + aCLSIDString + '\ProgID',HKEY_CLASSES_ROOT);
 DeleteRegKey('CLSID\' + aCLSIDString,HKEY_CLASSES_ROOT);
 DeleteRegKey(name + '\CLSID',HKEY_CLASSES_ROOT);
 DeleteRegKey(name + '\OPC\Vendor',HKEY_CLASSES_ROOT);
 DeleteRegKey(name + '\OPC',HKEY_CLASSES_ROOT);
 DeleteRegKey(name,HKEY_CLASSES_ROOT);

 try
  CoInitialize(nil);
  hr:=UnRegisterCLSIDInCategory(CLASS_DA2,CATID_OPCDAServer20);
  if hr <> 0 then
   ;

  hr:=UnCreateComponentCategory(CATID_OPCDAServer20,'MRD OPC Data Access');
  if hr <> 0 then
   ;
 finally
  CoUninitialize;
 end;
end;

end.
