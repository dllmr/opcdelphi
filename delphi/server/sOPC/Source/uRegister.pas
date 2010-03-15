//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uRegister;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComObj, StdCtrls, Registry,
  OPCDA, sOPC_TLB;

type
  TOPCAutoObjectFactory = class (TAutoObjectFactory)
  public
    procedure UpdateRegistry(Register: Boolean); override;
  end;

function RegisterOPCServer: boolean;
function UnRegisterOPCServer: boolean;

implementation

uses
  ActiveX,
  ComCat, uGlobals, uOPC;

procedure TOPCAutoObjectFactory.UpdateRegistry(Register: Boolean);
const
  ThreadStrs: array[TThreadingModel] of string =
    ('', 'Apartment', 'Free', 'Both'
    , 'Neutral'
    );
var
  ClassID, ProgID, ServerKeyName, ShortFileName: string;
  ClassKey: string;
  TypeLib: ITypeLib;
  TLibAttr: PTLibAttr;
begin
  // UpdateRegistry TComObjectFactory
  if Instancing = ciInternal then exit;
  ClassID := GUIDToString(inherited ClassID);
  ProgID := OPC.ServerName;
  ServerKeyName := 'CLSID\' + ClassID + '\' + ComServer.ServerKey;
  if Register then begin
    CreateRegKey('CLSID\' + ClassID, '', OPC.ServerDescription);
    ShortFileName := ComServer.ServerFileName;
    if AnsiPos(' ', ShortFileName) <> 0 then
      ShortFileName := ExtractShortPathName(ShortFileName);
    CreateRegKey(ServerKeyName, '', ShortFileName);
    if (ThreadingModel <> tmSingle) and IsLibrary then
      CreateRegKey(ServerKeyName, 'ThreadingModel',
        ThreadStrs[ThreadingModel]);
    if ProgID <> '' then begin
      CreateRegKey(ProgID, '', OPC.ServerDescription);
      CreateRegKey(ProgID + '\Clsid', '', ClassID);
      CreateRegKey('CLSID\' + ClassID + '\ProgID', '', ProgID);
    end;
  end else begin
    if ProgID <> '' then begin
      DeleteRegKey('CLSID\' + ClassID + '\ProgID');
      DeleteRegKey(ProgID + '\Clsid');
      DeleteRegKey(ProgID);
    end;
    DeleteRegKey(ServerKeyName);
    DeleteRegKey('CLSID\' + ClassID);
  end;

  // UpdateRegistry TTypedComObjectFactory
  ClassKey := 'CLSID\' + GUIDToString(inherited ClassID);
  if Register then begin
    TypeLib := ComServer.TypeLib;
    OleCheck(TypeLib.GetLibAttr(TLibAttr));
    try
      CreateRegKey(ClassKey + '\Version', '', Format('%d.%d',
        [TLibAttr.wMajorVerNum, TLibAttr.wMinorVerNum]));
      CreateRegKey(ClassKey + '\TypeLib', '', GUIDToString(TLibAttr.guid));
    finally
      TypeLib.ReleaseTLibAttr(TLibAttr);
    end;
  end else begin
    DeleteRegKey(ClassKey + '\TypeLib');
    DeleteRegKey(ClassKey + '\Version');
  end;
end;

function RegisterOPCServer: boolean;
var
  Registry: TRegistry;
  stGUID: string;
begin
  Result := False;
  stGUID := GUIDToString(CLASS_OPCDataAccess20);
  try
    Registry := TRegistry.Create;
    Registry.RootKey := HKEY_CLASSES_ROOT;
    // HKCR\CLSID\AppID
    if not Registry.OpenKey('\CLSID\' + stGUID, True) then exit;
    Registry.WriteString('AppID', stGUID);
    Registry.CloseKey;
    // HKCR\AppID
    if not Registry.OpenKey('\AppID\' + stGUID, True) then exit;
    Registry.WriteString('', OPC.ServerDescription);
    Registry.WriteString('RunAs', 'Interactive User');
    Registry.CloseKey;
    Registry.Free;
  except
    on E: Exception do begin
      OPCLogException('RegisterOPCServer', E);
      exit;
    end;
  end;
  if CreateComponentCategory(CATID_OPCDAServer20, 'OPC Daten Server V2.0') <> 0 then exit;
  if RegisterCLSIDInCategory(CLASS_OPCDataAccess20, CATID_OPCDAServer20) <> 0 then exit;
  Result := True;
end;

function UnRegisterOPCServer: boolean;
var
  Registry: TRegistry;
  stGUID: string;
begin
  Result := False;
  stGUID := GUIDToString(CLASS_OPCDataAccess20);
  UnRegisterCLSIDInCategory(CLASS_OPCDataAccess20, CATID_OPCDAServer20);
  try
    Registry := TRegistry.Create;
    Registry.RootKey := HKEY_CLASSES_ROOT;
    // delete key HKCR\AppID\GUID
    Registry.DeleteKey('\AppID\' + stGUID);
    // delete key HKCR\CLSID\GUID
    Registry.DeleteKey('\CLSID\' + stGUID);
    Registry.Free;
  except
    on E: Exception do begin
      OPCLogException('UnRegisterOPCServer', E);
      exit;
    end;
  end;
  Result := True;
end;

end.

