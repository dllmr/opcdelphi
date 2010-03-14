program OPCsrvrs;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

uses
  Classes, Forms, OPCDA, OPCenum;

{$R *.RES}

var
  OPCServerList: TOPCServerList;
  CATIDs: array of TGUID;
  ServerNames: TStringList;
  I: Integer;

// main program code

begin
  // among other things, this call makes sure that COM is initialized
  Application.Initialize;
  Writeln('========================================================');
  Writeln('OPC server enumerator, by Mike Dillamore, 2003-2009.    ');
  Writeln('Tested for compatibility with Delphi 5-7 and 2005-2010. ');
  Writeln('========================================================');
  Writeln;

  ServerNames := TStringList.Create;

  try
    SetLength(CATIDs, 0);
    OPCServerList := TOPCServerList.Create('', True, CATIDs);
    try
      OPCServerList.Update;
      ServerNames.AddStrings(OPCServerList.Items);
    finally
      OPCServerList.Free;
    end;
    Writeln('OPC servers from registry: -');
    for I := 0 to ServerNames.Count - 1 do
    begin
      Writeln(ServerNames[I]);
    end;

    Writeln;
    ServerNames.Clear;

    SetLength(CATIDs, 1);
    CATIDs[0] := CATID_OPCDAServer20;
    OPCServerList := TOPCServerList.Create('', False, CATIDs);
    try
      OPCServerList.Update;
      ServerNames.AddStrings(OPCServerList.Items);
    finally
      OPCServerList.Free;
    end;
    Writeln('OPC DA 2.0 servers from server browser: -');
    for I := 0 to ServerNames.Count - 1 do
    begin
      Writeln(ServerNames[I]);
    end;

  finally
    ServerNames.Free;
  end;
end.
