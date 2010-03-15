//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uLogging;

interface

uses
  Sysutils, Dialogs, Classes;

type
  TBufferItem = record
    Text: string;
  end;
  PBufferItem = ^TBufferItem;

  TLogging = class
  private
    mFileName : string;
    mUseTimeStamp: Boolean;
    mWriteBuffer: TStringList;
    mWriteInterval: Integer;
    function GenText(aText: string): string;

  public
    constructor Create(aFileName: string);
    procedure Init;
    destructor Destroy; override;
    procedure SetWriteIntervall(aNumber: Integer);
    procedure WriteWithTimeStamp(aTrue: Boolean);
    function WriteIntoFile(aText: string): Boolean;
  end;

implementation

uses
  uMessageDlg;

constructor TLogging.Create(aFileName: string);
begin
  mFileName := aFileName;
  mWriteBuffer := TStringList.Create;
  Init;
end;

procedure TLogging.Init;
begin
  mWriteInterval := 1;
  mUseTimeStamp := False;
end;

destructor TLogging.Destroy;
begin
  if mWriteBuffer <> nil then mWriteBuffer.free;
end;

procedure TLogging.SetWriteIntervall(aNumber: Integer);
begin
  mWriteInterval := aNumber;
  if mWriteInterval < 1 then begin
    mWriteInterval := 1;
  end;
end;

procedure TLogging.WriteWithTimeStamp(aTrue: Boolean);
begin
  mUseTimeStamp := aTrue;
end;

function TLogging.GenText(aText: string): string;
begin
  if mUseTimeStamp then begin
    Result := Format('%s->%s',[FormatDateTime('dd.mm.yyyy hh:nn:ss',now),aText]);
  end else begin
    Result := Format('%s',[aText]);
  end;
end;

function TLogging.WriteIntoFile(aText: String): Boolean;
var
  mFile: Textfile;
  IOError: integer;
begin
  Result := False;
  if (mWriteBuffer.Count+1) < mWriteInterval then begin
    mWriteBuffer.Add(GenText(aText));
    Result := True;
    exit;
  end else begin
    {$I-}
    AssignFile(mFile, mFileName);
    if FileExists(mFileName)
      then Append(mFile)
      else Rewrite(mFile);
    IOError := IOResult;
    if IOError = 0 then begin
      while mWriteBuffer.Count > 0 do begin
        Writeln(mFile,mWriteBuffer.Strings[0]);
        IOError := IOResult;
        if IOError > 0 then begin
          sMessageDlg(Format('WriteToLogFile 1 Error: %d', [IOError]), mtConfirmation, [mbYes], 0);
        end;
        mWriteBuffer.Delete(0);
      end;

      Writeln(mFile,GenText(aText));
      IOError := IOResult;
      if IOError > 0 then begin
        sMessageDlg(Format('WriteToLogFile 2 Error: %d', [IOError]), mtConfirmation, [mbYes], 0);
      end;

      CloseFile(mFile);
      IOError := IOResult;
      if IOError > 0 then begin
        sMessageDlg(Format('WriteToLogFile 3 Error: %d', [IOError]), mtConfirmation, [mbYes], 0);
      end;
    end else begin
      sMessageDlg(Format('WriteToLogFile 4 Error: %d', [IOError]), mtConfirmation, [mbYes], 0);
    end;
    {$I+}
  end;
end;

end.

