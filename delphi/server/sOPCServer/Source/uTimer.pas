//******************************************************************************
// sOPC created by Schmid IT-Management, http://www.schmid-itm.de/
//******************************************************************************
unit uTimer;

interface

uses
  Windows;

//******************************************************************************
type
  sTimer = class(TObject)
  private
    Freq: int64;
    Time: int64;
    FStopTime: int64;                         // Zeitmarke für Stopaufruf
    FRunning: boolean;

    function ReadTimer: int64;                // Zeitmarke holen
    function TimeDiff(Start: int64): int64;   // Zeitdifferenz in Ticks
    function msTimeDiff(Start: int64): int64; // Zeitdifferenz in Millisekunden [ms]

  public
    constructor Create(AutoStart: boolean = True);

    procedure Start;
    // Timer starten, Aufruf von Start stellt den Timer wieder
    // auf Null

    procedure Stop;
    // Timer stoppen, -> isRunning = False

    function isRunning: boolean;
    // True -> Timer läuft

    function msTime: double;
    // vergangene Zeit in Millisekunden [ms] mit µs Auflösung
    // z.B. 100 µs -> Rückgabewert 0.1

    function secTime: int64;
    // vergangene Zeit in Sekunden [sec]

    procedure Delay(ms: integer);
    // Zeitverzögerung in [ms]
  end;

//******************************************************************************
implementation

//******************************************************************************
function sTimer.ReadTimer: int64;
begin
  QueryPerformanceCounter(TLargeInteger(Result));
end;

function sTimer.TimeDiff(Start: int64): int64;
begin
  Result := 0;
  if not FRunning then begin
    if FStopTime > Start then Result := (FStopTime - Start);
  end else begin
    Result := (ReadTimer - Start);
  end;
end;

function sTimer.msTimeDiff(Start: int64): int64;
begin
  Result := (TimeDiff(Start) * 1000) div Freq;
end;

constructor sTimer.Create(AutoStart: boolean);
var
  aFreq: int64;
begin
  inherited Create;
  FRunning := False;
  Time := 0;
  FStopTime := 0;
  QueryPerformanceFrequency(aFreq);
  Freq := aFreq;
  if AutoStart then Start;
end;

procedure sTimer.Start;
begin
  FRunning := True;
  Time := ReadTimer;
end;

procedure sTimer.Stop;
begin
  // FStopTime wird nur gesetzt, wenn vorher Start aufgerufen wurde!
  if FRunning then FStopTime := ReadTimer;
  FRunning := False;
end;

function sTimer.isRunning: boolean;
begin
  Result := FRunning;
end;

function sTimer.msTime: double;
begin
  Result := (TimeDiff(Time) * 1000) / Freq;
end;

function sTimer.secTime: int64;
begin
  Result := TimeDiff(Time) div Freq;
end;

procedure sTimer.Delay(ms: integer);
var
  t: int64;
begin
  t := ReadTimer;
  while msTimeDiff(t) < ms do;
end;

//******************************************************************************
end.

