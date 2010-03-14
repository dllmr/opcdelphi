unit MSHTTPPost;

interface

uses
  classes;

procedure PostData(Url, Proxy, Username,Password, SoapAction: string;
                   const Request: TStream; Response: TStream);

implementation

uses
  wininet, windows, sysutils;

resourcestring
  SInvalidURL               = 'no or invalid url';
  SContentTypeUTF8          = 'text/xml; charset="utf-8"';
  SInvalidHTTPRequest       = 'invalid HTTP request';
  SContentTypeTextPlain     = 'text/plain';
  STextHtml                 = 'text/html';
  sTextXML                  = 'text/xml';
  SInvalidContentType       = 'invalid content type: %s - SOAP required "text/xml"';
  SHTTPSoapAction           = 'SOAPAction';
  SContentType              = 'Content-Type';
  SInvalidHTTPResponse      = 'Invalid HTTP response';
  SAgentName                = 'dOPC XML DA core client';

const
  MaxSinglePostSize = $8000;
  MaxStatusTest     = 4096;
  MaxContentType    = 256;

type
    THTTPMSPoster = class
    private
      FActive     : boolean;
      FProxy      : string;
      FURL        : string;
      FProxyByPass: string;
      FSoapAction : string;
      FUserName   : string;
      FPassword   : string;
      FAgent      : string;
      FInetRoot   : HINTERNET;
      FInetConnect: HINTERNET;
      FURLScheme  : Integer;
      FURLSite    : string;
      FURLHost    : string;
      FURLPort    : word;
      FConnectTimeout: integer;
      FSendTimeout   : integer;
      FReceiveTimeout: integer;
      function  Send(const ASrc: TStream): Integer;
      procedure Receive(Context: Integer; Resp: TStream);
      procedure Check(Error: Boolean; ShowSOAPAction: Boolean = False);
      procedure InitURL(const Value: string);
      procedure Connect(Value: Boolean);
    public
      constructor create;
      destructor destroy; override;
      procedure PostData(const Request: TStream; Response: TStream);
    end;



procedure PostData(Url, Proxy, Username,Password, SoapAction: string;
                   const Request: TStream; Response: TStream);
var
  P : THTTPMSPoster;
begin
  P := THTTPMSPoster.create;
  try
    P.FUrl        := Url;
    P.FProxy      := Proxy;
    P.FUsername   := Username;
    P.FPassword   := Password;
    P.FSoapAction := SoapAction;
    P.FAgent      := SAgentName;
    P.PostData(Request,Response);
  finally
    P.Free;
  end;
end;


constructor THTTPMSPoster.create;
begin
  inherited;
  FProxy      := '';
  FURL        := '';
  FProxyByPass:= '';
  FSoapAction := '';
  FUserName   := '';
  FPassword   := '';
  FAgent      := '';
  FConnectTimeout := 0;
  FSendTimeout    := 0;
  FReceiveTimeout := 0;
  FInetRoot   := nil;
  FInetConnect:= nil;
  FURLHost    := '';
  FURLPort    := 80;
  FURLScheme  := 0;
  FURLSite    := '';
end;


procedure THTTPMSPoster.Check(Error: Boolean; ShowSOAPAction: Boolean);
var
  ErrCode: Integer;
  S: string;
begin
  ErrCode := GetLastError;
  if Error and (ErrCode <> 0) then
  begin
    SetLength(S, 256);
    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_FROM_HMODULE, Pointer(GetModuleHandle('wininet.dll')),
      ErrCode, 0, PChar(S), Length(S), nil);
    SetLength(S, StrLen(PChar(S)));
    while (Length(S) > 0) and (S[Length(S)] in [#10, #13]) do
      SetLength(S, Length(S) - 1);
    raise Exception.CreateFmt('%s - URL:%s - SOAPAction:%s', [S, FURL, FSoapAction]);      { Do not localize }
  end;
end;



procedure THTTPMSPoster.Connect(Value: Boolean);
var
  AccessType: Integer;
  Flags: DWord;
begin
  if Value then
  begin
    if (FActive) then
      Exit;
    InitURL(FURL);
    { Proxy?? }
    if Length(FProxy) > 0 then
      AccessType := INTERNET_OPEN_TYPE_PROXY
    else
      AccessType := INTERNET_OPEN_TYPE_PRECONFIG;
     Flags:= 0;
    //Flags:= INTERNET_FLAG_ASYNC;

    { Also, could switch to new API introduced in IE4/Preview2}
    if InternetAttemptConnect(0) <> ERROR_SUCCESS then
      SysUtils.Abort;

    FInetRoot := InternetOpen(PChar(FAgent), AccessType, PChar(FProxy), PChar(FProxyByPass), Flags);
    Check(not Assigned(FInetRoot));
    try
      FInetConnect := InternetConnect(FInetRoot, PChar(FURLHost), FURLPort, PChar(FUserName),
        PChar(FPassword), INTERNET_SERVICE_HTTP, 0, 0);
      Check(not Assigned(FInetConnect));
      FActive := True;
    except
      InternetCloseHandle(FInetRoot);
      FInetRoot := nil;
      raise;
    end;
  end
  else
  begin
    if Assigned(FInetConnect) then
      InternetCloseHandle(FInetConnect);
    FInetConnect := nil;
    if Assigned(FInetRoot) then
      InternetCloseHandle(FInetRoot);
    FInetRoot := nil;
    FActive := False;
  end;
end;

procedure THTTPMSPoster.InitURL(const Value: string);
var
  URLComp: TURLComponents;
  P: PChar;
begin
  if Value <> '' then
  begin
    FillChar(URLComp, SizeOf(URLComp), 0);
    URLComp.dwStructSize := SizeOf(URLComp);
    URLComp.dwSchemeLength := 1;
    URLComp.dwHostNameLength := 1;
    URLComp.dwURLPathLength := 1;
    P := PChar(Value);
    InternetCrackUrl(P, 0, 0, URLComp);
    if not (URLComp.nScheme in [INTERNET_SCHEME_HTTP, INTERNET_SCHEME_HTTPS]) then
      raise Exception.CreateFmt(SInvalidURL, [Value]);
    FURLScheme := URLComp.nScheme;
    FURLPort := URLComp.nPort;
    FURLHost := Copy(Value, URLComp.lpszHostName - P + 1, URLComp.dwHostNameLength);
    FURLSite := Copy(Value, URLComp.lpszUrlPath - P + 1, URLComp.dwUrlPathLength);
  end
  else
    raise Exception.Create(SInvalidURL);
end;


function GetSOAPActionHeader(SoapAction: string): string;
begin
  if (SoapAction = '') then
    Result := SHTTPSoapAction + ':'
  else if (SoapAction = '""') then
    Result := SHTTPSoapAction + ': ""'
  else
    Result := SHTTPSoapAction + ': ' + '"' + SoapAction + '"';
end;


function THTTPMSPoster.Send(const ASrc: TStream): Integer;
var
  Request: HINTERNET;
  RetVal, Flags: DWord;
  P: Pointer;
  ActionHeader: string;
  ContentHeader: string;
  BuffSize, Len: Integer;
  INBuffer: INTERNET_BUFFERS;
  Buffer: TMemoryStream;
  StrStr: TStringStream;
begin
  { Connect }
  Connect(True);

  Flags := INTERNET_FLAG_KEEP_CONNECTION or INTERNET_FLAG_NO_CACHE_WRITE;

  if FURLScheme = INTERNET_SCHEME_HTTPS then
  begin
    Flags := Flags or INTERNET_FLAG_SECURE;
    Flags := Flags or (INTERNET_FLAG_IGNORE_CERT_CN_INVALID or
                       INTERNET_FLAG_IGNORE_CERT_DATE_INVALID);
  end;

  Request := nil;
  try
    Request := HttpOpenRequest(FInetConnect, 'POST', PChar(FURLSite), nil,
                               nil, nil, Flags, 0);
    Check(not Assigned(Request));

    { Timeouts }
    if FConnectTimeout > 0 then
      Check(InternetSetOption(Request, INTERNET_OPTION_CONNECT_TIMEOUT, Pointer(@FConnectTimeout), SizeOf(FConnectTimeout)));
    if FSendTimeout > 0 then
      Check(InternetSetOption(Request, INTERNET_OPTION_SEND_TIMEOUT, Pointer(@FSendTimeout), SizeOf(FSendTimeout)));
    if FReceiveTimeout > 0 then
      Check(InternetSetOption(Request, INTERNET_OPTION_RECEIVE_TIMEOUT, Pointer(@FReceiveTimeout), SizeOf(FReceiveTimeout)));

    { SOAPAction header }
    ActionHeader:= GetSOAPActionHeader(FSoapAction);
    HttpAddRequestHeaders(Request, PChar(ActionHeader), Length(ActionHeader), HTTP_ADDREQ_FLAG_ADD);

    { Content-Type }
    ContentHeader := Format('Content-Type: %s', [sContentTypeUTF8]);
    HttpAddRequestHeaders(Request, PChar(ContentHeader), Length(ContentHeader), HTTP_ADDREQ_FLAG_ADD);

    ASrc.Position := 0;
    BuffSize := ASrc.Size;
    if BuffSize > MaxSinglePostSize then
    begin
      Buffer := TMemoryStream.Create;
      try
        Buffer.SetSize(MaxSinglePostSize);

        { Init Input Buffer }
        INBuffer.dwStructSize := SizeOf(INBuffer);
        INBuffer.Next := nil;
        INBuffer.lpcszHeader := nil;
        INBuffer.dwHeadersLength := 0;
        INBuffer.dwHeadersTotal := 0;
        INBuffer.lpvBuffer := nil;
        INBuffer.dwBufferLength := 0;
        INBuffer.dwBufferTotal := BuffSize;
        INBuffer.dwOffsetLow := 0;
        INBuffer.dwOffsetHigh := 0;

        { Start POST }
        Check(not HttpSendRequestEx(Request, @INBuffer, nil,
                                    HSR_INITIATE or HSR_SYNC, cardinal(self)));
        try
          while True do
          begin
            { Calc length of data to send }
            Len := BuffSize - ASrc.Position;
            if Len > MaxSinglePostSize then
              Len := MaxSinglePostSize;
            { Bail out if zip.. }
            if Len = 0 then
              break;
            { Read data in buffer and write out}
            Len := ASrc.Read(Buffer.Memory^, Len);
            if Len = 0 then
              raise Exception.Create(SInvalidHTTPRequest);

            Check(not InternetWriteFile(Request, @Buffer.Memory^, Len, RetVal));

            RetVal := InternetErrorDlg(GetDesktopWindow(), Request, GetLastError,
              FLAGS_ERROR_UI_FILTER_FOR_ERRORS or FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS or
              FLAGS_ERROR_UI_FLAGS_GENERATE_DATA, P);
            case RetVal of
              ERROR_SUCCESS: ;
              ERROR_CANCELLED: SysUtils.Abort;
              ERROR_INTERNET_FORCE_RETRY: {Retry the operation};
            end;
          end;
        finally
          Check(not HttpEndRequest(Request, nil, 0, cardinal(self)));
        end;
      finally
        Buffer.Free;
      end;
    end else
    begin
      StrStr := TStringStream.Create('');
      try
        StrStr.CopyFrom(ASrc, 0);
        while True do
        begin
          Check(not HttpSendRequest(Request, nil, 0, @StrStr.DataString[1], Length(StrStr.DataString)));
          RetVal := InternetErrorDlg(GetDesktopWindow(), Request, GetLastError,
            FLAGS_ERROR_UI_FILTER_FOR_ERRORS or FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS or
            FLAGS_ERROR_UI_FLAGS_GENERATE_DATA, P);
          case RetVal of
            ERROR_SUCCESS: break;
            ERROR_CANCELLED: SysUtils.Abort;
            ERROR_INTERNET_FORCE_RETRY: {Retry the operation};
          end;
        end;
      finally
        StrStr.Free;
      end;
    end;
  except
    if (Request <> nil) then
      InternetCloseHandle(Request);
    Connect(False);
    raise;
  end;
  Result := Integer(Request);
end;



procedure  THTTPMSPoster.Receive(Context: Integer; Resp: TStream);
var
  Size, Downloaded, Status, Len, Index: DWord;
  S: string;
  ContentType: string;
begin
  Len := SizeOf(Status);
  Index := 0;
  { Handle error }
  if HttpQueryInfo(Pointer(Context), HTTP_QUERY_STATUS_CODE or HTTP_QUERY_FLAG_NUMBER,
    @Status, Len, Index) and (Status >= 300) and (Status <> 500) then
  begin
    Index := 0;
    Size := MaxStatusTest;
    SetLength(S, Size);
    if HttpQueryInfo(Pointer(Context), HTTP_QUERY_STATUS_TEXT, @S[1], Size, Index) then
    begin
      SetLength(S, Size);
      raise Exception.CreateFmt('%s (%d) - ''%s''', [S, Status, FURL]);
    end;
  end;

  { Ask for Content-Type }
  Size := MaxContentType;
  SetLength(ContentType, MaxContentType);
  HttpQueryInfo(Pointer(Context), HTTP_QUERY_CONTENT_TYPE, @ContentType[1], Size, Index);
  SetLength(ContentType, Size);

  { Read data }
  Len := 0;
  repeat
    Check(not InternetQueryDataAvailable(Pointer(Context), Size, 0, 0));
    if Size > 0 then
    begin
      SetLength(S, Size);
      Check(not InternetReadFile(Pointer(Context), @S[1], Size, Downloaded));
      Resp.Write(S[1], Size);
    end;
  until Size = 0;

  { Check that we have a valid content type}
  { Ideally, we would always check but there are several WebServers out there
    that send files with .wsdl extension with the content type 'text/plain' or
    'text/html' ?? }
   if SameText(ContentType, SContentTypeTextPlain) or
       SameText(ContentType, STextHtml) then
   raise Exception.CreateFmt(SInvalidContentType, [ContentType]);
end;


procedure THTTPMSPoster.PostData(const Request: TStream; Response: TStream);
var
  Context: Integer;
begin
    Connect(true);
    Context := Send(Request);
    try
     Receive(Context, Response);
    finally
      if Context <> 0  then
         InternetCloseHandle(Pointer(Context));
    end;
end;


destructor THTTPMSPoster.destroy;
begin
  Connect(false);
  inherited;
end;

end.
