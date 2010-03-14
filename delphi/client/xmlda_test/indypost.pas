unit indypost;

interface
uses
 classes;

procedure PostData(Url, Proxy, Username,Password: string;SoapAction: string;
                   const Request: TStream; Response: TStream);


implementation

uses
 Sysutils,
 IdHTTP,IdURI;          // http call with indy


const
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

procedure ParseURL(AURL: string; var VHost, VPort : string);
var
  URI: TIdURI;
  Index: Integer;
begin
  URI := TIdURI.Create(AURL);
  try
   // VProtocol := URI.Protocol;
    VHost := URI.Host;
    //VPath := URI.Path;
   // VDocument := URI.Document;
    VPort := URI.Port;
   // VBookmark := URI.Bookmark;
  finally
    URI.Free;
  end;
  { if fail then check for 'localhost:####' }
  if VHost = '' then
  begin
    Index := Pos(':', AURL);
    if Index > 0 then
    begin
      VHost := Copy(AURL, 1, Index-1);
      VPort := Copy(AURL, Index+1, MaxInt);
    end;
  end;
end;

procedure PostData(Url, Proxy, Username,Password: string;SoapAction: string;
                   const Request: TStream; Response: TStream);
var
  IndyHTTP: TIDHttp;
  ContentType: string;
  Host : string;
  Port : string;
begin
  IndyHTTP := TIDHttp.Create(nil);
  try
    IndyHttp.ReadTimeOut := 9000;
    IndyHttp.Host := URL;
//    if pos('https',lowercase(URL)) = 1 then
//      IndyHttp.IOHandler := TIdSSLIOHandlerSocket.Create(Nil);
    IndyHttp.Request.ContentType := 'text/xml';
    IndyHttp.Request.CustomHeaders.Add(SHTTPSoapAction + ': ' + '"' + SoapAction + '"');
   // IndyHttp.Request.CustomHeaders.Add('expect: 100-continue');
    IndyHttp.Request.Accept := '*/*';
    IndyHttp.Request.UserAgent := 'dOPC XML DA Client (http://www.kassl.de)';
    { Proxy support configuration }
    if Proxy <> '' then
    begin
      { first check for 'http://localhost:####' }
      ParseURL(Proxy, Host, Port);
      IndyHttp.ProxyParams.ProxyServer := Host;
      if Port <> '' then
        IndyHttp.ProxyParams.ProxyPort := StrToInt(Port);
      { If name/password is used in conjunction with proxy, it's passed
        along for proxy authentication }
      IndyHttp.ProxyParams.ProxyUsername := UserName;
      IndyHttp.ProxyParams.ProxyPassword := Password;
      if (UserName <> '') and (Password <> '') then
        IndyHttp.ProxyParams.BasicAuthentication := true;

    end else
    begin
      { no proxy with Username/Password implies basic authentication }

      IndyHttp.Request.Username := UserName;
      IndyHttp.Request.Password := Password;
      if (UserName <> '') and (Password <> '') then
         IndyHttp.Request.BasicAuthentication := true;
    end;
    // IndyHttp.HTTPOptions := IndyHttp.HTTPOptions  + [hoKeepOrigProtocol];

    IndyHttp.Post(URL, Request, Response);
    ContentType := IndyHttp.Response.RawHeaders.Values[SContentType];
    if Response.Size = 0 then
      raise Exception.Create(SInvalidHTTPResponse);
    if SameText(ContentType, SContentTypeTextPlain) or
       SameText(ContentType, STextHtml) then
    raise Exception.CreateFmt(SInvalidContentType, [ContentType]);
  finally
    if Assigned(IndyHttp.IoHandler) then
      IndyHttp.IOHandler.Free;
    FreeAndNil(IndyHTTP);
  end;
end;



end.
