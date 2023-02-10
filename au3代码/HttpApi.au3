#include-once

#include <AutoItConstants.au3>
#include <StringConstants.au3>
#include <FileConstants.au3>
#include <Debug.au3>
#include <Array.au3>
#include <WinAPIMisc.au3>
#include <WinAPIFiles.au3>

;================================================================================================================================
;
; HTTPAPI -  Microsoft HTTP Server API Wrapper
;
; Purpose
; The HTTP Server API enables applications to communicate over HTTP without using Microsoft Internet Information Server (IIS).
; Applications can register to receive HTTP requests for particular URLs, receive HTTP requests, and send HTTP responses. The
; HTTP Server API includes SSL support so that applications can exchange data over secure HTTP connections without IIS.
;
; Author:
; TheXman (https://www.autoitscript.com/forum/profile/23259-thexman/)
;
; Related Links:
; Start Page:        https://docs.microsoft.com/en-us/windows/win32/http/http-api-start-page
; API v2 Reference:  https://docs.microsoft.com/en-us/windows/win32/http/http-server-api-version-2-0-reference
;
;================================================================================================================================

;================================================================================================================================
; Modification Log
;
; 2022-06-04    v1.2.1
;               - Added _HTTPAPI_HttpReceiveRequestEntityBody(), which was authored by dreamscd.  The
;                 function returns additional request body entities.  This is required when POST and PUT requests have body data
;                 but the EntityChunkCount is not used. i.e. When sending POST requests using WinHttp.
;               - The process_post_request() function, in the example script, was modified to include the processing of additional
;                 request body entities, if they exist.  Thanks dreamscd.
;               - Corrected an issue in _HTTPAPI_GetKnownRequestHeaderValue(). It was reading value as wchar instead of char.
;               - Corrected an issue in _HTTPAPI_HttpSendHttpResponse().  Response status text was not getting set properly.
;               - Corrected a function header.  It listed a parameter that was not being used.
;
; 2022-03-19    v1.1.5
;               - Added _HTTPAPI_GetRequestBody().  As it implies, the function will return the request body.
;                 Thanks @sylremo for the suggestion.
;               - Added sample POST request processing to the example script.
;               - Corrected some incomplete and/or erroneous function header info.
;
; 2022-03-17    v1.1.4
;               - Corrected _HTTPAPI_HttpSendHttpResponse function header info.
;               - Added the ability for _HTTPAPI_HttpSendHttpResponse() to recognize and process a redirect response.
;                 This is achieved by setting $iStatusCode to 301 or 302 and setting $sReason to the target location URL.
;                 Thanks @mrwilly for the suggestion.
;               - Added a redirect example to the example_http_server_get_requests.au3 example script.
;
; 2021-05-08    v1.1.3
;               - Standardized/Beautify DATA_CHUNK structure unions
;               - Corrected $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FRAGMENT_CACHE_EX.
;                 it was a duplicate of FRAGMENT_CACHE.  I had to add the BYTE_RANGE struct.
;
; 2021-01-19    v1.1.2
;               - Fix an error in the $__HTTPAPI_gtagHTTP_BYTE_RANGE structure definition.
;                 - unit64 -> uint64
;
; 2020-05-31    v1.1.1
;               - Optimized Danyfirex's fix so that it works for all union'd DATA_CHUNK structures.
;
; 2020-05-30    v1.1.0
;               - An issue that was causing responses to fail when run in 32bit mode was found and fixed. Thanks Danyfirex!
;
; 2020-05-29    v1.0.0
;               - Initial Release
;
;================================================================================================================================

; #INDEX# =======================================================================================================================
; Title .........: HTTPAPI
; AutoIt Version : 3.3.14.5
; Language ......: English
; Description ...: Microsoft's HTTP Server API Wrapper
; Author(s) .....: TheXman  (Creator/Maintainer)
;                  dreamscd (_HTTPAPI_HttpReceiveRequestEntityBody)
; Dll(s) ........: httpapi.dll
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _HTTPAPI_GetKnownRequestHeaderValue($pKnownHeaders, $iKnownHeader)
; _HTTPAPI_GetLastErrorMessage()
; _HTTPAPI_GetRequestBody(ByRef $tRequestBuffer)
; _HTTPAPI_GetRequestHeaders($pRequestHeaders)
; _HTTPAPI_HttpAddUrlToUrlGroup($iUrlGroupId, $sUrl)
; _HTTPAPI_HttpCloseServerSession($iServerSessionId)
; _HTTPAPI_HttpCloseUrlGroup($iUrlGroupId)
; _HTTPAPI_HttpCreateRequestQueue()
; _HTTPAPI_HttpCreateServerSession()
; _HTTPAPI_HttpCreateUrlGroup($iServerSessionId)
; _HTTPAPI_HttpInitialize()
; _HTTPAPI_HttpReceiveHttpRequest($hRequestQueue)
; _HTTPAPI_HttpReceiveRequestEntityBody($hRequestQueue, $RequestId)
; _HTTPAPI_HttpRemoveUrlFromUrlGroup($iUrlGroupId, $sUrl, $iFlags = 0)
; _HTTPAPI_HttpSendHttpResponse($hRequestQueue, $iRequestId, $iStatusCode, $sReason, $sBody)
; _HTTPAPI_HttpShutdownRequestQueue($hRequestQueue)
; _HTTPAPI_HttpSetUrlGroupProperty($iUrlGroupId, $iProperty, $pInfo, $iInfoLength)
; _HTTPAPI_Startup($bDebug = False)
; _HTTPAPI_HttpTerminate()
; _HTTPAPI_UDFVersion()
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __HTTPAPI_CreateStringBuffer($sString, $bUnicode = False)
; __HTTPAPI_Shutdown()
; __HTTPAPI_ApiVersion($iMajorVersion, $iMinorVersion = 0)
; ===============================================================================================================================

;========================
; ENUMERATIONS
;========================
;Property Flag Enums (bit field)
Enum   $__HTTPAPI_HTTP_PROPERTY_FLAG_PRESENT_BIT

;Global HTTP_VERB Enums
Enum $__HTTPAPI_HttpVerbUnparsed, _
     $__HTTPAPI_HttpVerbUnknown, _
     $__HTTPAPI_HttpVerbInvalid, _
     $__HTTPAPI_HttpVerbOPTIONS, _
     $__HTTPAPI_HttpVerbGET, _
     $__HTTPAPI_HttpVerbHEAD, _
     $__HTTPAPI_HttpVerbPOST, _
     $__HTTPAPI_HttpVerbPUT, _
     $__HTTPAPI_HttpVerbDELETE, _
     $__HTTPAPI_HttpVerbTRACE, _
     $__HTTPAPI_HttpVerbCONNECT, _
     $__HTTPAPI_HttpVerbTRACK, _
     $__HTTPAPI_HttpVerbMOVE, _
     $__HTTPAPI_HttpVerbCOPY, _
     $__HTTPAPI_HttpVerbPROPFIND, _
     $__HTTPAPI_HttpVerbPROPPATCH, _
     $__HTTPAPI_HttpVerbMKCOL, _
     $__HTTPAPI_HttpVerbLOCK, _
     $__HTTPAPI_HttpVerbUNLOCK, _
     $__HTTPAPI_HttpVerbSEARCH, _
     $__HTTPAPI_HttpVerbMaximum

;Global HTTP_HEADER_ID Enums
Enum $__HTTPAPI_HttpHeaderCacheControl       = 0, _
     $__HTPPAPI_HttpHeaderConnection         = 1, _
     $__HTPPAPI_HttpHeaderDate               = 2, _
     $__HTPPAPI_HttpHeaderKeepAlive          = 3, _
     $__HTPPAPI_HttpHeaderPragma             = 4, _
     $__HTPPAPI_HttpHeaderTrailer            = 5, _
     $__HTPPAPI_HttpHeaderTransferEncoding   = 6, _
     $__HTPPAPI_HttpHeaderUpgrade            = 7, _
     $__HTPPAPI_HttpHeaderVia                = 8, _
     $__HTPPAPI_HttpHeaderWarning            = 9, _
     $__HTPPAPI_HttpHeaderAllow              = 10, _
     $__HTPPAPI_HttpHeaderContentLength      = 11, _
     $__HTPPAPI_HttpHeaderContentType        = 12, _
     $__HTPPAPI_HttpHeaderContentEncoding    = 13, _
     $__HTPPAPI_HttpHeaderContentLanguage    = 14, _
     $__HTPPAPI_HttpHeaderContentLocation    = 15, _
     $__HTPPAPI_HttpHeaderContentMd5         = 16, _
     $__HTPPAPI_HttpHeaderContentRange       = 17, _
     $__HTPPAPI_HttpHeaderExpires            = 18, _
     $__HTPPAPI_HttpHeaderLastModified       = 19, _
	 $__HTPPAPI_HttpHeaderAccept             = 20, _  ;REQUEST HEADERS
     $__HTPPAPI_HttpHeaderAcceptCharset      = 21, _
     $__HTPPAPI_HttpHeaderAcceptEncoding     = 22, _
     $__HTPPAPI_HttpHeaderAcceptLanguage     = 23, _
     $__HTPPAPI_HttpHeaderAuthorization      = 24, _
     $__HTPPAPI_HttpHeaderCookie             = 25, _
     $__HTPPAPI_HttpHeaderExpect             = 26, _
     $__HTPPAPI_HttpHeaderFrom               = 27, _
     $__HTPPAPI_HttpHeaderHost               = 28, _
     $__HTPPAPI_HttpHeaderIfMatch            = 29, _
     $__HTPPAPI_HttpHeaderIfModifiedSince    = 30, _
     $__HTPPAPI_HttpHeaderIfNoneMatch        = 31, _
     $__HTPPAPI_HttpHeaderIfRange            = 32, _
     $__HTPPAPI_HttpHeaderIfUnmodifiedSince  = 33, _
     $__HTPPAPI_HttpHeaderMaxForwards        = 34, _
     $__HTPPAPI_HttpHeaderProxyAuthorization = 35, _
     $__HTPPAPI_HttpHeaderReferer            = 36, _
     $__HTPPAPI_HttpHeaderRange              = 37, _
     $__HTPPAPI_HttpHeaderTe                 = 38, _
     $__HTPPAPI_HttpHeaderTranslate          = 39, _
     $__HTPPAPI_HttpHeaderUserAgent          = 40, _
     $__HTPPAPI_HttpHeaderRequestMaximum     = 41, _
     $__HTPPAPI_HttpHeaderAcceptRanges       = 20, _  ;RESPONSE HEADERS
     $__HTPPAPI_HttpHeaderAge                = 21, _
     $__HTPPAPI_HttpHeaderEtag               = 22, _
     $__HTPPAPI_HttpHeaderLocation           = 23, _
     $__HTPPAPI_HttpHeaderProxyAuthenticate  = 24, _
     $__HTPPAPI_HttpHeaderRetryAfter         = 25, _
     $__HTPPAPI_HttpHeaderServer             = 26, _
     $__HTPPAPI_HttpHeaderSetCookie          = 27, _
     $__HTPPAPI_HttpHeaderVary               = 28, _
     $__HTPPAPI_HttpHeaderWwwAuthenticate    = 29, _
     $__HTPPAPI_HttpHeaderResponseMaximum    = 30, _
     $__HTPPAPI_HttpHeaderMaximum            = 41

;Global Server Property Enums
Enum $__HTPPAPI_HttpServerAuthenticationProperty, _
     $__HTPPAPI_HttpServerLoggingProperty, _
     $__HTPPAPI_HttpServerQosProperty, _
     $__HTPPAPI_HttpServerTimeoutsProperty, _
     $__HTPPAPI_HttpServerQueueLengthProperty, _
     $__HTPPAPI_HttpServerStateProperty, _
     $__HTPPAPI_HttpServer503VerbosityProperty, _
     $__HTPPAPI_HttpServerBindingProperty, _
     $__HTPPAPI_HttpServerExtendedAuthenticationProperty, _
     $__HTPPAPI_HttpServerListenEndpointProperty, _
     $__HTPPAPI_HttpServerChannelBindProperty, _
     $__HTPPAPI_HttpServerProtectionLevelProperty

;Global Data Chunk Type Enums
Enum $__HTPPAPI_HttpDataChunkFromMemory, _
     $__HTPPAPI_HttpDataChunkFromFileHandle, _
     $__HTPPAPI_HttpDataChunkFromFragmentCache, _
     $__HTPPAPI_HttpDataChunkFromFragmentCacheEx, _
     $__HTPPAPI_HttpDataChunkMaximum

;Global Cache Policy Enums
Enum $__HTTPAPI_HttpCachePolicyNocache,         _
     $__HTTPAPI_HttpCachePolicyUserInvalidates, _
	 $__HTTPAPI_HttpCachePolicyTimeToLive

;Global Log Data Enum
Enum $__HTTPAPI_HttpLogDataTypeFields

;========================
; STRUCTURES
;========================
;Global HHTP Version Structure
Global $__HTTPAPI_gtagHTTP_VERSION = _
           "struct;"               & _
		   "ushort MajorVersion;"  & _
		   "ushort MinorVersion;"  & _
		   "endstruct;"

;Global HHTPAPI Version Structure and Constants
Global $__HTTPAPI_gtagHTTPAPI_VERSION    = _
           "struct;"                     & _
		   "ushort HttpApiMajorVersion;" & _
		   "ushort HttpApiMinorVersion;" & _
		   "endstruct;"

;Global Cooked URL Structure
Global $__HTTPAPI_gtagHTTP_COOKED_URL  = _
           "struct;"                   & _
		   "ushort FullUrlLength;"     & _
		   "ushort HostLength;"        & _
		   "ushort AbsPathLength;"     & _
		   "ushort QueryStringLength;" & _
		   "ptr    pFullUrl;"          & _ ;wstr
		   "ptr    pHost;"             & _ ;wstr
		   "ptr    pAbsPath;"          & _ ;wstr
		   "ptr    pQueryString;"      & _ ;wstr
		   "endstruct;"

;Global Transport Address Structure
Global $__HTTPAPI_gtagHTTP_TRANSPORT_ADDRESS = _
           "struct;"                         & _
		   "ptr  pRemoteAddress;"            & _
		   "ptr  pLocalAddress;"             & _
		   "endstruct;"

;Global Unkown Header Structure & Constant
Global $__HTTPAPI_gtagHTTP_UNKNOWN_HEADER = _
           "struct;"                    & _
		   "ushort  NameLength;"        & _
		   "ushort  RawValueLength;"    & _
		   "ptr     pName;"             & _ ;str
		   "ptr     pRawValue;"         & _ ;str
		   "endstruct;"
Const  $__HTTPAPI_HTTP_UNKNOWN_HEADER_SIZE = DllStructGetSize(DllStructCreate($__HTTPAPI_gtagHTTP_UNKNOWN_HEADER))

;Global Known Header Structure & Constant
Global $__HTTPAPI_gtagHTTP_KNOWN_HEADER = _
           "struct;"                    & _
		   "ushort  RawValueLength;"    & _
		   "ptr     pRawValue;"         & _ ;str
		   "endstruct;"
Const  $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE = DllStructGetSize(DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER))

;Global HTTP Request Headers Structure
Global $__HTTPAPI_gtagHTTP_REQUEST_HEADERS = _
           "struct;"                       & _
		   "ushort  UnknownHeaderCount;"   & _
		   "ptr     pUnknownHeaders;"      & _
           "ushort  TrailerCount;"         & _ ;Reserved, must be 0
           "ptr     pTrailers;"            & _ ;Reserved, must be NULL
		   StringFormat("byte KnownHeaders[%i];", $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE * $__HTPPAPI_HttpHeaderRequestMaximum) & _
		   "endstruct;"

;Global HTTP Response Headers Structure
Global $__HTTPAPI_gtagHTTP_RESPONSE_HEADERS = _
           "struct;"                       & _
		   "ushort  UnknownHeaderCount;"   & _
		   "ptr     pUnknownHeaders;"      & _
           "ushort  TrailerCount;"         & _ ;Reserved, must be 0
           "ptr     pTrailers;"            & _ ;Reserved, must be NULL
		   StringFormat("byte KnownHeaders[%i];", $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE * $__HTPPAPI_HttpHeaderResponseMaximum) & _
		   "endstruct;"

;Global Data Chunk Structures/Unions
;  This is a union in the api.  So special alignment is done to make it work in AutoIt.
;  Because the largest data type that folows the data chunk struct is a uint64, in 32-bit,
;  there needs to an additional 4 bytes of padding between the initial struct and the
;  unioned struct.
Global $__HTTPAPI_gtagHTTP_DATA_CHUNK = _
           "struct;"                  & _
		   "int  DataChunkType;"      & _
		   "endstruct;"

Global $__HTTPAPI_gtagHTTP_BYTE_RANGE = _
		   "struct;"                  & _
		   "uint64  StartingOffset;"  & _
		   "uint64  Length;"          & _
		   "endstruct;"
Const  $__HTTPAPI_gtagHTTP_BYTE_RANGE_TO_EOF = -1

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_MEMORY = _
           "struct;"                  & _
		   "ptr    pBuffer;"          & _
		   "ulong  pBufferLength;"    & _
		   "endstruct;"

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FILE_HANDLE = _
		   "struct;"                      & _
		   $__HTTPAPI_gtagHTTP_BYTE_RANGE & _
		   "handle   FileHandle;"         & _
		   "endstruct;"

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FRAGMENT_CACHE = _
		   "struct;"                      & _
		   "ushort  FragmentNameLength;"  & _ ;in bytes not including the NULL
		   "ptr     pFragmentName;"       & _ ;wstr
		   "endstruct;"

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FRAGMENT_CACHE_EX = _
		   "struct;"                      & _
		   $__HTTPAPI_gtagHTTP_BYTE_RANGE & _
		   "ptr     pFragmentName;"       & _ ;wstr
		   "endstruct;"

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_MEMORY = _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK & _
		   (@AutoItX64 ? "" : "ptr;")     & _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_MEMORY

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_FILE_HANDLE = _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK & _
		   (@AutoItX64 ? "" : "ptr;")     & _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FILE_HANDLE

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_FRAGMENT_CACHE = _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK & _
		   (@AutoItX64 ? "" : "ptr;")     & _
		   $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FRAGMENT_CACHE

Global $__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_FRAGMENT_CACHE_EX = _
           $__HTTPAPI_gtagHTTP_DATA_CHUNK & _
		   (@AutoItX64 ? "" : "ptr;")     & _
		   $__HTTPAPI_gtagHTTP_DATA_CHUNK_UNION_FROM_FRAGMENT_CACHE_EX

;Global SSL Info Structure
Global $__HTTPAPI_gtagHTTP_SSL_INFO           = _
		   "struct;"                          & _
		   "ushort  ServerCertKeySize;"       & _
		   "ushort  ConnectionKeySize;"       & _
		   "ulong   ServerCertIssuerSize;"    & _
		   "ulong   ServerCertSubjectSize;"   & _
		   "ptr     pServerCertIssuer;"       & _ ;str
		   "ptr     pServerCertSubject;"      & _ ;str
		   "struct* pClientCertInfo;"         & _
		   "ulong   SslClientCertNegotiated;" & _
		   "endstruct;"

;Global SSL Certificate Info Structure
Global $__HTTPAPI_gtagHTTP_SSL_CLIENT_CERT_INFO = _
		   "struct;"                            & _
		   "ulong   CertFlags;"                 & _
		   "ulong   CertEncodedSize;"           & _
		   "byte*   pCertEncoded;"              & _
		   "handle  Token;"                     & _
		   "boolean CertDeniedByMapper;"        & _
		   "endstruct;"

;Global HTTP Property Flags Structure
Global $__HTTPAPI_gtagHTTP_PROPERTY_FLAGS = _
		   "struct;"                      & _
		   "ulong Flags;"                 & _
		   "endstruct;"

;Global Binding Info Structure
Global $__HTTPAPI_gtagHTTP_BINDING_INFO       = _
		   "struct;"                          & _
		   $__HTTPAPI_gtagHTTP_PROPERTY_FLAGS & _
		   "handle  RequestQueueHandle;"      & _
		   "endstruct;"

;Global Cache Policy Structure
Global $__HTTPAPI_gtagHTTP_CACHE_POLICY = _
		   "struct;"                    & _
		   "int    Policy;"             & _
		   "ulong  SecondsToLive;"      & _
		   "endstruct;"

;Global LOG_DATA Structures
Global $__HTTPAPI_gtagHTTP_LOG_DATA = "struct; int LogDataType; endstruct;"

Global $__HTTPAPI_gtagHTTP_LOG_FIELDS_DATA = _
       $__HTTPAPI_gtagHTTP_LOG_DATA        & _
       "ushort  UserNameLength;"           & _
       "ushort  UriStemLength;"            & _
       "ushort  ClientIpLength;"           & _
       "ushort  ServerNameLength;"         & _
       "ushort  ServiceNameLength;"        & _
       "ushort  ServerIpLength;"           & _
       "ushort  MethodLength;"             & _
       "ushort  UriQueryLength;"           & _
       "ushort  HostLength;"               & _
       "ushort  UserAgentLength;"          & _
       "ushort  CookieLength;"             & _
       "ushort  ReferrerLength;"           & _
       "ptr     UserName;"                 & _
       "ptr     UriStem;"                  & _
       "ptr     ClientIp;"                 & _
       "ptr     ServerName;"               & _
       "ptr     ServiceName;"              & _
       "ptr     ServerIp;"                 & _
       "ptr     Method;"                   & _
       "ptr     UriQuery;"                 & _
       "ptr     Host;"                     & _
       "ptr     UserAgent;"                & _
       "ptr     Cookie;"                   & _
       "ptr     Referrer;"                 & _
       "ushort  ServerPort;"               & _
       "ushort  ProtocolStatus;"           & _
       "ulong   Win32Status;"              & _
       "int     MethodNum;"                & _
       "ushort  SubStatus;"                & _
	   "endstruct;"

;Global REQUEST Structures & Constants
Global $__HTTPAPI_gtagHTTP_REQUEST_V1    = _
           "struct;"                     & _
           "ulong    Flags;"             & _
           "uint64   ConnectionId;"      & _
           "uint64   RequestId;"         & _
           "uint64   UrlContext;"        & _
           $__HTTPAPI_gtagHTTP_VERSION   & _
           "int      Verb;"              & _
           "ushort   UnknownVerbLength;" & _
           "ushort   RawUrlLength;"      & _
           "ptr      pUnknownVerb;"      & _ ;str
           "ptr      pRawUrl;"           & _ ;str
           $__HTTPAPI_gtagHTTP_COOKED_URL        & _
           $__HTTPAPI_gtagHTTP_TRANSPORT_ADDRESS & _
           $__HTTPAPI_gtagHTTP_REQUEST_HEADERS   & _
           "uint64   BytesReceived;"     & _
           "ushort   EntityChunkCount;"  & _
           "ptr      pEntityChunks;"     & _
           "uint64   RawConnectionId;"   & _
           "ptr      pSslInfo;"          & _
           "endstruct;"

Global $__HTTPAPI_gtagHTTP_REQUEST_V2     = _
           $__HTTPAPI_gtagHTTP_REQUEST_V1 & _
           "struct;"                      & _
           "ushort  RequestInfoCount;"    & _
           "ptr     pRequestInfo;"        & _
           "endstruct;"
Const $__HTTPAPI_REQUEST_V2_SIZE     = DllStructGetSize(DllStructCreate($__HTTPAPI_gtagHTTP_REQUEST_V2))

;Global RESPONSE Structures & Constants
Global $__HTTPAPI_gtagHTTP_RESPONSE_V1   = _
           "struct;"                     & _
           "ulong    Flags;"             & _
           $__HTTPAPI_gtagHTTP_VERSION   & _
           "ushort   StatusCode;"        & _
           "ushort   ReasonLength;"      & _
           "ptr      pReason;"           & _
           $__HTTPAPI_gtagHTTP_RESPONSE_HEADERS  & _
           "ushort   EntityChunkCount;"  & _
           "ptr      pEntityChunks;"     & _
           "endstruct;"

Global $__HTTPAPI_gtagHTTP_RESPONSE_V2     = _
           "struct;"                       & _
           $__HTTPAPI_gtagHTTP_RESPONSE_V1 & _
           "ushort  ResponseInfoCount;"    & _
           "ptr     pResponseInfo;"        & _
           "endstruct;"
Const $__HTTPAPI_RESPONSE_V2_SIZE     = DllStructGetSize(DllStructCreate($__HTTPAPI_gtagHTTP_RESPONSE_V2))

;========================
; CONSTANTS
;========================
;HTTPAPI UDF Constant
Const $__HTTPAPI_UDF_VERSION = "1.2.1"

;Constant used by HttpRemoveUrlFromUrlGroup
Const $__HTTPAPI_HTTP_URL_FLAG_REMOVE_ALL = 0x00000001

;HTTPAPI Error Code Constants
Const $__HTTPAPI_ERROR_NO_ERROR          = 0, _
	  $__HTTPAPI_ERROR_FILE_NOT_FOUND    = 2, _
	  $__HTTPAPI_ERROR_ACCESS_DENIED     = 5, _
	  $__HTTPAPI_ERROR_INVALID_HANDLE    = 6, _
      $__HTTPAPI_ERROR_HANDLE_EOF        = 38, _
      $__HTTPAPI_ERROR_INVALID_PARAMETER = 87, _
      $__HTTPAPI_ERROR_ALREADY_EXISTS    = 183, _
      $__HTTPAPI_ERROR_MORE_DATA         = 234, _
      $__HTTPAPI_ERROR_NO_ACCESS         = 998, _
      $__HTTPAPI_ERROR_DLL_INIT_FAILED   = 1114, _
      $__HTTPAPI_ERROR_REVISION_MISMATCH = 1306

;Constants used by HttpInitialize() and HttpTerminate()
Const $__HTTPAPI_HTTP_INITIALIZE_SERVER = 0x00000001, _
      $__HTTPAPI_HTTP_INITIALIZE_CONFIG = 0x00000002

;Request flag constant for POST & PUT
Const $__HTTPAPI_HTTP_REQUEST_FLAG_MORE_ENTITY_BODY_EXISTS = 1

;Property Flag Constants (bit field)
Const $__HTTPAPI_HTTP_PROPERTY_FLAG_PRESENT = 0x1

;Global Request Header Names Array
Const $__HTTPAPI_gaRequestHeaderNames =  [ _
                                         "Cache-Control",       _
                                         "Connection",          _
                                         "Date",                _
                                         "Keep-Alive",          _
                                         "Pragma",              _
                                         "Trailer",             _
                                         "Transfer-Encoding",   _
                                         "Upgrade",             _
                                         "Via",                 _
                                         "Warning",             _
                                         "Allow",               _
                                         "Content-Length",      _
                                         "Content-Type",        _
                                         "Content-Encoding",    _
                                         "Content-Language",    _
                                         "Content-Location",    _
                                         "Content-MD5",         _
                                         "Content-Range",       _
                                         "Expires",             _
                                         "Last-Modified",       _
                                         "Accept",              _  ;REQUEST HEADERS
                                         "Accept-Charset",      _
                                         "Accept-Encoding",     _
                                         "Accept-Language",     _
                                         "Authorization",       _
                                         "Cookie",              _
                                         "Expect",              _
                                         "From",                _
                                         "Host",                _
                                         "If-Match",            _
                                         "If-Modified-Since",   _
                                         "If-None-Match",       _
                                         "If-Range",            _
                                         "If-Unmodified-Since", _
                                         "Max-Forwards",        _
                                         "Proxy-Authorization", _
                                         "Referer",             _
                                         "Range",               _
                                         "TE",                  _
                                         "Translate",           _
                                         "User-Agent"           _
                                         ]

;Global Response Header Names Array
Const $__HTTPAPI_gaResponseHeaderNames = [ _
                                         "Cache-Control",       _
                                         "Connection",          _
                                         "Date",                _
                                         "Keep-Alive",          _
                                         "Pragma",              _
                                         "Trailer",             _
                                         "Transfer-Encoding",   _
                                         "Upgrade",             _
                                         "Via",                 _
                                         "Warning",             _
                                         "Allow",               _
                                         "Content-Length",      _
                                         "Content-Type",        _
                                         "Content-Encoding",    _
                                         "Content-Language",    _
                                         "Content-Location",    _
                                         "Content-MD5",         _
                                         "Content-Range",       _
                                         "Expires",             _
                                         "Last-Modified",       _
                                         "Accept-Ranges",       _  ;RESPONSE HEADERS
                                         "Age",                 _
                                         "Etag",                _
                                         "Location",            _
                                         "Proxy-Authenticate",  _
                                         "Retry-After",         _
                                         "Server",              _
                                         "Set-Cookie",          _
                                         "Vary",                _
                                         "WWW-Authenticate"     _
                                         ]

;HTTP Request Flag Constants
Const $__HTTPAPI_HTTP_RECEIVE_REQUEST_FLAG_COPY_BODY  = 0x00000001, _
      $__HTTPAPI_HTTP_RECEIVE_REQUEST_FLAG_FLUSH_BODY = 0x00000002

;HTTP Response Flag Constants
Const $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_DISCONNECT     = 0x00000001, _
      $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_MORE_DATA      = 0x00000002, _
      $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_BUFFER_DATA    = 0x00000004, _
      $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_ENABLE_NAGLING = 0x00000008, _
      $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_PROCESS_RANGES = 0x00000020, _
      $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_OPAQUE         = 0x00000040

;HTTP API Version Constants (these emulate the macros in httpapi.dll)
Const  $__HTTPAPI_HTTPAPI_VERSION_1 = __HTTPAPI_ApiVersion(1), _
       $__HTTPAPI_HTTPAPI_VERSION_2 = __HTTPAPI_ApiVersion(2)

;HTTP Request Constants
Const $__HTTPAPI_REQUEST_BODY_SIZE = 4096

;======================
;Global Strings
;======================
Global $__HTTPAPI_gsLastErrorMessage = ""

;======================
;Global Flags
;======================
Global $__HTTPAPI_gbDebugging = False

;======================
;Global Handles
;======================
Global $__HTTPAPI_ghHttpApiDll = -1


; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_GetKnownRequestHeaderValue
; Description ...: Gets the specified known header value.
; Syntax ........: _HTTPAPI_GetKnownRequestHeaderValue($pKnownHeaders, $iKnownHeader)
; Parameters ....: $pKnownHeaders       A pointer to a Known Header structure.
;                  $iKnownHeader        An integer value identifying the Known Header to set. See global enums for named values.
; Return values .: Success:             A string containing vallue of known header.
;                  Failure:             "" and sets @error flag to non-zero.
;                  @error:              1 - Invalid $iKnownHeader
; Author ........: TheXman
; ===============================================================================================================================
Func _HTTPAPI_GetKnownRequestHeaderValue($pKnownHeaders, $iKnownHeader)

	Local $pHeader = Null
	Local $tHeader = ""

	;Validate parameters
	If $iKnownHeader < 0 Or $iKnownHeader >= $__HTPPAPI_HttpHeaderResponseMaximum Then Return SetError(1, 0, "")

	;Create a known header at correct memory location
	$pHeader = $pKnownHeaders + ($iKnownHeader * $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE)
	$tHeader = DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER, $pHeader)

	Return _WinAPI_GetString($tHeader.pRawValue, False)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_GetLastErrorMessage
; Description ...: Returns the last error message
; Syntax ........: _HTTPAPI_GetLastErrorMessage()
; Parameters ....: None
; Return values .: A string containing the last error message.
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; ===============================================================================================================================
Func _HTTPAPI_GetLastErrorMessage()

	_DebugOut(@CRLF & "Function: _HTTPAPI_GetLastErrorMessage")

	Return $__HTTPAPI_gsLastErrorMessage

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_GetRequestBody
; Description ...: Returns the request body.
; Syntax ........: _HTTPAPI_GetRequestBody($tRequestBuffer)
; Parameters ....: $tRequestBuffer      A structure containing the request.
; Return values .: Success:             A string containing the request body.
; Author ........: TheXman
; Related .......:
; Remarks .......:
; ===============================================================================================================================
Func _HTTPAPI_GetRequestBody(ByRef $tRequestBuffer)

	_DebugOut(@CRLF & "Function: _HTTPAPI_GetRequestBody")

	Local $tRequest = "", $tDataChunk = ""
	Local $sRequestBody = ""

	;Create struct that points to the request and body
	$tRequest = DllStructCreate( _
	                $__HTTPAPI_gtagHTTP_REQUEST_V2 & StringFormat("byte body[%i];", $__HTTPAPI_REQUEST_BODY_SIZE), _
					DllStructGetPtr($tRequestBuffer))

	;Parse request body from the buffer, if one exists
	If $tRequest.pEntityChunks > 0 Then
		$tDataChunk = DllStructCreate($__HTTPAPI_gtagHTTP_DATA_CHUNK, $tRequest.pEntityChunks)
		If $tDataChunk.DataChunkType = $__HTPPAPI_HttpDataChunkFromMemory Then
			$tDataChunk   = DllStructCreate($__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_MEMORY, $tRequest.pEntityChunks)
			$sRequestBody = _WinAPI_GetString($tDataChunk.pBuffer, False)
		EndIf
	EndIf

	Return $sRequestBody
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_GetRequestHeaders
; Description ...: Gets a list of the request headers.
; Syntax ........: _HTTPAPI_GetRequestHeaders($pRequestHeader)
; Parameters ....: $pRequestHeader      A pointer to the HTTP_REQUEST_HEADERS structure.
; Return values .: Success:             A 2D array containing the request headers.
; Author ........: TheXman
; Remarks .......: [n][0] = Header Name
;                  [n][1] = Header Value
; Related .......:
; Link ..........:
; ===============================================================================================================================
Func _HTTPAPI_GetRequestHeaders($pRequestHeaders)

	_DebugOut(@CRLF & "Function: _HTTPAPI_GetRequestHeaders")

	Local $aRequestHeaders[0][2]
	Local $tHeaderStruct = "", $tHeader = ""

	;Create http request headers struct
	$tHeaderStruct = DllStructCreate($__HTTPAPI_gtagHTTP_REQUEST_HEADERS, $pRequestHeaders)

	;Display headers
	For $i = 0 To $__HTPPAPI_HttpHeaderRequestMaximum - 1
		$tHeader = DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER, _
		                           DllStructGetPtr($tHeaderStruct, "KnownHeaders") + ($i * $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE))
		If $tHeader.RawValueLength > 0 Then
			_ArrayAdd($aRequestHeaders, $__HTTPAPI_gaRequestHeaderNames[$i] & "|" & _WinAPI_GetString($tHeader.pRawValue, False))
		EndIf
	Next

	For $i = 0 To $tHeaderStruct.UnknownHeaderCount - 1
		$tHeader = DllStructCreate($__HTTPAPI_gtagHTTP_UNKNOWN_HEADER, _
		                           DllStructGetData($tHeaderStruct, "pUnknownHeaders") + ($i * $__HTTPAPI_HTTP_UNKNOWN_HEADER_SIZE))
		_ArrayAdd($aRequestHeaders, _WinAPI_GetString($tHeader.pName, False) & "|" & _WinAPI_GetString($tHeader.pRawValue, False))
	Next

	_ArraySort($aRequestHeaders)

	Return $aRequestHeaders
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: HttpAddUrlToUrlGroup
; Description ...: Adds the specified URL to the URL Group identified by the URL Group ID.
; Syntax ........: _HTTPAPI_HttpAddUrlToUrlGroup($iUrlGroupId, $sUrl)
; Parameters ....: $iUrlGroupId         A uint64 integer containing the URL Group ID
;                  $sUrl                A string containing a properly formed UrlPrefix that identifies the URL to be registered.
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpaddurltourlgroup
; ===============================================================================================================================
Func _HTTPAPI_HttpAddUrlToUrlGroup($iUrlGroupId, $sUrl)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpAddUrlToUrlGroup")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpAddUrlToUrlGroup", _
					   "uint64", $iUrlGroupId, _
					   "wstr",   $sUrl, _
					   "uint64", 0, _
					   "ulong",  0)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from _HTTPAPI_HttpAddUrlToUrlGroup", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed To the function."
			Case $__HTTPAPI_ERROR_ACCESS_DENIED
				$__HTTPAPI_gsLastErrorMessage = "The calling process does not have permission to register the URL."
			Case $__HTTPAPI_ERROR_ALREADY_EXISTS
				$__HTTPAPI_gsLastErrorMessage = "The specified URL conflicts with an existing registration."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpCloseServerSession
; Description ...: Closes a server session.
; Syntax ........: _HTTPAPI_HttpCloseServerSession($iServerSessionId)
; Parameters ....: $iServerSessionId    A uint64 integer containing the Server Session ID
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpcloseserversession
; ===============================================================================================================================
Func _HTTPAPI_HttpCloseServerSession($iServerSessionId)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpCloseServerSession")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpCloseServerSession", _
					   "uint64", $iServerSessionId)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from _HTTPAPI_HttpCloseServerSession", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = _
					"The Server Session does not exist." & @CRLF & _
					"The application does not have permission to close the server session. " & _
					"Only the application that created the server session can close the session."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpCloseUrlGroup
; Description ...: Closes a URL group.
; Syntax ........: _HTTPAPI_HttpCloseServerSession($iUrlGroupId)
; Parameters ....: $iUrlGroupId         A uint64 integer containing the URL Group ID
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpcloseserversession
; ===============================================================================================================================
Func _HTTPAPI_HttpCloseUrlGroup($iUrlGroupId)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpCloseUrlGroup")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpCloseUrlGroup", _
					   "uint64", $iUrlGroupId)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from _HTTPAPI_HttpCloseUrlGroup", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the functiosn."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpCreateRequestQueue
; Description ...: Creates a new request queue.
; Syntax ........: _HTTPAPI_HttpCreateRequestQueue()
; Parameters ....: None
; Return values .: Success:             A handle to the new request queue.
;                  Failure:             0 and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpcreaterequestqueue
; ===============================================================================================================================
Func _HTTPAPI_HttpCreateRequestQueue()

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpCreateRequestQueue")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpCreateRequestQueue", _
					   "struct",  $__HTTPAPI_HTTPAPI_VERSION_2, _
					   "wstr",    Null, _
					   "ptr",     Null, _
					   "ulong",   0, _
					   "handle*", Null )
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, 0)
	EndIf

	_DebugReportVar("Result from HttpCreateRequestQueue", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the function."
			Case $__HTTPAPI_ERROR_REVISION_MISMATCH
				$__HTTPAPI_gsLastErrorMessage = "The version parameter is invalid or unsupported."
			Case $__HTTPAPI_ERROR_ALREADY_EXISTS
				$__HTTPAPI_gsLastErrorMessage = "The queue name already exists."
			Case $__HTTPAPI_ERROR_ACCESS_DENIED
				$__HTTPAPI_gsLastErrorMessage = "The calling process does not have permission to open the request queue."
			Case $__HTTPAPI_ERROR_DLL_INIT_FAILED
				$__HTTPAPI_gsLastErrorMessage = "The application has not called HttpInitialize prior to calling HttpCreateRequestQueue."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], 0)
	EndIf

	_DebugOut("Request Queue Handle = 0x" & Hex($aResult[5]))

	;All is good
	Return $aResult[5]

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpCreateServerSession
; Description ...: Creates a server session.
; Syntax ........: _HTTPAPI_HttpCreateServerSession()
; Parameters ....: None
; Return values .: Success:             A uint64 containing the server session ID.
;                  Failure:             0 and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpcreateserversession
; ===============================================================================================================================
Func _HTTPAPI_HttpCreateServerSession()

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpCreateServerSession")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpCreateServerSession", _
					   "struct",   $__HTTPAPI_HTTPAPI_VERSION_2, _
					   "uint64*", Null, _
					   "ulong",    0)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, 0)
	EndIf

	_DebugReportVar("Result from HttpCreateServerSession", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "The Flags parameter contains an unsupported value."
			Case $__HTTPAPI_ERROR_REVISION_MISMATCH
				$__HTTPAPI_gsLastErrorMessage = "The version passed is invalid or unsupported."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], 0)
	EndIf

	_DebugOut("Server Session Id = 0x" & Hex($aResult[2]) & " (" & $aResult[2] & ")")

	;All is good
	Return $aResult[2]

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpCreateUrlGroup
; Description ...: Creates a URL group.
; Syntax ........: _HTTPAPI_HttpCreateUrlGroup($iServerSessionId)
; Parameters ....: $iServerSessionId    A uint64 integer containing the Server Session ID
; Return values .: Success:             A uint64 containing the URL Group ID
;                  Failure:             0 and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpcreateurlgroup
; ===============================================================================================================================
Func _HTTPAPI_HttpCreateUrlGroup($iServerSessionId)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpCreateUrlGroup")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpCreateUrlGroup", _
					   "uint64",  $iServerSessionId, _
					   "uint64*", Null, _
					   "ulong",   0)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from HttpCreateUrlGroup", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the function."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	_DebugOut("Url Group Id = 0x" & Hex($aResult[2]) & " (" & $aResult[2] & ")")

	;All is good
	Return $aResult[2]

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpInitialize
; Description ...: Initialize HTTPAPI
; Syntax ........: _HTTPAPI_HttpInitialize()
; Parameters ....: None
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......: A call to this function is required before call any other HTTPAPI functions.
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpinitialize
; ===============================================================================================================================
Func _HTTPAPI_HttpInitialize()

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpInitialize")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpInitialize", _
					   "struct", $__HTTPAPI_HTTPAPI_VERSION_2, _
					   "ulong",  $__HTTPAPI_HTTP_INITIALIZE_SERVER, _
					   "ptr",    Null)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from HttpInitialize", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "The Flags parameter contains an unsupported value."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpReceiveHttpRequest
; Description ...: Receive a HTTP request.
; Syntax ........: _HTTPAPI_HttpReceiveHttpRequest($hRequestQueue)
; Parameters ....: $hRequestQueue       A handle to the request queue to receive the request from.
; Return values .: Success:             A request buffer that includes the request with the body of the request appended.
;                  Failure:             "" and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpreceivehttprequest
; ===============================================================================================================================
Func _HTTPAPI_HttpReceiveHttpRequest($hRequestQueue)

	_DebugOut(@CRLF & "Function: HttpReceiveHttpRequest")

	Local $tRequest       = DllStructCreate(StringFormat("byte request[%i]; byte body[%i];", _
	                                        $__HTTPAPI_REQUEST_V2_SIZE, $__HTTPAPI_REQUEST_BODY_SIZE))

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpReceiveHttpRequest", _
					   "handle",  $hRequestQueue, _
					   "uint64",  0, _
					   "ulong",   $__HTTPAPI_HTTP_RECEIVE_REQUEST_FLAG_COPY_BODY, _
					   "struct*", $tRequest, _
					   "ulong",   DllStructGetSize($tRequest), _
					   "ulong*",  Null, _
					   "ptr",     Null)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, "")
	EndIf

	_DebugReportVar("Result from HttpReceiveHttpRequest", $aResult)
	_DebugReportVar("Bytes Returned from HttpReceiveHttpRequest", $aResult[6])

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed To the function."
			Case $__HTTPAPI_ERROR_MORE_DATA
				$__HTTPAPI_gsLastErrorMessage = "More data than the buffer could hold still exists."
			Case $__HTTPAPI_ERROR_HANDLE_EOF
				$__HTTPAPI_gsLastErrorMessage = "The specified request has been fully retrieved."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], "")
	EndIf

	;All is good
	Return $tRequest

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpReceiveRequestEntityBody
; Description ...: Returns the request body.
; Syntax ........: _HTTPAPI_HttpReceiveRequestEntityBody($hRequestQueue, $RequestId)
; Parameters ....: $hRequestQueue       A handle to the request queue to receive the request from.
;                : $RequestId           This value is returned in the RequestId member of the HTTP_REQUEST structure by a call to
;                                       the HttpReceiveHttpRequest function.
; Return values .: Success:             A string containing the request body.
;                  Failure:             "" and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: dreamscd
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpreceiverequestentitybody
; ===============================================================================================================================

Func _HTTPAPI_HttpReceiveRequestEntityBody($hRequestQueue, $RequestId)

    _DebugOut(@CRLF & "Function: _HTTPAPI_HttpReceiveRequestEntityBody")

    Const $ENTITY_BUFFER_LENGTH = 2048

	Local $aResult        = ""
	Local $sRequestBody   = ""
	Local $iBytesReturned = 0
    Local $EntityBuffer   = DllStructCreate( StringFormat("byte EntityBuffer[%i];", $ENTITY_BUFFER_LENGTH) )
    Local $pEntityBuffer  = DllStructGetPtr($EntityBuffer)

    While 1
        ;Call API function
        $aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpReceiveRequestEntityBody", _
                           "handle",  $hRequestQueue, _
                           "uint64",  $RequestId, _
                           "ulong",   0, _
                           "struct*", $pEntityBuffer, _
                           "ulong",   DllStructGetSize($EntityBuffer), _
                           "ulong*",  0, _
                           "ptr",     Null)
        If @error Then
            $__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
            Return SetError(1, @error, "")
        EndIf

		$iBytesReturned = $aResult[6]

        ;Set last error message
        Switch $aResult[0]
            Case $__HTTPAPI_ERROR_NO_ERROR
                If $iBytesReturned > 0 Then $sRequestBody &= _WinAPI_GetString($pEntityBuffer, False)
			Case $__HTTPAPI_ERROR_HANDLE_EOF
                If $iBytesReturned > 0 Then $sRequestBody &= _WinAPI_GetString($pEntityBuffer, False)
                ExitLoop
            Case $__HTTPAPI_ERROR_INVALID_PARAMETER
                $__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed To the function."
                Return SetError(2, $aResult[0], "")
            Case $__HTTPAPI_ERROR_DLL_INIT_FAILED
                $__HTTPAPI_gsLastErrorMessage = "The calling application did not call HttpInitialize before calling this function."
                Return SetError(2, $aResult[0], "")
            Case Else
                $__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
                Return SetError(2, $aResult[0], "")
        EndSwitch

    WEnd

    Return $sRequestBody
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpRemoveUrlFromUrlGroup
; Description ...: Removes one, or all, of the URLs from the group.
; Syntax ........: _HTTPAPI_HttpRemoveUrlFromUrlGroup($iUrlGroupId, $sUrl, $iFlags = 0)
; Parameters ....: $iUrlGroupId         A uint64 integer containing the URL Group ID
;                  $sUrl                A string containing a properly formed UrlPrefix that identifies the URL to be removed.
;                  $iFlags              [optional] An integer qualifying the URL that is removed. Default is 0. See reamrks.
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Remarks .......: $iFlags can be one of the following values:
;                     $__HTTPAPI_HTTP_URL_FLAG_REMOVE_ALL    Removes all of the URLs currently registered with the URL Group. If
;                                                            this flag is used, then $sUrl must be null.
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpremoveurlfromurlgroup
; ===============================================================================================================================
Func _HTTPAPI_HttpRemoveUrlFromUrlGroup($iUrlGroupId, $sUrl, $iFlags = 0)

	_DebugOut(@CRLF & "Function: HttpRemoveUrlFromUrlGroup")

	Local $aResult[0]

	;If REMOVE_ALL flag was specified, then make sure $sUrl is null
	If BitAND($iFlags, $__HTTPAPI_HTTP_URL_FLAG_REMOVE_ALL) Then $sUrl = Null

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpRemoveUrlFromUrlGroup", _
					   "uint64", $iUrlGroupId, _
					   "wstr",   $sUrl, _
					   "uint",   $iFlags)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from HttpRemoveUrlFromUrlGroup", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed To the function."
			Case $__HTTPAPI_ERROR_ACCESS_DENIED
				$__HTTPAPI_gsLastErrorMessage = "The calling process does not have permission to register the URL."
			Case $__HTTPAPI_ERROR_FILE_NOT_FOUND
				$__HTTPAPI_gsLastErrorMessage = "The specified URL is not registered with the URL Group. "
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpSendHttpResponse
; Description ...: Send a HTTP reresponse.
; Syntax ........: _HTTPAPI_HttpSendHttpResponse($hRequestQueue, $iRequestId, $iStatusCode, $sReason, $sBody)
; Parameters ....: $hRequestQueue       A handle to the request queue from which the specified request was retrieved.
;                  $iRequestId          An identifier of the HTTP request to which this response corresponds.
;                  $iStatusCode         An integer representing the HTTP response status code
;                  $sReason             A string representing the HTTP response status code description or location URL if
;                                       $iStatusCode is 301 or 302 (redirect).
;                  $sBody               A string containing the response body.
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Remarks .......: If $iStatusCode is 301 or 302 (redirect), then it uses the value in $sReason as the redirect location URL.
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpsendhttpresponse
; ===============================================================================================================================
Func _HTTPAPI_HttpSendHttpResponse($hRequestQueue, $iRequestId, $iStatusCode, $sReason, $sBody)

	_DebugOut(@CRLF & "Function: HttpSendHttpResponse")

	Local $aResult[0]
	Local $pKnownHeader = Null
	Local $iFlags = 0

	Local $tResponse      = DllStructCreate($__HTTPAPI_gtagHTTP_RESPONSE_V2), _
	      $tCachePolicy   = DllStructCreate($__HTTPAPI_gtagHTTP_CACHE_POLICY), _
          $tLoggedData    = DllStructCreate($__HTTPAPI_gtagHTTP_LOG_FIELDS_DATA), _
          $tDataChunk     = DllStructCreate($__HTTPAPI_gtagHTTP_DATA_CHUNK_FROM_MEMORY), _
	      $tReason        = "", _
	      $tLocation      = "", _
	      $tBody          = "", _
	      $tContentType   = __HTTPAPI_CreateStringBuffer("text/html"), _
	      $tContentLength = __HTTPAPI_CreateStringBuffer(String(DllStructGetSize($tBody))), _
		  $tHeader        = ""


	;If redirect
	If ($iStatusCode = 301 Or $iStatusCode = 302) Then
		;Set reason based on redirect status code and location header value in $sReason
		$tReason = __HTTPAPI_CreateStringBuffer(($iStatusCode = 301 ? "Moved Permanently" : "Found"))

		$tLocation              = __HTTPAPI_CreateStringBuffer($sReason)
		$pKnownHeader           = DllStructGetPtr($tResponse, "KnownHeaders") + ($__HTPPAPI_HttpHeaderLocation * $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE)
		$tHeader                = DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER, $pKnownHeader)
		$tHeader.pRawValue      = DllStructGetPtr($tLocation)
		$tHeader.RawValueLength = DllStructGetSize($tLocation)
	Else
		;Set reason
		$tReason = __HTTPAPI_CreateStringBuffer($sReason)
	EndIf

	;Set HTTP body values, if body exists
	If StringLen($sBody) > 0 Then
		$tBody          = __HTTPAPI_CreateStringBuffer($sBody)
		$tContentType   = __HTTPAPI_CreateStringBuffer("text/html")
	    $tContentLength = __HTTPAPI_CreateStringBuffer(String(DllStructGetSize($tBody)))
	EndIf

	;Init response
	$tResponse.MajorVersion  = 1
	$tResponse.MinorVersion  = 1
	$tResponse.StatusCode    = $iStatusCode
	$tResponse.pReason       = DllStructGetPtr($tReason)
	$tResponse.ReasonLength  = DllStructGetSize($tReason)
	$tResponse.LogDataType   = $__HTTPAPI_HttpLogDataTypeFields

	;Add known header(s) to response
	If StringLen($sBody) > 0 Then
		$pKnownHeader           = DllStructGetPtr($tResponse, "KnownHeaders") + ($__HTPPAPI_HttpHeaderContentType * $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE)
		$tHeader                = DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER, $pKnownHeader)
		$tHeader.pRawValue      = DllStructGetPtr($tContentType)
		$tHeader.RawValueLength = DllStructGetSize($tContentType)

		$pKnownHeader           = DllStructGetPtr($tResponse, "KnownHeaders") + ($__HTPPAPI_HttpHeaderContentLength * $__HTTPAPI_HTTP_KNOWN_HEADER_SIZE)
		$tHeader                = DllStructCreate($__HTTPAPI_gtagHTTP_KNOWN_HEADER, $pKnownHeader)
		$tHeader.pRawValue      = DllStructGetPtr($tContentLength)
		$tHeader.RawValueLength = DllStructGetSize($tContentLength)
	EndIf

	;Close connection if not status code 200
	If $iStatusCode <> 200 Then $iFlags = $__HTTPAPI_HTTP_SEND_RESPONSE_FLAG_DISCONNECT

	;Add entity chunk (body) to response (if exists)
	If StringLen($sBody) > 0 Then
		$tDataChunk.pBuffer         = DllStructGetPtr($tBody)
		$tDataChunk.pBufferLength   = DllStructGetSize($tBody)
		$tResponse.EntityChunkCount = 1
		$tResponse.pEntityChunks    = DllStructGetPtr($tDataChunk)
	Else
		$tResponse.EntityChunkCount = 0
	EndIf

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpSendHttpResponse", _
					   "handle",  $hRequestQueue, _
					   "uint64",  $iRequestId, _
					   "ulong",   $iFlags, _          ;Flags
					   "struct*", $tResponse, _       ;Response
					   "struct*", $tCachePolicy, _    ;pCachePolicy
					   "ulong*",  Null, _             ;pBytesSent
					   "ptr",     Null, _             ;Reserved1
					   "ulong",   0, _                ;Reserved2
					   "struct*", Null, _             ;pOverlapped
					   "struct*", $tLoggedData)       ;pLoggedData
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from HttpSendHttpResponse", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the function."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpShutdownRequestQueue
; Description ...: Shuts down the specified request queue.
; Syntax ........: _HTTPAPI_HttpShutdownRequestQueue($hRequestQueue)
; Parameters ....: $hRequestQueue       A handle to the request queue.
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......: Stops queuing requests for the specified request queue process. Outstanding calls to HttpReceiveHttpRequest are canceled.
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpshutdownrequestqueue
; ===============================================================================================================================
Func _HTTPAPI_HttpShutdownRequestQueue($hRequestQueue)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpShutdownRequestQueue")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpShutdownRequestQueue", _
					   "handle", $hRequestQueue)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from _HTTPAPI_HttpShutdownRequestQueue", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the functiosn."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpSetUrlGroupProperty
; Description ...:
; Syntax ........: _HTTPAPI_HttpSetUrlGroupProperty($iUrlGroupId, $iProperty, $pInfo, $iInfoLength)
; Parameters ....: $iUrlGroupId         - An integer value containing the URL Group ID
;                  $iProperty           - An integer value containing the property to set.  See server property enums.
;                  $pInfo               - A pointer to the info/structure.
;                  $iInfoLength         - An integer containing the size (in bytes) of the info/structure.
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Remarks .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpseturlgroupproperty
; ===============================================================================================================================
Func _HTTPAPI_HttpSetUrlGroupProperty($iUrlGroupId, $iProperty, $pInfo, $iInfoLength)

	_DebugOut(@CRLF & "Function: _HTTPAPI_HttpSetUrlGroupProperty")

	Local $aResult[0]
	Local $tByteBuffer = ""

	If $__HTTPAPI_gbDebugging Then
		$tByteBuffer = DllStructCreate(StringFormat("byte data[%i]", $iInfoLength), $pInfo)
		_DebugOut("$tBindingInfo = " & $tByteBuffer.data)
	EndIf

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpSetUrlGroupProperty", _
					   "uint64", $iUrlGroupId, _
					   "int",    $iProperty, _
					   "ptr",    $pInfo, _
					   "ulong",  $iInfoLength)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from _HTTPAPI_HttpSetUrlGroupProperty", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "An invalid parameter was passed to the functiosn."
			Case $__HTTPAPI_ERROR_INVALID_HANDLE
				$__HTTPAPI_gsLastErrorMessage = "An invalid handle was passed to the function."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_Startup
; Description ...: Initialize HTTPAPI environment
; Syntax ........: __HTTPAPI_Startup([$bDebug = False])
; Parameters ....: $bDebug              [optional] Enable debug messages. Default is False.
; Return values .: Success:             True
;                  Failure:             False and sets the @error flag to non-zero
;                  @error:              1 - DllOpen failed
; Author ........: TheXman
; ===============================================================================================================================
Func _HTTPAPI_Startup($bDebug = False)

	;If dll not opened yet
	If $__HTTPAPI_ghHttpApiDll = -1 Then
		If $bDebug Then
			_DebugSetup("HTTPAPI", False, 5)
			$__HTTPAPI_gbDebugging = True
		EndIf

		_DebugOut(@CRLF & "==========================================================================================")
		_DebugOut(@CRLF & "Function: __HTTPAPI_Startup()")

		;Open dll
		$__HTTPAPI_ghHttpApiDll = DllOpen("httpapi.dll")
		If $__HTTPAPI_ghHttpApiDll = -1 Then
			$__HTTPAPI_gsLastErrorMessage = "Unable to open httpapi.dll"
			Return SetError(1, 0, False)
		EndIf

		;Set flag and register _HTTPAPI_ShutDown() on exit
		OnAutoItExitRegister("__HTTPAPI_ShutDown")
	EndIf

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_HttpTerminate
; Description ...: Cleans up resources used by the HTTP Server API
; Syntax ........: _HTTPAPI_HttpTerminate()
; Parameters ....: None
; Return values .: Success:             True
;                  Failure:             False and sets @error flag to non-zero.
;                  @error:              1 - DllCall error - @extended = @error
;                                       2 - Bad return code returned from function call - @extended = return code
; Author ........: TheXman
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/nf-http-httpterminate
; ===============================================================================================================================
Func _HTTPAPI_HttpTerminate()

	_DebugOut(@CRLF & "Function: __HTTPAPI_HttpTerminate")

	Local $aResult[0]

	;Call API function
	$aResult = DllCall($__HTTPAPI_ghHttpApiDll, "int", "HttpTerminate", _
					   "ulong",  $__HTTPAPI_HTTP_INITIALIZE_SERVER, _
					   "ptr",    Null)
	If @error Then
		$__HTTPAPI_gsLastErrorMessage = "DllCall function failed."
		Return SetError(1, @error, False)
	EndIf

	_DebugReportVar("Result from HttpTerminate", $aResult)

	;If bad return code from function
	If $aResult[0] <> 0 Then
		;Set last error message
		Switch $aResult[0]
			Case $__HTTPAPI_ERROR_INVALID_PARAMETER
				$__HTTPAPI_gsLastErrorMessage = "The Flags parameter contains an unsupported value."
			Case Else
				$__HTTPAPI_gsLastErrorMessage = "Unrecognized error - check winerror.h for more information."
		EndSwitch

		;Return with error
		Return SetError(2, $aResult[0], False)
	EndIf

	;All is good
	Return True

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _HTTPAPI_UDFVersion
; Description ...: Return the UDF file's version number
; Syntax ........: _HTTPAPI_UDFVersion()
; Parameters ....: None
; Return values .: A string version number.
; Author ........: TheXman
; ===============================================================================================================================
Func _HTTPAPI_UDFVersion()

	_DebugOut(@CRLF & "Function: _HTTPAPI_UDFVersion")

	Return $__HTTPAPI_UDF_VERSION

EndFunc

;===================================================  INTERNAL FUNCTIONS  =======================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __HTTPAPI_CreateStringBuffer
; Description ...: Creates a structure that contains a string.
; Syntax ........: __HTTPAPI_CreateStringBuffer($sString, $bUnicode = False)
; Parameters ....: $sString             A string.
;                  $bUnicode            [optional] A boolean value that specifies whether string is a wchar.
; Return values .: Success:             A structure containing the string
;                  Failure:             "" and sets @error flag to non-zero.
;                  @error:              1 - DllStructCreate error - @extended = @error
; Author ........: TheXman
; ===============================================================================================================================
Func __HTTPAPI_CreateStringBuffer($sString, $bUnicode = False)

	Local $tBuffer = ""

	;Create struct based on unicode flag
	If $bUnicode Then
		$tBuffer = DllStructCreate(StringFormat("wchar data[%i]", StringLen($sString)))
	Else
		$tBuffer = DllStructCreate(StringFormat("char data[%i]",  StringLen($sString)))
	EndIf
	If @error Then Return SetError(1, @error, "")

	;Copy string to struct
	$tBuffer.data = $sString

	;All is good, return structure
	Return $tBuffer

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __HTTPAPI_Shutdown
; Description ...: Closes down the HTTPAPI environment.
; Syntax ........: __HTTPAPI_Shutdown()
; Parameters ....: None
; Return values .: None
; Author ........: TheXman
; ===============================================================================================================================
Func __HTTPAPI_Shutdown()

	_DebugOut(@CRLF & "Function: __HTTPAPI_Shutdown")

	;If dll file is open, then close it
	If $__HTTPAPI_ghHttpApiDll <> -1 Then
		DllClose($__HTTPAPI_ghHttpApiDll)
		$__HTTPAPI_ghHttpApiDll = -1
	EndIf

EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __HTTPAPI_ApiVersion
; Description ...: Returns a populated HTTPAPI_VERSION structure with the specified values.
; Syntax ........: __HTTPAPI_ApiVersion($iMajorVersion, $iMinorVersion = 0)
; Parameters ....: $iMajorVersion       Integer containing the server API major version number.
;                  $iMinorVersion       [optional] Integer containing the server API minor version number. Default is 0.
; Return values .: Success:             An populated HTTPAPI_VERSION structure
;                  Failure:             "" and sets @error flag to non-zero.
;                                       @extended is set to @error or status code from function that failed.
;                  @error:              1 - Invalid $iMajorVersion
;                                       2 - Invalid $iMinorVersion
; Author ........: TheXman
; Remark ........:
; Link ..........: https://docs.microsoft.com/en-us/windows/win32/api/http/ns-http-httpapi_version
; ===============================================================================================================================
Func __HTTPAPI_ApiVersion($iMajorVersion, $iMinorVersion = 0)

	_DebugOut(@CRLF & "Function: __HTTPAPI_VERSION")

	Local $tHttpApiVersion = ""

	;If version is not valid, then return with error
	If $iMajorVersion < 1 Or $iMajorVersion > 2 Then Return SetError(1, 0, "")
	If $iMinorVersion <> 0                      Then Return SetError(2, 0, "")

	;Create struct and populate it
	$tHttpApiVersion                     = DllStructCreate($__HTTPAPI_gtagHTTPAPI_VERSION)
	$tHttpApiVersion.HttpApiMajorVersion = $iMajorVersion
	$tHttpApiVersion.HttpApiMinorVersion = $iMinorVersion

	;Return structure
	Return $tHttpApiVersion

EndFunc
