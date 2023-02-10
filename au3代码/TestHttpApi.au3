#include <Constants.au3>
#include <Array.au3>
#include <InetConstants.au3>
#include "HttpApi.au3"
Func httpapi_initialize()
	;Intialize the API
	_HTTPAPI_HttpInitialize()
	If @error Then
		MsgBox($MB_ICONERROR, "_HTTPAPI_HttpInitialize ERROR", _
			"@error =" & @error & "  @extended = " & @extended & @CRLF & @CRLF & _
			_HTTPAPI_GetLastErrorMessage())
		Exit 1
	EndIf
	OnAutoItExitRegister(httpapi_terminate)
EndFunc

Func httpapi_terminate()
	;Terminate the API
	_HTTPAPI_HttpTerminate()
	If @error Then
		MsgBox($MB_ICONERROR, "_HTTPAPI_HttpTerminate ERROR", _
			"@error =" & @error & "  @extended = " & @extended & @CRLF & @CRLF & _
			_HTTPAPI_GetLastErrorMessage())
	EndIf

	MsgBox($MB_ICONINFORMATION, "INFO", "HTTP Server has successfully terminated.")
EndFunc
httpapi_initialize()


ShellExecute("https://open.feishu.cn/open-apis/wiki/v2/spaces")
$data='"name","知识空间","description","知识空间描述"'
;Local $pRequestHeaders=
;_ArrayAdd($pRequestHeaders,'Content-Type: application/json')
;_ArrayAdd($pRequestHeaders,'Authorization: Bearer u-0_32LepVV5tFcgN_q5Q7Bz0lltOxlhm1OO0015say79n')
Local $pRequestHeaders='Authorization:Bearer u-0_32LepVV5tFcgN_q5Q7Bz0lltOxlhm1OO0015say79n'
$b=_HTTPAPI_GetRequestHeaders($pRequestHeaders)
MsgBox(0,0,$b)

;$a= _HTTPAPI_GetRequestBody($data)
 ;MsgBox(0,0,$a)