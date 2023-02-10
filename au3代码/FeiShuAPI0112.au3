$Temp_ID=7187632384194396163  ;֪ʶ�ռ�id
Global Const $HTTP_STATUS_OK = 200  
$URL = "https://open.feishu.cn/open-apis/wiki/v2/spaces"
$header1="Authorization"
$header2="Bearer u-3m6C2IPRh9cFDkSIvAeH.Y0llZqxlhYHO20044cay2sj"
$header3="Content-Type"
$header4="application/json"
$name='"name":"֪ʶ�ռ�0112/1pm",'
$discrib='"dscription":"֪ʶ�ռ�����:����֪ʶ�ռ�api�ĵ��ã���Ҫ�Զ���ȡtoken֮��ſ������Զ����ɣ�"'
$t=''&$name&$discrib&''
;$sBody='{'&$t&'}'
$sBody='{"name":"֪ʶ�ռ�0112/1:30pm","dscription":"֪ʶ�ռ�����:����֪ʶ�ռ�api�ĵ��ã���Ҫ�Զ���ȡtoken֮��ſ������Զ����ɣ�"}'
	
	
HttpPost($URL,$sBody)

;ʹ��ת˲���ŵ��Խ�С����
$app_id="cli_a34a74d63a395013";����Ӧ��֮���Ψһ��ʶ
$app_secret="5AxOX3OVfGE7DVV9IO6iIhPxv4HllErF";Ӧ���ܳ�
;���������ǻ�ȡ  app_access_token ������������token �ı�Ҫ   token��ʱЧ



;�������JSON��ʽ   ����ά��



;���巽��
;;�ϴ��ļ�  �ȵ������ƿռ�����
Func PostFile($URL,$FilePath)


	EndFunc

;;��ȡaccess token  �����ȡ��Լ�  �и���Ҫ��ȷһ��   https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal
Func GetToken($URL,$appid,$appsecret)
	

	EndFunc

;;���ϴ��ļ��������ƿռ䣬����Ҫ�Ӹ����ƿռ䵽֪ʶ�ռ�ָ����id�����
Func ToYunSpace()       
	
	
	EndFunc




Func CreatePoj($car,$chair,$shape)
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")  
	$oHTTP.Open("POST", "https://open.feishu.cn/open-apis/wiki/v2/spaces/:space_id/nodes", False)  
	
	EndFunc

Func HttpPost($sURL, $sBody)  
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")  
	$oHTTP.Open("POST", $sURL, False)  
	If (@error) Then Return SetError(1, 0, 0)  
		$oHTTP.SetRequestHeader($header1,$header2)  
		$oHTTP.SetRequestHeader($header3,$header4)  
		$oHTTP.Send($sBody)  
	MsgBox(0,0,$oHTTP.ResponseText) 
	If (@error) Then Return SetError(2, 0, 0)  
	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)  
	Return SetError(0, 0, $oHTTP.ResponseText)  

	EndFunc  


Func HttpGet($sURL, $sData = "")  
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")  
	$oHTTP.Open("GET", $sURL & "?" & $sData, False)  
	If (@error) Then Return SetError(1, 0, 0)  
		$oHTTP.Send()  
	If (@error) Then Return SetError(2, 0, 0)  
	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)  
	Return SetError(0, 0, $oHTTP.ResponseText)  
	
	EndFunc  



;;;;  ��� json
;~ #noTrayIcon
opt('mustDeclareVars',1)
opt('trayIconDebug',1)
#include "WinHttp2.au3"
#include "JSON.au3"
#include "JSON_Translate.au3" ; examples of translator functions, includes JSON_pack and JSON_unpack


Func _JSONGet($json, $path, $seperator = ".")
    Local $seperatorPos,$current,$next,$l
    $seperatorPos = StringInStr($path, $seperator)

    If $seperatorPos > 0 Then
         $current = StringLeft($path, $seperatorPos - 1)
         $next = StringTrimLeft($path, $seperatorPos + StringLen($seperator) - 1)
    Else
         $current = $path
         $next = ""
    EndIf

    If _JSONIsObject($json) Then
         $l = UBound($json, 1)
	     For $i = 0 To $l - 1
	    	 If $json[$i][0] == $current Then
		    	 If $next == "" Then
		    		 return $json[$i][1]
		    	 Else
		    		 return _JSONGet($json[$i][1], $next, $seperator)
		    	 EndIf
		      EndIf
         Next
    ElseIf IsArray($json) And UBound($json, 0) == 1 And UBound($json, 1) > $current Then
        If $next == "" Then
			return $json[$current]
        Else
            return _JSONGet($json[$current], $next, $seperator)
        EndIf
	EndIf

	return $_JSONNull
EndFunc

Global $sGet = HttpGet("https://ip-json.rhcloud.com/json")
;FileWrite("ip.json", $sGet)
;Local $data = FileRead(@ScriptDir & "\ip.json")
MsgBox(0,"",$sGet)
;Local $s3='{"site": "http://ip-json.rhcloud.com", "city": "Wuxi", "region_name": "Jiangsu", "region": "04", "area_code": 0, "time_zone": "Asia/Shanghai", "longitude": 120.28859710693359, "metro_code": 0, "country_code3": "CHN", "latitude": 31.568899154663086, "postal_code": null, "dma_code": 0, "country_code": "CN", "country_name": "China", "q": "122.193.143.35"}'
Local $json=_JSONDecode($sGet)
;Local $json2=_JSONDecode($s3)

;query this object
Local $city = _JSONGet($json, "city")
Local $province = _JSONGet($json, "region_name")
Local $timezone = _JSONGet($json2, "time_zone")

msgbox(0,default,$city&$province&$timezone)