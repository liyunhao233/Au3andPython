$Temp_ID=7187632384194396163  ;知识空间id
Global Const $HTTP_STATUS_OK = 200  
$URL = "https://open.feishu.cn/open-apis/wiki/v2/spaces"
$header1="Authorization"
$header2="Bearer u-3m6C2IPRh9cFDkSIvAeH.Y0llZqxlhYHO20044cay2sj"
$header3="Content-Type"
$header4="application/json"
$name='"name":"知识空间0112/1pm",'
$discrib='"dscription":"知识空间描述:这是知识空间api的调用，需要自动获取token之后才可以算自动生成！"'
$t=''&$name&$discrib&''
;$sBody='{'&$t&'}'
$sBody='{"name":"知识空间0112/1:30pm","dscription":"知识空间描述:这是知识空间api的调用，需要自动获取token之后才可以算自动生成！"}'
	
	
HttpPost($URL,$sBody)

;使用转瞬即逝的自建小程序
$app_id="cli_a34a74d63a395013";生成应用之后的唯一标识
$app_secret="5AxOX3OVfGE7DVV9IO6iIhPxv4HllErF";应用密匙
;以上两个是获取  app_access_token 或者其他两个token 的必要   token有时效



;请求体的JSON格式   后期维护



;定义方法
;;上传文件  先到个人云空间里面
Func PostFile($URL,$FilePath)


	EndFunc

;;获取access token  这个获取相对简单  切割需要精确一点   https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal
Func GetToken($URL,$appid,$appsecret)
	

	EndFunc

;;先上传文件到个人云空间，还需要从个人云空间到知识空间指定的id下面的
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



;;;;  变成 json
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