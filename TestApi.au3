
;客户端!!!!!!!! 请先运行服务端... dummy!!
Local $g_IP = "192.168.101.137"

; 开始 TCP 服务
;==============================================
TCPStartup()

; 创建一个套接字(SOCKET)连接到已经存在的服务器
;==============================================
Local $socket = TCPConnect($g_IP, 3000)
If $socket == -1 Then 
	Exit
Else
	MsgBox(0,0,$socket)
EndIf

TCPShutdown()