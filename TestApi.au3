
;�ͻ���!!!!!!!! �������з����... dummy!!
Local $g_IP = "192.168.101.137"

; ��ʼ TCP ����
;==============================================
TCPStartup()

; ����һ���׽���(SOCKET)���ӵ��Ѿ����ڵķ�����
;==============================================
Local $socket = TCPConnect($g_IP, 3000)
If $socket == -1 Then 
	Exit
Else
	MsgBox(0,0,$socket)
EndIf

TCPShutdown()