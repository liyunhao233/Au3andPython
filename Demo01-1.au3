#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#Include <File.au3>
#include <Excel.au3>
#include "FuncList.au3"
#include "FuncList2.au3"

#include <Date.au3>


$hMainGUI = GUICreate("Normalize and Release-0210-3pm", @DesktopWidth/2, @DesktopHeight / 2-50,  -1, -1, -1, $WS_EX_ACCEPTFILES) ;������
GUICtrlCreateLabel("����ճ�����뼴��Ҫ�����Ĺ������������·�����ļ���·�������ǵ����ļ���������Ҳ�ɽ��ļ�����ק���¿��дﵽͬ��Ч��:", 50, 50)    ; creat label
$path = GUICtrlCreateInput("", 130, 100, @DesktopWidth/3, 40)    ; creat input box
GUICtrlSetState($path, $GUI_DROPACCEPTED) ;����ָ���ؼ���״̬.

Local $iOKButton = GUICtrlCreateButton("�汾����",400, 270, 100,30)     ; creat publish button
Local $iOKButton2 = GUICtrlCreateButton("����", 150, 160, 100,30)   ;����ʩ��֮���
Local $iOKButton4 = GUICtrlCreateButton("ͬ������",400, 160, 100,30)   ;����ʩ��֮���
;Local $iOKButton3 = GUICtrlCreateButton("������Ŀ", 10, 10, 90)   ;������ʼ��Ŀ  ��Ҫ�����ѵ�Ľ���
Local $iOKButton5 = GUICtrlCreateButton("Ԥ����", 150, 270, 100,30)   ;Ԥ����
GUICtrlCreateLabel("RELEASED-->>ʩ����  �������ʱһ��������", 400, 310)
GUICtrlCreateLabel("ʩ����-->>RELEASED  �㷢����ť", 150, 200)
GUICtrlCreateLabel("ͬ�����µ�����֪ʶ����", 400, 200)
GUICtrlCreateLabel("Ԥ��������֪ʶ��  ֪ͨ��Ա", 150, 310)

GUISwitch($hMainGUI)
GUISetState(@SW_SHOW,$hMainGUI)
HotKeySet("{ESC}", "Terminate")



Local $bMsg = 0
While 1
        $bMsg = GUIGetMsg(1)
             Select
		Case $bMsg[0] = $iOKButton

			Local $check=MsgBox(1, "��������ȷ��", "�Ƿ�ʼ����"&GUICtrlRead($path)&" ?")
			If $check==1 Then
			CheckFile(GUICtrlRead($path))	;�ȼ���Ƿ������������    R->ʩ
			GetHash(GUICtrlRead($path))  ;��excel�м�¼����ǰ�Ĺ�ϣֵ
			AddRow(GUICtrlRead($path))		;�ڱ����Ӧ�мӰ汾��
			CopyDir(GUICtrlRead($path))  ;���Ƴ�����һ���ļ�
			StartUpdate(GUICtrlRead($path)) ;�ļ���������ļ���ǰ׺
			
			$code=$cmdline[1]    ;��code��python
			$Code=StringTrimLeft($code,9)   ;��code
			MsgBox(0,0,$Code)
			SendToPython2(GUICtrlRead($path),$Code)
	
			MsgBox(0, "INFO", "�����ɹ���")
			EndIf
		Case $bMsg[0] = $iOKButton2
			Local $check=MsgBox(1, "��������ȷ��", "�Ƿ񷢲�"&GUICtrlRead($path)&" ?")    ;ʩ---��R
			If $check==1 Then
			CheckFile2(GUICtrlRead($path))   		;����Ƿ���Ϸ����������ļ��������ļ�����
			GetNewHash(GUICtrlRead($path))             ;��ȡ����¼��ϣֵ     1������ʩ����Ĺ�ϣ����֮ǰrelease�Ĺ�ϣֵ�Ƚ�
			CopyDir2(GUICtrlRead($path)) 			;�ȸ��Ƴ�����һ�����ļ�
			Released(GUICtrlRead($path)) 			;�޸��ļ�������   �Լ��ض��ļ�������
			
			$code=$cmdline[1]    ;��code��python
			$Code=StringTrimLeft($code,9)   ;��code
			MsgBox(0,0,$Code)
			SendToPython(GUICtrlRead($path),$Code)
			
			MsgBox(0, "INFO", "�����ɹ���")
			ExitLoop
		EndIf
		#cs
		Case $bMsg[0] = $iOKButton3
			GUISetState(@SW_HIDE,$hMainGUI)
			ShowGui()
			GUISetState(@SW_SHOW,$hMainGUI)
		#ce
		Case $bMsg[0] = $iOKButton4
			
			$code=$cmdline[1]    ;��code��python
			$Code=StringTrimLeft($code,9)   ;��code
			MsgBox(0,0,$Code)
			UpDataFile(GUICtrlRead($path),$Code)
			
			MsgBox(1,'INFO','ͬ����������ɣ�')
			Local $check=MsgBox(1, 'INFO', "��������ɣ���Ϣ�Ѵ�������˳�")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		
		Case $bMsg[0] = $iOKButton5
		
			$code=$cmdline[1]    ;��code��python   ÿ����ť��codeֵ��һ��  ���� good
			$Code=StringTrimLeft($code,9)   ;��code
			MsgBox(0,0,$Code)
			
			PrePost(GUICtrlRead($path),$Code)
			MsgBox(1,'INFO','Ԥ������ɣ�')
			Local $check=MsgBox(1, 'INFO', "��������ɣ���Ϣ�Ѵ�������˳�")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		
		Case $bMsg[0] = $GUI_EVENT_CLOSE And $bMsg[1] = $hMainGUI
			Local $check=MsgBox(1, 'INFO', "�Ƿ��˳���")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		EndSelect
		
	
			;ExitLoop
			;EndSelect

WEnd 







