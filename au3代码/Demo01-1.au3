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


$hMainGUI = GUICreate("Normalize and Release-0210-3pm", @DesktopWidth/2, @DesktopHeight / 2-50,  -1, -1, -1, $WS_EX_ACCEPTFILES) ;主窗口
GUICtrlCreateLabel("复制粘贴输入即将要发布的工作产物包所在路径（文件夹路径，不是单个文件），或者也可将文件夹拖拽至下框中达到同样效果:", 50, 50)    ; creat label
$path = GUICtrlCreateInput("", 130, 100, @DesktopWidth/3, 40)    ; creat input box
GUICtrlSetState($path, $GUI_DROPACCEPTED) ;调整指定控件的状态.

Local $iOKButton = GUICtrlCreateButton("版本升级",400, 270, 100,30)     ; creat publish button
Local $iOKButton2 = GUICtrlCreateButton("发布", 150, 160, 100,30)   ;发布施工之后的
Local $iOKButton4 = GUICtrlCreateButton("同步更新",400, 160, 100,30)   ;发布施工之后的
;Local $iOKButton3 = GUICtrlCreateButton("创建项目", 10, 10, 90)   ;创建初始项目  需要放在难点的角落
Local $iOKButton5 = GUICtrlCreateButton("预发布", 150, 270, 100,30)   ;预发布
GUICtrlCreateLabel("RELEASED-->>施工中  升级需耗时一分钟左右", 400, 310)
GUICtrlCreateLabel("施工中-->>RELEASED  点发布按钮", 150, 200)
GUICtrlCreateLabel("同步更新到飞书知识库上", 400, 200)
GUICtrlCreateLabel("预发布更新知识库  通知人员", 150, 310)

GUISwitch($hMainGUI)
GUISetState(@SW_SHOW,$hMainGUI)
HotKeySet("{ESC}", "Terminate")



Local $bMsg = 0
While 1
        $bMsg = GUIGetMsg(1)
             Select
		Case $bMsg[0] = $iOKButton

			Local $check=MsgBox(1, "升级二次确认", "是否开始升级"&GUICtrlRead($path)&" ?")
			If $check==1 Then
			CheckFile(GUICtrlRead($path))	;先检测是否符合升级条件    R->施
			GetHash(GUICtrlRead($path))  ;在excel中记录解锁前的哈希值
			AddRow(GUICtrlRead($path))		;在表格相应列加版本号
			CopyDir(GUICtrlRead($path))  ;复制出另外一个文件
			StartUpdate(GUICtrlRead($path)) ;文件夹下面的文件加前缀
			
			$code=$cmdline[1]    ;传code给python
			$Code=StringTrimLeft($code,9)   ;纯code
			MsgBox(0,0,$Code)
			SendToPython2(GUICtrlRead($path),$Code)
	
			MsgBox(0, "INFO", "升级成功！")
			EndIf
		Case $bMsg[0] = $iOKButton2
			Local $check=MsgBox(1, "发布二次确认", "是否发布"&GUICtrlRead($path)&" ?")    ;施---》R
			If $check==1 Then
			CheckFile2(GUICtrlRead($path))   		;检测是否符合发布条件（文件数量、文件名）
			GetNewHash(GUICtrlRead($path))             ;获取并记录哈希值     1、计算施工后的哈希，与之前release的哈希值比较
			CopyDir2(GUICtrlRead($path)) 			;先复制出另外一个大文件
			Released(GUICtrlRead($path)) 			;修改文件夹名字   以及特定文件的属性
			
			$code=$cmdline[1]    ;传code给python
			$Code=StringTrimLeft($code,9)   ;纯code
			MsgBox(0,0,$Code)
			SendToPython(GUICtrlRead($path),$Code)
			
			MsgBox(0, "INFO", "发布成功！")
			ExitLoop
		EndIf
		#cs
		Case $bMsg[0] = $iOKButton3
			GUISetState(@SW_HIDE,$hMainGUI)
			ShowGui()
			GUISetState(@SW_SHOW,$hMainGUI)
		#ce
		Case $bMsg[0] = $iOKButton4
			
			$code=$cmdline[1]    ;传code给python
			$Code=StringTrimLeft($code,9)   ;纯code
			MsgBox(0,0,$Code)
			UpDataFile(GUICtrlRead($path),$Code)
			
			MsgBox(1,'INFO','同步更新已完成！')
			Local $check=MsgBox(1, 'INFO', "操作已完成，消息已传达，正在退出")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		
		Case $bMsg[0] = $iOKButton5
		
			$code=$cmdline[1]    ;传code给python   每个按钮的code值不一样  ！！ good
			$Code=StringTrimLeft($code,9)   ;纯code
			MsgBox(0,0,$Code)
			
			PrePost(GUICtrlRead($path),$Code)
			MsgBox(1,'INFO','预发布完成！')
			Local $check=MsgBox(1, 'INFO', "操作已完成，消息已传达，正在退出")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		
		Case $bMsg[0] = $GUI_EVENT_CLOSE And $bMsg[1] = $hMainGUI
			Local $check=MsgBox(1, 'INFO', "是否退出？")
			If $check ==1 Then
			ExitLoop
			;Else ContinueLoop
			EndIf
		EndSelect
		
	
			;ExitLoop
			;EndSelect

WEnd 







