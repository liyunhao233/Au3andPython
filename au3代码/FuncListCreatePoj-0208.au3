#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>

Global $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
ShowGui($timestamp)

;通过复制一个存在于公共盘上面的xlsx文件。   步骤：1、先复制，再打开。2、打开后进行添加信息操作
Func ShowGui($timestamp)
$GUI=GUICreate("输入项目信息",450,300)
$BT1=GUICtrlCreateButton("确定",170,250,100,20)

$car=GUICtrlCreateInput("",60,20,200,25)
$chair=GUICtrlCreateInput("",60,55,200,25)
$shape=GUICtrlCreateInput("",60,90,200,25)

GUICtrlCreateLabel("整车厂",10,25,40,30)
GUICtrlCreateLabel("整椅厂",10,60,50,30)
GUICtrlCreateLabel("车型",10,95,40,30)


GUICtrlCreateLabel("A/B ：代表芯片来源，A:国产芯片 B：进口芯片		01/02/03/04:代表座椅的位置，01：前左，02：前右，03：后左，04：后右",270,60,100,100)
GUISetState(@SW_ENABLE,$GUI)
GUISetState(@SW_SHOW,$GUI)
GUICtrlCreateGroup("芯片类型", 30, 130, 180, 50)
$Ca=GUICtrlCreateRadio("A:国产",50,145,50,30)
$Cb=GUICtrlCreateRadio("B:进口",130,145,50,30)
GUICtrlSetState($Ca, $GUI_CHECKED)
$C='国产'
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("1/2/3排", 30, 190, 200, 50)
$Aa=GUICtrlCreateRadio("1排",50,200,50,30)
$Ab=GUICtrlCreateRadio("2排",100,200,50,30)
$Ac=GUICtrlCreateRadio("3排",150,200,50,30)
GUICtrlSetState($Aa, $GUI_CHECKED)
$A='01'
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("左/右驾", 200, 190, 180, 50)
$Ba=GUICtrlCreateRadio("左驾",250,200,50,30)
$Bb=GUICtrlCreateRadio("右驾",300,200,50,30)
GUICtrlSetState($Ba, $GUI_CHECKED)
$B='左驾'
GUICtrlCreateGroup("", -99, -99, 1, 1)


Local $msg=0
While 1
	$msg=GUIGetMsg(1)
	Select
		
	Case $msg[0] = $Aa
		$A='01'
	Case $msg[0] = $Ab
		$A='02'
	Case $msg[0] = $Ac
		$A='03'	
	Case $msg[0] = $Ba
		$B='左驾'	
	Case $msg[0] = $Bb
		$B='右驾'
	Case $msg[0] = $Ca
		$C='国产'
	Case $msg[0] = $Cb
		$C='进口'
		
	Case $msg[0] = $BT1
	If(StringIsAlNum(GUICtrlRead($car))=0 Or StringIsAlNum(GUICtrlRead($shape))=0 Or StringIsAlNum(GUICtrlRead($chair))=0 ) Then
			MsgBox(16,'error','还有信息未填写！')
	
	Else
			Showinfo($car,$chair,$shape,$A,$B,$C,$timestamp) ;车、椅子、车型、排、左右驾、芯片类型
			;CreateFile($car,$chair,$shape,$A,$B,$C)
			CopyExcel($car,$chair,$shape,$A,$B,$C,$timestamp) ;车厂名-座椅厂-车型-芯片-左右驾-多少排
			
			;;获取创建项目中的信息，并且利用API创建一个文件夹
			;CreatePoj($car,$chair,$shape)    ;格式：Greely_YF_CS1E_v001
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			;GUISetState(@SW_SHOW,$hMainGUI)
			
			ExitLoop
	EndIf
	Case $msg[0] =$GUI_EVENT_CLOSE And $msg[1] = $GUI
			Local $check=MsgBox(1, "INFO", "是否退出？")
			If $check==1 Then
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			;GUISetState(@SW_SHOW,$hMainGUI)
			ExitLoop
			EndIf
		EndSelect
WEnd

$car1=GUICtrlRead($car)
$chair1=GUICtrlRead($chair)
$shape1=GUICtrlRead($shape)
;调用python 在知识空间中创建大的项目号
$code=$cmdline[1]    ;传code给python
;MsgBox(0,0,$code)   ;开始处理code格式
$CodeAndId=StringTrimLeft($code,7)
;MsgBox(0,3,$CodeAndId)

$xlsxpath='\\192.168.2.100\100\版本迭代\'&$car1&'_'&$chair1&'_'&$shape1&'_'&$C&$A&'_'&$B&'\'&'V000'&'-'&$timestamp&'_施工中'&'\'&'V000_'&$timestamp&'_A13Calc_施工中'&'.xlsx'
ShellExecute("\\192.168.2.100\100\0208CreatePoj\0208CreatePoj.exe",$car1&'_'&$chair1&'_'&$shape1&'_'&$A&'_'&$B&'_'&$C&'>'&$xlsxpath&'>'&$CodeAndId)  


EndFunc


Func Showinfo($aa,$bb,$cc,$d,$e,$f,$timestamp)	;车厂名-座椅厂-车型-芯片-左右驾-多少排
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	MsgBox(0,'创建初始项目',GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中')
	
	
	Local $oExcel = _Excel_Open(False, True, Default, Default, True)
	Local $oWorkbook = _Excel_BookNew($oExcel)
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中'&'\'&'V000_'&$timestamp&'_A13Calc_施工中'&'.xlsx')   ;看起来自动创建xlsx文件已经解决   12/12  18:18
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中'&'\'&'V000_'&$timestamp&'_A10BOM_施工中'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中'&'\'&'V000_'&$timestamp&'_A12BOMtoERP_施工中'&'.xlsx')  
	_Excel_Close($oExcel)
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	
EndFunc


Func CopyExcel($aa,$bb,$cc,$d,$e,$f,$timestamp) ;车厂名-座椅厂-车型-芯片-左右驾-多少排
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	FileChangeDir('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A01_客户标准'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A02_通信协议'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A30_3D数模'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A32_2D总成内部'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A33_2D总成客户'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A50_软件源代码'&'_施工中')
	DirCreate('V000_'&$timestamp&'_'&'A51_治具软件代码'&'_施工中')
	;MsgBox(0,0,'文件夹')
	FileOpen('V000_'&$timestamp&'_'&'A09_ECR说明'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A11_BOM工具中间件'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B30_发给供应商数模'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C30_能够释放给客户的3D数模，必须去掉内部结构'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C20_发给客户的TRpdf'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A31_stp/igs格式数模'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C33_能够释放给客户的图纸'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A34_线束图纸'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A35_CC/SC'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A36_DFMEA'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A40_Pads产物'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B40_何种方式给供应商'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A41_SMT元件的坐标文件'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A42_SMT元件参数'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A43_eICD(电装图、线束原理图)'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C43_释放给客户图纸'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A52_产品Hex'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B52_用何种方式发给工厂'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A53_治具软件hex'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A54_代码测试checklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A60_试产checklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A70_功能测试checklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A71_整车联调checklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A72_压力测试checklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A73_真实DVchecklist'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A80_DVP-结构-给客户看'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A81_DVP-电子-给客户看'&'_施工中.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A82_DV报告'&'_施工中.txt',2)
	
	
	$path='\\192.168.2.100\100\版本迭代\模板文件\A00 Draft 230105.xlsx'
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$path)
	$oWorkbook.Sheets(1).Cells(1,2).Value =GUICtrlRead($aa)
	$oWorkbook.Sheets(1).Cells(1,5).Value =GUICtrlRead($bb)
	$oWorkbook.Sheets(1).Cells(1,8).Value =GUICtrlRead($cc)
	$oWorkbook.Sheets(1).Cells(1,11).Value =$f
	$oWorkbook.Sheets(1).Cells(1,14).Value =$e&$d
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_施工中'&'\'&'V000_'&$timestamp&'_A00Card_施工中'&'.xlsx')
	;_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&'_A00Card_RELEASED'&'.xlsx')
	MsgBox(0,0,'操作结束！')
	_Excel_Close($oExcel)
	
EndFunc












