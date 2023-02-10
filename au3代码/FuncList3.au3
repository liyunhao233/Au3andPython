#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
;Show()
Func Show()
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
$C='A'
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
$B='左'
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
		$B='左'	
	Case $msg[0] = $Bb
		$B='右'
	Case $msg[0] = $Ca
		$C='A'
	Case $msg[0] = $Cb
		$C='B'
		
	Case $msg[0] = $BT1
	If(StringIsAlNum(GUICtrlRead($car))=0 Or StringIsAlNum(GUICtrlRead($shape))=0 Or StringIsAlNum(GUICtrlRead($chair))=0 ) Then
			MsgBox(32,'test','还有信息未填写！')
	
	Else
			Showinfo($car,$chair,$shape,$A,$B,$C) ;车、椅子、车型、排、左右驾、芯片类型
			CreateFile($car,$chair,$shape,$A,$B,$C)
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			;GUISetState(@SW_SHOW,$hMainGUI)
			
			ExitLoop
	EndIf
	Case $msg[0] =$GUI_EVENT_CLOSE And $msg[1] = $GUI
			MsgBox(32, "INFO", "是否退出？")
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			GUISetState(@SW_SHOW,$hMainGUI)
			ExitLoop
		EndSelect
WEnd

EndFunc



Func Showinfo($aa,$bb,$cc,$d,$e,$f)	;车厂名-座椅厂-车型-芯片-左右驾-多少排
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	MsgBox(0,'创建初始项目',GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
EndFunc


Func CreateFile($aa,$bb,$cc,$d,$e,$f) ;车、椅子、车型、排、左右驾、芯片类型
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	FileChangeDir('\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A01_客户标准'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A02_通信协议'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A30_3D数模'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A32_2D总成内部'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A33_2D总成客户'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A50_软件源代码'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A51_治具软件代码'&'-RELEASED')
	;MsgBox(0,0,'文件夹')
	FileOpen('V000_'&$timestamp&'_'&'A09_ECR说明'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A11_BOM工具中间件'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B30_发给供应商数模'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C30_能够释放给客户的3D数模，必须去掉内部结构'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C20_发给客户的TRpdf'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A31_stp/igs格式数模'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C33_能够释放给客户的图纸'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A34_线束图纸'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A35_CC/SC'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A36_DFMEA'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A40_Pads产物'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B40_何种方式给供应商'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A41_SMT元件的坐标文件'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A42_SMT元件参数'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A43_eICD(电装图、线束原理图)'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C43_释放给客户图纸'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A52_产品Hex'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B52_用何种方式发给工厂'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A53_治具软件hex'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A54_代码测试checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A60_试产checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A70_功能测试checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A71_整车联调checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A72_压力测试checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A73_真实DVchecklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A80_DVP-结构-给客户看'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A81_DVP-电子-给客户看'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A82_DV报告'&'-RELEASED.txt',2)
	

	;MsgBox(0,0,'txt代替文件')
	Local $oExcel = _Excel_Open(Default, True, Default, Default, True)
	Local $oWorkbook = _Excel_BookNew($oExcel)
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A13Calc-RELEASED'&'.xlsx')   ;看起来自动创建xlsx文件已经解决   12/12  18:18
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A10BOM-RELEASED'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A12BOMtoERP-RELEASED'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A00Card-RELEASED'&'.xlsx')   
	
	;MsgBox(0,0,'stop2')
	#cs
	Global $xlEdgeBottom = 9 ; XlBordersIndex enumeration. Border at the bottom of the range.
	Global $xlContinuous = 1 ; XlLineStyle Enumeration. Continuous line.
	;Global $xlThin = 3 ; XlBorderWeight Enumeration. Continuous line. Thin line.
	With $oWorkbook.Sheets(1).Range("B2").Borders($xlEdgeBottom)
		.LineStyle = $xlContinuous
		.Size = 13
		.ColorIndex = 4 ; Index value into the current color palette, or as one of the XlColorIndex constants.
	EndWith
	#ce
	
	;Local $oWorkbook = _Excel_BookOpen('V000_'&$timestamp&'_A00Card-RELEASED'&'.xlsx','\\192.168.2.100\100\版本迭代\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
	;插入填入表格的数据
	$oWorkbook.Sheets(1).Cells(1,2).Value =GUICtrlRead($aa)
	$oWorkbook.Sheets(1).Cells(1,4).Value =GUICtrlRead($bb)
	$oWorkbook.Sheets(1).Cells(1,6).Value =GUICtrlRead($cc)
	$oWorkbook.Sheets(1).Cells(1,8).Value =$f
	$oWorkbook.Sheets(1).Cells(1,10).Value =$e
	$oWorkbook.Sheets(1).Cells(1,12).Value =$d
	;MsgBox(32,'a','stop')
	$oWorkbook.Sheets(1).Cells(1,1).Value ="整车厂"
	$oWorkbook.Sheets(1).Cells(1,3).Value ="整椅厂"
	$oWorkbook.Sheets(1).Cells(1,5).Value ="车型"
	$oWorkbook.Sheets(1).Cells(1,7).Value ="芯片类型"
	$oWorkbook.Sheets(1).Cells(1,9).Value ="左/右驾"
	$oWorkbook.Sheets(1).Cells(1,11).Value ="一/二/三排"
	$oWorkbook.Sheets(1).Cells(3,3).Value ="默认值"
	;插入初始表格的初始数据
	$oWorkbook.Sheets(1).Cells(4,1).Value ="创建时间"
	$oWorkbook.Sheets(1).Cells(6,1).Value ="工作产物"  
	$oWorkbook.Sheets(1).Cells(7,1).Value ="A00"				
	$oWorkbook.Sheets(1).Cells(8,1).Value ="A01"				
	$oWorkbook.Sheets(1).Cells(9,1).Value ="A02"				
	$oWorkbook.Sheets(1).Cells(10,1).Value ="A09"				
	$oWorkbook.Sheets(1).Cells(12,1).Value ="A10"				
	$oWorkbook.Sheets(1).Cells(13,1).Value ="A11"				
	$oWorkbook.Sheets(1).Cells(14,1).Value ="A12"				
	$oWorkbook.Sheets(1).Cells(15,1).Value ="A13"				
	$oWorkbook.Sheets(1).Cells(17,1).Value ="A20"				
	$oWorkbook.Sheets(1).Cells(18,1).Value ="C20"				
	$oWorkbook.Sheets(1).Cells(20,1).Value ="A30"				
	$oWorkbook.Sheets(1).Cells(21,1).Value ="B30"				
	$oWorkbook.Sheets(1).Cells(22,1).Value ="C30"				
	$oWorkbook.Sheets(1).Cells(23,1).Value ="A31"				
	$oWorkbook.Sheets(1).Cells(24,1).Value ="A32"				
	$oWorkbook.Sheets(1).Cells(25,1).Value ="A33"				
	$oWorkbook.Sheets(1).Cells(26,1).Value ="C33"				
	$oWorkbook.Sheets(1).Cells(27,1).Value ="A34"				
	$oWorkbook.Sheets(1).Cells(28,1).Value ="A35"				
	$oWorkbook.Sheets(1).Cells(29,1).Value ="A36"				
	$oWorkbook.Sheets(1).Cells(31,1).Value ="A40"				
	$oWorkbook.Sheets(1).Cells(32,1).Value ="B40"				
	$oWorkbook.Sheets(1).Cells(33,1).Value ="A41"				
	$oWorkbook.Sheets(1).Cells(34,1).Value ="A42"				
	$oWorkbook.Sheets(1).Cells(35,1).Value ="A43"				
	$oWorkbook.Sheets(1).Cells(36,1).Value ="C43"				
	$oWorkbook.Sheets(1).Cells(38,1).Value ="A50"				
	$oWorkbook.Sheets(1).Cells(39,1).Value ="A51"				
	$oWorkbook.Sheets(1).Cells(40,1).Value ="A52"				
	$oWorkbook.Sheets(1).Cells(41,1).Value ="B52"				
	$oWorkbook.Sheets(1).Cells(42,1).Value ="A53"				
	$oWorkbook.Sheets(1).Cells(43,1).Value ="A54"				
	$oWorkbook.Sheets(1).Cells(45,1).Value ="A60"				
	$oWorkbook.Sheets(1).Cells(47,1).Value ="A70"				
	$oWorkbook.Sheets(1).Cells(48,1).Value ="A71"				
	$oWorkbook.Sheets(1).Cells(49,1).Value ="A72"				
	$oWorkbook.Sheets(1).Cells(50,1).Value ="A73"				
	$oWorkbook.Sheets(1).Cells(52,1).Value ="A80"				
	$oWorkbook.Sheets(1).Cells(53,1).Value ="A81"				
	$oWorkbook.Sheets(1).Cells(54,1).Value ="A82"				
	$oWorkbook.Sheets(1).Cells(56,1).Value ="A90"				


	$oWorkbook.Sheets(1).Cells(6,2).Value ="产品属性卡"
	$oWorkbook.Sheets(1).Cells(7,2).Value ="产品参数/属性/变更履历/交付物汇总账目总表"
	$oWorkbook.Sheets(1).Cells(8,2).Value ="客户输入的技术规范/需求定义/标准"
	$oWorkbook.Sheets(1).Cells(9,2).Value ="通信协议、LIN/CAN MATRIX"
	$oWorkbook.Sheets(1).Cells(10,2).Value ="ECR说明"
	
	$oWorkbook.Sheets(1).Cells(12,2).Value ="BOM - 人类能方便看的格式"
	$oWorkbook.Sheets(1).Cells(13,2).Value ="BOM工具中间文件"
	$oWorkbook.Sheets(1).Cells(14,2).Value ="能够直接导入ERP的BOM格式"
	$oWorkbook.Sheets(1).Cells(15,2).Value ="皮料用量计算表"
	
	$oWorkbook.Sheets(1).Cells(17,2).Value ="技术方案"
	$oWorkbook.Sheets(1).Cells(18,2).Value ="发给客户版的TR pdf"
	
	$oWorkbook.Sheets(1).Cells(20,2).Value ="3D数模"
	$oWorkbook.Sheets(1).Cells(21,2).Value ="需要发给供应商的数模"
	$oWorkbook.Sheets(1).Cells(22,2).Value ="能够释放给客户的3D数模，必须去掉内部结构"
	$oWorkbook.Sheets(1).Cells(23,2).Value ="stp/igs格式数模"
	$oWorkbook.Sheets(1).Cells(24,2).Value ="2D总成图 - 内部"
	$oWorkbook.Sheets(1).Cells(25,2).Value ="2D总成图 - 给客户版"
	$oWorkbook.Sheets(1).Cells(26,2).Value ="能够释放给客户的图纸"
	$oWorkbook.Sheets(1).Cells(27,2).Value ="线束图纸"
	$oWorkbook.Sheets(1).Cells(28,2).Value ="CC/SC"
	$oWorkbook.Sheets(1).Cells(29,2).Value ="DFMEA"
	
	$oWorkbook.Sheets(1).Cells(31,2).Value ="Pads电子设计的工作产物(PCB/PCBA/原理图)"
	$oWorkbook.Sheets(1).Cells(32,2).Value ="什么方式发给供应商"
	$oWorkbook.Sheets(1).Cells(33,2).Value ="SMT元件的坐标文件"
	$oWorkbook.Sheets(1).Cells(34,2).Value ="SMT元件参数(需要额外手改) - 要与BOM对应"
	$oWorkbook.Sheets(1).Cells(35,2).Value ="eICD(电装图、线束原理图)"
	$oWorkbook.Sheets(1).Cells(36,2).Value ="能够释放给客户的图纸"
	
	$oWorkbook.Sheets(1).Cells(38,2).Value ="软件源代码"
	$oWorkbook.Sheets(1).Cells(39,2).Value ="治具软件代码"
	$oWorkbook.Sheets(1).Cells(40,2).Value ="产品hex文件"
	$oWorkbook.Sheets(1).Cells(41,2).Value ="什么方式发送给工厂?"
	$oWorkbook.Sheets(1).Cells(42,2).Value ="治具软件hex文件"
	$oWorkbook.Sheets(1).Cells(43,2).Value ="代码测试checklist"
	
	$oWorkbook.Sheets(1).Cells(45,2).Value ="试产checklist"
	
	$oWorkbook.Sheets(1).Cells(47,2).Value ="功能测试checklist"
	$oWorkbook.Sheets(1).Cells(48,2).Value ="整车联调checklist"
	$oWorkbook.Sheets(1).Cells(49,2).Value ="压力测试checklist"
	$oWorkbook.Sheets(1).Cells(50,2).Value ="真实DV checklist"

	$oWorkbook.Sheets(1).Cells(52,2).Value ="DVP-结构-用于给客户看"
	$oWorkbook.Sheets(1).Cells(53,2).Value ="DVP-电子-用于给客户看"
	$oWorkbook.Sheets(1).Cells(54,2).Value ="DV报告"
	
	$oWorkbook.Sheets(1).Cells(56,2).Value ="评审checklist"
	;产品功能要求
	$oWorkbook.Sheets(1).Cells(59,2).Value ="产品功能要求"
	$oWorkbook.Sheets(1).Cells(60,2).Value ="LIN/CAN"
	$oWorkbook.Sheets(1).Cells(61,2).Value ="记忆"
	$oWorkbook.Sheets(1).Cells(62,2).Value ="OTA"
	
	$oWorkbook.Sheets(1).Cells(64,1).Value ="1"
	$oWorkbook.Sheets(1).Cells(64,2).Value ="产品性能要求"
	$oWorkbook.Sheets(1).Cells(65,1).Value ="1.1"
	$oWorkbook.Sheets(1).Cells(65,2).Value ="充气"
	$oWorkbook.Sheets(1).Cells(66,2).Value ="充气高度"
	$oWorkbook.Sheets(1).Cells(67,2).Value ="充气速度"
	$oWorkbook.Sheets(1).Cells(68,1).Value ="1.2"
	$oWorkbook.Sheets(1).Cells(68,2).Value ="噪音"
	
	
	
	
	
	
	
	
	
	
	
	
	$oWorkbook.Sheets(1).Range('A1:A159').Interior.ColorIndex = 16   ;深灰
	$oWorkbook.Sheets(1).Range('A1:A159').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('B1:B159').Interior.ColorIndex = 16   ;深灰
	$oWorkbook.Sheets(1).Range('B1:B159').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('C7:Z56').Interior.ColorIndex = 15   ;浅灰
	$oWorkbook.Sheets(1).Range('C7:Z56').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('A7:A56').Font.Size=14   ;指定范围 字体大小
	$oWorkbook.Sheets(1).Range("B1:B99").Columns.AutoFit   ;指定列 自适应长度
	MsgBox(0,0,'创建项目完成！')
	_Excel_Close($oExcel)

EndFunc

#cs
#include <Excel.au3>
;FileOpen('C:\Users\Administrator\Desktop\2.xlsx',2+8)
Local $oExcel = _Excel_Open()
Local $sWorkbook = "C:\Users\Administrator\Desktop\2.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, Default, Default, True)

Local $oExcel = _Excel_Open()
Local $oWorkbook = _Excel_BookNew($oExcel)
_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\1\card.xlsx')   ;看起来自动创建xlsx文件已经解决   12/12  18:18
_Excel_Close($oExcel)
#ce
