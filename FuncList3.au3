#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
;Show()
Func Show()
$GUI=GUICreate("������Ŀ��Ϣ",450,300)
$BT1=GUICtrlCreateButton("ȷ��",170,250,100,20)

$car=GUICtrlCreateInput("",60,20,200,25)
$chair=GUICtrlCreateInput("",60,55,200,25)
$shape=GUICtrlCreateInput("",60,90,200,25)

GUICtrlCreateLabel("������",10,25,40,30)
GUICtrlCreateLabel("���γ�",10,60,50,30)
GUICtrlCreateLabel("����",10,95,40,30)


GUICtrlCreateLabel("A/B ������оƬ��Դ��A:����оƬ B������оƬ		01/02/03/04:�������ε�λ�ã�01��ǰ��02��ǰ�ң�03������04������",270,60,100,100)
GUISetState(@SW_ENABLE,$GUI)
GUISetState(@SW_SHOW,$GUI)
GUICtrlCreateGroup("оƬ����", 30, 130, 180, 50)
$Ca=GUICtrlCreateRadio("A:����",50,145,50,30)
$Cb=GUICtrlCreateRadio("B:����",130,145,50,30)
GUICtrlSetState($Ca, $GUI_CHECKED)
$C='A'
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("1/2/3��", 30, 190, 200, 50)
$Aa=GUICtrlCreateRadio("1��",50,200,50,30)
$Ab=GUICtrlCreateRadio("2��",100,200,50,30)
$Ac=GUICtrlCreateRadio("3��",150,200,50,30)
GUICtrlSetState($Aa, $GUI_CHECKED)
$A='01'
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("��/�Ҽ�", 200, 190, 180, 50)
$Ba=GUICtrlCreateRadio("���",250,200,50,30)
$Bb=GUICtrlCreateRadio("�Ҽ�",300,200,50,30)
GUICtrlSetState($Ba, $GUI_CHECKED)
$B='��'
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
		$B='��'	
	Case $msg[0] = $Bb
		$B='��'
	Case $msg[0] = $Ca
		$C='A'
	Case $msg[0] = $Cb
		$C='B'
		
	Case $msg[0] = $BT1
	If(StringIsAlNum(GUICtrlRead($car))=0 Or StringIsAlNum(GUICtrlRead($shape))=0 Or StringIsAlNum(GUICtrlRead($chair))=0 ) Then
			MsgBox(32,'test','������Ϣδ��д��')
	
	Else
			Showinfo($car,$chair,$shape,$A,$B,$C) ;�������ӡ����͡��š����Ҽݡ�оƬ����
			CreateFile($car,$chair,$shape,$A,$B,$C)
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			;GUISetState(@SW_SHOW,$hMainGUI)
			
			ExitLoop
	EndIf
	Case $msg[0] =$GUI_EVENT_CLOSE And $msg[1] = $GUI
			MsgBox(32, "INFO", "�Ƿ��˳���")
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			GUISetState(@SW_SHOW,$hMainGUI)
			ExitLoop
		EndSelect
WEnd

EndFunc



Func Showinfo($aa,$bb,$cc,$d,$e,$f)	;������-���γ�-����-оƬ-���Ҽ�-������
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	MsgBox(0,'������ʼ��Ŀ',GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
EndFunc


Func CreateFile($aa,$bb,$cc,$d,$e,$f) ;�������ӡ����͡��š����Ҽݡ�оƬ����
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	FileChangeDir('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A01_�ͻ���׼'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A02_ͨ��Э��'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A30_3D��ģ'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A32_2D�ܳ��ڲ�'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A33_2D�ܳɿͻ�'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A50_���Դ����'&'-RELEASED')
	DirCreate('V000_'&$timestamp&'_'&'A51_�ξ��������'&'-RELEASED')
	;MsgBox(0,0,'�ļ���')
	FileOpen('V000_'&$timestamp&'_'&'A09_ECR˵��'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A11_BOM�����м��'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B30_������Ӧ����ģ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C30_�ܹ��ͷŸ��ͻ���3D��ģ������ȥ���ڲ��ṹ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C20_�����ͻ���TRpdf'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A31_stp/igs��ʽ��ģ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C33_�ܹ��ͷŸ��ͻ���ͼֽ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A34_����ͼֽ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A35_CC/SC'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A36_DFMEA'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A40_Pads����'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B40_���ַ�ʽ����Ӧ��'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A41_SMTԪ���������ļ�'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A42_SMTԪ������'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A43_eICD(��װͼ������ԭ��ͼ)'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C43_�ͷŸ��ͻ�ͼֽ'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A52_��ƷHex'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B52_�ú��ַ�ʽ��������'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A53_�ξ����hex'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A54_�������checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A60_�Բ�checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A70_���ܲ���checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A71_��������checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A72_ѹ������checklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A73_��ʵDVchecklist'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A80_DVP-�ṹ-���ͻ���'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A81_DVP-����-���ͻ���'&'-RELEASED.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A82_DV����'&'-RELEASED.txt',2)
	

	;MsgBox(0,0,'txt�����ļ�')
	Local $oExcel = _Excel_Open(Default, True, Default, Default, True)
	Local $oWorkbook = _Excel_BookNew($oExcel)
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A13Calc-RELEASED'&'.xlsx')   ;�������Զ�����xlsx�ļ��Ѿ����   12/12  18:18
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A10BOM-RELEASED'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A12BOMtoERP-RELEASED'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED'&'\'&'V000_'&$timestamp&'_A00Card-RELEASED'&'.xlsx')   
	
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
	
	;Local $oWorkbook = _Excel_BookOpen('V000_'&$timestamp&'_A00Card-RELEASED'&'.xlsx','\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'-RELEASED')
	;���������������
	$oWorkbook.Sheets(1).Cells(1,2).Value =GUICtrlRead($aa)
	$oWorkbook.Sheets(1).Cells(1,4).Value =GUICtrlRead($bb)
	$oWorkbook.Sheets(1).Cells(1,6).Value =GUICtrlRead($cc)
	$oWorkbook.Sheets(1).Cells(1,8).Value =$f
	$oWorkbook.Sheets(1).Cells(1,10).Value =$e
	$oWorkbook.Sheets(1).Cells(1,12).Value =$d
	;MsgBox(32,'a','stop')
	$oWorkbook.Sheets(1).Cells(1,1).Value ="������"
	$oWorkbook.Sheets(1).Cells(1,3).Value ="���γ�"
	$oWorkbook.Sheets(1).Cells(1,5).Value ="����"
	$oWorkbook.Sheets(1).Cells(1,7).Value ="оƬ����"
	$oWorkbook.Sheets(1).Cells(1,9).Value ="��/�Ҽ�"
	$oWorkbook.Sheets(1).Cells(1,11).Value ="һ/��/����"
	$oWorkbook.Sheets(1).Cells(3,3).Value ="Ĭ��ֵ"
	;�����ʼ���ĳ�ʼ����
	$oWorkbook.Sheets(1).Cells(4,1).Value ="����ʱ��"
	$oWorkbook.Sheets(1).Cells(6,1).Value ="��������"  
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


	$oWorkbook.Sheets(1).Cells(6,2).Value ="��Ʒ���Կ�"
	$oWorkbook.Sheets(1).Cells(7,2).Value ="��Ʒ����/����/�������/�����������Ŀ�ܱ�"
	$oWorkbook.Sheets(1).Cells(8,2).Value ="�ͻ�����ļ����淶/������/��׼"
	$oWorkbook.Sheets(1).Cells(9,2).Value ="ͨ��Э�顢LIN/CAN MATRIX"
	$oWorkbook.Sheets(1).Cells(10,2).Value ="ECR˵��"
	
	$oWorkbook.Sheets(1).Cells(12,2).Value ="BOM - �����ܷ��㿴�ĸ�ʽ"
	$oWorkbook.Sheets(1).Cells(13,2).Value ="BOM�����м��ļ�"
	$oWorkbook.Sheets(1).Cells(14,2).Value ="�ܹ�ֱ�ӵ���ERP��BOM��ʽ"
	$oWorkbook.Sheets(1).Cells(15,2).Value ="Ƥ�����������"
	
	$oWorkbook.Sheets(1).Cells(17,2).Value ="��������"
	$oWorkbook.Sheets(1).Cells(18,2).Value ="�����ͻ����TR pdf"
	
	$oWorkbook.Sheets(1).Cells(20,2).Value ="3D��ģ"
	$oWorkbook.Sheets(1).Cells(21,2).Value ="��Ҫ������Ӧ�̵���ģ"
	$oWorkbook.Sheets(1).Cells(22,2).Value ="�ܹ��ͷŸ��ͻ���3D��ģ������ȥ���ڲ��ṹ"
	$oWorkbook.Sheets(1).Cells(23,2).Value ="stp/igs��ʽ��ģ"
	$oWorkbook.Sheets(1).Cells(24,2).Value ="2D�ܳ�ͼ - �ڲ�"
	$oWorkbook.Sheets(1).Cells(25,2).Value ="2D�ܳ�ͼ - ���ͻ���"
	$oWorkbook.Sheets(1).Cells(26,2).Value ="�ܹ��ͷŸ��ͻ���ͼֽ"
	$oWorkbook.Sheets(1).Cells(27,2).Value ="����ͼֽ"
	$oWorkbook.Sheets(1).Cells(28,2).Value ="CC/SC"
	$oWorkbook.Sheets(1).Cells(29,2).Value ="DFMEA"
	
	$oWorkbook.Sheets(1).Cells(31,2).Value ="Pads������ƵĹ�������(PCB/PCBA/ԭ��ͼ)"
	$oWorkbook.Sheets(1).Cells(32,2).Value ="ʲô��ʽ������Ӧ��"
	$oWorkbook.Sheets(1).Cells(33,2).Value ="SMTԪ���������ļ�"
	$oWorkbook.Sheets(1).Cells(34,2).Value ="SMTԪ������(��Ҫ�����ָ�) - Ҫ��BOM��Ӧ"
	$oWorkbook.Sheets(1).Cells(35,2).Value ="eICD(��װͼ������ԭ��ͼ)"
	$oWorkbook.Sheets(1).Cells(36,2).Value ="�ܹ��ͷŸ��ͻ���ͼֽ"
	
	$oWorkbook.Sheets(1).Cells(38,2).Value ="���Դ����"
	$oWorkbook.Sheets(1).Cells(39,2).Value ="�ξ��������"
	$oWorkbook.Sheets(1).Cells(40,2).Value ="��Ʒhex�ļ�"
	$oWorkbook.Sheets(1).Cells(41,2).Value ="ʲô��ʽ���͸�����?"
	$oWorkbook.Sheets(1).Cells(42,2).Value ="�ξ����hex�ļ�"
	$oWorkbook.Sheets(1).Cells(43,2).Value ="�������checklist"
	
	$oWorkbook.Sheets(1).Cells(45,2).Value ="�Բ�checklist"
	
	$oWorkbook.Sheets(1).Cells(47,2).Value ="���ܲ���checklist"
	$oWorkbook.Sheets(1).Cells(48,2).Value ="��������checklist"
	$oWorkbook.Sheets(1).Cells(49,2).Value ="ѹ������checklist"
	$oWorkbook.Sheets(1).Cells(50,2).Value ="��ʵDV checklist"

	$oWorkbook.Sheets(1).Cells(52,2).Value ="DVP-�ṹ-���ڸ��ͻ���"
	$oWorkbook.Sheets(1).Cells(53,2).Value ="DVP-����-���ڸ��ͻ���"
	$oWorkbook.Sheets(1).Cells(54,2).Value ="DV����"
	
	$oWorkbook.Sheets(1).Cells(56,2).Value ="����checklist"
	;��Ʒ����Ҫ��
	$oWorkbook.Sheets(1).Cells(59,2).Value ="��Ʒ����Ҫ��"
	$oWorkbook.Sheets(1).Cells(60,2).Value ="LIN/CAN"
	$oWorkbook.Sheets(1).Cells(61,2).Value ="����"
	$oWorkbook.Sheets(1).Cells(62,2).Value ="OTA"
	
	$oWorkbook.Sheets(1).Cells(64,1).Value ="1"
	$oWorkbook.Sheets(1).Cells(64,2).Value ="��Ʒ����Ҫ��"
	$oWorkbook.Sheets(1).Cells(65,1).Value ="1.1"
	$oWorkbook.Sheets(1).Cells(65,2).Value ="����"
	$oWorkbook.Sheets(1).Cells(66,2).Value ="�����߶�"
	$oWorkbook.Sheets(1).Cells(67,2).Value ="�����ٶ�"
	$oWorkbook.Sheets(1).Cells(68,1).Value ="1.2"
	$oWorkbook.Sheets(1).Cells(68,2).Value ="����"
	
	
	
	
	
	
	
	
	
	
	
	
	$oWorkbook.Sheets(1).Range('A1:A159').Interior.ColorIndex = 16   ;���
	$oWorkbook.Sheets(1).Range('A1:A159').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('B1:B159').Interior.ColorIndex = 16   ;���
	$oWorkbook.Sheets(1).Range('B1:B159').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('C7:Z56').Interior.ColorIndex = 15   ;ǳ��
	$oWorkbook.Sheets(1).Range('C7:Z56').Borders.ColorIndex = 1
	
	$oWorkbook.Sheets(1).Range('A7:A56').Font.Size=14   ;ָ����Χ �����С
	$oWorkbook.Sheets(1).Range("B1:B99").Columns.AutoFit   ;ָ���� ����Ӧ����
	MsgBox(0,0,'������Ŀ��ɣ�')
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
_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\1\card.xlsx')   ;�������Զ�����xlsx�ļ��Ѿ����   12/12  18:18
_Excel_Close($oExcel)
#ce
