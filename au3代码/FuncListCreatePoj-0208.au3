#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>

Global $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
ShowGui($timestamp)

;ͨ������һ�������ڹ����������xlsx�ļ���   ���裺1���ȸ��ƣ��ٴ򿪡�2���򿪺���������Ϣ����
Func ShowGui($timestamp)
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
$C='����'
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
$B='���'
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
		$B='���'	
	Case $msg[0] = $Bb
		$B='�Ҽ�'
	Case $msg[0] = $Ca
		$C='����'
	Case $msg[0] = $Cb
		$C='����'
		
	Case $msg[0] = $BT1
	If(StringIsAlNum(GUICtrlRead($car))=0 Or StringIsAlNum(GUICtrlRead($shape))=0 Or StringIsAlNum(GUICtrlRead($chair))=0 ) Then
			MsgBox(16,'error','������Ϣδ��д��')
	
	Else
			Showinfo($car,$chair,$shape,$A,$B,$C,$timestamp) ;�������ӡ����͡��š����Ҽݡ�оƬ����
			;CreateFile($car,$chair,$shape,$A,$B,$C)
			CopyExcel($car,$chair,$shape,$A,$B,$C,$timestamp) ;������-���γ�-����-оƬ-���Ҽ�-������
			
			;;��ȡ������Ŀ�е���Ϣ����������API����һ���ļ���
			;CreatePoj($car,$chair,$shape)    ;��ʽ��Greely_YF_CS1E_v001
			GUISetState(@SW_DISABLE,$GUI)
			GUISetState(@SW_HIDE,$GUI)
			;GUISetState(@SW_SHOW,$hMainGUI)
			
			ExitLoop
	EndIf
	Case $msg[0] =$GUI_EVENT_CLOSE And $msg[1] = $GUI
			Local $check=MsgBox(1, "INFO", "�Ƿ��˳���")
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
;����python ��֪ʶ�ռ��д��������Ŀ��
$code=$cmdline[1]    ;��code��python
;MsgBox(0,0,$code)   ;��ʼ����code��ʽ
$CodeAndId=StringTrimLeft($code,7)
;MsgBox(0,3,$CodeAndId)

$xlsxpath='\\192.168.2.100\100\�汾����\'&$car1&'_'&$chair1&'_'&$shape1&'_'&$C&$A&'_'&$B&'\'&'V000'&'-'&$timestamp&'_ʩ����'&'\'&'V000_'&$timestamp&'_A13Calc_ʩ����'&'.xlsx'
ShellExecute("\\192.168.2.100\100\0208CreatePoj\0208CreatePoj.exe",$car1&'_'&$chair1&'_'&$shape1&'_'&$A&'_'&$B&'_'&$C&'>'&$xlsxpath&'>'&$CodeAndId)  


EndFunc


Func Showinfo($aa,$bb,$cc,$d,$e,$f,$timestamp)	;������-���γ�-����-оƬ-���Ҽ�-������
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	MsgBox(0,'������ʼ��Ŀ',GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e)
	DirCreate('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����')
	
	
	Local $oExcel = _Excel_Open(False, True, Default, Default, True)
	Local $oWorkbook = _Excel_BookNew($oExcel)
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����'&'\'&'V000_'&$timestamp&'_A13Calc_ʩ����'&'.xlsx')   ;�������Զ�����xlsx�ļ��Ѿ����   12/12  18:18
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����'&'\'&'V000_'&$timestamp&'_A10BOM_ʩ����'&'.xlsx')   
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����'&'\'&'V000_'&$timestamp&'_A12BOMtoERP_ʩ����'&'.xlsx')  
	_Excel_Close($oExcel)
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	
EndFunc


Func CopyExcel($aa,$bb,$cc,$d,$e,$f,$timestamp) ;������-���γ�-����-оƬ-���Ҽ�-������
	;Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	FileChangeDir('\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A01_�ͻ���׼'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A02_ͨ��Э��'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A30_3D��ģ'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A32_2D�ܳ��ڲ�'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A33_2D�ܳɿͻ�'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A50_���Դ����'&'_ʩ����')
	DirCreate('V000_'&$timestamp&'_'&'A51_�ξ��������'&'_ʩ����')
	;MsgBox(0,0,'�ļ���')
	FileOpen('V000_'&$timestamp&'_'&'A09_ECR˵��'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A11_BOM�����м��'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B30_������Ӧ����ģ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C30_�ܹ��ͷŸ��ͻ���3D��ģ������ȥ���ڲ��ṹ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C20_�����ͻ���TRpdf'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A31_stp/igs��ʽ��ģ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C33_�ܹ��ͷŸ��ͻ���ͼֽ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A34_����ͼֽ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A35_CC/SC'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A36_DFMEA'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A40_Pads����'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B40_���ַ�ʽ����Ӧ��'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A41_SMTԪ���������ļ�'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A42_SMTԪ������'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A43_eICD(��װͼ������ԭ��ͼ)'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'C43_�ͷŸ��ͻ�ͼֽ'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A52_��ƷHex'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'B52_�ú��ַ�ʽ��������'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A53_�ξ����hex'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A54_�������checklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A60_�Բ�checklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A70_���ܲ���checklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A71_��������checklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A72_ѹ������checklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A73_��ʵDVchecklist'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A80_DVP-�ṹ-���ͻ���'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A81_DVP-����-���ͻ���'&'_ʩ����.txt',2)
	FileOpen('V000_'&$timestamp&'_'&'A82_DV����'&'_ʩ����.txt',2)
	
	
	$path='\\192.168.2.100\100\�汾����\ģ���ļ�\A00 Draft 230105.xlsx'
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$path)
	$oWorkbook.Sheets(1).Cells(1,2).Value =GUICtrlRead($aa)
	$oWorkbook.Sheets(1).Cells(1,5).Value =GUICtrlRead($bb)
	$oWorkbook.Sheets(1).Cells(1,8).Value =GUICtrlRead($cc)
	$oWorkbook.Sheets(1).Cells(1,11).Value =$f
	$oWorkbook.Sheets(1).Cells(1,14).Value =$e&$d
	_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&GUICtrlRead($aa)&'_'&GUICtrlRead($bb)&'_'&GUICtrlRead($cc)&'_'&$f&$d&'_'&$e&'\'&'V000'&'-'&$timestamp&'_ʩ����'&'\'&'V000_'&$timestamp&'_A00Card_ʩ����'&'.xlsx')
	;_Excel_BookSaveAs($oWorkbook,'\\192.168.2.100\100\�汾����\'&'_A00Card_RELEASED'&'.xlsx')
	MsgBox(0,0,'����������')
	_Excel_Close($oExcel)
	
EndFunc












