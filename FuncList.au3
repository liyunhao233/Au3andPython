#cs
	��ʼ����
	���ļ�Ϊ		RELEASED --->>ʩ���� ������ť��������	�ű��ļ�
	RELEASED --->>ʩ����ʱ���Ȱ�RELEASED�ļ����еĹ�ϣֵ��¼��������Card.xlsx�У����뵽�̶��õ�ĳһ�л�һ�У���Ҫ����
	�������������Ҫ���õĺ���
	
#ce
#include <Array.au3>
#include <Date.au3>

Global $hasharray[1]=[""]    ;���hashֵ��ȫ������

Global $FileNum=37 ;��ǰ���ļ���������ļ�������

Global $NewDir
Func StartUpdate($searchdir)  			;��ʼ����
FileChangeDir($searchdir)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN   ;202211241406   ��ʱ�������Ϊȫ�ֱ���  Ŀǰ��������̫�У������ǻ����������ַ��ͻ
$search = FileFindFirstFile($searchdir & "\*.*")     
 If $search = -1 Then return -1                   
 While 1
 $file = FileFindNextFile($search)         ;������һ���ļ�
 If @error Then                                   
  FileClose($search)                           
  return                                                   
EndIf
;$count=StringLeft($file,3) 
$mid_file=StringSplit($file,".")  			;;;;;;;;;;;;;;;;;;;;;wait to modify
If @error=1 Then 
	;MsgBox(32,"���ļ���",$file)  		;;;;;;;; �����ļ���     ��ʱ�Ȳ�����      ���ļ���
	;$DirName1=StringTrimRight($searchdir,9)
	;FileChangeDir($DirName1)
		$f=StringLeft($file,4)				;;�ļ��д�����8��
		$f=StringRight($f,3)
		$f+=1
		$f=StringFormat("%03d",$f)	;��ʽ��һ��
		$file2=StringReplace($file,2,$f)    ;;;;;;;;;;���ļ������ֵ��޸�    ���滻�𣿣�����
		;MsgBox(0,"replace",$file2)
		;��ʱ���
		$file2=StringReplace($file2,6,$timestamp)
		;MsgBox(0,"replace-timestamp",$file2)
		$file2=StringTrimRight($file2,9)
		;MsgBox(0,"replace--delet",$file2)
	DirCreate($NewDir &"\"&$file2&"-ʩ����")
	DirCopy($searchdir&"\"&$file, $NewDir &"\"&$file2&"-ʩ����",1)
	;MsgBox(32,"�������Ƶ��ļ���",$searchdir&"\"&$file)
	;MsgBox(32,"�ļ����Ѹ���",$NewDir &"\"&$file2&"-ʩ����")
Else
		;$file2 = $mid_file[1]&"_ʩ����"&"."&$mid_file[2]
		
						;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;���ļ����ֵ��޸�   �����滻���� 13��10 ���޸�  v+�汾��+���+����+��׺   13�� 11��20�޸�
						;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;11/24 �ѱ��    A0x+v+�汾��+ʱ���+���+����+��׺
		$mid_file[1]=StringTrimRight($mid_file[1],9)  ;ɾreleased
		;MsgBox(0,"replace--delet",$mid_file[1])
		$file2 = $mid_file[1]&"_ʩ����"&"."&$mid_file[2]
		$f=StringLeft($file,4)
		$f=StringRight($f,3)
		$f+=1
		$f=StringFormat("%03d",$f)	
		$file2=StringReplace($file2,2,$f)
		$file2=StringReplace($file2,6,$timestamp)     
		;MsgBox(0,"aa",$file2)
		FileCopy($searchdir & "\" & $file, $NewDir&"\"& $file2)
EndIf

WEnd
EndFunc

Func CopyDir($DirName)		;�ļ��еĸ���  �� ����
$DirName=StringTrimLeft($DirName,2)	;12/7 �޸�·��    ������ǰ�����˫б��
$array=StringSplit($DirName,"\")
$b=$array[0]
$file=$array[$b]
$b=StringTrimRight($DirName,StringLen($file))

$v=StringLeft($file,4)   ;�и���ַ�������
$v=StringRight($v,3)
$v+=1
$v=StringFormat("%03d",$v)
;MsgBox(0,"",$v)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
$file=StringReplace($file,2,$v)
$file=StringReplace($file,6,$timestamp)   
$file=StringTrimRight($file,9)
$filename=$file&'-ʩ����'
;MsgBox(0,"",$filename)
 $NewDir='\\'&$b&$filename   ;���·�������� \\
DirCreate($NewDir)
;MsgBox(32,"��ͷ���ļ�","���ļ����Ѹ���")
EndFunc


Func CheckFile($FilePath)   							;У���ļ���  �ļ���  �ļ�����
	$FilePath=StringTrimLeft($FilePath,2)																;12/7 �޸ĵ��������У�·����Ҫ����
	;MsgBox(64,"�ļ�У��","����У���ļ�������")
	$array=StringSplit($FilePath,"\")
	$a=Int($array[0])

	$Check=StringInStr($array[$a],"RELEASED",0)		; !!!!!!!!!!!!!  fix  error

	;MsgBox(0,"cccc",$Check)
	If $Check = 0 Then											; Check Main   Dir
		MsgBox(16,"error","�����ļ�������,�����˳�!")
		Exit
		EndIf
	
	;MsgBox(0,"��ȡ�ļ�������",$array[$a])
	$version=StringLeft($array[$a],4)  ;����߿�ʼȡ4��
	$version=StringRight($version,3)
	;MsgBox(0,"�汾��",$version) 
	$file = FileFindFirstFile("\\"&$FilePath&"\*.*")  
	;MsgBox(0,"Ѱ���ļ�",$file) 
	If $file  = -1 Then MsgBox(16,"error","�ļ��ж�û��ȥ")
	$num =0
	While 1
		$file1 = FileFindNextFile($file) 
		If @error Then ExitLoop
		;If @extended Then 
		$v=StringLeft($file1 ,4)
		$v=StringRight($v,3)
		If Not($version = $v) Then ExitLoop
		$num+=1
		;MsgBox(1,"debug",$file1)
			;ContinueLoop
        ;EndIf
	WEnd
	FileClose($file)
	MsgBox(0,"�ܹ��ļ�����",$num)
	
	if not($num=37)Then       ;������ $FileNum  ��������������
		MsgBox(16,"error","�ļ�������汾�Ŵ��󣬳����˳���")
		Exit ;�˳�ȫ���ű�
	Elseif $num=37 Then
		$excelname=$file  
		MsgBox(0,"�ļ�У��","������ȷ��������һ������" )
		FileSetAttrib('\\'&$FilePath,"-R",1)  
		;MsgBox(0,"�����ļ��ѽ���" ,'\\'&$FilePath)
	EndIf
EndFunc


Func AddRow($FilePath)    								;����֮��  ����Ӧ��λ�ý�������һ��
	$search1 = FileFindFirstFile($FilePath & "\*.*")   ;�����ļ��У���ȡxlsx�����ļ�
	While 1 
		$file = FileFindNextFile($search1)
		If @error Then ExitLoop
		if Not StringInStr($file,'card')=0 Then   ;�ҵ�Card�ļ���11/21 15:30 ��ʵ��
			$excelname=$file  
		EndIf
WEnd
		$excelname1=$FilePath&"\"&$excelname
		$Split=StringLeft($excelname,4)
		$Split=StringRight($Split,3)    ;��ȡ��Ҫ������,����λ������һЩ
		;MsgBox(0,"split",$Split)
		$Split1= Int($Split)
		$Split1=$Split1+4   		 ;�����ٴ˴��޸�λ����
		;MsgBox(0,"",$Split1)
		$oExcel = _Excel_Open(False, True, Default, Default, True)
		$oWorkbook=_Excel_BookOpen($oExcel,$excelname1)
		
		;MsgBox(0,0,'open excel')
		;$oWorkbook.Sheets(1).Range("F3").Interior.ColorIndex = 3		  ;;;;;;;������ʾ
		$oWorkbook.ActiveSheet.unProtect("a") 						;�Ѿ�����һ��  ���ӵ����н�����������Ϊ����״̬
		$oWorkbook.Sheets(1).Cells(3,$Split1).Value ="V:"&$Split1-3
		$oWorkbook.Sheets(1).Cells(4,$Split1).Value = @MON&@MDAY&'|'&@HOUR&':'&@MIN             ; TimeStamp
		
		trans($Split1)
		$col=$end
		trans($Split1-1)
		$col_pre=$end
		
		$oExcel.Range($col_pre&':'&$col_pre).Locked = True
		$oExcel.Range($col&':'&$col).Locked = False						;1���Ƚ�������  2�����ض�����Ϊ����״̬ 3������Ϊ����
		$oWorkbook.ActiveSheet.Protect("a") 
		
		
MsgBox(64,"INFO","�¼��� " & $col & " ��������")

_Excel_Close($oExcel)

EndFunc  

Global $end
Func trans($num)			;����ת������ĸ����Ϊ  ��Ҫ��excel.range���������ض�����
	$c=Mod($num,26)
	If $num <=26 Then
	$l=$num+64
	$end=Chr($l)
	;MsgBox(0,"С��26",$end)
	ElseIf $num >26 And Not($c=0) Then   ;��Ҫ������ת��ĸ�߼�
	$d=$num/26
	;MsgBox(0,"����1",$d)
	$e=int($d)
	$e+=64
	$l=$c+64
	;MsgBox(0,"��Ӧ��ĸ",Chr($l))
	$end=Chr($e)&Chr($l)
	;MsgBox(0,"����",$end)
	;$e+=64
	;$e=Chr($e)
Else  
		$d=$num/26
		;MsgBox(0,"����2",$d)
		$e=int($d)
		$e-=1
		$e+=64	
		;MsgBox(0,"��һ��",$e)		
		$c="Z"
		$end=Chr($e)&$c
		;MsgBox(0,"����2",$end)
EndIf
EndFunc


#cs
	1������ΪR----��ʩ
	2��ֻ��Ҫ������Ϊʩ����֮ǰ��released�е������ļ�hashֵ�浽card�ļ��м���

#ce

Global $count   ;�ӱ�����
Global $dircount=0
Global $hasharray3[1]=[""]


Func GetHash($searchdir)      ;��ȡ�ļ��Ĺ�ϣֵ
	FindExcel($searchdir)
	FileChangeDir($searchdir)      ;һ��Ҫת������Ӧ���ļ�·��
	$search = FileFindFirstFile($searchdir & "\*.*")
	If $search = -1 Then return -111 
	While 1
	$file = FileFindNextFile($search)         ;������һ���ļ�
	If @error Then
		FileClose($search)
		;MsgBox(0,0,"look input long hash")
		$oExcel = _Excel_Open(False, True, Default, Default, True)
		$oWorkbook=_Excel_BookOpen($oExcel,$excelname1)
		$oWorkbook.Sheets(1).unProtect("a") 
		For $i=1 To UBound($hasharray3)-1
			$oWorkbook.Sheets(1).Cells($i+290,2).Value =$hasharray3[$i]
		Next
		;_ArrayDisplay($hasharray3)
		_Excel_Close($oExcel)
		return
	;EndIf
	
	ElseIf @extended=1 Then   
		;MsgBox(0,"Dir",$file)
		$dircount=$dircount+1     ;���������ļ��еļ�����
		GetDirHash($searchdir,$file)   ;��ȡ���ļ�������Ĺ�ϣֵ��������浽excel�������
		
	
	ElseIf  (StringInStr($file,'Card')=0) Then   ;���ֿյ����ݵ������������  �ж�������ͻ   �б�����0�±�
		$sData = _MD5ForFile1($file)
		_ArrayAdd($hasharray3,$sData)
		ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
		$count=$count+1
	EndIf
	;ConsoleWrite($file&@CRLF)
	WEnd
	
EndFunc


Func GetDirHash($searchdir,$file)   ;��ȡ���ļ������ļ��Ĺ�ϣֵ���ұ��浽������   �˴����ļ�������Ҫ����Ҫ���涨
	$a=$searchdir &"\"&$file
	$s = FileFindFirstFile($a& "\*.*")     
	;MsgBox(0,'get dir ',$s)  
	Global $Tarray[1]=[""]   ;���ļ��й�ϣ�����������ʱ���飬Ϊ��װ��ÿ����̫���ӱ��С�
	$count=0
	While 1
		$f = FileFindNextFile($s)  
		If @error Then                                   
			ExitLoop
		EndIf
	;MsgBox(0,0,$f)
		;$excelname1=$searchdir&"\"&$excelname	
		$sData1 = _MD5ForFile1($searchdir&"\"&$file&"\"&$f)
		_ArrayAdd($Tarray,$sData1)
		$count=$count+1
		ConsoleWrite("Resultmmm: " & $sData1 & @CRLF & @CRLF)
	WEnd
		AddToNewSheet($count,$Tarray)
EndFunc

Func FindExcel($searchdir)
$search = FileFindFirstFile($searchdir & "\*.*")
If $search = -1 Then return -1 
While 1
$file = FileFindNextFile($search)         ;������һ���ļ�
If @error Then                                   
FileClose($search)  
ExitLoop
EndIf
if Not (StringInStr($file,"card") = 0) Then
	Global $excelname1=$file
	;MsgBox(0,0,$excelname1)
	EndIf
WEnd
$excelname1=$searchdir&"\"&$excelname1
;MsgBox(0,'excelname',$excelname1)
EndFunc

Func AddToNewSheet($count,$Tarray)  ;���½����ı���  �����ϣֵ   �˴��߼���̫��ȷ��ÿ�ζ����±��ǲ��Եģ���Ҫȷ�����ӱ�����
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname1) ;·��������
	;MsgBox(0,2,$excelname1)
	Local $aWorkSheets = _Excel_SheetList($oWorkbook)
	Local $sheetnum=0
	For $IDX = 1 To UBound($aWorkSheets) -1
			$sheetnum=$sheetnum+1
		Next
	If $sheetnum<7 Then
	_Excel_SheetAdd($oWorkbook,Default,False)    ;����м�sheet�ӱ�� �������Ĭ���ӱ���������ڵڶ���λ��
	EndIf
	;MsgBox(0,3,'test')											;�������һ��ʼ��Ҫ�趨���ж����ļ��У����ٸ��ӱ��ÿ���ӱ�����ȷ���أ���
	For $i=1 To Number($count)
	$oWorkbook.Sheets($dircount+1).Cells($i+1,2).Value =$Tarray[$i]
	;MsgBox(0,0,$Tarray[$i])
	Next
	_Excel_Close($oExcel)
EndFunc



Func SendToPython2($Readpath,$Code)

	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,9)
	$path=$version&'_'&$time&'_'&'A13Calc_Released'
	
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0209ToUp\0209ToUp.exe",$Excelpath&'>'&$FileName&'>'&$Code)
	
EndFunc 





;����Ϊ���������ļ�MD5��ϣֵ
Func _MD5ForFile1($sFile)

    Local $a_hCall = DllCall("kernel32.dll", "hwnd", "CreateFileW", _                       ;DllCall  ����windowsϵͳ��   
            "wstr", $sFile, _
            "dword", 0x80000000, _ ; GENERIC_READ
            "dword", 3, _ ; FILE_SHARE_READ|FILE_SHARE_WRITE
            "ptr", 0, _
            "dword", 3, _ ; OPEN_EXISTING
            "dword", 0, _ ; SECURITY_ANONYMOUS
            "ptr", 0)

    If @error Or $a_hCall[0] = -1 Then
        Return SetError(1, 0, "")
    EndIf

    Local $hFile = $a_hCall[0]       ;�ŵ�$hFile

    $a_hCall = DllCall("kernel32.dll", "ptr", "CreateFileMappingW", _
            "hwnd", $hFile, _   ;$hFile ���õ��µ�DllCall 
            "dword", 0, _ ; default security descriptor
            "dword", 2, _ ; PAGE_READONLY
            "dword", 0, _
            "dword", 0, _
            "ptr", 0)

    If @error Or Not $a_hCall[0] Then
        DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)
        Return SetError(2, 0, "")
    EndIf

    DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)

    Local $hFileMappingObject = $a_hCall[0]   ;�ŵ�$hFileMappingObject 

    $a_hCall = DllCall("kernel32.dll", "ptr", "MapViewOfFile", _
            "hwnd", $hFileMappingObject, _
            "dword", 4, _ ; FILE_MAP_READ
            "dword", 0, _
            "dword", 0, _
            "dword", 0)

    If @error Or Not $a_hCall[0] Then
        DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
        Return SetError(3, 0, "")
    EndIf

    Local $pFile = $a_hCall[0]
    Local $iBufferSize = FileGetSize($sFile)

    Local $tMD5_CTX = DllStructCreate("dword i[2];" & _  ;$tMD5_CTX  �����
            "dword buf[4];" & _
            "ubyte in[64];" & _
            "ubyte digest[16]")

    DllCall("advapi32.dll", "none", "MD5Init", "ptr", DllStructGetPtr($tMD5_CTX))

    If @error Then
        DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
        DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
        Return SetError(4, 0, "")
    EndIf

    DllCall("advapi32.dll", "none", "MD5Update", _
            "ptr", DllStructGetPtr($tMD5_CTX), _
            "ptr", $pFile, _
            "dword", $iBufferSize)

    If @error Then
        DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
        DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
        Return SetError(5, 0, "")
    EndIf

    DllCall("advapi32.dll", "none", "MD5Final", "ptr", DllStructGetPtr($tMD5_CTX))

    If @error Then
        DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
        DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
        Return SetError(6, 0, "")
    EndIf

    DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
    DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)

    Local $sMD5 = Hex(DllStructGetData($tMD5_CTX, "digest"))

    Return SetError(0, 0, $sMD5)

EndFunc   ;==>_MD5ForFile















