#cs
	ȷ�Ϸ���
	���ļ�Ϊ	 ʩ����-->>RELEASED- ȷ�Ϸ����������ܽű��ļ�
	
	�������������Ҫ���õĺ���
	
#ce

Global $FileNum =37  ;��ǰ���ļ���������ļ�������
Global $NewDir
Global $oldarray[1]=[""]
Global $hasharray2[1]=[""]
Func Released($searchdir)  	;ȷ�Ϸ���      ��д������ݷ�ʽ�Լ����Ƹı��˵��ļ�����ʱ���������ļ�������ģ�12/21    16��00  �Ѱ������ļ��� 12/23
	Local $page=1
	FindExcel2($searchdir)
	;��Ϊexcel�еľ�hash�Լ��»�ȡ��hash�����ж�
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname2)   ; xlsl�ļ���$excelname2
	$oldbigdata=_Excel_RangeRead($oWorkbook,1,'B291:B319')      ;��������������   �����hash������û��  �˷�Χ�������ܻ���Ҫ�仯
	_ArrayAdd($oldbigarray,$oldbigdata)
	;MsgBox(0,1,"show two array")
	;_ArrayDisplay($oldbigarray)
	;������ݷ�ʽһ��ʼ������Ĵ��ļ����ƻ�����lnk�����������ļ��е�ʱ��Ӧ���������е��ӱ��Աȣ�
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	$search = FileFindFirstFile($searchdir & "\*.*")     
	If $search = -1 Then return -1    ;�������ļ����֮����˳�
	$j=0	  ;������  ��ʱ�ô˼�����   Ϊ�˸�����������б���
While 1
	$file = FileFindNextFile($search)         ;������һ���ļ�
If @error Then
	FileClose($search)  
	FileSetAttrib($NewDir,"+R",1)     ;������Ϊֻ��
	MsgBox(0,"������Ϊֻ��",$NewDir)
	_Excel_Close($oExcel)
	ExitLoop
	EndIf
If @extended Then           ;�ҵ����ļ��еģ����Դ���һ�������������ļ���hash��old hash���Ƚϣ��󲿷����Ϊ���ƣ�
	;MsgBox(0,"Dir",$searchdir&'\'&$file)
	$page=$page+1  ;ÿ�ҵ�һ�����ļ��У���Ҫ����Ӧ���ӱ��Ҳ����
	Comparehash($searchdir&'\'&$file,$page,$oWorkbook)    ;1�����ļ��е�·��,
	ContinueLoop
ElseIf $j<30 Then    ;����8�����ļ���37-8=29
	;MsgBox(0,1,$file)   ;���ƻ�����lnk��ͬ��Ҳ��ͨ��hashȥ�ж�
	$mid_file=StringSplit($file,".")  
	;�и��ļ���������Щ����lnk����
	$mid_file[1]=StringTrimRight($mid_file[1],4)
	$mid_file[1]=$mid_file[1]&"_Released"
	$mid_file[1]=StringReplace($mid_file[1],6,$timestamp)
	$file2 = $mid_file[1]&"."&$mid_file[2]		
	;For $j=1 To UBound($Newarray)-1
	;MsgBox(0,888,$j)
	If($Newarray[$j]==$oldbigarray[$j]) And (StringInStr($file2,"card")=0) Then   ;;;;;�����߼�������12/16
	;MsgBox(0,"equl",$j)
	FileCreateShortcut($searchdir & "\" & $file,$NewDir&"\"& $mid_file[1]&'.lnk')    ;0105   ������ݷ�ʽ������������V000  ����λ����Ӧ��һ��Ŀ¼��
	;$j=$j+1
	;MsgBox(0,"jjjj",$j)
	ElseIf  Not ($Newarray[$j]==$oldbigarray[$j]) And (StringInStr($file2,"card")=0) Then
	;copy�߼���Ҫ����
	;MsgBox(0,'����Ҫcopy',$file2)
	FileCopy($searchdir & "\" & $file, $NewDir&"\"& $file2)
	;$j=$j+1
	Else
	FileCopy($searchdir & "\" & $file, $NewDir&"\"& $file2)
	;$j=$j+1
	EndIf
;Next
	$j=$j+1
	EndIf
WEnd
EndFunc

Func  Comparehash($FilePath,$page,$oWorkbook) 
	Global $sheetarray[1]=[""]
	FileChangeDir($FilePath)
	;$oExcel = _Excel_Open(Default, True, Default, Default, True)
	;$oWorkbook=_Excel_BookOpen($oExcel,'\\192.168.2.100\100\�汾����\sbsb_a_a_A01_��\V003-202212211330-ʩ����\V003_202212211330_A00Card_ʩ����.xlsx')
	$sheetdata=_Excel_RangeRead($oWorkbook,$page)  ;��ȡָ���ӱ��е�������
	_ArrayAdd($sheetarray,$sheetdata)
	;_ArrayDisplay($sheetarray)
	GetNewDirHash($FilePath,$sheetarray) ;��·���Ѿ��ǽ��뵽���ļ���������   ֱ�ӿ�ʼ��������
EndFunc

Func GetNewDirHash($a,$old)   ;��ȡ���ļ������ļ��Ĺ�ϣֵ���ұ��浽������
	Local $NewTarray[1]=[""]   ;���ļ��й�ϣ�����������ʱ���飬Ϊ��װ��ÿ����ͬ���ӱ��У�ÿ�ε��ô˷���������NewTarray
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	;MsgBox(0,33,$a)
	$s = FileFindFirstFile($a& "\*.*")      
	$count=0   
	While 1
		$f = FileFindNextFile($s)  
		If @error Then  
			;_ArrayDisplay($NewTarray)
			ExitLoop
		EndIf
		;MsgBox(0,119,$f)  
		;$Newcount=$Newcount+1  
		$sData1 = _MD5ForFile2($a&"\"&$f)
		_ArrayAdd($NewTarray,$sData1)     ;ֻҪ�������ļ��У�������е��ļ������뵽New����������  12/21
		ConsoleWrite("Resultmmm: " & $sData1 & @CRLF & @CRLF)
	WEnd
	;_ArrayDisplay($NewTarray)
	$a=StringTrimLeft($a,2)					;�и�
	$array=StringSplit($a,"\")
	$b=$array[0]
	$file=$array[$b]
	;MsgBox(32,"�и���֮���",$file)  			 ;�е������ȵ���
	$file2=StringTrimRight($file,4) 
	$file2=$file2&"-Released"
	$file2=StringReplace($file2,6,$timestamp) 
	;MsgBox(0,999,_ArrayMaxIndex($NewTarray))

	If _ArrayMaxIndex($NewTarray)=0 Then   ;����һ���߼��жϣ����ǰ���ļ����ǿյģ���ֱ�Ӹ��Ƴ��µĿ��ļ���
		DirCreate($NewDir &"\"&$file2)
		;MsgBox(0,45,'create empty dir')
		EndIf
	;If (UBound($NewTarray)>0) Then   ;��Щ���ļ����ǿյ�
	For $j=1 To UBound($NewTarray)-1    ;ÿ��ȡ��һ�����ļ��������
			If Not($old[$j]=$NewTarray[$j]) Then 
				;MsgBox(0,0,"not equal")     ;ֻ�����ļ������������Ҫ�����жϣ����ļ��Լ�С�ļ������Ǵ���lnk
				;ֻҪ�ǲ���ȣ���ô��copy���ؼ���������copy��
				DirCreate($NewDir &"\"&$file2)
				DirCopy('\\'&$a, $NewDir &"\"&$file2,1)  
				;MsgBox(0,44,'copyfile')
			Else 
				;MsgBox(0,77,'equal')   ;��ȵĻ�����ô��lnk
				FileCreateShortcut('\\'&$a,$NewDir &"\"&$file2&'.lnk')   ;NewDir�����е�����
			EndIf
		Next
	;EndIf
EndFunc

Func CopyDir2($DirName)		;�ļ��еĸ���  �� ���� ��		;;;;;;�����ڴ˴��ټӵ��߼��жϣ����� lnk�Լ�����   12\9
$DirName=StringTrimLeft($DirName,2)	;12/7 �޸�·��
$array=StringSplit($DirName,"\")
$b=$array[0]
$file=$array[$b]
;MsgBox(0,'file',$file)
$b=StringTrimRight($DirName,StringLen($file))    ;$b���г����ļ�����
;MsgBox(0,'b',$b)
$file=StringTrimRight($file,4)   ;��   '-ʩ����' ɾ��
;MsgBox(0,'$file2',$file)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN  ;��ʱ���
$file=StringReplace($file,6,$timestamp)
;MsgBox(0,'$file3',$file)

$filename=$file&'-RELEASED'
;MsgBox(0,"filename",$filename)
 $NewDir='\\'&$b&$filename
DirCreate($NewDir)
;MsgBox(0,"newcount",$Newcount)
;BigHashPD()
			;��������ļ�֮��ͨ��
;MsgBox(32,"��ͷ���ļ�","���ļ����Ѹ���")
EndFunc

Func CheckFile2($FilePath)   		;У���ļ��е�����  �Լ�  �ļ�����  �����ļ����е��ļ�����û�м���֤ 11/25���Ѽ� 12\2��
	;MsgBox(64,"debug","����У�鷢���ļ�������")
	$FilePath=StringTrimLeft($FilePath,2)			;12/7 �޸�·��������������~~
	$array=StringSplit($FilePath,"\")
	$a=Int($array[0])
	$Check=StringInStr($array[$a],"ʩ����",0)		; !!!!!!!!!!!!!  fix  error
	MsgBox(0,"debug",$Check)
	If $Check = 0 Then											; Check Main   Dir
		MsgBox(16,"error","�����ļ�������,�����˳�!")
		Exit
		EndIf
	;MsgBox(0,"��ȡ�ļ�������",$array[$a])
	$version=StringLeft($array[$a],4)   ;V001_202212151424_1_ʩ����
	$version=StringRight($version,3)
	;$version=Int($version)
	;MsgBox(0,"�汾��",$version)
	$file = FileFindFirstFile('\\'&$FilePath&"\*.*")  
	If $file  = -1 Then MsgBox(16,"error","�ļ��ж�û��ȥ")
	$num =0
	While 1
		$file1 = FileFindNextFile($file) 
		If @error Then ExitLoop
		;If @extended Then 
		$v=StringLeft($file1 ,4)
		$v=StringRight($v,3)      ;BUG---��һ�ļ���ΪV002   ��v02X�������ļ������ж�Ϊ��ȷ�ģ���Ϊ�и��ԭ�� 11/25 13:30
		;$v=Int($v)
		$num+=1    ;����һ��ִ�е�λ�ã��������жϸ��Ƿ�Χ������    11/25
		If Not(StringCompare($version,$v)=0) Then $num-=1  ;�淶�Ƚ��ϸ��ж��ļ����ֲ�ͬ���˳��������û���ж��������ļ����˳�
		;MsgBox(1,"debug",$file1)
			;ContinueLoop
        ;EndIf
	WEnd
	FileClose($file)
	;MsgBox(0,"�ܹ��ļ�����",$num)
	if not($num=$FileNum)Then 
		MsgBox(16,"error","�ļ�������汾�Ŵ��󣬳����˳���")
		Exit ;�˳�ȫ���ű�
	Elseif $num=12 Then
		$excelname=$file  
		MsgBox(0,"debug","������ȷ��������һ������" )
	EndIf
EndFunc

#cs
	1����¼Rleased����card�е�hash�������һ������A�С�
	2������Rleased�����hash
	3����A��Rleased���¼������hash���бȽ�
#ce


Global $Newcount=0
Global $Newdircount=0
Global $Newarray[1]=[""]
Global $oldbigarray[1]=[""]


Func FindExcel2($searchdir)   ;�ҵ�card�ı���ļ�
$search = FileFindFirstFile($searchdir & "\*.*")
If $search = -1 Then return -1 
While 1
$file = FileFindNextFile($search)         ;������һ���ļ�
If @error Then                                   
FileClose($search)  
ExitLoop
EndIf
if Not (StringInStr($file,"card") = 0) Then
	Global $excelname2=$file
	;MsgBox(0,0,$excelname1)
	EndIf
WEnd
$excelname2=$searchdir&"\"&$excelname2
;MsgBox(0,'find',$excelname2)
EndFunc

;GetNewHash($searchdir)     ;��  ʩ �� R  ʱ����ȡʩ���������ļ��Ĺ�ϣֵ

Func GetNewHash($searchdir)      ;��ȡ�ļ��Ĺ�ϣֵ
	FileChangeDir($searchdir)      ;һ��Ҫת������Ӧ���ļ�·��
	$search = FileFindFirstFile($searchdir & "\*.*")
	If $search = -1 Then return -1 
	While 1
	$file = FileFindNextFile($search)         ;������һ���ļ�
	If @error Then
		FileClose($search)
		return
	EndIf
	
	If @extended=1 Then   
		;MsgBox(0,"Dir",$file)
		$Newdircount=$Newdircount+1     ;���������ļ��еļ�����    ���ļ�������
		;GetNewDirHash($searchdir,$file)   ;��ȡ���ļ�������Ĺ�ϣֵ����������´�����������
		ContinueLoop
		
	ElseIf  (StringInStr($file,'card')=0) Then   ;���ֿյ����ݵ������������  �ж�������ͻ   �б�����0�±�
		$sData = _MD5ForFile2($file)    ;�˴��������ϣֵ��
		_ArrayAdd($Newarray,$sData)
		ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
		$Newcount=$Newcount+1
		;�ڴ˼�  ��ȡ�ļ�������Ĺ�ϣֵ����������´�����������
	EndIf
	WEnd
	
EndFunc

Func GetOldarray($i,$searchdir)  ;��ȡ��һ�ε�hash���ݣ����浽oldarray�У��������ļ����жϣ��Ƿ�����lnk�ļ�
	FindExcel2($searchdir)
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname2)
	;$oWorkbook.Sheets(2).unProtect("a")
	Global $oldarray[1]=[""]
	;MsgBox(0,0,'GetOldarray'&$i)
	;MsgBox(0,11,$excelname2)
	;MsgBox(0,12,$i)
	$olddata=_Excel_RangeRead($oWorkbook,Number($i)+1)  ;��Ҫ�����ض��Ķ�ȡ��Χ������     12/7    �����޷�ȷ�����ļ���������ļ�����  12/21
	_ArrayAdd($oldarray,$olddata)

	If (UBound($NewTarray)-1>0) Then
	For $j=1 To UBound($NewTarray)-1    ;ÿ��ȡ��һ�����ļ��������
			If Not($oldarray[$j]==$NewTarray[$j]) Then 
				MsgBox(0,0,"not equal")     ;ֻ�����ļ������������Ҫ�����жϣ����ļ��Լ�С�ļ������Ǵ���lnk
			EndIf
		Next
	EndIf
	;$oldarray=_Excel_RangeRead($oWorkbook,2)
	Global $NewTarray[1]=[""]
	_Excel_Close($oExcel)
EndFunc

;SendToPython('\\192.168.2.100\100\�汾����\kkkk_t_sdfg_����01_���\V000-202302091024_ʩ����')

Func SendToPython($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"ʩ����")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_ʩ����'
	
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0209ToRel\0209ToRel.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","�ļ�������ӦΪʩ����")
	EndIf
	
EndFunc 



Func UpDataFile($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"ʩ����")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_ʩ����'
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0219UpDataFile\0219UpDataFile.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","�ļ�������ӦΪʩ����")
	EndIf
	
EndFunc 


Func PrePost($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"ʩ����")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_ʩ����'
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0210PrePost\0210PrePost.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","�ļ�������ӦΪʩ����")
	EndIf
	
EndFunc 


#cs   ԭ���Ļ�ȡ�ɹ�ϣֵ����
Func Match()
	ReadHash($oldfile)  ;�ȶ�ȡǰһ���ļ�excel����Ĺ�ϣֵ
	pd($hasharray2,$oldarray)
	
EndFunc

Func GetNewHash($searchdir)      ;��ȡ�ļ��Ĺ�ϣֵ
FileChangeDir($searchdir)      ;һ��Ҫת������Ӧ���ļ�·��
$search = FileFindFirstFile($searchdir & "\*.*")
If $search = -1 Then return -1 
While 1
$file = FileFindNextFile($search)         ;������һ���ļ�
If @error Then                                   
FileClose($search)  
ExitLoop
EndIf
If @extended=1 Then       ;�����ļ��е�ʱ��    �����һ���յ�result�Ĺ�ϣֵ   �������߼�������  ��δ�޸� (11/17 09:30)  ���޸���11/17 10:00)
;MsgBox(32,"�����ļ���",$file)
GetDirHash2($searchdir,$file)       
EndIf
ConsoleWrite($file&@CRLF)
if (StringInStr($file,"card") = 0) and (StringInStr($file,"dir") = 0)Then   ;���ֿյ����ݵ������������  �ж�������ͻ   �б�����0�±�
$sData = _MD5ForFile2($file)
_ArrayAdd($hasharray2,$sData)
ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
EndIf
WEnd
_ArrayDisplay($hasharray2)
EndFunc

Func GetDirHash2($searchdir,$file)   ;��ȡ���ļ������ļ��Ĺ�ϣֵ���ұ��浽������
$a=$searchdir &"\"&$file
$s = FileFindFirstFile($a& "\*.*")     
;MsgBox(0,0,$s)  
While 1
$f = FileFindNextFile($s)  
If @error Then                                   
ExitLoop
EndIf
;MsgBox(0,0,$f)  
$sData1 = _MD5ForFile2($searchdir&"\"&$file&"\"&$f)
_ArrayAdd($hasharray2,$sData1)
ConsoleWrite("Resultmmmmmm: " & $sData1 & @CRLF & @CRLF)
WEnd
EndFunc
#ce

Func _MD5ForFile2($sFile)

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



