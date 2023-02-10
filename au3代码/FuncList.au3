#cs
	开始升级
	该文件为		RELEASED --->>施工中 升级按钮函数功能	脚本文件
	RELEASED --->>施工中时，先把RELEASED文件所有的哈希值记录下来，到Card.xlsx中，放入到固定好的某一列或一行，需要上锁
	存放所有主窗口要调用的函数
	
#ce
#include <Array.au3>
#include <Date.au3>

Global $hasharray[1]=[""]    ;存放hash值的全局数组

Global $FileNum=37 ;当前大文件夹下面的文件总数量

Global $NewDir
Func StartUpdate($searchdir)  			;开始升级
FileChangeDir($searchdir)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN   ;202211241406   将时间戳定义为全局变量  目前看起来不太行，可能是会产生变量地址冲突
$search = FileFindFirstFile($searchdir & "\*.*")     
 If $search = -1 Then return -1                   
 While 1
 $file = FileFindNextFile($search)         ;查找下一个文件
 If @error Then                                   
  FileClose($search)                           
  return                                                   
EndIf
;$count=StringLeft($file,3) 
$mid_file=StringSplit($file,".")  			;;;;;;;;;;;;;;;;;;;;;wait to modify
If @error=1 Then 
	;MsgBox(32,"是文件夹",$file)  		;;;;;;;; 跳过文件夹     暂时先不复制      子文件夹
	;$DirName1=StringTrimRight($searchdir,9)
	;FileChangeDir($DirName1)
		$f=StringLeft($file,4)				;;文件夹从左切8个
		$f=StringRight($f,3)
		$f+=1
		$f=StringFormat("%03d",$f)	;格式换一下
		$file2=StringReplace($file,2,$f)    ;;;;;;;;;;对文件夹名字的修改    能替换吗？？？？
		;MsgBox(0,"replace",$file2)
		;加时间戳
		$file2=StringReplace($file2,6,$timestamp)
		;MsgBox(0,"replace-timestamp",$file2)
		$file2=StringTrimRight($file2,9)
		;MsgBox(0,"replace--delet",$file2)
	DirCreate($NewDir &"\"&$file2&"-施工中")
	DirCopy($searchdir&"\"&$file, $NewDir &"\"&$file2&"-施工中",1)
	;MsgBox(32,"即将复制的文件夹",$searchdir&"\"&$file)
	;MsgBox(32,"文件夹已复制",$NewDir &"\"&$file2&"-施工中")
Else
		;$file2 = $mid_file[1]&"_施工中"&"."&$mid_file[2]
		
						;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;对文件名字的修改   采用替换！！ 13：10 已修改  v+版本号+序号+名称+后缀   13号 11：20修改
						;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;11/24 已变更    A0x+v+版本号+时间戳+序号+名称+后缀
		$mid_file[1]=StringTrimRight($mid_file[1],9)  ;删released
		;MsgBox(0,"replace--delet",$mid_file[1])
		$file2 = $mid_file[1]&"_施工中"&"."&$mid_file[2]
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

Func CopyDir($DirName)		;文件夹的复制  和 改名
$DirName=StringTrimLeft($DirName,2)	;12/7 修改路径    公共盘前面会有双斜杠
$array=StringSplit($DirName,"\")
$b=$array[0]
$file=$array[$b]
$b=StringTrimRight($DirName,StringLen($file))

$v=StringLeft($file,4)   ;切割的字符串数量
$v=StringRight($v,3)
$v+=1
$v=StringFormat("%03d",$v)
;MsgBox(0,"",$v)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
$file=StringReplace($file,2,$v)
$file=StringReplace($file,6,$timestamp)   
$file=StringTrimRight($file,9)
$filename=$file&'-施工中'
;MsgBox(0,"",$filename)
 $NewDir='\\'&$b&$filename   ;最后路径加上了 \\
DirCreate($NewDir)
;MsgBox(32,"外头大文件","大文件夹已复制")
EndFunc


Func CheckFile($FilePath)   							;校验文件夹  文件名  文件数量
	$FilePath=StringTrimLeft($FilePath,2)																;12/7 修改到公共盘中，路径需要更改
	;MsgBox(64,"文件校验","进入校验文件函数中")
	$array=StringSplit($FilePath,"\")
	$a=Int($array[0])

	$Check=StringInStr($array[$a],"RELEASED",0)		; !!!!!!!!!!!!!  fix  error

	;MsgBox(0,"cccc",$Check)
	If $Check = 0 Then											; Check Main   Dir
		MsgBox(16,"error","升级文件名出错,程序退出!")
		Exit
		EndIf
	
	;MsgBox(0,"获取文件夹名字",$array[$a])
	$version=StringLeft($array[$a],4)  ;从左边开始取4个
	$version=StringRight($version,3)
	;MsgBox(0,"版本号",$version) 
	$file = FileFindFirstFile("\\"&$FilePath&"\*.*")  
	;MsgBox(0,"寻找文件",$file) 
	If $file  = -1 Then MsgBox(16,"error","文件夹都没进去")
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
	MsgBox(0,"总共文件数量",$num)
	
	if not($num=37)Then       ;数字与 $FileNum  传不进来？？？
		MsgBox(16,"error","文件数量或版本号错误，程序退出！")
		Exit ;退出全部脚本
	Elseif $num=37 Then
		$excelname=$file  
		MsgBox(0,"文件校验","数量正确，继续下一步操作" )
		FileSetAttrib('\\'&$FilePath,"-R",1)  
		;MsgBox(0,"整体文件已解锁" ,'\\'&$FilePath)
	EndIf
EndFunc


Func AddRow($FilePath)    								;升级之后  在相应的位置进行新增一列
	$search1 = FileFindFirstFile($FilePath & "\*.*")   ;遍历文件夹，获取xlsx类型文件
	While 1 
		$file = FileFindNextFile($search1)
		If @error Then ExitLoop
		if Not StringInStr($file,'card')=0 Then   ;找到Card文件，11/21 15:30 已实现
			$excelname=$file  
		EndIf
WEnd
		$excelname1=$FilePath&"\"&$excelname
		$Split=StringLeft($excelname,4)
		$Split=StringRight($Split,3)    ;获取需要加哪列,可以位移量大一些
		;MsgBox(0,"split",$Split)
		$Split1= Int($Split)
		$Split1=$Split1+4   		 ;可以再此处修改位移量
		;MsgBox(0,"",$Split1)
		$oExcel = _Excel_Open(False, True, Default, Default, True)
		$oWorkbook=_Excel_BookOpen($oExcel,$excelname1)
		
		;MsgBox(0,0,'open excel')
		;$oWorkbook.Sheets(1).Range("F3").Interior.ColorIndex = 3		  ;;;;;;;高亮显示
		$oWorkbook.ActiveSheet.unProtect("a") 						;已经加了一列  将加的那列解锁，其他的为保护状态
		$oWorkbook.Sheets(1).Cells(3,$Split1).Value ="V:"&$Split1-3
		$oWorkbook.Sheets(1).Cells(4,$Split1).Value = @MON&@MDAY&'|'&@HOUR&':'&@MIN             ; TimeStamp
		
		trans($Split1)
		$col=$end
		trans($Split1-1)
		$col_pre=$end
		
		$oExcel.Range($col_pre&':'&$col_pre).Locked = True
		$oExcel.Range($col&':'&$col).Locked = False						;1、先解锁保护  2、将特定列设为不锁状态 3、再设为保护
		$oWorkbook.ActiveSheet.Protect("a") 
		
		
MsgBox(64,"INFO","新加列 " & $col & " 且已锁定")

_Excel_Close($oExcel)

EndFunc  

Global $end
Func trans($num)			;数字转换到字母是因为  需要在excel.range里面锁定特定的列
	$c=Mod($num,26)
	If $num <=26 Then
	$l=$num+64
	$end=Chr($l)
	;MsgBox(0,"小于26",$end)
	ElseIf $num >26 And Not($c=0) Then   ;主要的数字转字母逻辑
	$d=$num/26
	;MsgBox(0,"除数1",$d)
	$e=int($d)
	$e+=64
	$l=$c+64
	;MsgBox(0,"对应字母",Chr($l))
	$end=Chr($e)&Chr($l)
	;MsgBox(0,"最终",$end)
	;$e+=64
	;$e=Chr($e)
Else  
		$d=$num/26
		;MsgBox(0,"除数2",$d)
		$e=int($d)
		$e-=1
		$e+=64	
		;MsgBox(0,"第一个",$e)		
		$c="Z"
		$end=Chr($e)&$c
		;MsgBox(0,"最终2",$end)
EndIf
EndFunc


#cs
	1、以下为R----》施
	2、只需要将解锁为施工中之前把released中的所有文件hash值存到card文件中即可

#ce

Global $count   ;子表数量
Global $dircount=0
Global $hasharray3[1]=[""]


Func GetHash($searchdir)      ;获取文件的哈希值
	FindExcel($searchdir)
	FileChangeDir($searchdir)      ;一定要转换到相应的文件路径
	$search = FileFindFirstFile($searchdir & "\*.*")
	If $search = -1 Then return -111 
	While 1
	$file = FileFindNextFile($search)         ;查找下一个文件
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
		$dircount=$dircount+1     ;遍历到的文件夹的计数器
		GetDirHash($searchdir,$file)   ;获取子文件夹里面的哈希值，并按序存到excel主表格中
		
	
	ElseIf  (StringInStr($file,'Card')=0) Then   ;出现空的数据的是这里的问题  判断条件冲突   列表里面0下标
		$sData = _MD5ForFile1($file)
		_ArrayAdd($hasharray3,$sData)
		ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
		$count=$count+1
	EndIf
	;ConsoleWrite($file&@CRLF)
	WEnd
	
EndFunc


Func GetDirHash($searchdir,$file)   ;获取子文件夹中文件的哈希值并且保存到数组中   此处对文件夹数量要求需要做规定
	$a=$searchdir &"\"&$file
	$s = FileFindFirstFile($a& "\*.*")     
	;MsgBox(0,'get dir ',$s)  
	Global $Tarray[1]=[""]   ;子文件夹哈希内容里面的临时数组，为了装到每个不太的子表中。
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
$file = FileFindNextFile($search)         ;查找下一个文件
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

Func AddToNewSheet($count,$Tarray)  ;在新建立的表中  填入哈希值   此处逻辑不太正确，每次都加新表，是不对的，需要确定好子表数量
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname1) ;路径有问题
	;MsgBox(0,2,$excelname1)
	Local $aWorkSheets = _Excel_SheetList($oWorkbook)
	Local $sheetnum=0
	For $IDX = 1 To UBound($aWorkSheets) -1
			$sheetnum=$sheetnum+1
		Next
	If $sheetnum<7 Then
	_Excel_SheetAdd($oWorkbook,Default,False)    ;表格中加sheet子表格 表格名，默认子表格名，放在第二个位置
	EndIf
	;MsgBox(0,3,'test')											;最好是在一开始就要设定好有多少文件夹，多少个子表格，每个子表格如何确定呢？？
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





;以下为单纯计算文件MD5哈希值
Func _MD5ForFile1($sFile)

    Local $a_hCall = DllCall("kernel32.dll", "hwnd", "CreateFileW", _                       ;DllCall  调用windows系统的   
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

    Local $hFile = $a_hCall[0]       ;放到$hFile

    $a_hCall = DllCall("kernel32.dll", "ptr", "CreateFileMappingW", _
            "hwnd", $hFile, _   ;$hFile 调用到新到DllCall 
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

    Local $hFileMappingObject = $a_hCall[0]   ;放到$hFileMappingObject 

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

    Local $tMD5_CTX = DllStructCreate("dword i[2];" & _  ;$tMD5_CTX  这个是
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















