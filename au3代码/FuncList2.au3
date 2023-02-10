#cs
	确认发布
	该文件为	 施工中-->>RELEASED- 确认发布函数功能脚本文件
	
	存放所有主窗口要调用的函数
	
#ce

Global $FileNum =37  ;当前大文件夹下面的文件总数量
Global $NewDir
Global $oldarray[1]=[""]
Global $hasharray2[1]=[""]
Func Released($searchdir)  	;确认发布      重写创建快捷方式以及复制改变了的文件（暂时不包括大文件夹里面的）12/21    16：00  已包含子文件夹 12/23
	Local $page=1
	FindExcel2($searchdir)
	;此为excel中的旧hash以及新获取的hash进行判断
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname2)   ; xlsl文件用$excelname2
	$oldbigdata=_Excel_RangeRead($oWorkbook,1,'B291:B319')      ;读表格里面的数据   大表格的hash看起来没错  此范围后续可能还需要变化
	_ArrayAdd($oldbigarray,$oldbigdata)
	;MsgBox(0,1,"show two array")
	;_ArrayDisplay($oldbigarray)
	;创建快捷方式一开始把外面的大文件复制或者是lnk，当遇到子文件夹的时候，应该是与表格中的子表格对比？
	Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN
	$search = FileFindFirstFile($searchdir & "\*.*")     
	If $search = -1 Then return -1    ;基本的文件溢出之后的退出
	$j=0	  ;计数器  暂时用此计数器   为了给两个数组进行遍历
While 1
	$file = FileFindNextFile($search)         ;查找下一个文件
If @error Then
	FileClose($search)  
	FileSetAttrib($NewDir,"+R",1)     ;整体设为只读
	MsgBox(0,"整体设为只读",$NewDir)
	_Excel_Close($oExcel)
	ExitLoop
	EndIf
If @extended Then           ;找到子文件夹的，可以创建一个方法，给子文件中hash与old hash做比较，大部分情况为复制，
	;MsgBox(0,"Dir",$searchdir&'\'&$file)
	$page=$page+1  ;每找到一个子文件夹，需要将相应的子表格也更改
	Comparehash($searchdir&'\'&$file,$page,$oWorkbook)    ;1、子文件夹的路径,
	ContinueLoop
ElseIf $j<30 Then    ;除了8个大文件，37-8=29
	;MsgBox(0,1,$file)   ;复制或者是lnk，同样也是通过hash去判断
	$mid_file=StringSplit($file,".")  
	;切割文件名并且有些会用lnk代替
	$mid_file[1]=StringTrimRight($mid_file[1],4)
	$mid_file[1]=$mid_file[1]&"_Released"
	$mid_file[1]=StringReplace($mid_file[1],6,$timestamp)
	$file2 = $mid_file[1]&"."&$mid_file[2]		
	;For $j=1 To UBound($Newarray)-1
	;MsgBox(0,888,$j)
	If($Newarray[$j]==$oldbigarray[$j]) And (StringInStr($file2,"card")=0) Then   ;;;;;复制逻辑出问题12/16
	;MsgBox(0,"equl",$j)
	FileCreateShortcut($searchdir & "\" & $file,$NewDir&"\"& $mid_file[1]&'.lnk')    ;0105   创建快捷方式，如何添加沿用V000  并定位到相应上一级目录中
	;$j=$j+1
	;MsgBox(0,"jjjj",$j)
	ElseIf  Not ($Newarray[$j]==$oldbigarray[$j]) And (StringInStr($file2,"card")=0) Then
	;copy逻辑需要更改
	;MsgBox(0,'即将要copy',$file2)
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
	;$oWorkbook=_Excel_BookOpen($oExcel,'\\192.168.2.100\100\版本迭代\sbsb_a_a_A01_左\V003-202212211330-施工中\V003_202212211330_A00Card_施工中.xlsx')
	$sheetdata=_Excel_RangeRead($oWorkbook,$page)  ;获取指定子表中的数据了
	_ArrayAdd($sheetarray,$sheetdata)
	;_ArrayDisplay($sheetarray)
	GetNewDirHash($FilePath,$sheetarray) ;该路径已经是进入到子文件夹里面了   直接开始遍历即可
EndFunc

Func GetNewDirHash($a,$old)   ;获取子文件夹中文件的哈希值并且保存到数组中
	Local $NewTarray[1]=[""]   ;子文件夹哈希内容里面的临时数组，为了装到每个不同的子表中，每次调用此方法回重置NewTarray
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
		_ArrayAdd($NewTarray,$sData1)     ;只要碰到了文件夹，会把所有的文件都传入到New数组里面吗？  12/21
		ConsoleWrite("Resultmmm: " & $sData1 & @CRLF & @CRLF)
	WEnd
	;_ArrayDisplay($NewTarray)
	$a=StringTrimLeft($a,2)					;切割
	$array=StringSplit($a,"\")
	$b=$array[0]
	$file=$array[$b]
	;MsgBox(32,"切割完之后的",$file)  			 ;切到合理长度的了
	$file2=StringTrimRight($file,4) 
	$file2=$file2&"-Released"
	$file2=StringReplace($file2,6,$timestamp) 
	;MsgBox(0,999,_ArrayMaxIndex($NewTarray))

	If _ArrayMaxIndex($NewTarray)=0 Then   ;另外一种逻辑判断，如果前面文件夹是空的，会直接复制出新的空文件夹
		DirCreate($NewDir &"\"&$file2)
		;MsgBox(0,45,'create empty dir')
		EndIf
	;If (UBound($NewTarray)>0) Then   ;有些子文件夹是空的
	For $j=1 To UBound($NewTarray)-1    ;每获取完一次子文件夹里面的
			If Not($old[$j]=$NewTarray[$j]) Then 
				;MsgBox(0,0,"not equal")     ;只是子文件夹里面的吗，需要两次判断，大文件以及小文件，都是创建lnk
				;只要是不相等，那么就copy，关键是在哪里copy呢
				DirCreate($NewDir &"\"&$file2)
				DirCopy('\\'&$a, $NewDir &"\"&$file2,1)  
				;MsgBox(0,44,'copyfile')
			Else 
				;MsgBox(0,77,'equal')   ;相等的话，那么就lnk
				FileCreateShortcut('\\'&$a,$NewDir &"\"&$file2&'.lnk')   ;NewDir可能有点问题
			EndIf
		Next
	;EndIf
EndFunc

Func CopyDir2($DirName)		;文件夹的复制  和 改名 √		;;;;;;可以在此处再加点逻辑判断，创建 lnk以及复制   12\9
$DirName=StringTrimLeft($DirName,2)	;12/7 修改路径
$array=StringSplit($DirName,"\")
$b=$array[0]
$file=$array[$b]
;MsgBox(0,'file',$file)
$b=StringTrimRight($DirName,StringLen($file))    ;$b是切除了文件名的
;MsgBox(0,'b',$b)
$file=StringTrimRight($file,4)   ;把   '-施工中' 删除
;MsgBox(0,'$file2',$file)
Local $timestamp=@YEAR&@MON&@MDAY&@HOUR&@MIN  ;加时间戳
$file=StringReplace($file,6,$timestamp)
;MsgBox(0,'$file3',$file)

$filename=$file&'-RELEASED'
;MsgBox(0,"filename",$filename)
 $NewDir='\\'&$b&$filename
DirCreate($NewDir)
;MsgBox(0,"newcount",$Newcount)
;BigHashPD()
			;创建完大文件之后，通过
;MsgBox(32,"外头大文件","大文件夹已复制")
EndFunc

Func CheckFile2($FilePath)   		;校验文件夹的名字  以及  文件数量  （对文件夹中的文件好像没有加验证 11/25，已加 12\2）
	;MsgBox(64,"debug","进入校验发布文件函数中")
	$FilePath=StringTrimLeft($FilePath,2)			;12/7 修改路径到公共盘上面~~
	$array=StringSplit($FilePath,"\")
	$a=Int($array[0])
	$Check=StringInStr($array[$a],"施工中",0)		; !!!!!!!!!!!!!  fix  error
	MsgBox(0,"debug",$Check)
	If $Check = 0 Then											; Check Main   Dir
		MsgBox(16,"error","升级文件名出错,程序退出!")
		Exit
		EndIf
	;MsgBox(0,"获取文件夹名字",$array[$a])
	$version=StringLeft($array[$a],4)   ;V001_202212151424_1_施工中
	$version=StringRight($version,3)
	;$version=Int($version)
	;MsgBox(0,"版本号",$version)
	$file = FileFindFirstFile('\\'&$FilePath&"\*.*")  
	If $file  = -1 Then MsgBox(16,"error","文件夹都没进去")
	$num =0
	While 1
		$file1 = FileFindNextFile($file) 
		If @error Then ExitLoop
		;If @extended Then 
		$v=StringLeft($file1 ,4)
		$v=StringRight($v,3)      ;BUG---若一文件名为V002   则v02X的所有文件都会判断为正确的，因为切割的原因 11/25 13:30
		;$v=Int($v)
		$num+=1    ;换了一个执行的位置，看起来判断覆盖范围更大了    11/25
		If Not(StringCompare($version,$v)=0) Then $num-=1  ;规范比较严格，判断文件名字不同则退出，会出现没有判断完所有文件就退出
		;MsgBox(1,"debug",$file1)
			;ContinueLoop
        ;EndIf
	WEnd
	FileClose($file)
	;MsgBox(0,"总共文件数量",$num)
	if not($num=$FileNum)Then 
		MsgBox(16,"error","文件数量或版本号错误，程序退出！")
		Exit ;退出全部脚本
	Elseif $num=12 Then
		$excelname=$file  
		MsgBox(0,"debug","数量正确，继续下一步操作" )
	EndIf
EndFunc

#cs
	1、记录Rleased里面card中的hash，存放在一个数组A中。
	2、计算Rleased里面的hash
	3、将A与Rleased重新计算出的hash进行比较
#ce


Global $Newcount=0
Global $Newdircount=0
Global $Newarray[1]=[""]
Global $oldbigarray[1]=[""]


Func FindExcel2($searchdir)   ;找到card的表格文件
$search = FileFindFirstFile($searchdir & "\*.*")
If $search = -1 Then return -1 
While 1
$file = FileFindNextFile($search)         ;查找下一个文件
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

;GetNewHash($searchdir)     ;当  施 到 R  时，获取施工后所有文件的哈希值

Func GetNewHash($searchdir)      ;获取文件的哈希值
	FileChangeDir($searchdir)      ;一定要转换到相应的文件路径
	$search = FileFindFirstFile($searchdir & "\*.*")
	If $search = -1 Then return -1 
	While 1
	$file = FileFindNextFile($search)         ;查找下一个文件
	If @error Then
		FileClose($search)
		return
	EndIf
	
	If @extended=1 Then   
		;MsgBox(0,"Dir",$file)
		$Newdircount=$Newdircount+1     ;遍历到的文件夹的计数器    子文件夹数量
		;GetNewDirHash($searchdir,$file)   ;获取子文件夹里面的哈希值，并按序存新创建的数组中
		ContinueLoop
		
	ElseIf  (StringInStr($file,'card')=0) Then   ;出现空的数据的是这里的问题  判断条件冲突   列表里面0下标
		$sData = _MD5ForFile2($file)    ;此处是求出哈希值。
		_ArrayAdd($Newarray,$sData)
		ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
		$Newcount=$Newcount+1
		;在此加  获取文件夹里面的哈希值，并按序存新创建的数组中
	EndIf
	WEnd
	
EndFunc

Func GetOldarray($i,$searchdir)  ;获取上一次的hash数据，并存到oldarray中，并进行文件的判断，是否生成lnk文件
	FindExcel2($searchdir)
	$oExcel = _Excel_Open(False, True, Default, Default, True)
	$oWorkbook=_Excel_BookOpen($oExcel,$excelname2)
	;$oWorkbook.Sheets(2).unProtect("a")
	Global $oldarray[1]=[""]
	;MsgBox(0,0,'GetOldarray'&$i)
	;MsgBox(0,11,$excelname2)
	;MsgBox(0,12,$i)
	$olddata=_Excel_RangeRead($oWorkbook,Number($i)+1)  ;需要更改特定的读取范围！！！     12/7    但是无法确定子文件夹里面的文件数量  12/21
	_ArrayAdd($oldarray,$olddata)

	If (UBound($NewTarray)-1>0) Then
	For $j=1 To UBound($NewTarray)-1    ;每获取完一次子文件夹里面的
			If Not($oldarray[$j]==$NewTarray[$j]) Then 
				MsgBox(0,0,"not equal")     ;只是子文件夹里面的吗，需要两次判断，大文件以及小文件，都是创建lnk
			EndIf
		Next
	EndIf
	;$oldarray=_Excel_RangeRead($oWorkbook,2)
	Global $NewTarray[1]=[""]
	_Excel_Close($oExcel)
EndFunc

;SendToPython('\\192.168.2.100\100\版本迭代\kkkk_t_sdfg_国产01_左驾\V000-202302091024_施工中')

Func SendToPython($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"施工中")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_施工中'
	
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0209ToRel\0209ToRel.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","文件名错误，应为施工中")
	EndIf
	
EndFunc 



Func UpDataFile($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"施工中")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_施工中'
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0219UpDataFile\0219UpDataFile.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","文件名错误，应为施工中")
	EndIf
	
EndFunc 


Func PrePost($Readpath,$Code)
	
	$aDays=StringSplit($Readpath,"\")
	$version=StringLeft($aDays[7],4)	;v00x
	$pd=StringRight($aDays[7],3)
	If StringCompare($pd,"施工中")=0 Then
	$time=StringTrimLeft($aDays[7],5)
	$time=StringTrimRight($time,4)
	$path=$version&'_'&$time&'_'&'A13Calc_施工中'
	$Excelpath=$Readpath&"\"&$path
	;For $i = 1 To $aDays[0] ; Loop through the array returned by StringSplit to display the individual values.
       ; MsgBox(4096, "Example1", "$aDays[" & $i & "] - " & $aDays[$i])
   ; Next
	;MsgBox(0,'excelpath',$Excelpath)
	$FileName=$version&'_'&$aDays[6]
	ShellExecute("\\192.168.2.100\100\0210PrePost\0210PrePost.exe",$Excelpath&'>'&$FileName&'>'&$Code)  
	Else
		MsgBox(0,"error","文件名错误，应为施工中")
	EndIf
	
EndFunc 


#cs   原来的获取旧哈希值方法
Func Match()
	ReadHash($oldfile)  ;先读取前一个文件excel上面的哈希值
	pd($hasharray2,$oldarray)
	
EndFunc

Func GetNewHash($searchdir)      ;获取文件的哈希值
FileChangeDir($searchdir)      ;一定要转换到相应的文件路径
$search = FileFindFirstFile($searchdir & "\*.*")
If $search = -1 Then return -1 
While 1
$file = FileFindNextFile($search)         ;查找下一个文件
If @error Then                                   
FileClose($search)  
ExitLoop
EndIf
If @extended=1 Then       ;遍历文件夹的时候    会出现一个空的result的哈希值   可能是逻辑的问题  还未修复 (11/17 09:30)  已修复（11/17 10:00)
;MsgBox(32,"捕获文件夹",$file)
GetDirHash2($searchdir,$file)       
EndIf
ConsoleWrite($file&@CRLF)
if (StringInStr($file,"card") = 0) and (StringInStr($file,"dir") = 0)Then   ;出现空的数据的是这里的问题  判断条件冲突   列表里面0下标
$sData = _MD5ForFile2($file)
_ArrayAdd($hasharray2,$sData)
ConsoleWrite("Result: " & $sData & @CRLF & @CRLF)
EndIf
WEnd
_ArrayDisplay($hasharray2)
EndFunc

Func GetDirHash2($searchdir,$file)   ;获取子文件夹中文件的哈希值并且保存到数组中
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



