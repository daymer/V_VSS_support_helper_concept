#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
If Not IsAdmin() Then
    MsgBox($MB_SYSTEMMODAL, "Warning","Admin rights are not detected."& @CR& 'You need to be admin to run this tool')
	Exit
EndIf
#include <Date.au3>
#RequireAdmin
$ProgrammDir = 'C:\VeeamVssHelper\'
$temp_dir = @TempDir&'\VeeamVssHelper\'
Global $PCName = @ComputerName
If Not FileExists($ProgrammDir) Then DirCreate($ProgrammDir)
If Not FileExists($temp_dir) Then DirCreate($temp_dir)
  $file_time = StringRegExpReplace(_NowDate(),'[:/]','-')
ConsoleWrite($file_time&@CR)
$list_writers_command = $temp_dir&'listwriters.txt'
If Not FileExists($list_writers_command) Then
   _FileCreate ($list_writers_command)
   FileWriteLine($list_writers_command, 'list writers')
   FileWriteLine($list_writers_command, 'exit')
   EndIf
$writers_log = $ProgrammDir&'VSS_writers_'&$PCName&'-'&$file_time&'.log'
If FileExists($writers_log) Then FileDelete ($writers_log)
#Region vss list writers
$cmd_pid = Run("C:\WINDOWS\system32\cmd.exe")
WinWaitActive("C:\WINDOWS\system32\cmd.exe")
$title_1st = WinGetTitle('[LAST]')
send('DISKSHADOW.EXE /S '&$list_writers_command&' /l '&$writers_log & "{ENTER}")
$title_2nd = WinGetTitle('[LAST]')
While 1
   $title = WinGetTitle('[LAST]')
ConsoleWrite($title&@CR)
if $title = $title_2nd Then
   Sleep(100)
   ContinueLoop
EndIf
if $title = $title_1st Then
   ProcessClose($cmd_pid)
   ExitLoop
EndIf
WEnd
#EndRegion
#Region writer array
$hFile = FileOpen($writers_log, 32)
$sText = FileRead($hFile)
;ConsoleWrite(@error)
$aLines = StringSplit($sText, @CRLF, 1)
Global $writersid[1][2]
 $writersid[0][0] = 'name'
  $writersid[0][1] = 'id'
;_ArrayDisplay($aLines)
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $target = stringInStr($aLines[$i], '		- Writer ID   = ')
   if  $target = 0 Then  ContinueLoop

Local $Temp_writer[1][2]
$writer = StringRegExpReplace($aLines[$i -1],'[^0-9a-zA-Z_]', '')
$writer = StringTrimLeft($writer, 6)
$writerid = StringTrimLeft($aLines[$i], 18)
$writerid = StringRegExpReplace($writerid,'[\r\n\t]+', '')
$Temp_writer[0][0] = $writer
$Temp_writer[0][1] = $writerid
ConsoleWrite($Temp_writer[0][0]&@cr)
ConsoleWrite($Temp_writer[0][1]&@cr)

_ArrayConcatenate($writersid, $Temp_writer)
Next
;_ArrayDisplay($writersid)
FileClose($hFile)
#EndRegion
#Region VSS_volumes
$volums_log = $ProgrammDir&'VSS_volumes_'&$PCName&'-'&$file_time&'.log'
If FileExists($volums_log) Then FileDelete ($volums_log)
   $cmd = 'vssadmin list volumes > '&$volums_log
   $cmd_pid = RunWait(@ComSpec & " /c " & $cmd, "", @SW_HIDE)
#EndRegion
  #Region  VSS_volumes array
$hFile = FileOpen($volums_log, 0)
$sText = FileRead($hFile)
ConsoleWrite(@error)
$aLines = StringSplit($sText, @CRLF, 1)
;_ArrayDisplay($aLines)
Global $volume[1][2]
 $volume[0][0] = ' Volume path:'
  $volume[0][1] = 'Volume name:'
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $target = stringInStr($aLines[$i], 'Volume name:')
   if  $target = 0 Then  ContinueLoop
Local $Temp_volume[1][2]
$path = StringTrimLeft($aLines[$i -1], 12)
$name = StringTrimLeft($aLines[$i], 16)
;$name = StringRegExpReplace($name,'[^0-9a-zA-Z_}\?:]', '')
$Temp_volume[0][0] = $path
$Temp_volume[0][1] = $name
ConsoleWrite($Temp_volume[0][0]&@cr)
ConsoleWrite($Temp_volume[0][1]&@cr)

_ArrayConcatenate($volume, $Temp_volume)
Next
;_ArrayDisplay($volume)
FileClose($hFile)
   #EndRegion
   #Region VSS_providers
$providers_log = $ProgrammDir&'VSS_providers_'&$PCName&'-'&$file_time&'.log'
If FileExists($providers_log) Then FileDelete ($providers_log)
   $cmd = 'vssadmin list providers > '&$providers_log
   $cmd_pid = RunWait(@ComSpec & " /c " & $cmd, "", @SW_HIDE)
#EndRegion
   #Region  VSS_providers array
$hFile = FileOpen($providers_log, 0)
$sText = FileRead($hFile)
ConsoleWrite(@error)
$aLines = StringSplit($sText, @CRLF, 1)
;_ArrayDisplay($aLines)
Global $provider[1][2]
 $provider[0][0] = 'provider name:'
  $provider[0][1] = 'provider id:'
For $i = 1 To $aLines[0] Step +1
   If $aLines[$i] = '' Then
	  ContinueLoop
   EndIf
  $target = stringInStr($aLines[$i], 'Provider Id:')
   if  $target = 0 Then  ContinueLoop
Local $Temp_provider[1][2]
$name = StringTrimLeft($aLines[$i -2], 16)
$name = StringTrimRight($name,1)
$id = StringTrimLeft($aLines[$i], 16)
;$name = StringRegExpReplace($name,'[^0-9a-zA-Z_}\?:]', '')
$Temp_provider[0][0] = $name
$Temp_provider[0][1] = $id
ConsoleWrite($Temp_provider[0][0]&@cr)
ConsoleWrite($Temp_provider[0][1]&@cr)

_ArrayConcatenate($provider, $Temp_provider)
Next
;_ArrayDisplay($provider)
FileClose($hFile)
   #EndRegion



If FileExists($temp_dir) Then DirRemove($temp_dir)
   If FileExists($list_writers_command) Then FileDelete($temp_dir)

