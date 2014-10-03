#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <GUIListBox.au3>
#include <date.au3>
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
If Not IsAdmin() Then
    MsgBox($MB_SYSTEMMODAL, "Warning","Admin rights are not detected."& @CR& 'You need to be admin to run this tool')
	Exit
EndIf
#include <Date.au3>
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
;_ArrayDisplay($provider)
;_ArrayDisplay($volume)
;_ArrayDisplay($writersid)
local $iRows_provider = UBound($provider, 1) -1
ConsoleWrite ('ConsoleWrite ($iRows_provider&@cr) '&$iRows_provider&@cr)
local $iRows_volume = UBound($volume, 1) -1
local $iRows_writersid = UBound($writersid, 1) -1
if $iRows_provider = 0 or $iRows_writersid = 0 or  $iRows_volume= 0 Then
   MsgBox(4096, 'Error', 'Nothing found - general error"')
   Exit
   EndIf
$provider_to_display =''
For $i = 1 To $iRows_provider Step +1
  ConsoleWrite ($iRows_provider&@cr)
$provider_to_display &= $provider[$i][0]&'|'
ConsoleWrite ($provider_to_display&@cr)
Next
$volume_to_display =''
For $i = 1 To $iRows_volume Step +1

Local $Temp_array_accaunt[1][2]

$volume_to_display &= $volume[$i][0]&'|'
ConsoleWrite ($volume_to_display&@cr)
Next
$writersid_to_display=''
For $i = 1 To $iRows_writersid Step +1

$writersid_to_display &= $writersid[$i][0]&'|'
Next

Example($writersid_to_display, $volume_to_display, $provider_to_display)
Func Example($writersid_to_display, $volume_to_display, $provider_to_display)
    Local $msg, $mylist, $file2
#Region ### START Koda GUI section ### Form=c:\users\dremsama\desktop\koda\templates\tabbed pages.kxf
$dlgTabbed = GUICreate("Vss helper", 413, 390, 299, 218)
GUISetIcon("", -1)
$PageControl1 = GUICtrlCreateTab(8, 8, 396, 320)
$TabSheet1 = GUICtrlCreateTabItem("Writers to validate")
#Region writers
 $Button_ok1 = GUICtrlCreateButton("&OK", 322, 275, 75, 25)
 $file2 = GUICtrlCreateInput("", 22, 277, 284, 22)
    $listwiters = GUICtrlCreateList("", 28, 37, 361, 216,BitOR($LBS_STANDARD, $LBS_EXTENDEDSEL))
    GUICtrlSetLimit(-1, 200)
    GUICtrlSetData(-1, $writersid_to_display)
    GUISetState()
#EndRegion
$TabSheet2 = GUICtrlCreateTabItem("Provider to use")
#Region writers
 $Button_ok2 = GUICtrlCreateButton("&OK", 322, 275, 75, 25)
 $file3 = GUICtrlCreateInput("", 22, 277, 284, 22)
    $listprovaiders = GUICtrlCreateList("", 28, 37, 361, 216,BitOR($LBS_STANDARD, $LBS_EXTENDEDSEL))
    GUICtrlSetLimit(-1, 200)
    GUICtrlSetData(-1, $provider_to_display)
    GUISetState()
#EndRegion
$TabSheet3 = GUICtrlCreateTabItem("Volums to use")
#Region writers
 $Button_ok3 = GUICtrlCreateButton("&OK", 322, 275, 75, 25)
 $file4 = GUICtrlCreateInput("", 22, 277, 284, 22)
    $listvolums = GUICtrlCreateList("", 28, 37, 361, 216,BitOR($LBS_STANDARD, $LBS_EXTENDEDSEL))
    GUICtrlSetLimit(-1, 200)
    GUICtrlSetData(-1, $volume_to_display)
    GUISetState()
#EndRegion
GUICtrlCreateTabItem("")
$Button1 = GUICtrlCreateButton("&Start", 166, 344, 75, 25)
$Button2 = GUICtrlCreateButton("&Cancel", 246, 344, 75, 25)
$Button3 = GUICtrlCreateButton("&Help", 328, 344, 75, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
Global $local_array_writersid[0][2]
Global $local_array_provaiders[0][2]
Global $local_array_volums[0][2]
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
	Case $GUI_EVENT_CLOSE
	        Exit
				Case $Button2
	        Exit
			Case $listwiters
                GUICtrlSetData($file2, GUICtrlRead($listwiters))
			 case $Button_ok1
				ConsoleWrite(GUICtrlRead($listwiters)&@cr)
				$writer_name = GUICtrlRead($listwiters)
				$found = _ArraySearch($writersid, $writer_name)
				ConsoleWrite($writersid[$found][1]&@cr)
				     Local $Temp_array_writersid[1][2]
					 $Temp_array_writersid[0][0] = $writersid[$found][0]
					 $Temp_array_writersid[0][1] =$writersid[$found][1]
					_ArrayConcatenate( $local_array_writersid, $Temp_array_writersid)

Case $listprovaiders
                GUICtrlSetData($file3, GUICtrlRead($listprovaiders))
			 case $Button_ok2
				$provider_name = GUICtrlRead($listprovaiders)
				$found = _ArraySearch($provider, $provider_name)
								     Local $Temp_array_writersid[1][2]
					 $Temp_array_writersid[0][0] = $provider[$found][0]
					 $Temp_array_writersid[0][1] = $provider[$found][1]
					_ArrayConcatenate( $local_array_provaiders, $Temp_array_writersid)

				Case $listvolums
                GUICtrlSetData($file4, GUICtrlRead($listvolums))
			 case $Button_ok3
				$volums_name = GUICtrlRead($listvolums)
				$found = _ArraySearch($volume, $volums_name)
								     Local $Temp_array_writersid[1][2]
					 $Temp_array_writersid[0][0] = $volume[$found][0]
					 $Temp_array_writersid[0][1] = $volume[$found][1]
					_ArrayConcatenate( $local_array_volums, $Temp_array_writersid)

				 case $Button1

					Global $local_array_volums_col = UBound($local_array_volums, 1)
					Global $local_array_provaiders_col = UBound($local_array_provaiders, 1)
					Global $local_array_writersid_col = UBound($local_array_writersid, 1)
					if $local_array_volums_col = 0 Then
					   MsgBox(4096, 'Error', 'No volums selected')
					ElseIf $local_array_writersid_col = 0 Then
					   MsgBox(4096, 'Error', 'No writers selected')
					   ElseIf $local_array_provaiders_col = 0 Then
					   MsgBox(4096, 'Error', 'No providers selected')
					   Else
GUIDelete($dlgTabbed)
ExitLoop
EndIf
	EndSwitch
 WEnd
 EndFunc   ;==>Example
$ProgrammDir = 'C:\VeeamVssHelper\'
$temp_dir = @TempDir&'\VeeamVssHelper\'
Global $PCName = @ComputerName
If Not FileExists($temp_dir) Then DirCreate($temp_dir)
  $file_time = StringRegExpReplace(_NowDate(),'[:/]','-')
ConsoleWrite($file_time&@CR)
$diskshadow_command = $temp_dir&'diskshadow.txt'
ConsoleWrite($diskshadow_command&@CR)
If  FileExists($diskshadow_command) Then
   FileDelete($diskshadow_command)
   EndIf
If Not FileExists($diskshadow_command) Then
   _FileCreate ($diskshadow_command)
   FileWriteLine($diskshadow_command, 'set context persistent')
   FileWriteLine($diskshadow_command, 'set verbose on')
   FileWriteLine($diskshadow_command, 'set metadata c:\VeeamVssHelper\metadata.cab')
   FileWriteLine($diskshadow_command, 'begin backup')
   ConsoleWrite('$local_array_volums_col '&$local_array_volums_col&@CR)
   for $i= 0 to $local_array_volums_col - 1 Step + 1
	  $volume_tempo = StringRegExpReplace($local_array_volums[$i][0],'[\]','')
	 ;$volume_tempo = StringTrimRight($local_array_volums[$i][0], 1)

   FileWriteLine($diskshadow_command, 'add volume '& $volume_tempo&' provider '&$local_array_provaiders[0][1])
   ;FileWriteLine($diskshadow_command, 'add volume'& $volume_tempo&' alias VolumeC')
   ConsoleWrite('add volume'& $volume_tempo&' provider '&$local_array_provaiders[0][1]&@CR)
Next
 ConsoleWrite('$local_array_writersid '&$local_array_volums_col&@CR)
for $i= 0 to $local_array_writersid_col - 1 Step + 1
   FileWriteLine($diskshadow_command, 'writer verify '& $local_array_writersid[$i][1])
   ConsoleWrite('writer verify '& $local_array_writersid[$i][1]&@CR)
   Next
   FileWriteLine($diskshadow_command, 'create')
   FileWriteLine($diskshadow_command, 'end backup')
   FileWriteLine($diskshadow_command, 'exit')
   EndIf
$output_log = $ProgrammDir&'VSS_output_'&$PCName&'-'&$file_time&'.log'
If FileExists($output_log) Then FileDelete ($output_log)
#Region vss list writers
$cmd_pid = Run("C:\WINDOWS\system32\cmd.exe")
WinWaitActive("C:\WINDOWS\system32\cmd.exe")
$title_1st = WinGetTitle('[LAST]')
send('DISKSHADOW.EXE /S '&$diskshadow_command&' /l '&$output_log & "{ENTER}")
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



If FileExists($temp_dir) Then DirRemove($temp_dir)
   If FileExists($list_writers_command) Then FileDelete($temp_dir)

