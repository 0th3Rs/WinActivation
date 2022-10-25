#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=CompleteActivationx86.Exe
#AutoIt3Wrapper_Outfile_x64=CompleteActivationx64.Exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

$colItems = ""
$Output=""
$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", 0x30)

If IsObj($colItems) then
   For $objItem In $colItems
      $Output = $objItem.Caption & " " & $objItem.Version & @CRLF
   Next
Else
   Msgbox(0,"WMI Output","No WMI Objects Found for class: " & "Win32_OperatingSystem" )
Endif

#Region ### START Koda GUI section ### Form=c:\users\wa007\desktop\formgui.kxf
$FormGUI = GUICreate("Windows Activation System", 615, 298, 192, 124)
$Title = GUICtrlCreateLabel("Windows Activation System", 48, 0, 526, 75)
GUICtrlSetFont(-1, 30, 800, 0, "Niramit")
$Subtitle = GUICtrlCreateLabel("Click your version of Windows below", 144, 80, 321, 29)
GUICtrlSetFont(-1, 15, 400, 0, "MS Sans Serif")
$Home = GUICtrlCreateButton("Home", 80, 144, 75, 25)
$HomeN = GUICtrlCreateButton("Home N", 80, 176, 75, 25)
$HomeSL = GUICtrlCreateButton("Home SL", 80, 240, 75, 25)
$HomeCS = GUICtrlCreateButton("Home CS", 80, 208, 75, 25)
$Pro = GUICtrlCreateButton("Pro", 208, 144, 75, 25)
$ProN = GUICtrlCreateButton("Pro N", 208, 176, 75, 25)
$EDU = GUICtrlCreateButton("EDU", 336, 144, 75, 25)
$EDUN = GUICtrlCreateButton("EDU N", 336, 176, 75, 25)
$Enterprise = GUICtrlCreateButton("Enterprise", 456, 144, 75, 25)
$EnterpriseN = GUICtrlCreateButton("Enterprise N", 456, 176, 75, 25)
$WinVer = GUICtrlCreateLabel($Output, 136, 112, 300, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $Home
			RunWait(@ComSpec & " /c slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $HomeN
			RunWait(@ComSpec & " /c slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $HomeSL
			RunWait(@ComSpec & " /c slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $HomeCS
			RunWait(@ComSpec & " /c slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $Pro
			RunWait(@ComSpec & " /c slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $ProN
			RunWait(@ComSpec & " /c slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $EDU
			RunWait(@ComSpec & " /c slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $EDUN
			RunWait(@ComSpec & " /c slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $Enterprise
			RunWait(@ComSpec & " /c slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
		Case $EnterpriseN
			RunWait(@ComSpec & " /c slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /skms kms8.msguides.com", "", @SW_HIDE)
			RunWait(@ComSpec & " /c slmgr /ato", "", @SW_HIDE)
	EndSwitch
WEnd