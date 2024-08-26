#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <Timers.au3>
#include <Misc.au3>
#AutoIt3Wrapper_icon=1630669841070.ico
HotKeySet("{F1}", "RestartProcess")
HotKeySet("{ESC}","ExitScript")

#RequireAdmin
Opt("WinTitleMatchMode", 2)

; Initialize variables
Global $sConfigFile = @AppDataDir & "\AuctionConfig\.env"
Global $sSaleNumber, $sBidders, $sUsername, $sPassword, $sAuctionCenter, $sStatus, $sExeDirectory, $sCertificateSelection
Global $aBidderNames = ["kolkata1", "kolkata2", "guwahatileaf", "guwahatidust", "teaserver", "coimbatore", "cochin", "coonoor", "siliguri"]
Global $idSave, $idEscLabel, $idF1Label
Global $firstSuccesfulData = False  ; New variable to track first successful data

; New global variable for log file
Global $sLogFile = @AppDataDir & "\AuctionConfig\auction_log.txt"

$hGUI = GUICreate("Configuration", 450, 620)

; Sale Number
GUICtrlCreateLabel("Sale Number:", 20, 20, 100, 20)
$idSaleNumber = GUICtrlCreateInput("", 130, 18, 280, 25)

; Bidders
GUICtrlCreateLabel("Bidders:", 20, 55, 100, 20)
$idBidders = GUICtrlCreateCheckboxes(130, 55, 280, 180)

; Username
GUICtrlCreateLabel("Username:", 20, 245, 100, 20)
$idUsernameInput = GUICtrlCreateInput("", 130, 243, 280, 25)

; Password
GUICtrlCreateLabel("Password:", 20, 280, 100, 20)
$idPasswordInput = GUICtrlCreateInput("", 130, 278, 280, 25, $ES_PASSWORD)

; PIN
GUICtrlCreateLabel("PIN:", 20, 315, 100, 20)
$idPinInput = GUICtrlCreateInput("", 130, 313, 280, 25, $ES_PASSWORD)

; Auction Center
GUICtrlCreateLabel("Auction Center:", 20, 350, 100, 20)
$idAuctionCenter = GUICtrlCreateInput("", 130, 348, 280, 25)

; Status
GUICtrlCreateLabel("Status:", 20, 385, 100, 20)
$idStatusDI = GUICtrlCreateCheckbox("DI", 130, 383, 60, 20)

; Certificate Selection
GUICtrlCreateGroup("Certificate Selection", 240, 385, 200, 50)
$idCertificateAuto = GUICtrlCreateRadio("Automatic", 250, 405, 80, 20)
$idCertificateManual = GUICtrlCreateRadio("Manual", 340, 405, 80, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group

; Start Button
$idSave = GUICtrlCreateButton("Start", 130, 450, 200, 35)

; New Labels
$idEscLabel = GUICtrlCreateLabel("Press ESC to exit program", 20, 520, 400, 20)
GUICtrlSetState($idEscLabel, $GUI_HIDE)
$idF1Label = GUICtrlCreateLabel("Press F1 to restart program", 20, 550, 400, 20)
GUICtrlSetState($idF1Label, $GUI_HIDE)

GUISetState(@SW_SHOW)

; Load configuration on startup
LoadConfig()

; New function to write log entries
Func WriteLog($sMessage)
    Local $hFile = FileOpen($sLogFile, $FO_APPEND)
    If $hFile = -1 Then
        ConsoleWrite("Unable to open log file for writing." & @CRLF)
        Return
    EndIf
    
    Local $sLogEntry = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " - " & $sMessage
    FileWriteLine($hFile, $sLogEntry)
    FileClose($hFile)
EndFunc

; Main loop
While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            ExitScript()
        
        Case $idSave
            If SaveConfig() Then
                ; Disable the Start button
                GUICtrlSetState($idSave, $GUI_DISABLE)
                
                ; Show, highlight and enlarge the ESC and F1 labels
                ShowAndHighlightLabels()
                
                OpenAppAndLogin()
            EndIf
    EndSwitch
WEnd

Func ShowAndHighlightLabels()
    ; Show the labels
    GUICtrlSetState($idEscLabel, $GUI_SHOW)
    GUICtrlSetState($idF1Label, $GUI_SHOW)
    
    ; Highlight and enlarge the labels
    GUICtrlSetColor($idEscLabel, 0xFF0000)  ; Red color
    GUICtrlSetColor($idF1Label, 0xFF0000)  ; Red color
    GUICtrlSetFont($idEscLabel, 12, 800)  ; Increase size to 12, 800 is bold
    GUICtrlSetFont($idF1Label, 12, 800)  ; Increase size to 12, 800 is bold
EndFunc

Func GUICtrlCreateCheckboxes($iLeft, $iTop, $iWidth, $iHeight)
    Local $aCheckboxes[UBound($aBidderNames)]
    
    For $i = 0 To UBound($aBidderNames) - 1
        $aCheckboxes[$i] = GUICtrlCreateCheckbox($aBidderNames[$i], $iLeft, $iTop + ($i * 20), $iWidth, 20)
    Next
    
    Return $aCheckboxes
EndFunc

Func SaveConfig()
    WriteLog("Saving configuration")
    ; Store all entered values in variables
    $sSaleNumber = GUICtrlRead($idSaleNumber)
    $sUsername = GUICtrlRead($idUsernameInput)
    $sPassword = GUICtrlRead($idPasswordInput)
    Global $sPin = GUICtrlRead($idPinInput)
    $sAuctionCenter = GUICtrlRead($idAuctionCenter)
    
    Local $aSelectedBidders[0]
    For $i = 0 To UBound($idBidders) - 1
        If BitAND(GUICtrlRead($idBidders[$i]), $GUI_CHECKED) Then
            _ArrayAdd($aSelectedBidders, $aBidderNames[$i])
        EndIf
    Next
    $sBidders = _ArrayToString($aSelectedBidders, ",")
    
    $sStatus = GUICtrlRead($idStatusDI) = $GUI_CHECKED ? "DI" : "All"
    
    $sCertificateSelection = GUICtrlRead($idCertificateAuto) = $GUI_CHECKED ? "Auto" : "Manual"
    
    Local $sConfig = "SALE_NUMBER=" & $sSaleNumber & @CRLF & _
                     "BIDDERS=" & $sBidders & @CRLF & _
                     "USERNAME=" & $sUsername & @CRLF & _
                     "PASSWORD=" & $sPassword & @CRLF & _
                     "PIN=" & $sPin & @CRLF & _
                     "AUCTION_CENTER=" & $sAuctionCenter & @CRLF & _
                     "STATUS=" & $sStatus & @CRLF & _
                     "CERTIFICATE_SELECTION=" & $sCertificateSelection & @CRLF & _
                     "EXE_DIRECTORY=" & $sExeDirectory

    ; Ensure the directory exists
    Local $sConfigDir = @AppDataDir & "\AuctionConfig"
    If Not FileExists($sConfigDir) Then
        DirCreate($sConfigDir)
    EndIf

    Local $hFile = FileOpen($sConfigFile, 2)  ; 2 = Overwrite mode
    If $hFile = -1 Then
        WriteLog("Error: Unable to open file for writing")
        MsgBox(48, "Error", "Unable to open file for writing.")
        Return False
    EndIf
    
    FileWrite($hFile, $sConfig)
    FileClose($hFile)
    
    WriteLog("Configuration saved successfully")
    Return True
EndFunc

Func LoadConfig()
    WriteLog("Loading configuration")
    If FileExists($sConfigFile) Then
        Local $aConfig = FileReadToArray($sConfigFile)
        Local $aSavedBidders = []
        
        For $sLine In $aConfig
            $aParts = StringSplit($sLine, "=", 2)
            If UBound($aParts) = 2 Then
                Switch $aParts[0]
                    Case "SALE_NUMBER"
                        $sSaleNumber = $aParts[1]
                        GUICtrlSetData($idSaleNumber, $sSaleNumber)
                    Case "BIDDERS"
                        $sBidders = $aParts[1]
                        $aSavedBidders = StringSplit($sBidders, ",", 2)
                    Case "USERNAME"
                        $sUsername = $aParts[1]
                        GUICtrlSetData($idUsernameInput, $sUsername)
                    Case "PASSWORD"
                        $sPassword = $aParts[1]
                        GUICtrlSetData($idPasswordInput, $sPassword)
                    Case "PIN"
                        $sPin = $aParts[1]
                        GUICtrlSetData($idPinInput, $sPin)
                    Case "AUCTION_CENTER"
                        $sAuctionCenter = $aParts[1]
                        GUICtrlSetData($idAuctionCenter, $sAuctionCenter)
                    Case "STATUS"
                        $sStatus = $aParts[1]
                        If $sStatus = "DI" Then
                            GUICtrlSetState($idStatusDI, $GUI_CHECKED)
                        Else
                            GUICtrlSetState($idStatusDI, $GUI_UNCHECKED)
                        EndIf
                    Case "CERTIFICATE_SELECTION"
                        $sCertificateSelection = $aParts[1]
                        If $sCertificateSelection = "Auto" Then
                            GUICtrlSetState($idCertificateAuto, $GUI_CHECKED)
                        Else
                            GUICtrlSetState($idCertificateManual, $GUI_CHECKED)
                        EndIf
                    Case "EXE_DIRECTORY"
                        $sExeDirectory = $aParts[1]
                EndSwitch
            EndIf
        Next
        
        ; Set bidder checkboxes
        For $i = 0 To UBound($idBidders) - 1
            If _ArraySearch($aSavedBidders, $aBidderNames[$i]) <> -1 Then
                GUICtrlSetState($idBidders[$i], $GUI_CHECKED)
            Else
                GUICtrlSetState($idBidders[$i], $GUI_UNCHECKED)
            EndIf
        Next
        
        WriteLog("Configuration loaded successfully")
    Else
        WriteLog("No saved configuration found")
        MsgBox(48, "Information", "No saved configuration found. Please enter the configuration.")
    EndIf
EndFunc

; Open the application and 
Func OpenAppAndLogin()
    WriteLog("Attempting to open application and login")
    Local $sAppExecutable = "PANAuctioneer.exe"
    Local $sCommand = 'cmd.exe /C "cd /d ' & $sExeDirectory & ' && ' & $sAppExecutable & '"'
    
    Local $iPID = Run($sCommand, "", @SW_HIDE)
    
    If @error Then
        WriteLog("Failed to run the application. Error code: " & @error)
        MsgBox(48, "Error", "Failed to run the application. Error code: " & @error)
    Else
        WriteLog("Application launched successfully")
        Sleep(3000)
        Login($sUsername, $sPassword, $sAuctionCenter)
        Sleep(2000)
        DetectAndHandlePasswordExpirePopup()
        Sleep(15000)
        InitialLoading()
        
        While 1
            ContinuousProcessing()  ; This will return if a disconnection is detected
            If Not CheckForDisconnection() Then
                WriteLog("Exiting script due to non-disconnection error")
                ExitScript()  ; Exit if ContinuousProcessing ended for a reason other than disconnection
            EndIf
            WriteLog("Disconnection detected and handled, restarting process")
        WEnd
    EndIf
EndFunc

; Login function
Func Login($sUsername, $sPassword, $sAuctionCenter)
    WriteLog("Attempting to login with username: " & $sUsername & " and auction center: " & $sAuctionCenter)
    WinActivate("Auction")
    ;Enter the username
    ControlSend("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.EDIT\.app\..*)$; INSTANCE:1]", $sUsername)
    ;Enter the password
    ControlSend("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.EDIT\.app\..*)$; INSTANCE:2]", $sPassword)
    ;Click Incremental Checkbox
    ControlClick("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:1]")
    ;Set Auction Center
    For $i = 1 To 10
        ControlSend("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:1]", "{down}")
    Next
    Local $upCount
    Switch $sAuctionCenter
        Case "guwahatidust"
            $upCount = 5
        Case "cochin"
            $upCount = 8
        Case "coonoor"
            $upCount = 6
        Case "coimbatore"
            $upCount = 7
        Case "teaserver"
            $upCount = 0
        Case "guwahatileaf"
            $upCount = 4
        Case "kolkata1"
            $upCount = 3
        Case "siliguri"
            $upCount = 1
        Case "kolkata2"
            $upCount = 2
    EndSwitch
    For $i = 1 To $upCount
        ControlSend("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:1]", "{up}")
	Next
    ;Click Login Button
    ControlClick("Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:2]")
    WriteLog("Login attempt completed")
EndFunc

Func IsControlPresent($timeout)
    WriteLog("Checking if control is present")
    Local $startTime = TimerInit()
    If WinExists("e-Auction") Then WinActivate("e-Auction")
    Do
        Local $handle = ControlGetHandle("","","[REGEXPCLASS:^(WindowsForms10\.Window\.8\.app\..*)$; INSTANCE:16]")
        If $handle <> 0 Then
            WriteLog("Control found")
            Return $handle
        EndIf
        
        Sleep(1000)
    Until TimerDiff($startTime) >= $timeout
    
    WriteLog("Control not found within timeout")
    Return 0
EndFunc

Func WaitForPleaseWaitToDisappear()
    WriteLog("Waiting for 'Please Wait' to disappear")
    Local $startTime = TimerInit()
    Local $timeout = 600000  ; 10 minutes in milliseconds
    
    While TimerDiff($startTime) < $timeout
        If WinExists("e-Auction") Then WinActivate("e-Auction")
        
        Local $handle = ControlGetHandle("","","[REGEXPCLASS:WindowsForms10\.STATIC\.app\..*]")
        If $handle = 0 Or Not StringInStr(ControlGetText("", "", $handle), "Please Wait") Then
            WriteLog("'Please Wait' disappeared")
            Return True  ; "Please Wait" not found, exit the function and continue the script
        Else
            Sleep(1000)  ; "Please Wait" found, sleep for 1 second
        EndIf
    WEnd
    WriteLog("WaitForPleaseWaitToDisappear: Timeout reached after 10 minutes.")
    Return False
EndFunc

Func InitialLoading()
    WriteLog("Starting initial loading")
    WinActivate("e-Auction")
    Local $timeout = 600000  ; 30 seconds
    Local $handle = IsControlPresent($timeout)
    If $handle <> 0 Then
        Sleep(5000)
        OpenDO()
    Else
        IsControlPresent($timeout)
    EndIf
    WriteLog("Initial loading completed")
EndFunc

Func CheckDealBook()
    WriteLog("Checking deal book")
    Local $startTime = TimerInit()
    Local $timeout = 600000  ; 30 seconds
    If WinExists("e-Auction") Then WinActivate("e-Auction")
    Do
        Local $handle = ControlGetHandle("","","[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:6]")
        If $handle <> 0 Then
            WriteLog("Deal book found")
            Return True
        EndIf
        
        Sleep(1000)
    Until TimerDiff($startTime) >= $timeout
    WriteLog("Deal book not found within timeout")
    Return False
EndFunc

Func OpenDO()
    WriteLog("Opening DO")
    If WinExists("e-Auction") Then
        WinActivate("e-Auction")
        Send("{F11}")
        Sleep(5000)
    EndIf
    WriteLog("DO opened")
EndFunc

Func SetSaleNumber($targetSaleNumber)
    WriteLog("Setting sale number to " & $targetSaleNumber)
    If Not WinActive("e-Auction") Then
        WinActivate("e-Auction")
    EndIf
    Sleep(5000)
    
    ControlFocus("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:4]")
    
    ; Go to the top of the list
    Send("{HOME}")
    Sleep(500)
    
    Local $currentSaleNumber = ""
    Local $attempts = 0
    Local $maxAttempts = 100
    Local $targetFound = False
    Local $closestLowerNumber = ""
    
    ; Search from the top
    While $attempts < $maxAttempts
        $currentSaleNumber = ControlGetText("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:4]")
        
        If $currentSaleNumber = $targetSaleNumber Then
            $targetFound = True
            ExitLoop
        ElseIf Number($currentSaleNumber) < Number($targetSaleNumber) And ($closestLowerNumber = "" Or Number($currentSaleNumber) > Number($closestLowerNumber)) Then
            $closestLowerNumber = $currentSaleNumber
        EndIf
        
        Send("{DOWN}")
        Sleep(100)
        
        $attempts += 1
    WEnd
    
    ; Select the appropriate number
    If $targetFound Then
        $finalSelection = $targetSaleNumber
    ElseIf $closestLowerNumber <> "" Then
        $finalSelection = $closestLowerNumber
    Else
        $finalSelection = $currentSaleNumber  ; Fallback to the last number in the list if no lower number found
    EndIf
    
    ; Set the selected number
    Send("{HOME}")
    Sleep(100)
    While ControlGetText("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:4]") <> $finalSelection
        Send("{DOWN}")
        Sleep(100)
    WEnd
    
    Send("{ENTER}")
    
    Sleep(500)
    $finalSelection = ControlGetText("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:4]")
    
    If $finalSelection = $targetSaleNumber Then
        WriteLog("Successfully set Sale Number to " & $targetSaleNumber)
    ElseIf $finalSelection = $closestLowerNumber Then
        WriteLog("Target Sale Number " & $targetSaleNumber & " not found. Selected closest lower number: " & $finalSelection)
    Else
        WriteLog("Failed to set Sale Number. Current value: " & $finalSelection)
    EndIf
    
    Return $finalSelection
EndFunc

Func SetBidder($bidder)
    WriteLog("Setting bidder to " & $bidder)
    WinActivate("e-Auction")
    For $i = 1 To 10
        ControlSend("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:3]", "{down}")
    Next
    Local $upCount
    Switch $bidder
        Case "kolkata2"
            $upCount = 1
        Case "teaserver"
            $upCount = 2
        Case "coimbatore"
            $upCount = 3
        Case "coonoor"
            $upCount = 4
        Case "cochin"
            $upCount = 5
        Case "siliguri"
            $upCount = 6
        Case "guwahatidust"
            $upCount = 7
        Case "kolkata1"
            $upCount = 8
        Case Else
            $upCount = 0
    EndSwitch
    For $i = 1 To $upCount
        ControlSend("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:3]", "{up}")
    Next
    WriteLog("Bidder set to " & $bidder)
EndFunc

Func CheckForDisconnection()
    WriteLog("Checking for disconnection")
    If WinExists("E-Auction System", "You have been disconnected. Please relogin") Then
        WinActivate("E-Auction System", "You have been disconnected. Please relogin")
        WriteLog("Disconnection detected. Restarting process...")
        ; Click the OK button on the popup
        Send("{ENTER}")
        ControlClick("E-Auction System", "You have been disconnected. Please relogin", "[CLASS:BUTTON;INSTANCE:1")
        RestartProcess()
        Sleep(1000) ; Wait for the popup to close
        Return True
    EndIf
    WriteLog("No disconnection detected")
    Return False
EndFunc

Func RestartProcess()
    $firstSuccesfulData = False
    WriteLog("Restarting process")
    
    ; Force close the application
    Local $sWindowTitle = "e-Auction"
    Local $bClosed = False
    
    ; Find all windows matching the title
    Local $aWinList = WinList($sWindowTitle)
    
    If $aWinList[0][0] > 0 Then
        For $i = 1 To $aWinList[0][0]
            Local $hWnd = $aWinList[$i][1]
            Local $iPID = WinGetProcess($hWnd)
            
            ; Try to close the window gracefully
            WinClose($hWnd)
            
            ; Wait for up to 5 seconds for the window to close
            Local $iTimeout = 5000
            Local $hTimer = TimerInit()
            While WinExists($hWnd) And TimerDiff($hTimer) < $iTimeout
                Sleep(100)
            WEnd
            
            ; If the window still exists, force close it
            If WinExists($hWnd) Then
                ; Use WinKill to forcefully close the window
                WinKill($hWnd)
                
                ; If WinKill fails, use taskkill on the process
                If WinExists($hWnd) And $iPID <> -1 Then
                    RunWait(@ComSpec & " /c taskkill /F /PID " & $iPID, "", @SW_HIDE)
                EndIf
            EndIf
        Next
        
        ; Final check and process termination
        $aWinList = WinList($sWindowTitle)
        If $aWinList[0][0] > 0 Then
            WriteLog("Warning: Unable to close all instances of " & $sWindowTitle & " using window methods. Attempting process termination.")
            
            ; Try to terminate all instances of PANAuctioneer.exe
            RunWait(@ComSpec & " /c taskkill /F /IM PANAuctioneer.exe", "", @SW_HIDE)
            
            Sleep(2000)  ; Wait for 2 seconds
            
            ; Final check after process termination
            $aWinList = WinList($sWindowTitle)
            If $aWinList[0][0] = 0 Then
                $bClosed = True
            Else
                WriteLog("Error: Failed to close all instances of " & $sWindowTitle)
            EndIf
        Else
            $bClosed = True
        EndIf
    Else
        $bClosed = True  ; No windows found, consider it closed
    EndIf
    
    If $bClosed Then
        WriteLog("All instances of " & $sWindowTitle & " have been successfully closed")
    Else
        WriteLog("Warning: Unable to close all instances of " & $sWindowTitle & ". Attempting to continue anyway.")
    EndIf
    
    ; Wait a moment to ensure the application is fully closed
    Sleep(2000)
    
    ; Restart the application
    Local $startTime = TimerInit()
    Local $timeout = 60000 ; 60 seconds timeout
    
    While 1
        OpenAppAndLogin()
        
        ; Check if the application started successfully
        If WinExists("e-Auction") Then
            WriteLog("Application restarted successfully")
            ExitLoop
        EndIf
        
        Sleep(5000) ; Wait 5 seconds before trying again
    WEnd
    
    ; Re-enable the Start button and reset label appearance
    GUICtrlSetState($idSave, $GUI_ENABLE)
    ResetLabels()
    WriteLog("Process restarted")
EndFunc

Func ResetLabels()
    WriteLog("Resetting labels")
    ; Hide the labels
    GUICtrlSetState($idEscLabel, $GUI_HIDE)
    GUICtrlSetState($idF1Label, $GUI_HIDE)
    
    ; Reset color and font (in case they're shown again later)
    GUICtrlSetColor($idEscLabel, 0x000000)  ; Black color
    GUICtrlSetColor($idF1Label, 0x000000)  ; Black color
    GUICtrlSetFont($idEscLabel, 9, 400)  ; 9 is the size, 400 is normal weight
    GUICtrlSetFont($idF1Label, 9, 400)
    WriteLog("Labels reset")
EndFunc

Func ExitScript()
    WriteLog("Exiting script")
    ; Re-enable the Start button and reset labels before exiting
    GUICtrlSetState($idSave, $GUI_ENABLE)
    ResetLabels()
    
    GUIDelete($hGUI)
    ProcessClose(@AutoItPID)
EndFunc

Func SetStatus($status)
    WriteLog("Setting status to " & $status)
    WinActivate("e-Auction")
    For $i = 1 To 10
        ControlSend("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:2]", "{down}")
    Next
    Local $upCount
    Switch $status
        Case "All"
            $upCount = 5
        Case "DI"
            $upCount = 3
    EndSwitch
    For $i = 1 To $upCount
        ControlSend("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.COMBOBOX\.app\..*)$; INSTANCE:2]", "{up}")
    Next
    WriteLog("Status set to " & $status)
EndFunc

Func HandleSecurityDialog()
    WriteLog("Handling security dialog")
    Local $hWnd = WinWait("Windows Security", "", 1000)
    If $hWnd Then
        WinActivate($hWnd)
        Send("!o")  ; Send Alt+O
        Sleep(1000)
        WriteLog("Handled Windows Security dialog")
    EndIf
EndFunc

Func ClickRefresh()
    WriteLog("Clicking Refresh")
    WinActivate("e-Auction")
    ControlClick("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:6]")
    Sleep(5000)
    WriteLog("Refresh clicked")
EndFunc

Func ClickSelectUnselectAll()
    WriteLog("Clicking Select/Unselect All")
    WinActivate("e-Auction")
    ControlClick("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:4]")
    Sleep(1000)
    WriteLog("Select/Unselect All clicked")
EndFunc

Func ClickGetPDF()
    WriteLog("Clicking Get PDF")
    WinActivate("e-Auction")
    ControlClick("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:3]")
    Sleep(1000)
    WriteLog("Get PDF clicked")
EndFunc

Func HandleDeliveryInstructionsPopup()
    WriteLog("Handling Delivery Instructions popup")
    Local $timeout = 600000  ; 10 seconds timeout
    Local $start = TimerInit()
    
    While TimerDiff($start) < $timeout
        If WinExists("Delivery") Then
            WinActivate("Delivery")
            Send("{ENTER}")
            WriteLog("Handled Delivery Instructions popup")
Sleep(1000)
            Return
        EndIf
        Sleep(100)
    WEnd
    WriteLog("Delivery Instructions popup not found within timeout")
EndFunc

Func ClickUploadPDF()
    WriteLog("Clicking Upload PDF")
    WinActivate("e-Auction")
    ControlClick("e-Auction", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:2]")
    Sleep(1000)
    WriteLog("Upload PDF clicked")
EndFunc

Func CloseFileUploader()
    WriteLog("Closing File Uploader")
    Local $timeout = 600000  ; 10 seconds timeout
    Local $start = TimerInit()
    
    While TimerDiff($start) < $timeout
        If WinExists("File Uploader") Then
            Sleep(1000)
            WinActivate("File Uploader")
            ControlClick("File Uploader", "", "[REGEXPCLASS:^(WindowsForms10\.BUTTON\.app\..*)$; INSTANCE:1]")
            WinWaitClose("File Uploader", "", 10)
            Sleep(1000)
            WriteLog("File Uploader closed")
            Return
        EndIf
        Sleep(100)
    WEnd
    WriteLog("File Uploader not found within timeout")
EndFunc

Func EnterUserPin()
    WriteLog("Entering user PIN")
    ; Wait for the "Verify User PIN" window to appear
    WinWait("Verify User PIN", "", 1000)  ; Wait up to 10 seconds for the window to appear
    
    If WinExists("Verify User PIN") Then
        ; Activate the "Verify User PIN" window
        WinActivate("Verify User PIN")
        
        ; Wait for the window to be active
        WinWaitActive("Verify User PIN", "", 1000)  ; Wait up to 5 seconds for the window to become active
        
        ; Send the PIN
        Send($sPin)  ; We're using the global $sPin variable here
        
        ; Add a small delay to ensure the PIN is entered completely
        Sleep(500)
        
        ; Click the Login button
        ControlClick("Verify User PIN", "", "[Class:Button;Instance:2]")
        
        ; Wait for the PIN verification window to close
        WinWaitClose("Verify User PIN", "", 1000)  ; Wait up to 10 seconds for the window to close
        
        WriteLog("User PIN entered successfully")
        Return True  ; PIN entered successfully
    Else
        WriteLog("Failed to enter PIN: Verify User PIN window not found")
        Return False  ; Failed to enter PIN
    EndIf
EndFunc

Func ContinuousProcessing()
    $firstSuccesfulData = False
    WriteLog("Starting continuous processing")
    Local $aSaleNumbers = [$sSaleNumber, $sSaleNumber, $sSaleNumber-1, $sSaleNumber, $sSaleNumber-2, $sSaleNumber, $sSaleNumber-1]  ; Fixed array of sale numbers
    Local $currentIndex = 0
    Local $lastActivityTime = TimerInit()
    Local $inactivityTimeout = 10 * 60 * 1000  ; 10 minutes in milliseconds

    While 1
        If CheckForDisconnection() Then
            WriteLog("Disconnection detected. Moving to next bidder.")
            ContinueLoop
        EndIf

        WinActivate("e-Auction")
        SetStatus($sStatus)
        $currentSaleNumber = $aSaleNumbers[$currentIndex]
        ProcessDealBook($currentSaleNumber)

        $currentIndex += 1
        If $currentIndex >= UBound($aSaleNumbers) Then
            $currentIndex = 0  ; Reset to the beginning of the array
        EndIf
       
        ; Check for inactivity
        If TimerDiff($lastActivityTime) > $inactivityTimeout Then
            WriteLog("Bot inactive for more than 10 minutes. Moving to next bidder.")
            ContinueLoop
        EndIf
        
        ; Reset the activity timer after each cycle
        $lastActivityTime = TimerInit()

        Sleep(5000)  ; Wait 5 seconds before starting the next cycle
    WEnd
EndFunc

Func DetectAndHandlePopup()
    WriteLog("Detecting and handling popup")
    ; Wait for the popup window to appear, timeout after 5 seconds
    Local $hWnd = WinWait("Delivery Instructions / Advice", "Data Not found", 5)
    
    If $hWnd Then
        WriteLog("Popup detected")
        
        ; Click the OK button
        ControlFocus($hWnd, "", "OK")
        send("{ENTER}")
        ControlClick($hWnd, "", "OK")
        Return True
    Else
        WriteLog("No popup detected")
        Return False
    EndIf
EndFunc

Func noDO()
    WriteLog("Detecting and handling no DO popup")
    ; Wait for the popup window to appear, timeout after 5 seconds
    Local $hWnd = WinWait("Delivery Instructions / Advice", "Sorry, you have not selected any Delivery Instructions", 5)
    
    If $hWnd Then
        WriteLog("Popup detected")
        
        ; Click the OK button
        ControlFocus($hWnd, "", "OK")
        send("{ENTER}")
        ControlClick($hWnd, "", "OK")
        Return True
    Else
        WriteLog("No popup detected")
        Return False
    EndIf
EndFunc

Func DetectAndHandlePasswordExpirePopup()
    WriteLog("Detecting and handling password expiry popup")
    ; Wait for the popup window to appear, timeout after 5 seconds
    Local $hWnd = WinWait("TeaAuction", "Your password will expire", 5)
    
    If $hWnd Then
        WriteLog("Password expiry popup detected")
        
        ; Click the OK button
        ControlClick($hWnd, "", "OK")
        
        Return True
    Else
        WriteLog("No password expiry popup detected")
        Return False
    EndIf
EndFunc

Func ProcessDealBook($currentSaleNumber)
    Local $aBidders = StringSplit($sBidders,",",2)
    WinActivate("Auction")
    SetSaleNumber($currentSaleNumber)
    For $i = 0 To UBound($aBidders) - 1
        If CheckForDisconnection() Then
            WriteLog("Disconnection detected. Moving to next bidder.")
            ContinueLoop
        EndIf
        WinActivate("Auction")
        SetBidder($aBidders[$i])
        ClickRefresh()
        If Not WaitForPleaseWaitToDisappear() Then
            WriteLog("Timeout waiting for 'Please Wait' to disappear. Moving to next bidder.")
            ContinueLoop
        EndIf
        If DetectAndHandlePopup() Then
            WriteLog("No data found for bidder: " & $aBidders[$i] & ". Moving to next bidder.")
            ContinueLoop
        EndIf
        ClickSelectUnselectAll()
        ClickGetPDF()
        If noDO() Then
            WriteLog("No data found for bidder: " & $aBidders[$i] & ". Moving to next bidder.")
            ContinueLoop
        EndIf
        HandleDeliveryInstructionsPopup()
        If Not $firstSuccesfulData Then
            If $sCertificateSelection = "Auto" Then
                HandleSecurityDialog()
            Else
                MsgBox(0, "Certificate", "Please select the correct certificate")
            EndIf
            Sleep(2000)
            If Not EnterUserPin() Then
                WriteLog("Failed to enter PIN. Moving to next bidder.")
                ContinueLoop
            EndIf
            $firstSuccesfulData = True
        EndIf
        WinActivate("e-Auction")
        If CheckForDisconnection() Then
            WriteLog("Disconnection detected. Moving to next bidder.")
            ContinueLoop
        EndIf
        
        HandleDeliveryInstructionsPopup()
        Sleep(1000)
        WinActivate("e-Auction")
        ClickUploadPDF()
        HandleDeliveryInstructionsPopup()
        Sleep(5000)
        CloseFileUploader()
        WinActivate("e-Auction")

        If CheckForDisconnection() Then
            WriteLog("Disconnection detected. Moving to next bidder.")
            ContinueLoop
        EndIf
    Next
    WriteLog("Deal book processing completed for sale number: " & $currentSaleNumber)
EndFunc