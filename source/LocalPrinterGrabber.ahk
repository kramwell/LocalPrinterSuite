;Written by KramWell.com - 10/JAN/2017
;This tool is useful if you want to check multiple systems and view/remove their locally installed printers over a network.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

global PrinterVBSPath :=, LPSPath :=, ComputersAdded :=

LoadSettings()

Gui, LocalPrinterGrabber:Default
Gui, LocalPrinterGrabber:Add, Button, Default x10 y10 gAddFile, Add List
Gui, LocalPrinterGrabber:Add, Button, x65 y10 gClearList, Clear

Gui, LocalPrinterGrabber:Add, Button, x180 y10 gStartPrinterGrab, Run Process

Gui, LocalPrinterGrabber:Add, ListView, x10 y40 w250 h280 vComputerName, Name|Status|Local

Gui, LocalPrinterGrabber:Add, Text, x10 y330 w250 vTotalFount, Found 0 Results

Gui, LocalPrinterGrabber:Add, Text, x10 y350 w250 vFoundPrinters,

Gui, LocalPrinterGrabber:Add, Button, x185 y330 gOpenViewer, Open Viewer

Gui, LocalPrinterGrabber:show, w270 h370, Local Printer Grabber

return

OpenViewer:
Run LocalPrinterViewer.ahk
return

StartPrinterGrab:

Gui, LocalPrinterGrabber:Default
Gui, ListView, ComputerName	

;here we would loop listview, get asset and check if file exist, if so then grab printer info

if (LV_GetCount() != "0"){

dhw := A_DetectHiddenWindows
DetectHiddenWindows On
Run "%ComSpec%" /k,, Hide, pid
while !(hConsole := WinExist("ahk_pid" pid))
	Sleep 10
DllCall("AttachConsole", "UInt", pid)
DetectHiddenWindows %dhw%
objShell := ComObjCreate("WScript.Shell")

CurrentCheckCount :=
row1 :=

	Loop % LV_GetCount()
	{
			CurrentCheckCount++	
			
    LV_GetText(PcName, A_Index)
	Ping4(PcName)
		If (ErrorLevel){
			LV_Modify(A_Index, "Col2", "OFF", A_Space)
		}Else{
			;grab printer info and save to csv?

			SaveListViewLineNumberToVar := A_Index
			
			;grab printer info
			;PrinterVBSPath
			objExec := objShell.Exec("cscript " PrinterVBSPath "  -l -s \\" PcName)
			While !objExec.Status
			Sleep 100
			
			
	
			LocalPrinterResult := objExec.StdOut.ReadAll() ;read the output at once
			
			LocalPrinterResultArray := StrSplit(LocalPrinterResult, "`r`n")
				
				
				
				
					Loop % LocalPrinterResultArray.MaxIndex()
					{
						if (LocalPrinterResultArray[A_Index] != ""){

						;catch output and save to csv
					
							LineText := LocalPrinterResultArray[A_Index]
							
							If InStr(LineText, "Access is denied"){
							MsgBox, 4,, Access Denied for user %A_UserName%`nDo you want to continue?, 7
								IfMsgBox, No
								{
								ExitApp
								}
							}
							
							If InStr(LineText, "Number of local printers and connections enumerated"){
							numbers := RegExReplace(LineText, "[^0-9]")
							;FileAppend, %A_Index%`r`n, dmp.txt
							LV_Modify(SaveListViewLineNumberToVar, "Col2", "ON", numbers)
							
							}
						
							If InStr(LineText, "Printer name"){
							StringTrimLeft, OutputVar, LineText, 13
							row1 := row1 . PcName . "," . OutputVar
							}

							If InStr(LineText, "Share name"){
							StringTrimLeft, OutputVar, LineText, 11
							row1 := row1 . "," . OutputVar
							}

							If InStr(LineText, "Driver name"){
							StringTrimLeft, OutputVar, LineText, 12
							row1 := row1 . "," . OutputVar
							}

							If InStr(LineText, "Port name"){
							StringTrimLeft, OutputVar, LineText, 10
							row1 := row1 . "," . OutputVar							
							}

							If InStr(LineText, "Comment"){
							StringTrimLeft, OutputVar, LineText, 13
							row1 := row1 . "," . OutputVar . ",`r"
							}	
						
						
						} ;end if LocalPrinterResultArray = blank
					} ;end loop max index
			
			LV_ModifyCol()
		} ;end if pc found
			GuiControl,,FoundPrinters, Checking %CurrentCheckCount% of %ComputersAdded% Computers	
	} ;end loop count
	
	;msgbox % row1
				
				if (CurrentCheckCount != ""){
				FileAppend, %row1%, %LPSPath% ;write to csv file		
				
				;here we could sort by portName
				
;				; The following example sorts the contents of a file:
;				FileRead, Contents, %LPSPath%
;				if not ErrorLevel  ; Successfully loaded.
;				{
;				Sort, Contents, -nk5
;				FileDelete, %LPSPath%
;				FileAppend, %Contents%, %LPSPath%
;				Contents =  ; Free the memory.
;				}

	DllCall("FreeConsole")
	Process Exist, %pid%
	if (ErrorLevel == pid){
	Process Close, %pid%
	}					
				
				
				MsgBox % "File written to " LPSPath
				}
				
}else{
MsgBox % "No values to process."
}

return

ClearList:
LV_Delete() ;delete all cells
GuiControl,,TotalFount,
return

AddFile:

FileSelectFile, SelectedFile, 3, , Open a file,
if (SelectedFile != ""){

LV_Delete() ;delete all cells

Loop, read, %SelectedFile%
	{
	
	if (A_LoopReadLine != ""){
	StringReplace, NewStr, A_LoopReadLine, %A_SPACE%, , All
	
	LV_Add("",NewStr)	
	ComputersAdded ++
	}
		
	} ;end loop read
LV_ModifyCol() 
}

GuiControl,,TotalFount, Found %ComputersAdded% Resuts

return

LoadSettings(){

LPSPath := A_ScriptDir . "\LPS.kw"

;here we look for printer mgr ;should assume win7 for now.
PrinterVBSPath := A_WinDir . "\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs"

if !FileExist(PrinterVBSPath)
	{
MsgBox % "prnmngr.vbs has not been found."
	ExitApp
	}

if FileExist(LPSPath)
	{
MsgBox, 4,, LPS.kw exists!`nDo you want to delete?
IfMsgBox Yes
FileDelete, %LPSPath%

	}	
	
}

Ping4(Addr, Timeout := 1024) {

;setdefaulttimeout

   ; ICMP status codes -> http://msdn.microsoft.com/en-us/library/aa366053(v=vs.85).aspx
   ; WSA error codes   -> http://msdn.microsoft.com/en-us/library/ms740668(v=vs.85).aspx
   Static WSADATAsize := (2 * 2) + 257 + 129 + (2 * 2) + (A_PtrSize - 2) + A_PtrSize
   ;OrgAddr := Addr
   ;Result := ""
   ; -------------------------------------------------------------------------------------------------------------------
   ; Initiate the use of the Winsock 2 DLL
   VarSetCapacity(WSADATA, WSADATAsize, 0)
   If (Err := DllCall("Ws2_32.dll\WSAStartup", "UShort", 0x0202, "Ptr", &WSADATA, "Int")) {
      ErrorLevel := "WSAStartup failed with error " . Err
      Return ""
   }
   If !RegExMatch(Addr, "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") { ; Addr contains a name
      If !(HOSTENT := DllCall("Ws2_32.dll\gethostbyname", "AStr", Addr, "UPtr")) {
         DllCall("Ws2_32.dll\WSACleanup") ; Terminate the use of the Winsock 2 DLL
         ErrorLevel := "gethostbyname failed with error " . DllCall("Ws2_32.dll\WSAGetLastError", "Int")
         Return ""
      }
      PAddrList := NumGet(HOSTENT + 0, (2 * A_PtrSize) + 4 + (A_PtrSize - 4), "UPtr")
      PIPAddr   := NumGet(PAddrList + 0, 0, "UPtr")
      Addr := StrGet(DllCall("Ws2_32.dll\inet_ntoa", "UInt", NumGet(PIPAddr + 0, 0, "UInt"), "UPtr"), "CP0")
   }
   INADDR := DllCall("Ws2_32.dll\inet_addr", "AStr", Addr, "UInt") ; convert address to 32-bit UInt
   If (INADDR = 0xFFFFFFFF) {
      ErrorLevel := "inet_addr failed for address " . Addr
      Return ""
   }
   ; Terminate the use of the Winsock 2 DLL
   DllCall("Ws2_32.dll\WSACleanup")
   ; -------------------------------------------------------------------------------------------------------------------
   HMOD := DllCall("LoadLibrary", "Str", "Iphlpapi.dll", "UPtr")
   Err := ""
   If (HPORT := DllCall("Iphlpapi.dll\IcmpCreateFile", "UPtr")) { ; open a port
      REPLYsize := 32 + 8
      VarSetCapacity(REPLY, REPLYsize, 0)
      If DllCall("Iphlpapi.dll\IcmpSendEcho", "Ptr", HPORT, "UInt", INADDR, "Ptr", 0, "UShort", 0
                                            , "Ptr", 0, "Ptr", &REPLY, "UInt", REPLYsize, "UInt", Timeout, "UInt") {
      }
      Else
         Err := "IcmpSendEcho failed with error " . A_LastError
      DllCall("Iphlpapi.dll\IcmpCloseHandle", "Ptr", HPORT)
   }
   Else
      Err := "IcmpCreateFile failed to open a port!"
   DllCall("FreeLibrary", "Ptr", HMOD)
   ; -------------------------------------------------------------------------------------------------------------------
   If (Err) {
      ErrorLevel := Err
      Return ""
   }
   ErrorLevel := 0
   Return
}

LocalPrinterGrabberGuiClose:
Gui,LocalPrinterGrabber:destroy
ExitApp
return