;Written by KramWell.com - 10/JAN/2017
;This tool is useful if you want to check multiple systems and view/remove their locally installed printers over a network.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Global PrinterVBSPath :=, PathToCSVFile:=, ComputerName :=, PrinterName :=, ShareName :=, DriverName :=, PortName :=, Comment :=

LoadSettings()

;Asset,PrinterName,ShareName,DriverName,PortName,Comment
Gui, Font, s10
Gui, Add, TreeView, x10 y15 h400 w180 gMyTreeView

Gui, Add, Text, x230 y20 w200,Computer Name:
Gui, Add, Text, x230 y40 w200 vCOMPUTERNAME,

Gui, Add, Text, x230 y80 w200, Printer Name:
Gui, Add, Text, x230 y100 w200 vPRINTERNAME,

Gui, Add, Text, x230 y140 w200, Driver Name:
Gui, Add, Text, x230 y160 w200 vDRIVERNAME

Gui, Add, Text, x230 y200 w200, Port Name:
Gui, Add, Text, x230 y220 w200 vPORTNAME,

Gui, Add, Text, x230 y260 w200, Share Name:
Gui, Add, Text, x230 y280 w200 vSHARENAME,

Gui, Add, Text, x230 y320 w200, Comment:
Gui, Add, Text, x230 y340 w200 vCOMMENT,

Gui, Add, Button, x220 y380 gDeletePrinter, Delete Printer
Gui, Add, Button, x350 y380 gOpenPrinter, Open Print Manager

LoadTreeView()

Gui, Show, w520 h425 , LocalPrinter Viewer 0.64  ; Show the window and its TreeView.

return

LoadSettings(){

;here we look for printer mgr ;should assume win7 and above for now.
PrinterVBSPath := A_WinDir . "\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs"

if !FileExist(PrinterVBSPath)
	{
MsgBox % "prnmngr.vbs has not been found."
	ExitApp
	}

PathToCSVFile := A_ScriptDir "\LPS.kw"	

if !FileExist(PathToCSVFile)
	{
MsgBox % "csv file has not been found."
	ExitApp
	}
	
}

LoadTreeView(){

field5Change :=
countit := 0

TV_Delete()

Loop, read, %PathToCSVFile%
	{
    Loop, parse, A_LoopReadLine, CSV
		{
		field%a_index%=%A_LoopField%
		}
	
	;field2 = Printer name
	;field3 = Share name
	;field4 = Driver name
	;field5 = Port name
	;field6 = Comment
	
	if (field5 != field5Change){
	countit++
	P%countit% := TV_Add(field5)
	field5Change := field5
	;msgbox % P1
	}
	fieldFull := field1 ; A_Space field2 A_Space field3 A_Space field4 A_Space field5 A_Space field6
	TV_Add(fieldFull, P%countit%)
	
	} ;end loop read

}

GetInfoFromCSVFile(ComputerNameToFind,PortNameToFind){

FoundResult :=

Loop, read, %PathToCSVFile%
	{
    Loop, parse, A_LoopReadLine, CSV
		{
		field%a_index%=%A_LoopField%
		}
	
	;field2 = Printer name
	;field3 = Share name
	;field4 = Driver name
	;field5 = Port name
	;field6 = Comment

	;fieldFull := field1 ; A_Space field2 A_Space field3 A_Space field4 A_Space field5 A_Space field6
	
	if (field1 = ComputerNameToFind AND field5 = PortNameToFind){	
	FoundResult := 1
	break
	}
	
	} ;end loop read

		if (FoundResult = 1){
			;msgbox % field1 A_Space field2 A_Space field3 A_Space field4 A_Space field5 A_Space field6

		;populate feilds
		
		GuiControl,,COMPUTERNAME, %field1%
		GuiControl,,PRINTERNAME, %field2%
		GuiControl,,SHARENAME, %field3%
		GuiControl,,DRIVERNAME, %field4%
		GuiControl,,PORTNAME, %field5%
		GuiControl,,COMMENT, %field6%		
		
		ComputerName := field1
		PrinterName := field2
		ShareName := field3
		DriverName := field4
		PortName := field5
		Comment := field6

		}else{
			MsgBox % "Info not available, try reloading the app."
		}
	
}

RunProgram(TYPEOF){

dhw := A_DetectHiddenWindows
DetectHiddenWindows On
Run "%ComSpec%" /k,, Hide, pid
while !(hConsole := WinExist("ahk_pid" pid))
	Sleep 10
DllCall("AttachConsole", "UInt", pid)
DetectHiddenWindows %dhw%
objShell := ComObjCreate("WScript.Shell")

	if (TYPEOF = "PDELETE"){
		TempOptional := """" . PrinterName . """"
		objExec := objShell.Exec("cscript " PrinterVBSPath "  -d -s \\" ComputerName " -p " TempOptional)
	}

	While !objExec.Status
	Sleep 100

	MessageOutPutResult := objExec.StdOut.ReadAll() ;read the output at once
			
DllCall("FreeConsole")
Process Exist, %pid%
	if (ErrorLevel == pid){
	Process Close, %pid%
	}				
	
	Return MessageOutPutResult
}

DeleteCSVLine(){

PrinterTmpVBSPath := A_Temp . "\lp.kw"

;check if tmp file exists
if !FileExist(PrinterTmpVBSPath)
	{
	FileDelete, %PrinterTmpVBSPath%
	}


;loop csv file and remove selected entry
Loop, read, %PathToCSVFile%
	{
    Loop, parse, A_LoopReadLine, CSV
		{
		field%a_index%=%A_LoopField%
		}
		
	if (field1 = ComputerName AND field2 = PrinterName AND field5 = PortName){
	}else{
	
	FileAppend, %A_LoopReadLine%`r, %PrinterTmpVBSPath% ;this is the new file to replace csv with.
	}
	}

	FileCopy, %PrinterTmpVBSPath%, %PathToCSVFile%, 1
	FileDelete, %PrinterTmpVBSPath%
}

DeletePrinter:
;msgbox % ComputerName A_Space PrinterName A_Space ShareName A_Space DriverName A_Space PortName A_Space Comment

;here we would delete the driver based on pulled results, then delete from csv and refresh list.

if (PrinterName != ""){


MsgBox, 4, , Are you sure you want to delete the printer`n%PrinterName% at %ComputerName%?
IfMsgBox No
    return

ReturnValue := RunProgram("PDELETE")
;could return value and say if 5: then remove from csv and msg user, otherwise 

ReturnValueArray := StrSplit(ReturnValue, "`r`n")	

				if (ReturnValueArray.MaxIndex() = 5){
				
				DeleteCSVLine()
				LoadTreeView()
				
				
				;deleted ok, here we create new file with the line removed -
				;refresh the list.
				
					MsgBox % ReturnValueArray[4]
					;Return
					
				ComputerName :=
				PrinterName :=
				ShareName :=
				DriverName :=
				PortName :=
				Comment :=
							
				GuiControl,,COMPUTERNAME,
				GuiControl,,PRINTERNAME,
				GuiControl,,SHARENAME,
				GuiControl,,DRIVERNAME,
				GuiControl,,PORTNAME,
				GuiControl,,COMMENT,
					
					
				}

				if (ReturnValueArray.MaxIndex() = 9){
					MsgBox % ReturnValueArray[4] ;error
					;Return
				}

;msgbox % ReturnValue A_Space ReturnValueArray.MaxIndex()


		
}		
		
return

OpenPrinter:
DeleteCSVLine()
return

MyTreeView:  ; This subroutine handles user actions (such as clicking).

if (A_GuiEvent = "DoubleClick")
{
TV_GetText(SelectedItemText, A_EventInfo)

    ParentID := TV_GetParent(A_EventInfo)
    if (ParentID != 0){
	
	TV_GetText(ParentText, ParentID)
	
    ;msgbox % ParentText A_Space SelectedItemText
	;here we search csv for values and display results in right pane
	GetInfoFromCSVFile(SelectedItemText,ParentText)
	
	
	} ;end if parent id is 0

} ;end if double click on treeview

return

GuiClose:  ; Exit the script when the user closes the TreeView's GUI window.
ExitApp	