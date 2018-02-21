#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;------------Terminate script
Escape::
	ExitApp
Return

;---------Hot key to run the script
^j::
	;-----------------------------------------------------
	excelTitle = pozicijos	;	EXCEL FILE NAME
	;-----------------------------------------------------
	queryNo = 01			;	QUERY NUMBER
	;-----------------------------------------------------
	operator = Rita			;	QUERY OPERATOR
	;----------------------------------------------------
	RunExcel(excelTitle)
	list := CopyDataFromExcel(1, 100)
	Close(excelTitle)
	RunETP(operator, queryNo)
	FillInData(list)
	ChromePageWait()
	MsgBox, ,SpeedScript, Done!
	ExitApp 
return

;--------------------------------------EXCEL PART-----------------------------------
;--------A function that opens an excel file with the name specified in the function parameter
;--------file must have xlsx extension!
RunExcel(name)
{
	Run, %name%.xlsx, , UseErrorLevel
	if ErrorLevel
	{
		MsgBox, The file %name%.xlsx does not exist
	}
	SetTitleMatchMode 2
	WinWait, %name%
}

;--------A function that copies the values of the first two excel file columns
;--------Parameters:
;--------startRow - a row number indicating where to start copying columns
;--------max - the maximum amount of rows to be copied if no empty cell is encountered
;--------Return value:
;--------data - a string containing pairs of part number and amount values where each individual element is separated by '|'
;---pastaba: startRow adresa reiktu paduot kad veliau galima butu tiesiog ji vel paduot
CopyDataFromExcel(startRow, max)
{
	data = 
	localRow = 1
	partNoCell = A%startRow%
	amountCell = B%startRow%
	partNo := CopyDataFromCell(partNoCell)
	amount := CopyDataFromCell(amountCell)
	StringLen, length, partNo
	;if partNo cell is not empty or row number is less or equal to the maximum given row amount
	;we add the partNo and ampunt pair to the return result
	while (length > 2 && localRow <= max)
	{
		data = %data%|%partNo%|%amount%
		startRow++
		localRow++
		partNoCell = A%startRow%
		amountCell = B%startRow%
		partNo := CopyDataFromCell(partNoCell)
		amount := CopyDataFromCell(amountCell)
		StringLen, length, partNo
	}
	StringReplace, data, data, |, ;get rid of the first '|'
	;MsgBox, %data%	
	return data
}

;--------A function that copies the contents of the specified cell and returns them
CopyDataFromCell(cell)
{
	GoToCell(cell)
	clipboard = 
	Send, {ctrl down}c{ctrl up}
	ClipWait
	copiedData = %clipboard%
	return copiedData
}

;----------A function that moves focus to the specified cell
GoToCell(cell)
{
	SetTitleMatchMode 2 			;----A window's title can contain WinTitle anywhere inside it to be a match.
	WinActivate Microsoft Excel
	send {f5}
	sleep, 100
	send %cell%
	send {enter}
}

;-----------A function that closes the window containing the given title
Close(title)
{
	;SetTitleMatchMode 2
	WinClose, %title%
}

;------------------------------------ETP PART---------------------------------	
;---------A function that opens the ETP website and prepares it for the query input
;---------Parameters:
;---------name - the operator name
;---------queryNo - number of the query
RunETP(name, queryNo)
{
	Run, https://www.etpbonomi.it/eng/preventivi/start
	ChromePageWait()
	PressTab(25)
	SendInput %name%
	Send {Tab}
	SendInput %queryNo%
	PressTab(2)
	Send {Enter}
	ChromePageWait()	
	PressTab(31)
}

;---------A function that waits till the webpage is loaded
ChromePageWait()
{
	;tikrinam ar dar kairej rodo, kad krauna
	x = 399
	y = 256
	color = 0xFFFFFF
	seconds = 10
	start := A_TickCount
	while ( (seconds = -1) || (A_TickCount-start <= seconds*1000) )
	{
	  PixelGetColor, Loaded, x, y, RGB
	  if Loaded = color
		 break
	}
	ErrorLevel := 1 ; Page failed to load
	;tikrinam ar dar sukasi ratukas
	while (A_Cursor = "AppStarting")
		continue
	Sleep, 100
}

;--------A funtion that fills in the data copied from the excel file to the website and submits the query
FillInData(data)
{
	SetTitleMatchMode 2
	ifWinExist ETP
	{
		WinActivate
		Loop, Parse, data, |
		{
			SendInput %A_LoopField%
			;sleep, 100
			;MsgBox, %A_LoopField%
			send {Tab}
			sleep,  100
		}
		Send {Enter}
	}
}

;------------------------------UTILITIES---------------------------
;----------A function that presses the Tab key the specified amount of time
PressTab(count)
{
	if count is integer
	{
		counter = 0
		Loop
		{
			counter += 1
			if (counter>count)
				break
			
			Send, {Tab}
		}
	}
	else
	{
		MsgBox, , Error, variable passed to function PressTab is not of type integer. Variable value: %count%
		ExitApp
	}
}	
