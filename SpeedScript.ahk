#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;--------Nutraukti scripto vykdyma
Escape::
	ExitApp
Return

;---------Paleisti scripta
^j::
	;-----------------------------------------------------
	excelTitle = pozicijos	;	EXCELIO FAILO PAVADINIMAS
	;-----------------------------------------------------
	queryNo = 01			;	UZKLAUSOS NUMERIS
	;-----------------------------------------------------
	RunExcel(excelTitle)
	list := CopyDataFromExcel(1, 100)
	CloseExcel(excelTitle)
	RunETP("Rita", queryNo)
	FillInData(list)
	ChromePageWait()
	MsgBox, ,SpeedScript, Done!
	ExitApp 
return

;---------------------------------EXCELIO DALIS-----------------------------
;--------Funkcija, kuri paleidzia parametre name nurodytu vardu esanti excel faila
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

;--------Funkcija, kuri kopijuoja pirmus du excelio stulpelius kol sutinka tuscia cele arba nukopijuoja max eiluciu
;--------Parametrai
;--------startRow - eilute, nuo kurios pradedama kopijuoti
;--------max - maksimalus kopijuojamu eiluciu skaicius, kuri pasiekus nustojama kopijuoti

;---pastaba: startRow adresa reiktu paduot kad veliau galima butu tiesiog ji vel paduot
CopyDataFromExcel(startRow, max)
{
	data = 
	localRow = 1
	partNoCell = A%startRow%
	amountCell = B%startRowRow%
	partNo := CopyDataFromCell(partNoCell)
	amount := CopyDataFromCell(amountCell)
	StringLen, length, partNo
	;jei yra detale tai prideam prie galutinio rezultato
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
	StringReplace, data, data, |, ;nukerpam pirma |
	;MsgBox, %data%	
	return data
}

;--------Funkcija, kuri nukopijuoje parametre nurodytos celles duoneis ir grazina kvietejui
CopyDataFromCell(cell)
{
	GoToCell(cell)
	clipboard = 
	Send, {ctrl down}c{ctrl up}
	ClipWait
	copiedData = %clipboard%
	return copiedData
}

;----------Funkcija, kuri nueina i nurodyta excelio cele
GoToCell(cell)
{
	SetTitleMatchMode 2 			;----A window's title can contain WinTitle anywhere inside it to be a match.
	WinActivate Microsoft Excel
	send {f5}
	sleep, 100
	send %cell%
	send {enter}
}

;-----------Funkcija, kuri uzdaro excelio langa
CloseExcel(title)
{
	;SetTitleMatchMode 2
	WinClose, %title%
}

;------------------------------ETP DALIS-----------------------------
;---------Funkcija, kuri atidaro ETP svetaine ir paruosia uzklausos vedimui	
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

;----------Funkcija, kuri palaukia kol uzsikrauna chrome puslapis (reikia dar taisyt)
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


;----------Funkcija, kuri supildo duomenis i ETP ir pasubmitina
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
		;Send {Enter}
	}
}

;---------------------UTILITIES--------------------
;----------Funkcija, kuri spaudzia tab klavisa tiek kartu, kiek nurodyta count parametre
PressTab(count)
{
	if count is integer
	{
		counter = 0
		Loop
		{
			;-----Ejimui per cikla ir jo nutraukimui reikalingos operacijos
			counter += 1
			if (counter>count)
				break
			
			Send, {Tab}
			;Sleep, 100
		}
	}
	else
		MsgBox, , Error, variable passed to function PressTab is not of type integer. Variable value: %count%
}	