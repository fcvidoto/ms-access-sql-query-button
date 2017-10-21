#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Persistent
SetTitleMatchMode, 2 ; partial name of the file
DetectHiddenWindows, on
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; #Warn  ; Enable warnings to assist with detecting common errors.
; ______________________________________
; // tray icon menu
mountMenu()
; ______________________________________
; // globals
global accessAddress := ; access database address
global acc := ; access instance
global filename := ; 
global winPressed := ;
global recentFilesAddress := [] ;
; ______________________________________
; // hotkeys
#n:: main(false) ; selects a new db and create a query
+esc:: ExitApp
; ______________________________________
; // functions
main(onlySelectsDB) { ; sends the command to create a new query
	if (accessAddress == "")
		accessAddress := fileDialogAccess() ; get the database path
		if (accessAddress == "")
			return
	; --------------------------------------
	acc := ComObjGet(accessAddress)
	filename := acc.currentproject.name
	mountMenu() ; mount tray menu
	; --------------------------------------
	if (acc.visible == false)
		acc.visible := true
	WinWait, %filename%
	WinActivate, %filename%
	; --------------------------------------
	; method overload -> exits the app ! Just activate the database
	if (onlySelectsDB) 
		return
	; --------------------------------------
	Run, %A_AHKPath% "%A_ScriptDir%\secondthread.ahk"
	acc.docmd.runcommand(603) ; accmdnewobjectdesignquery ' create new query
	acc.docmd.runcommand(184) ; accmdsqlview ' switch to sql view
	acc.docmd.setwarnings(true) 
	return
}
; --------------------------------------
fileDialogAccess() {
	FileSelectFile,accessAddress,2,,Selects the Access Database,Access Database (*.accdb) ; get the access file
	recentFilesAddress.Insert(accessAddress)
	return accessAddress
} 
; --------------------------------------
mountMenu() { ; mount tray menu
	menu, tray, DeleteAll
	menu, tray, Icon, favicon.ico 
	menu, tray, NoStandard ; remove standard menu
	Menu, tray, Add, Selects a New Database, selectsAccessDB
	Menu, tray, Add, Activate Selected DB, showsAccessDB
	recentFilesMountMenu() ; adds temp files submenu	
	Menu, tray, Add, Recent Files, :recentFilesTemp
	Menu, tray, Add, Unload Selected Database, unloadSelectedDatabase
	Menu, tray, Add, Reload Script, reloadMe
	Menu, tray, Add, Quit, exitNewQuery
}
; --------------------------------------
recentFilesMountMenu() { ; adds temp files to submenu
	if (recentFilesAddress[1] == "") ; if is empty sends a empty item
		Menu, recentFilesTemp, Add,, recentFilesTemp
	for k,v in recentFilesAddress ; if its not empy, mount the recent files
		Menu, recentFilesTemp, Add, %v%, recentFilesTemp
}
; ______________________________________
; // sub-labels
selectsAccessDB:
	accessAddress := fileDialogAccess() ; access database address
	main(true) ; sends the command to create a new query
return
; --------------------------------------
showsAccessDB:
	if (filename != "") {
		WinActivate, %filename%
	} else {
		MsgBox, No database selected
	}
return
; --------------------------------------
recentFiles:
return
; --------------------------------------
recentFilesTemp: ; selects the recent file
	accessAddress := A_ThisMenuItem
	main(true) ; sends the command to create a new query
return
; --------------------------------------
unloadSelectedDatabase:
	filename := ;
	accessAddress := ; access database address
	acc := ; access instance
return
; --------------------------------------
reloadMe:
	Reload
return
; --------------------------------------
exitNewQuery:
	ExitApp
return
