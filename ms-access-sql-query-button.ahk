/* 
================================================================================================
	
	* version: 
					1.10 [22/10/2017] (dd/mm/yyyy)
	* info:
					contact:			- king-of-hearts				
					email: 				- fcvidoto@hotmail.com
					download:			- https://github.com/fcvidoto/ms-access-sql-query-button/archive/master.zip
					forum-topic: 	- https://autohotkey.com/boards/viewtopic.php?f=6&t=38707
					github: 			- https://github.com/fcvidoto/ms-access-sql-query-button

================================================================================================ 
*/
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force
#Persistent
SetTitleMatchMode, 2 ; partial name of the file
DetectHiddenWindows, on
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; #Warn  ; Enable warnings to assist with detecting common errors.

; // tray icon menu
mountMenu() ; 

global accessAddress := ; access database address
global acc := ; access instance
global filename := ; 
global winPressed := ;
global recentFilesAddress := [] ;

; // hotkeys
#n:: main(false,true) ; selects a new db and create a query
#!n:: main(false,false) ; selects a new db and create a query
+esc:: ExitApp

; // functions
main(onlySelectsDB, createSQLquery) { ; sends the command to create a new query
	if (accessAddress == "")
		accessAddress := fileDialogAccess() ; get the database path
		if (accessAddress == "")
			return
	
	acc := ComObjGet(accessAddress)
	filename := acc.currentproject.name
	mountMenu() ; mount tray menu
	
	if (acc.visible == false)
		acc.visible := true
	WinWait, %filename%
	WinActivate, %filename%
	
	if (onlySelectsDB) ; !!method overload -> exits the app ! Just activate the database
		return
	
	Run, %A_AHKPath% "%A_ScriptDir%\lib-msacc\secondthread.ahk"
	acc.docmd.runcommand(603) ; accmdnewobjectdesignquery ' create new query
	if (createSQLquery) ; <--- method overload to create SQL Query
		acc.docmd.runcommand(184) ; accmdsqlview ' switch to sql view

	acc.docmd.setwarnings(true) 
	return
}

cleanGlobals() {
	filename := ;
	accessAddress := ; access database address
	acc := ; access instance
}

exportAllTable() { ; exporta all tables
	allTables := ; emptys the clipboard
	for table in acc.currentdb.tabledefs {
		tableExpt := table.name
		StringLeft, strLeft, tableExpt, 4
		
		if (strLeft != "MSys") {
			allTables .= "`r`n ------------------------------- "
			allTables .= "`r`n'" . table.name . "'"
			for field in table.fields
				allTables .= "`r`n   " . field.name . ","
		}
	}
	Clipboard := allTables
	ClipWait
	opensNotePadAndPaste()
	return
}

exportAllQuery() { ; export all queries from the database
	allQueries := ; emptys the clipboard
	for query in acc.currentdb.querydefs {
		allQueries .= "`r`n ------------------------------- "
		allQueries .= "`r`n'" . query.name . "'"
		allQueries .= "`r`n" . query.sql 
	}
	Clipboard := allQueries
	ClipWait
	opensNotePadAndPaste()
	return
}

opensNotePadAndPaste() { ; create a new instance of notepad and paste
	run, Notepad
	WinWaitActive, ahk_class Notepad
	send, ^v ; paste from clipboard
	Clipboard := ; cleans the clipboard
	ClipWait
	return
}

fileDialogAccess() { ; dialog to open the access database
	FileSelectFile,accessAddress,2,,Selects the Access Database,Access Database (*.accdb) ; get the access file
	if (accessAddress != "")
		recentFilesAddress.Insert(accessAddress)
	return accessAddress
} 

mountMenu() { ; mount tray menu
	menu, tray, DeleteAll
	menu, tray, Icon, favicon.ico 
	menu, tray, NoStandard ; remove standard menu
	Menu, tray, Add, Selects a New Database, selectsAccessDB
	Menu, tray, Add, Activate Database, showsAccessDB
	Menu, tray, Add ; line separator ; --------------------------
	Menu, tray, Add, New SQL Query, createsSQLquery
	Menu, tray, Add, New SQL Design, createsDesignQuery
	Menu, tray, Add ; line separator ; --------------------------
	Menu, tray, Add, Export All Tables, exportTableStructure
	Menu, tray, Add, Export All Queries, exportSQLStructure
	Menu, tray, Add ; line separator ; --------------------------
	recentFilesMountMenu() ; adds temp files submenu	
	Menu, tray, Add, Recent Files, :recentFilesTemp
	Menu, tray, Add ; line separator ; --------------------------
	Menu, tray, Add, Unload Selected Database, unloadSelectedDatabase
	Menu, tray, Add, Reload Script, reloadMe
	Menu, tray, Add, Quit, exitNewQuery
}

recentFilesMountMenu() { ; adds temp files to submenu
	if (recentFilesAddress[1] == "") ; if is empty sends a empty item
		Menu, recentFilesTemp, Add,, recentFilesTemp
	for k,v in recentFilesAddress ; if its not empy, mount the recent files
		Menu, recentFilesTemp, Add, %v%, recentFilesTemp
}

; // sub-labels
createsSQLquery:
	main(false,true) ; sends the command to create a new query
return

createsDesignQuery:
	main(false,false) ; sends the command to create a new query
return

exportTableStructure:
	if (filename == "") {
		MsgBox, No database selected
		return
	}
	exportAllTable() ; exporta all tables
return

exportSQLStructure:
	if (filename == "") {
		MsgBox, No database selected
		return
	}
	exportAllQuery() ; export all queries from the database
return

selectsAccessDB:
	cleanGlobals() ; clean globals
	main(true,false) ; selects the database
return

showsAccessDB:
	if (filename != "") {
		WinActivate, %filename%
	} else {
		MsgBox, No database selected
	}
return

recentFiles:
return

recentFilesTemp: ; selects the recent file
	cleanGlobals() ; clean globals
	accessAddress := A_ThisMenuItem
	main(true,false) ; sends the command to create a new query
return

unloadSelectedDatabase: 
	cleanGlobals() ; clean globals
return

reloadMe:
	Reload
return

exitNewQuery:
	ExitApp
return
