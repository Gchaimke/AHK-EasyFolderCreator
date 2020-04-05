#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#EscapeChar '

FormatTime, Year,A_now, yy

;Month creation
Month:=A_MM

InitSettings(){
;check if settings.ini exists
	if FileExist("settings.ini")
		Goto readGlobals
	else
		Goto writeGlobals

;Write GLOBALS to settings.ini if not exists
writeGlobals:
	msgbox,,Alert,New Settings created!
	FileAppend,,settings.ini
	IniWrite, Flex2,settings.ini,Globals,allProjects
	IniWrite,2,settings.ini,Flex2,zeroNum
	IniWrite,F'%tSN'%,settings.ini,Flex2,startFrom
	IniWrite,07,settings.ini,Flex2,midlle
	IniWrite,'%year'%,settings.ini,Flex2,endWith
	IniWrite,1,settings.ini,Flex2,template
return

;Read GLOBALS from settings.ini
readGlobals:
	IniRead,allProjects,settings.ini,Globals,allProjects
return
}


;==========Main GUI START============
InitSettings()
createMainGui:
Gui, guiMain:Default
Gui, Font, s10, Verdana

Gui, Add, GroupBox, w200 h420,Script settings
Gui, Add, Text,x20 y30,Select Project

IniRead,allProjects,settings.ini,Globals,allProjects
Gui, Add, DropDownList,w180 vProjName gSlectedItem sort uppercase,%allProjects%

Gui, Add, Text,, How many folders?
Gui, Add, Edit, vCount
Gui, Add, UpDown, Range1-20, 1
Gui, Add, Text,, First Serial Number?
Gui, Add, Edit, vFirst
Gui, Add, Text,, Current Month:
Gui, Add, Edit, vMonthMidlle
Gui, Add, UpDown,gUDUpdate Range01-12, %Month%
Gui, Add, Text,, Current Year:
Gui, Add, Edit, vYearMidlle
Gui, Add, UpDown,gUDUpdate Range18-99, %Year%
Gui, Add, Text,
Gui, Add, Button,w180 h30 gStart,Start
Gui, Add, Button,w180 h30 yp+180 gOpenSettings,Add New Project
Gui, Add, ListView,x250 y10 w350 r12 vStatusText gMyListView,Serial number
Gui, Add, edit,w350 vSelectedFile 
Gui, Add, Button,w150 gSelectExcel, Excel Template
Gui, Add, Text,
Gui, Add, edit,w350 vSelectedFolder 
Gui, Add, Button,w150 gSelectFolder, Output Folder
Gui , Add, StatusBar ,, Please Select project...
GuiControl, Disable, MonthMidlle
GuiControl, Disable, YearMidlle
GuiControl, Disable, Start
Gui, Show, ,Easy Folder Creator
return

guiMainGuiClose:
ExitApp

;==========Main GUI END============

;==========Settings GUI START============
OpenSettings:
Gui, guiMain: Destroy
Gui, guiSettings:Default
Gui, Font, s10, Verdana
Gui, Add, Text,, Project Name
Gui, Add, edit,w350 vProjectName 
Gui, Add, Text,, Number of zero in serial number template 0001-MM-YY = 3 zero
Gui, Add, edit,w350 vZeroNum
Gui, Add, Text,, Template options:
Gui, Add, Text,,'%tSN'% = counter number XXX
Gui, Add, Text,,'%A_MM'% = month number MM
Gui, Add, Text,, '%Year'% = year number YY
Gui, Add, Text,, Examples: F'%tSN'%-'%A_MM'%-'%Year'% , PCL-'%A_MM'%'%Year'%-P'%tSN'%
Gui, Add, Text,, Template Start:
Gui, Add, edit,w350 vStartsFrom
Gui, Add, Text,, Template middle:
Gui, Add, edit,w350 vMidlle 
Gui, Add, Text,, Template End:
Gui, Add, edit,w350 vEndWith
Gui, Add, Text,, Template kind with seperators = 1, without = 2.
Gui, Add, edit,w350 vTemplateKind
Gui, Add, Button,w180 h30 gAddProject,Create New
Gui, Show, ,Create new project
return

guiSettingsButtonOK:
guiSettingsGuiClose:
guiSettingsGuiEscape:
Gui Destroy  ; Destroy the settings box.
Goto createMainGui
return
;==========Settings GUI END============

;Add new project to ini file
AddProject:
	Gui, Submit,NoHide
	IniWrite,%ZeroNum%,settings.ini,%ProjectName%,zeroNum
	IniWrite,%StartsFrom%,settings.ini,%ProjectName%,startFrom
	IniWrite,%Midlle%,settings.ini,%ProjectName%,midlle
	IniWrite,%EndWith%,settings.ini,%ProjectName%,endWith
	IniWrite,%TemplateKind%,settings.ini,%ProjectName%,template
	msgbox,,Success!,Project added to settings file.
	IniRead,allProjects,settings.ini,Globals,allProjects
	IniWrite,%allProjects%|%ProjectName%,settings.ini,Globals,allProjects
	Gui Destroy  ; Destroy the about box.
	Goto createMainGui
return

;select Excel template
SelectExcel:
FileSelectFile, SelectedFile, 3, , Open a file, Text Documents (*.xlsx)
if SelectedFile =
    SB_SetText("The user didn't select any file.")
else
	GuiControl,, SelectedFile, %SelectedFile%
return

;select output folder
SelectFolder:
FileSelectFolder, myFolder,,,Select folder for creation
if myFolder =
    SB_SetText("The user didn't select any folder.")
else
    GuiControl,, SelectedFolder, %myFolder%
return

;Update GUI on UpDown changes
UDUpdate:
	startString := GetString(ProjName,First)
	SB_SetText("First Serial Number: "startString ". Press Start to create folders.")
return
;After project selected, get print parameters
SlectedItem:
	InputBox, StartCount, Start number, Please enter first serial number
	uStart := StartCount+0
	GuiControl,,First,%uStart%
	InputBox, foldersCount, How many folders?, Please enter How many folders
	GuiControl,,Count,%foldersCount%
	Gui, Submit,NoHide
	GuiControl,,SelectedFolder, %A_Desktop%\%ProjName%
	GuiControl,,SelectedFile, %A_ScriptDir%\Templates\%ProjName%.xlsx
	GuiControl,,Midlle,%Month%
	startString := GetString(ProjName,First)
	SB_SetText("First Serial Number: "startString ". Press Start to create folders.")
	GuiControl, Enable, MonthMidlle
	GuiControl, Enable, YearMidlle
	GuiControl, Enable, Start
return

;After Start button pressed, create directories with files
Start:
Gui, Submit,NoHide
msgbox,4,,I will create %Count% folders in %SelectedFolder%
	IfMsgBox Yes
	{
		Loop %Count% {
			tNum:=First
			tCurrentNum := GetString(ProjName,tNum)
			LV_Add("", tCurrentNum)
			tNum := First++
			FileCreateDir, %SelectedFolder%\%tCurrentNum%
			FileCopy, %SelectedFile%, %SelectedFolder%\%tCurrentNum%\%tCurrentNum%.xlsx
			}
		SB_SetText("All folders are created!")
	}else{
		SB_SetText("Creation canceled.")
	}
return

MyListView:
if (A_GuiEvent = "DoubleClick")
{
    LV_Delete()
}
return


GetString(ProjName,number){

	GuiControlget , Month, ,MonthMidlle
	GuiControlget , Year, ,YearMidlle
	
	;D-fend month and year creation
	monthEnum:=Object(1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,"A",11,"B",12,"C")
	Df_MM:=monthEnum[Month]
	Df_year:=Substr(Year+1,0)
	
	if(Month<10)
		Month:=Max1Zero(Month)
		
	
	IniRead,zeroNum,settings.ini,%ProjName%,zeroNum
	if(zeroNum==1)
		tSN := Max1Zero(number)
	if(zeroNum==2)
		tSN := Max2Zero(number)
	if(zeroNum==3)
		tSN := Max3Zero(number)
	
	;Transform - Performs miscellaneous math functions, bitwise operations, and tasks. 
	;Deref - Expands variable references and escape sequences contained inside other variables.
	
	IniRead,startFrom,settings.ini,%ProjName%,startFrom
	Transform,startFrom,deref,%startFrom%
	
	IniRead,midlle,settings.ini,%ProjName%,midlle
	Transform,midlle,deref,%midlle%
	
	IniRead,endWith,settings.ini,%ProjName%,endWith
	Transform,endWith,deref,%endWith%

	;msgbox,,DEBUG,%startFrom%-%midlle%-%endWith%
	IniRead,template,settings.ini,%ProjName%,template
	
	if(template==1){
		return startFrom "-" midlle "-" endWith
	}else if(template==2){
		return startFrom "" midlle "" endWith
	}else{
		return startFrom "-" midlle "-" endWith
	}
}

Max1Zero(num){
	if(num+0<10)
		return "0"num
	else
	 return num
}

Max2Zero(num){
	if(num+0<10)
		return "00"num
	else if(num+0>=10 and num+0<100)
		return "0"num
	else
		return num
}

Max3Zero(num){
	if(num+0<10)
		return "000"num
	else if(num+0>=10 and num+0<100)
		return "00"num
	else if(num+0>=100)
		return "0"num
	else
		return num
}
