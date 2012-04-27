Critical
#Include %A_ScriptDir%\fobgoblin.ini
SetTitleMatchMode, 2
SetWorkingDir, %A_ScriptDir%

Loop, parse, applist, |
	app%A_Index% = %A_LoopField%

iniread, lastxpos, wellnet.ini, FOB, LastXPos, 350
iniread, lastypos, wellnet.ini, FOB, LastYPos, 350
iniread, agentname, wellnet.ini, A&A, AgentName, %A_Space%
iniread, fobletterlocation, wellnet.ini, FOB, fobletterlocation, %A_Space%

Gui, Add, Text, x26 y11 w60 h20 , Application:
Gui, Add, GroupBox, x16 y37 w570 h130 , User information
Gui, Add, Text, x26 y61 w60 h20 vFirstNameText, First name:
Gui, Add, Text, x236 y61 w70 h20 vMiddleNameText, Middle:
Gui, Add, Text, x416 y61 w60 h20 vLastNameText, Last:
Gui, Add, Text, x26 y81 w60 h20 vPositionText, Position:
Gui, Add, Text, x236 y81 w60 h20 vPhoneText, Phone:
Gui, Add, Text, x416 y81 w60 h20 vFaxText, Fax:
Gui, Add, Text, x26 y101 w70 h20 vOrgNameText, Organization:
Gui, Add, Text, x301 y101 w70 h20 vAddressText, Address:
Gui, Add, Text, x26 y121 w70 h20 vCityText, City:
Gui, Add, Text, x186 y121 w50 h20 vProvText, Province:
Gui, Add, Text, x282 y121 w60 h20 vPostalText, Postal Code:
Gui, Add, Text, x406 y121 w70 h20 vEmailText, E-mail:
Gui, Add, Text, x26 y141 w70 h20 vSecretWordText, Secret word:
Gui, Add, Text, x306 y141 w70 h20 vPromptText, Prompt:
Gui, Add, GroupBox, x16 y172 w370 h70 , RSA Information
Gui, Add, Text, x26 y196 w70 h20 Lowercase Limit20 vUserIDText, User ID:
Gui, Add, Text, x206 y196 w70 h20 vSerialNoText, Serial Number:
Gui, Add, Text, x26 y216 w70 h20 vTokenTypeText, Token Type:
Gui, Add, Text, x206 y216 w70 h20 vFobSourceText, Source:
Gui, Add, GroupBox, x391 y172 w195 h70 , Automation Settings
Gui, Add, GroupBox, x16 y247 w570 h70 , Request Information
Gui, Add, Text, x26 y271 w80 h20 vAgentNameText, Completed by:
Gui, Add, Text, x236 y271 w70 h20 vRequestNoText, Request #:
Gui, Add, DropDownList, x96 y7 w180 gSetAppNums vAppNum AltSubmit, %applist%
Gui, Add, Radio, x286 y11 w60 h20 vActionType gEnableRSAOption Checked, Create
;Gui, Add, Radio, x356 y11 w60 h20 gDisableRSAOption, Amend
;Gui, Add, Radio, x426 y11 w60 h20 gDisableRSAOption, Delete
Gui, Add, Button, x486 y11 w100 h20, Clear
Gui, Add, Edit, x96 y57 w80 h20 vFirstname gUserIdUpdate, 
Gui, Add, Edit, x176 y57 w50 h20 vPrefname gUserIdUpdate, 
Gui, Add, Edit, x276 y57 w130 h20 vMiddlename, 
Gui, Add, Edit, x446 y57 w130 h20 vLastname gUserIdUpdate, 
Gui, Add, Edit, x96 y77 w130 h20 vPosition, 
Gui, Add, Edit, x276 y77 w130 h20 vPhone, 
Gui, Add, Edit, x446 y77 w130 h20 vFax, 
Gui, Add, Edit, x96 y97 w195 h20 vOrgName, 
Gui, Add, Edit, x346 y97 w230 h20 vAddress, 
Gui, Add, Edit, x96 y117 w80 h20 vCity, 
Gui, Add, Edit, x236 y117 w30 h20 vProv, 
Gui, Add, Edit, x346 y117 w50 h20 vPostal, 
Gui, Add, Edit, x446 y117 w130 h20 vEmail, 
Gui, Add, Edit, x96 y137 w195 h20 vSecretword, 
Gui, Add, Edit, x346 y137 w230 h20 vPrompt, 
Gui, Add, Edit, x96 y192 w100 h20 vUserid Lowercase, 
Gui, Add, Edit, x276 y192 w100 h20 +Number vSerialno, 
Gui, Add, DropDownList, x96 y212 w100 vTokentype AltSubmit, Hardware||Software
Gui, Add, DropDownList, x276 y212 w100 vFobSource AltSubmit, IBM Inventory||Client Inventory
Gui, Add, CheckBox, x401 y192 w60 h20 vDoIMAC checked, IMAC
Gui, Add, CheckBox, x481 y192 w90 h20 vDoEventTracker checked, Event Tracker
Gui, Add, CheckBox, x401 y212 w60 h20 vDoRSA checked, RSA
Gui, Add, CheckBox, x481 y212 w60 h20 vDoRUNA checked, RUNA
Gui, Add, CheckBox, x481 y212 w100 h20 vDoLetter, Fob letter
GuiControl, hide, doletter
Gui, Add, Edit, x96 y267 w130 h20 vAgentname, %agentname%
Gui, Add, Edit, x301 y267 w105 h20 vRequestNo, 
Gui, Add, Text, x416 y271 w60 h20 , Requestor:
Gui, Add, Edit, x476 y267 w100 h20 vRequestor,
Gui, Add, Text, x26 y292 w550 h20 , * Note: You must enter your full name or IBM e-mail address in order for Event Tracker to work.
Gui, Add, Button, x196 y322 w100 h30 Default, Submit
Gui, Add, Button, x306 y322 w100 h30 , Cancel
Gui, Show, x%lastxpos% y%lastypos% h361 w606, FobGoblin
Return

;======== Gui subroutines ========
UserIdUpdate:
	gui, submit, nohide
	guicontrol,,userid, % AppName == "CGI" ? SubStr(prefname ? prefname . "." . lastname : firstname . "." . lastname,1,20) : SubStr(prefname ? prefname . lastname : firstname . lastname,1,20)
return

SetAppNums:
	gui, submit, nohide
	appname := app%appnum%
	if appnum between %GOAL% and %GOAH%
		invoice = 1
	else if appnum between %HealthL% and %HealthH%
	{
		securitydomain = ahw_alberta
		invoice = 2
	}
	else if appnum between %RHAL% and %RHAH%
		invoice = 3
	else if appnum between %WellnetL% and %WellnetH%
	{
		SecurityDomain = wnt
		invoice = 4
	}
	else
	{
		msgbox,4096,Error,Error 85: invalid application.`n`nDebug info: %appname%,%appnum%,%goal%,%goah%,%healthl%,%healthh%,%rhal%,%rhah%,%wellnetl%,%wellneth%`n`nClick OK to continue.
		return
	}
	imacapp := SubStr(appname,1,4)
	ShowAllFields()
	if appname = Netcare
	{
		GuiControl,,DoRUNA,0
		HideFields("DoRUNA","Position","Phone","Fax","email","secretword","prompt")
		ShowFields("DoLetter")
		GuiControl,,DoLetter,1
	}
	else if appname = CGI
	{
		HideFields("city","orgname","prov","postal","address","position","phone","fax")
		GuiControl,,city,Edmonton
		GuiControl,,orgname,CGI
		GuiControl,,prov,AB
		GuiControl,,postal,T5J 3N6
		GuiControl,,address,10303 - Jasper Ave, Suite 800
	}




	if appname contains AADL
		securitydomain = aadl
	else if appname = cgi
		securitydomain = ahw_cgi
	else if appname contains RHA1
		securitydomain = Chinook
	else if appname contains rha2
		securitydomain = rha_02
	else if appname contains rha3
		securitydomain = rha_04
	else if appname contains rha4
		securitydomain = DTHR
	else if appname contains RHA5
		securitydomain = EAST_CENTRAL
	else if appname contains RHA6
		securitydomain = Capital
	else if appname contains RHA7
		securitydomain = Aspen
	else if appname = ACB
		securitydomain = ACB
	else if appname = RSHIP
		securitydomain = RSHIP
	else if appname contains IBM
		securitydomain = IBM
	else
		securitydomain = WNT
return

EnableRSAOption:
	GuiControl,Enable,DoRSA
	GuiControl,Enable,DoRUNA
	GuiControl,,DoRSA,1
	GuiControl,,DoRUNA,1
	if AppName = Netcare
		gosub SetAppNums
return

DisableRSAOption:
	GuiControl,,DoRSA,0
	GuiControl,,DoRUNA,0
	GuiControl,Disable,DoRSA
	GuiControl,Disable,DoRUNA
return


ButtonCancel:
GuiClose:
	ExitApp
return

ButtonSubmit:
	Gui, Submit, nohide
	if DoImac + DoEventTracker + DoRSA + DoRUNA == 0
		msgbox,4096,,You have not chosen to automate anything.
	else if not firstname or not lastname or not appnum or not userid or not serialno
		msgbox,4096,,The following fields are mandatory:`n`n`tApplication`n`tFirst Name`n`tLast Name`n`tUser ID`n`tSerial Number
	else if (not address or not city or not prov or not orgname or not postal) and doletter and appname == "Netcare"
		msgbox,4096,,You must enter an address.
	else
		goto DoThings
return
;=================================

DoThings:
if appname == "Netcare" and doletter and not fobletterlocation ;This should not happen unless the user deletes the fobletterlocation key from the .ini
{
	FileSelectFile,fobletterlocation,3,%A_ScriptDir%\ANP form letter - keyfob.dot,Please select your ANP fob letter,*.dot
	if ErrorLevel
	{
		msgbox,4132,,The ANP form letter for keyfobs could not be found. Would you like to try again? Clicking no will prevent a fob letter from being filled out.
		IfMsgBox, Yes
			goto DoThings
		else
			DoLetter = 0
	}
	else
		iniwrite, %fobletterlocation%, wellnet.ini, FOB, fobletterlocation
}
WingetPos, lastxpos, lastypos,,,FobGoblin
IniWrite, %lastxpos%, wellnet.ini, FOB, LastXPos
IniWrite, %lastypos%, wellnet.ini, FOB, LastYPos
IniWrite, %AppName%, wellnet.ini, A&A, DefaultApp
IniWrite, %AgentName%, wellnet.ini, A&A, AgentName
Gui, hide
if DoRSA
{
	IfWinNotExist, RSA Security Console: -
		MsgBox, 4096,, Please open the RSA 7.1 Security Console.
	WinWait, RSA Security Console: -	
	WinActivate
	Loop
	{
		WinWaitActive
		click % ImagePos("rsalogo")
		ImagePos("done")
		sleep 500
		send {tab 11}{enter}
		sleep 100
		ImagePos("done")
		IfWinExist, Add New User with Options
			break
	}
	click % ImagePos("next")
	WinWaitActive, Add New User -
	ImagePos("done")
	sleep 500
	send +{tab 2}
	sleep 200
	send -%securitydomain%{enter}
	ImagePos("done")
	send {tab 2}
	send %firstname%
	If prefname
		send %A_Space%(%prefname%)
	send {tab}%middlename%{tab}%lastname%{tab}%userid%{tab}%email%
	click % ImagePos("savenext")
	WinWaitActive RSA Security Console: - Add User Group Membership
	ImagePos("done")
	click % ImagePos("where", wherex, wherey)
	sleep 500
	send notes{tab}c{tab} ; search criteria: notes contains
	sleep 200
	send %appname%{enter}
	ImagePos("done")
	click % ImagePos("selectall", SelectAllX, SelectAllY)
	click % ImagePos("savenext")
	WinWaitActive, RSA Security Console: - Assign SecurID Tokens
	ImagePos("done")
	click % ImagePos("where", wherex, wherey)
	sleep 500
	send serial{tab}c{tab}%serialno%{enter}
	ImagePos("done")
	click % ImagePos("selectall", SelectAllX, SelectAllY)
	click % ImagePos("savefinish")
	sleep 500
}

if DoIMAC
{
	ShipDate=%A_MMM%-%A_DD%-%A_YYYY%
	run, c:\Notes\Notes.exe ahwshared!!\health\sftp_challenge_response.nsf
	WinWaitActive, IMAC User Catalog
	send !co
	WinWaitActive, (Untitled) - 
	sleep, 1000
	send {tab}%ShipDate%{tab}%userID%{tab}
	sleep 500
	send {down %invoice%}{tab}
	sleep 500
	send %imacapp%{tab}
	sleep 500
	send {up 2}{up %ActionType%}{tab}
	sleep 500
	send {up 3}{tab}
	sleep 500
	imacfobsource := fobsource-1
	send {right %imacfobsource%}
	sleep 500
	send {tab}
	imactokentype := tokentype-1
	send {right %imactokentype%}
	sleep 500
	send {tab}
	sleep 500
	send %SerialNo%
	sleep 500
	send {tab}
	sleep 500
	send {space}
	WinWaitNotActive, (Untitled) - 
	sleep 500
}

if DoEventTracker
{
	if ActionType = 1
		etaction = add
	else if ActionType = 2
		etaction = modify
	else if ActionType = 3
		etaction = delete
	run, c:\Notes\Notes.exe i_dir\idrequest.nsf	
	WinWaitActive, IDAdmin Event Tracker -
	sleep 500
	send !2
	WinWaitActive, IDAdmin Event Tracker - ID Request - 
	sleep 500
	send AHW{tab}Canada{tab 2}%requestno%{tab}%requestno%{tab 2}%etaction% %appname% RSA token for %FirstName% %LastName%{tab}r{tab}%requestor%{tab}%FirstName% %LastName%{tab}%appname%{tab}n{tab 2}1{del}{tab}web{tab}%etaction%{tab 2}1{tab 3}%etaction% %appname% RSA token for %FirstName% %LastName%{tab}%agentname%;{tab}closed
	MsgBox,4128,,Please verify that the information in Event Tracker is correct, then click submit.
	WinActivate, IDAdmin Event Tracker - ID Request - 
	WinWaitNotActive, IDAdmin Event Tracker - ID Request -
}

if DoRUNA
{
	run, c:\Notes\Notes.exe ahwshared!!\health\wellnetruna.nsf
	WinWaitActive, Wellnet Remote User Net Access
	sleep, 1000
	send, !cp
	WinWaitActive, Application for User Registration - 
	sleep 500
	send {tab 2}{space down}
	sleep 200
	send {space up}{tab 2}%firstname%
	if prefname
		send %A_Space%(%prefname%)
	send % Substr(middlename,1,1)
	send {tab}%lastname%{tab 2}%position%{tab}%phone%{tab}%fax%{tab}%orgname%{tab}
	if invoice = 3
		send %imacapp%
	send {tab}%orgname%{tab}%address%{tab}%city%{tab}%prov%{tab}%postal%{tab}%email%{tab}
	if invoice = 3
		send VPN
	else
		send %appname%
	send {tab 2}%secretword%{tab}%prompt%^{end}+{tab}
	if invoice = 3
		send `,vpn
	send +{tab}
	if invoice in 1,4
		send wellnet - firewall
	else if invoice = 3
		send %imacapp% - VPN Server
	else if invoice = 2
		send Wellnet - VPN
	send +{tab 4}{left}+{tab}{left}+{tab 2}%shipdate%+{tab 3}%agentname%
	if requestno
		send %A_Space%created as per request %requestno%
	send +{tab 2}%serialno%+{tab}%userid%{esc}
	WinWaitActive, ahk_class #32770, Do you want to save this new document?
	sleep 200
	send {enter}	
}

if DoLetter
{

}
EndIfDoLetter:

GuiControl,,FirstName
GuiControl,,PrefName
GuiControl,,MiddleName
GuiControl,,LastName
GuiControl,,Position
GuiControl,,Email
GuiControl,,SecretWord
GuiControl,,Prompt
GuiControl,,userid
GuiControl,,SerialNo, % SerialNo+1
Gui,Show
return


ButtonClear:
GuiControl,,FirstName
GuiControl,,PrefName
GuiControl,,MiddleName
GuiControl,,LastName
GuiControl,,Position
GuiControl,,Phone
GuiControl,,Fax
GuiControl,,OrgName
GuiControl,,Address
GuiControl,,City
GuiControl,,Prov
GuiControl,,Postal
GuiControl,,Email
GuiControl,,SecretWord
GuiControl,,Prompt
GuiControl,,userid
GuiControl,,SerialNo
GuiControl,,RequestNo
GuiControl,,Requestor
return


ImagePos(ImageName, OffsetX = 0, OffsetY = 0)
{
	ErrorLevel = 99
	WinGetPos,WinX,WinY,WinWi,WinHi
	Loop
	{
		sleep 100
		ImageSearch, clickx, clicky, WinX, WinY, WinWi, Winhi, %ImageName%.bmp
		If Errorlevel = 0
			break
		ImageSearch, clickx, clicky, WinX, WinY, WinWi, Winhi, smooth%ImageName%.bmp
		If Errorlevel = 0
			break
    }
	clickx += OffsetX
	clicky += OffsetY
	clicks = %clickx% %clicky%
	return clicks
}

HideFields(fields*)
{
	global
	for index,fieldname in fields
	{
		GuiControl, hide, %fieldname%
		GuiControl, hide, %fieldname%text
	}
	fields.Remove(1,fields.MaxIndex())
}

ShowFields(fields*)
{
	global
	for index,fieldname in fields
	{
		GuiControl, show, %fieldname%
		GuiControl, show, %fieldname%text
	}
	fields.Remove(1,fields.MaxIndex())
}

ShowAllFields()
{
	global
	ShowFields("Firstname","Prefname","Middlename","LastName","Position","Phone","Fax","OrgName","Address","city","Prov","Postal","email","SecretWord","Prompt","Userid","serialno","TokenType","TokenSource","DoIMAC","DoEventTracker","DoRSA","DoRUNA","AgentName","RequestNo","Requestor")
	GuiControl,,DoLetter,0
	HideFields("DoLetter")
}