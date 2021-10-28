' ============================================================================== 
' Script.........: [DynamicDataDefinitions.vbs] 
' Author.Email...: [me@benburke.dev] 
' Version........: 7.0 
' Date Written...........: 31-Jul-2009 
' Updated.................: 29-Oct-2021
'
' One Line Description: This is a simple include file for all dynamic data 
' 			definitions that are in use across different modules.
' 
'
' - the standard template is not in use for this include file.
'
Dim WshShell		' WSH Shell Object, to access process context 
Dim objEnv		' Object containing Process level
dim ITS_DATA 		' Will contain environment variable value for ITS_DATA		
dim ITS_LOG		' 	ditto
dim ITS_PROCEDURES	' 	ditto
dim USERNAME		' will contain windows username
dim USERDOMAIN		' sometimes it's good to know where we logged on
dim ASP			' detect if we are executing in ASP script engine
dim objRootDSE		' Object used to get the default domain naming context
dim strDomainName	' get the current naming context for use in AD queries and updates
dim strHostname
dim WshNetwork   	' get the hostname from WshNetwork
dim strMailFrom		' combine user@host in a string

On error resume next 	' turn off default error handling - each line must do it's own 
'On error goto 0


Set WshShell = CreateObject("WScript.Shell")

if err.number <> 0 then
	
	dieifbroken "Failure to create Shell"
	
end if


set objRootDSE = getobject("LDAP://RootDSE")

strDomainName = objRootDSE.get("DefaultNamingContext")

set objRootDSE = nothing


ASP = request.servervariables("script_name")

if err.number <> 0 then

	ASP = False
else
	ASP = True
end if

err.clear




Set objEnv = WshShell.Environment("Process")

if err.number <> 0 then
	dieifbroken "Failure to create objEnv"
	'wscript.quit
end if

if Trim(objENV("ITS_LOG")) = "" then
	
	dieifbroken "ITS_LOG not defined as Environment Variable"
	
else
	ITS_LOG = Trim(objENV("ITS_LOG"))
	speakifalive ITS_LOG & " is log "
	
end if

if Trim(objENV("ITS_DATA")) = "" then
	
	dieifbroken "ITS_DATA not defined as Environment Variable"
	
else
	ITS_DATA = Trim(objENV("ITS_DATA"))
	speakifalive  ITS_DATA & " is data "
end if


if Trim(objENV("ITS_PROCEDURES")) = "" then
	
	dieifbroken "ITS_PROCEDURES not defined as Environment Variable"

else
	ITS_PROCEDURES = Trim(objENV("ITS_PROCEDURES"))
	
	speakifalive ITS_PROCEDURES & " is procedures "
end if




if not ASP then

	if Trim(objENV("USERNAME")) = "" then
		
		dieifbroken "USERNAME not defined as Environment Variable"
	
	else
	
		USERNAME = Trim(objENV("USERNAME"))
		speakifalive USERNAME & " is username "
		
	end if

else
	

	USERNAME = request.servervariables("LOGON_USER")
	
	USERNAME = mid(USERNAME,instr(USERNAME,"\")+1,len(USERNAME)-instr(USERNAME,"\"))
	
	
	speakifalive USERNAME & " is username "
	
end if	


Set WshNetwork = WScript.CreateObject("WScript.Network")


strHostname = WshNetwork.ComputerName

strMailFrom = trim(USERNAME) & "@" & trim(strHostname)


sub dieifbroken(msg)

	if ASP then
	
		response.write "<br>" & msg & "</br>"
		response.end
	else
	
		wscript.echo msg
		wscript.quit
		
	end if


end sub



sub speakifalive(msg)

	if ASP then
	
		if debuglevel > 0 then response.write "<br>" & msg & "</br>"

	else
	
		wscript.echo msg
		
	end if


end sub

on error goto 0 	' turn on default error handling
