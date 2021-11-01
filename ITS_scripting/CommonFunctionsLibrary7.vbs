' ============================================================================== 
' Script.........: [CommonFunctionsLibrary.vbs] 
' Author.Email...: [me@benburke.dev] 
' Version........: 7.0 
' This model version...........: 20-Jan-2011 
' Updated......................: 29-Oct-2021
'
' One Line Description: All shared functions are declared in this file.
'
' They have their own comments sections.
'
' This version (6) introduces shared libraries between VBS and ASP. 
' All VBS scripts that inherit this library run as WSF files, using the built in include (rather than a manual include)
' 
' 
' For  revision history, go to end of file 
' (this saves the interpreter from 'reading' the comments.


'Option Explicit - allow the callers declaration options to rule
'On Error Goto 0 - and their error handlers/
' ============================================================================== 
'  Subroutines/Functions/Classes 
' ============================================================================== 
' -------------------------------------------------------------------------- 
'  Subroutine.....: Logger
'  Purpose........: Write a logfile in a standard location with a standard format
'  Arguments......: The message to write to the logfile
'  Example........: logger("Failure to open database connection " & dbName)
'  Requirements...: Environment variable ITS_LOGS
'  Notes..........: This routine will create a logfile via fso, if not already
'			open and write the records with timestamps
'

sub Logger(LogText)




	if not IsObject(fsoLogger) then 
		set fsoLogger = createobject("scripting.FileSystemObject")
		writeresponse "FSO initialised for logger"
	end if
	

	
	if not IsObject(tsoLogger) then 
		set tsoLogger = fsoLogger.OpenTextFile(logfilename,2,True,0)
		writeresponse "TSO logger opened logfile & LogFileSpec"
	end if

	
	tsoLogger.writeline (LogText)

end sub


' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: Initialise
'  Purpose........: General purpose initialise
'  Arguments......: none
'  Example........: Initialise()
'  Requirements...: WSH
'  Notes..........: set scriptname variable and trigger logging
' -------------------------------------------------------------------------- 

sub Initialise ()

	dim scriptarray

	
	ScriptTimeStamp = datepart("yyyy",now) & right("0" & datepart("m",now()),2) 
	ScriptTimeStamp = ScriptTimeStamp  & right("0" & datepart("d",now()),2)   
	ScriptDateStamp = ScriptTimeStamp
	ScriptTimeStamp = ScriptTimeStamp  & right("0" & datepart("h",now()),2)   
	ScriptTimeStamp = ScriptTimeStamp  & right("0" & datepart("n",now()),2)   

	
	if ASP then 
	
		scriptname = request.servervariables("path_info")
	
	
		scriptarray = split(scriptname,"/")
	
		scriptname = scriptarray(ubound(scriptarray))
	
		scriptarray = split(scriptname,".")
	
		scriptname = scriptarray(ubound(scriptarray)-1)
		
	else
	
		scriptname = left(wscript.scriptname,len(wscript.scriptname)-4)

	
	end if
	

	logfilename = ITS_LOG & "\" & scriptname & "_" & username & "_" & ScriptTimeStamp & ".log"
	
	message "Scriptname Initialised with logfile " & logfilename

	
end sub

Function DownloadFile(ImageURL, SavePath, Milliseconds)

	Dim WinHttpReq
	Dim StartTime
	Dim EndTime
	Dim RunTime
	Dim ReqSuccess
	Dim t ' timer

	t = Timer

	message "Downloading " & ImageURL & " to " & SavePath


	Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
	WinHttpReq.Open "GET", ImageURL, False
	WinHttpReq.send

	If WinHttpReq.Status = 200 Then
		Set MyStream = CreateObject("ADODB.Stream")
		MyStream.Open
		MyStream.Type = 1
		MyStream.Write WinHttpReq.responseBody
		MyStream.SaveToFile SavePath, 2 ' 1: Don't overwrite, 2: Overwrite
		MyStream.Close
		ReqSuccess = true
	End If

	if ReqSuccess Then

		RunTime = datediff("s",StartTime, EndTime)
			' Int() behaves exactly like Floor() function, i.e. it returns the biggest integer lower than function's argument
		temp = Int(t)

		Milliseconds = Int((t-temp) * 1000)

		message "Run time is " & Milliseconds & " Milliseconds"

		DownloadFile = true 

	end if

End Function

' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: Cleanup
'  Purpose........: General purpose Cleanup
'  Arguments......: none
'  Example........: Cleanup
'  Requirements...: WSH
'  Notes..........: set scriptname variable and trigger logging
' -------------------------------------------------------------------------- 

sub Cleanup ()

	tsoLogger.close
	set tsoLogger = nothing
	set fsoLogger = nothing
	
end sub




' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: Message
'  Purpose........: General purpose message handler - calls Logger as required
'  Arguments......: ScriptName as string, message text as string
'  Example........: writeresponse scriptname, "My hovercraft is full of eels"")
'  Requirements...: WSH
'  Notes..........: This is the basis of all message output from scripts, except alarms
' -------------------------------------------------------------------------- 

sub message (messagetext)

	dim message 
	
	Message = now() & "--->"& scriptname & "--->" & MessageText
'	writeresponse message
	writeresponse message
	Logger message
	
end sub

function DoPings(ipaddress, status, response_time)

    dim wmistring,colItems,objItem

    wmistring = "Select * From win32_PingStatus where address='" & ipaddress & "'" 

    ' todo debug wscript.echo "WmiString ", wmistring

    Set colItems = objWMIService.ExecQuery(wmistring)

	
    
    for each objItem in colItems
        status = objItem.statuscode
        response_time = objItem.ResponseTime
    next

	if status = 0 Then
		' todo Debug
		DoPings = True 
	else
		DoPings = False 
	end if

end function

function Append_to_file(full_file_path,record)
' this function is meant to add a single record to a given files
' a poor choice of you need to do thousands of append_results_to_log
	dim localTS

	message full_file_path
	set localTS = fsoLogger.OpenTextFile(full_file_path,ForAppending, true) ' create if not exists

	localTS.WriteLine(record)

	localTS.Close()


end function


function GetProcessEnvironmentVariable(variable_name)
' returns environment variable '

	dim localValue, wshProcessEnv
	Set wshShell = CreateObject( "WScript.Shell" )
	Set wshProcessEnv = wshShell.Environment( "PROCESS" )
	'todo Debug WScript.Echo "SYSTEM:  TEMP=" & wshSystemEnv( "TEMP" )

	localValue = wshProcessEnv(variable_name)

	Set wshProcessEnv = Nothing
	Set wshShell     = Nothing

	GetProcessEnvironmentVariable = localValue

end function

' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Subroutine.....: ping_average
'  Purpose........: calls WMI to execute pings to an IP
'  Arguments......: target_ip, loop_count,average_value
'  Example........: 
'  Requirements...: ITS_FRAMEWORK, WMI
'  Notes..........: 
' -------------------------------------------------------------------------- 

function ping_average(target_ip, loop_count, average_value) 

	dim objWMIService
	dim colItems
	dim objItem
	dim Counter ' loop_count from parameters
	dim AverageResponse ' returned average
	dim status, response_time ' wmi status
	dim lstatus ' function call status


	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

	loop_count =cInt(loop_count) ' became a string as a parameter and broke the while loop 

	Counter = 0
	AverageResponse = 0 
	Do while Counter < loop_count
		' todo Debug wscript.echo "Counter " & counter & " loop count " & loop_count
		lstatus = DoPings (target_ip , status, response_time)
		'if 	DoPings( target_ip , status, response_time) then 

		if lstatus <> 0 then   
			counter = counter + 1
			AverageResponse = ( (AverageResponse + response_time) / counter )
		else
			
			ping_average = False
		
			exit function

		end if
	Loop

	' todo Debug wscript.echo "Average was " & AverageResponse
	average_value = round(AverageResponse,4)
	ping_average = true

end function


' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: SendSMTP
'  Purpose........: Send an email
'  Arguments......: (To,Subject,Body,[attachment_file_name)
'  Example........: SendSMTP (prodSupportEmail,"Error detected", "Please fix this")
'  Requirements...: CDONTS, network access 
'  Notes..........: The subject will be prepended by the datetime and scriptname. 
'			Files attached if they exist
'			Errors are handled inline and returned to the caller
' -------------------------------------------------------------------------- 


Function SendSMTP(ToStr, Subject, Body, aFilename)

	dim objEmail
	
	
	subject = now() & "--->"& scriptname & "--->" & Subject 
	
	message "Sending mail to: " & ToStr & ", subject: " & subject
	
	on error resume next


	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = strMailFrom
	objEmail.To = ToStr
	objEmail.Subject = Subject
	objEmail.Textbody = Body
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
	        smtpServer 
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	
	if aFilename <> "" then
	
		if fsoLogger.FileExists(aFilename) then 
		
			message "Attaching " & aFilename
		
			objEmail.AddAttachment aFilename
			
		else
		
			message "Attachment " & aFilename & " does not exist"
			
		end if
		

	end if

		
	objEmail.Configuration.Fields.Update

	objEmail.Send
	
	if err.number = 0 then SendSMTP = True
	
	on error goto 0




end function



' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: SendSMS
'  Purpose........: Call the Message Media api
'  Arguments......: msisdn - mobile number. message - 160 characters, replytoemail
'  Example........: SendSms (0418674002 , "It's coffee time"
'  Requirements...: WSH
'  Notes..........: This is the basis of all message output from scripts, except alarms
' -------------------------------------------------------------------------- 
Function SendSMS (MobileNumber, strMessage, ReplyToEmail)

	sendsms = true

	dim http
	dim URL
	dim SMSResponse 

	' Build URL
	URL = "http://www.messagenet.com.au/dotnet/lodge.asmx/LodgeSMSMessageWithReply?Username=" & mmUsername & "&Pwd=" & mmPassword & "&PhoneNumber=" & MobileNumber & "&PhoneMessage=" & strMessage & "&ReplyType=EMAIL&ReplyPath=" & ReplyToEmail

	debugwrite 2, "URL is " & URL

	on error resume next
	
	' Create HTTP object
	Set Http = CreateObject("Microsoft.XmlHttp")

	if err.number <> 0 then
		
		message "Failed to create MS XMLHTTP object"
	
		on error goto 0

		exit function
	end if
	
	' Create connection
	debugwrite 2, "Doing GET "
	
	http.open "GET", URL, False 

	if err.number <> 0 then
		
		message "Failed to GET url =" & URL

		on error goto 0
	
		exit function
	end if
	
	' Send URL data
	debugwrite 2, "Doing Send"
	
	http.send
	
	if err.number <> 0 then
		
		message "Failed in XML HTTP Send"
	
		on error goto 0
		
		exit function
	end if
	
	
	on error goto 0

	' Get the result
	SMSResponse = http.responseText

	' Parse result
	if instr(SMSResponse, "Message sent successfully")  then 
		
		message "Message Sent successfully"
		
	   	SendSMS = True 
	else

	   	message "Failed SMS response is " & SMSResponse 
	   	
	   	SendSMS = False 
	
	end if

	set http = nothing 

End Function 




' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: Array to String
'  Purpose........: convert a one dimensional array to string
'  Arguments......: arrayname
'  Example........: str = arrtoStr(fred)
'  Requirements...: WSH
'  Notes..........: yeah...
' -------------------------------------------------------------------------- 
Public function arrtostr(arrayname)
    	Dim i
    	Dim tmpstr
    	tmpstr = CStr(arrayname(LBound(arrayname)))
    	For i=LBound(arrayname)+1 To UBound(arrayname)
    		tmpstr = tmpstr & "," & CStr(arrayname(i))
    	Next
    	arrtostr=tmpstr
End function
    
    
    
' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: ConnectAdoDb
'  Purpose........: Connect to generic adodb database
'  Arguments......: connectionstring
'  Example........: str = ConnectAdoDb("driver={mssql.1;database=fred;)
'  Requirements...: WSH
'  Notes..........: 
' -------------------------------------------------------------------------- 
Function ConnectAdoDb(ConnectionString, cnx)

    	                                                                               
        on Error Resume Next                                                           
       
        Set	cnx = createobject("ADODB.Connection")                             
                                                                                       
        if err.number <> 0 then                                                        

		message	"Failed	to connect to database described in: " & ConnectionString  

		cnx = NULL                                                                 

		on error goto 0                                                            

		exit function                                                              

        end if                                                                         
                                                                                       
        cnx.open ConnectionString
                                                                                       
        if err.number <> 0 then       
                                                         
		message	"Failed	to Open	to database described in: " & ConnectionString     

		cnx = NULL                                                                 

		on error goto 0                                                            

		exit function                                                              

        end if                                                                         
                                                                                       
        On error goto 0                                                                
                                                                                       
        message "Connection opened to SQL server using " &  ConnectionString         
                                                                                       
        ConnectAdoDB = True                                                             

End Function




Function ConnectEXCEL(cnx, ExcelPath)

    Dim cnxstr

	'cnxstr = "Driver={Microsoft Excel Driver (*.xls)}; DBQ=" & ExcelPath  & ";ReadOnly=0;"
	
	' Never understood why, but some windows installations has different MDAC librarys. Sometimes this one works, where the above doesn't
	cnxstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& ExcelPath  &";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
	'cnxstr = "DSN=qh1"
	
	message "Connecting to Excel databased using " & cnxstr
	
	on error resume next
	
	Set cnx = createobject("ADODB.Connection")
	
	if err.number <> 0 then

	    	message "Failed to Connect to spreadsheet described in: " & cnxstr

	   	cnx = NULL

	    	on error goto 0

	    	exit function

	end if
	
	    
	cnx.open cnxstr


	if err.number <> 0 then

	    	message "Failed to Open to Spreadsheet described in: " & cnxstr

	   	cnx = NULL

	    	on error goto 0

	    	exit function

	end if


	message "Connection opened to " & excelPath
    
    	on error goto 0
    	
    	ConnectEXCEL = True


End Function


Function ConnectMSaccess(cnx, MSaccessPath)

    Dim cnxstr

	cnxstr = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & MSaccessPath & ";Uid=Admin;Pwd=;"
	
	message "Connecting to MSaccess databased using " & cnxstr
	
	on error resume next
	
	Set cnx = createobject("ADODB.Connection")
	
	if err.number <> 0 then

	    	message "Failed to Connect to spreadsheet described in: " & cnxstr

	   	cnx = NULL

	    	on error goto 0

	    	exit function

	end if
	
	    
	cnx.open cnxstr


	if err.number <> 0 then

	    	message "Failed to Open to Spreadsheet described in: " & cnxstr

	   	cnx = NULL

	    	on error goto 0

	    	exit function

	end if


	message "Connection opened to " & MSaccessPath
    
    	on error goto 0
    	
    	ConnectMSaccess = True


End Function



' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: MakeRecordSet
'  Purpose........: creates an ado recordset based on a connection string and some sql
'  Arguments......: recordset_to_return, sql, connectionname
'  Example........: status = Makerecordset(rsContext, "select * from fred", cnx_object)
'  Requirements...: WSH
'  Notes..........: returns boolean
' -------------------------------------------------------------------------- 
Function MakeRecordSet(rsx,strSQL,ConnectionName, ReadOnly)

	Dim errorstring
	Dim LogString
	
	Set rsx = createobject("ADODB.Recordset")

	On error resume next

	debugwrite 1, "Makerecordset--->Cnx = " & ConnectionName & "--->SQL=" & strSQL
	
	rsx.ActiveConnection =  ConnectionName
	rsx.CursorLocation = adUseServer

	if err.number <> 0 then

		errorString = err.number

		Message "Failed to access database Active connection", errorstring

		MakeRecordSet = False
		
		on error goto 0

		exit function

	end if

	if ReadOnly then

		rsx.CursorType = adOpenStatic
		rsx.LockType = adLockReadOnly
	
	else
	
		rsx.CursorType = adOpenKeyset
		rsx.LockType = adLockOptimistic
		
	end if 
	
	rsx.Source = strSQL

	rsx.Open

	if err.number <> 0 then

		LogString = "Error-->" & err.description 

		message "Failed to access database -->" & LogString

		MakeRecordSet = False
		
		on error goto 0

		exit function


	end if

	debugwrite 1, "Record Set Created Successfully - SQL = " & strSQL

	on error goto 0
	
	MakeRecordSet = True


End Function


' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: readtable
'  Purpose........: utility to read a whole table and display the rows
'  Arguments......: tablename connectioncontext
'  Example........: status = readtable("fred",cnx_object)
'  Requirements...: WSH
'  Notes..........:  
' -------------------------------------------------------------------------- '
Function ReadTable(tablename,ConnectionContext,title)
	Dim rsMap
	Dim sql
	Dim i
	
	
	sql = "select * from "& tablename 
	
	message "Opening " & tablename

	
	status = MakeRecordSet(rsMap,sql,ConnectionContext, True)
	
	if not (status) then
		message "Failed to create a recordset"
		exit Function
	end if
	
	
		
	
	message "Processing " & tablename
	
	status = displayrs(rsMap,title)
	
	rsmap.close
	ReadTable = True
	
End Function 

' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: displayrs
'  Purpose........: utility to display all the rows in a recordset
'  Arguments......: recordset
'  Example........: status = displayrs(rsObject)
'  Requirements...: WSH or ASP
'  Notes..........:  
' -------------------------------------------------------------------------- '

function displayrs(rs,title)

	dim i
	

	
	
	if ASP then 
	
	
		if rs.recordcount > 0 then 
	
			rs.movefirst

			response.write "<H4>" & title & ": </H4>"
	
			response.write "<table border=1><div align=center>"
	
			response.write "<TR>"
			
			for i = 0 to rs.fields.count -1
			
				response.write "<TD><B>" & rs.fields(i).name & "</TD>"
			
			next
			
			response.write "</TR>"
			
		else
		
			response.write "<H4> " & title & " -- No records ! </H4> "
		
		
		end if 
		
	end if
	
	do while not rs.EOF
	
		if ASP then response.write "<TR>"
		
		for i = 0 to rs.fields.count -1
		
			if ASP then response.write "<TD>" & rs.fields(i) & "</TD>"
				
			
			message rs.fields(i).name & "---->" & rs.fields(i).value
	
		next
		
		if ASP then response.write "</TR>"
		
		rs.movenext
	loop
	
	if ASP then response.write "</table>"

end function



' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Function.......: outputRsTofile
'  Purpose........: utility to dump a whole recordset to a flat file in its_data
'  Arguments......: recordsetname, filename (in its_data) 
'  Example........: status = readtable("fred",cnx_object)
'  Requirements...: WSH
'  Notes..........:  - toDo, parameterise delimiter
' -------------------------------------------------------------------------- '
function outputRsTofile(rs, dataset, flatPath)
	

	' bail out if we are switched off
	if not WriteFlatFile then
		outputRstofile = True 
		exit function
	end if
	
	
	Dim i , j
	Dim tso
	Dim oLine
	Dim oFile
	
	oFile = flatPath & "\" & dataset & "_" & ScriptDateStamp & ".csv"
	
	message "Doing flat file output for " & dataset	
	
	set tso = fsoLogger.CreateTextFile(oFile,True)
	
	oLine = ""
	
	For i = 0 to rs.fields.count - 1
			
		oLine = oLine & rs.fields(i).name & ","
		
	Next
	
	' remove last Pipe
	
	oLine = left(oLine, len(oLine)-1)
	
	' write header oLine
	
	tso.writeline(oLine)
	
	oLine = ""
		
	do while not rs.EOF
	
		status = write_a_row_flat(rs,tso)
		
		PrintPercentage j, rs.recordcount
		
		j = j + 1
		
		rs.Movenext
				
	loop
	
	tso.close
	
	set tso = nothing

	outputRstofile = True 

end function 


function write_a_row_flat(rs,tso)

	dim i
	dim oLine
	
	For i = 0 to rs.fields.count - 1
				
'	if rs.fields(i).name = "Request_Prefix" then  message "Flatfile conversion of Request_Prefix" & rs.fields(i)

		if IsNumeric(rs.fields(i)) then
		
			oLine = oLine & cstr(rs.fields(i)) & ","
			
			
		else
		
			oLine = oLine & rtrim(rs.fields(i)) & ","
		
		end if
	
	Next

	' remove last Pipe

	oLine = left(oLine, len(oLine)-1)
	
	' write header line
	
	tso.writeline(oLine)
	

	write_a_row = True

end function


' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Sub ......: debugwrite
'  Purpose........: one routine for 4 levels of debug
'  Notes..........: debug level routines
' -------------------------------------------------------------------------- '

sub debugwrite(level, pmessage)

	

	if level <= DebugLevel then 
	
		message "dbx-"& cstr(level) & ">" & cstr(pmessage)
	
	end if

end sub

'return an integer date as dd/m/yyyy
Function ConvertIntDateToDate(StrIntDate)

	Dim DateStr

	DateStr = mid(StrIntDate,5,4) & "/" & mid(StrIntDate,3,2) & "/" & mid(StrIntDate,1,2)
	
	DateStr =  Cdate(DateStr)
	
	ConvertIntDateToDate = DateStr
	

end function



' On every 10 percent boundary of current out of total, tell us where we are up to
sub PrintPercentage(current, total)

	Dim percent
	
	
	percent = current/total * 100
	
	if (percent = 10)  or _
	(percent = 20)  or _
	(percent = 30)  or _
	(percent = 40)  or _
	(percent = 50)  or _
	(percent = 60)  or _
	(percent = 70)  or _
	(percent = 80)  or _
	(percent = 90)  or _
	(percent = 100)  then
	
		writeresponse FormatPercent(current/total)
	end if

end sub

' produce a copy of a given recordset
Function CloneRS(rsSource,rsTemp)   
      
    Dim F   
  
    Set rsTemp = createobject("ADODB.Recordset") 
      
    For Each F In rsSource.Fields 
      
        If F.Type <> adChapter Then 
          
          
	        rsTemp.Fields.Append F.Name, F.Type, F.DefinedSize, F.Attributes And adFldIsNullable 
          
              With rsTemp(F.Name) 
                
                      .Precision = F.Precision 
                      .NumericScale = F.NumericScale 
                        
              End With 
          
        End If 
      
    Next   
          
    CloneRS = True 
    
      
End Function 





' ----------------------------------------------------------------------------- 
' adapted all output statements for ASP compatibility
' ----------------------------------------------------------------------------- 
sub writeresponse(message)

	
	if ASP then

		if debuglevel > 0 then response.write "<br>" & message & "</br>"
	else
	
		wscript.echo message
	
	end if
end sub


' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 
'  Subroutine.....: SendSMS2Group
'  Purpose........: Look up the specified Group Name in the default AD
'			if a member has a mobile number, send specified text
'			as SMS
'  Arguments......: rs - a recordset containing Common Names and mobilenumbers
'		    strMessage - a message string
'		    strReplyTo - an email address for SMS replies
'  Example........: SendSMS2Group("managers","Please come to meeting"
'  Requirements...: 
'  Notes..........: 
' -------------------------------------------------------------------------- 
' -------------------------------------------------------------------------- 

function SendSMS2Group (rs,pMessagestr,strReplyTo)

	dim pReplyTo
	dim MobileNumber
	
	if strReplyTo <> "" then 
		
		pReplyTo = strReplyTo
		
	else
	
		pReplyTo = replyTo ' the constant
		
	end if
	
	message "Message length is " & len(pMessagestr)
		
	rs.movefirst
	
	do while not rs.EOF
	
		MobileNumber = rs.fields("MobileNumber")
		
		message "Sending to " & rs.fields("CommonName")

'		if true then 
		if  SendSMS(MobileNumber,pMessageStr,replyto) then
	
			message "Successfully sent"
		
		else
	
			message "Sending failed"
			
			exit function
			
		end if
		
		rs.movenext
		
	loop
	
	

	SendSMS2Group = True
	
end function

'
' Read the AD group members and build a recordset containing Cn and Mobile
function GetGroupMobiles(rs,strGroupname)

	dim ldapstring 
	dim objGroup
	dim objUser 
	
	on error resume next 
	
	ldapstring = "LDAP://CN=" & StrGroupname & ",CN=Users," & strDomainName
	
	set objGroup = getobject(ldapstring)
	
	if err.number <> 0  then
	
		message "No such group " & groupname
		
		exit function
		
	end if
	
	on error goto 0
	
	
	set rs = createobject("ADODB.recordset")
	
	rs.fields.append "CommonName", adVarchar, 50
	rs.fields.append "MobileNumber", adVarchar, 50
	
	rs.open

	message "Group " & StrGroupname & " has " & objGroup.members.count & " members"
	
	for each objUser in objGroup.members
	
		rs.addnew
		
		rs.fields("CommonName") = objUser.name
		rs.fields("MobileNumber") = objUser.mobile
		
	
	next
	
	
	set objGroup = nothing
	set objUser = nothing	

	rs.movefirst
	
	GetGroupMobiles = True
	
end function 


' write a standard HTML page header
sub writeheader(title)

	if not ASP then exit sub
	
	response.write "<HTML>"
	response.write "<HEAD>"
	response.write "<TITLE>" & Title & "</TITLE>"
	response.write "<BODY BACKGROUND=myimage.jpg>"
	response.write "<DIV align = center >"


end sub


sub writetrailer()

	response.write "</BODY>"
	response.write "</DIV>"
	response.write "</HTML>"


end sub
'