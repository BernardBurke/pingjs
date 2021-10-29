option explicit

dim objWMIService
dim colItems
dim objItem
dim Counter
dim AverageResponse
dim status, timeout

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Counter = 0
AverageResponse = 0 

function DoPings(ipaddress, status, timeout)

    dim wmistring 

    wmistring = "Select * From win32_PingStatus where address='" & ipaddress & "'" 

    'wscript.echo "WmiString ", wmistring

    Set colItems = objWMIService.ExecQuery(wmistring)
    
    for each objItem in colItems
        status = objItem.statuscode
        timeout = objItem.ResponseTime
    next

end function

Do while Counter < 100
    ' wscript.echo "Counter " & counter 
    DoPings "1.1.1.1",status,timeout 
    counter = counter + 1
    AverageResponse = ( (AverageResponse + timeout) / counter )
Loop

wscript.echo "Average was " & AverageResponse
