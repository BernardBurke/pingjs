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

function DoPings(status, timeout)

    Set colItems = objWMIService.ExecQuery("Select * From win32_PingStatus where address='1.1.1.1'")
    
    for each objItem in colItems
        status = objItem.statuscode
        timeout = objItem.ResponseTime
    next

end function

Do while Counter < 100
    ' wscript.echo "Counter " & counter 
    DoPings status,timeout 
    counter = counter + 1
    AverageResponse = ( (AverageResponse + timeout) / counter )
Loop

wscript.echo "Average was " & AverageResponse
