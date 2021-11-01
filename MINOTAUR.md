# Framework for client-side testing and logging

This vbscript can be used to keep records of the ping times to a given server.

It uses the ITS Script architecture described at https://github.com/BernardBurke/Bscripting

I'm intending to merge the repos over time.


## Setup

Customise its_local_setup.cmd for ITS (as per architecture doc linked above)

The provided CMD should work fine out of the box (as there is no data and logfiles go to TEMP)


```
set fulp=%cd%
set ITS_ROOT=%fulp%\ITS_scripting
set ITS_DATA=%ITS_ROOT%\data
set ITS_PROCEDURES=%ITS_ROOT%
set ITS_LOG=%TEMP%\log
mkdir %ITS_LOG%

```


### Sample usage and output

```
cscript ITS_scripting\minotaur.wsf /debug:1 /ping:www.integralife.com /loop:100 /log  

Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

C:\Users\ben\AppData\Local\Temp\log is log 
C:\Users\ben\Documents\repos\pingjs\ITS_scripting\data is data 
C:\Users\ben\Documents\repos\pingjs\ITS_scripting is procedures
ben is username
29/10/2021 1:39:48 PM--->minotaur--->Scriptname Initialised with logfile C:\Users\ben\AppData\Local\Temp\log\minotaur_ben_202110291339.log
FSO initialised for logger
TSO logger opened logfile & LogFileSpec
29/10/2021 1:39:48 PM--->minotaur--->Appending to C:\Users\ben\AppData\Local\Temp\minotaur_ben.log
29/10/2021 1:39:48 PM--->minotaur--->Ping Loops 100
29/10/2021 1:39:48 PM--->minotaur--->Ping Target www.integralife.com
29/10/2021 1:39:48 PM--->minotaur--->Debuglevel set to 1
29/10/2021 1:39:48 PM--->minotaur--->ping_target www.integralife.com
29/10/2021 1:39:48 PM--->minotaur--->About to do 100 executions
29/10/2021 1:39:50 PM--->minotaur--->Success with average 0.0809
29/10/2021 1:39:50 PM--->minotaur--->C:\Users\ben\AppData\Local\Temp\minotaur_ben.log
29/10/2021 1:39:50 PM--->minotaur--->Appended to C:\Users\ben\AppData\Local\Temp\minotaur_ben.log

```

### Typical output after multiple executions

```

ben,29/10/2021 1:23:09 PM,0.092,www.integralife.com,100
ben,29/10/2021 1:23:39 PM,0.011,1.1.1.1,1000
ben,29/10/2021 1:24:16 PM,0.01,www.integralife.com,1000
ben,29/10/2021 1:30:06 PM,0.01,www.integralife.com,1000
ben,29/10/2021 1:39:50 PM,0.0809,www.integralife.com,100
ben,29/10/2021 1:43:53 PM,1.7679,techcrunch.com,100

```

# Next features 

- http GET without Cache (check a webservers health)
This has just been implemented! Example below
- rest service call with timings (GET and POST)
- customise **message** function in CommonFunctionsLibrary to post to splunk and zabbix


### Example with http GET timings 

```
cscript ITS_scripting\minotaur.wsf /debug:1 /ping:www.integralife.com /loop:10 /log /download:"https://deadline.com/wp-content/uploads/2019/05/amc.jpg" /downloadfile:%TEMP%/fred.jpg
```

### Log contents

```
ben,2/11/2021 9:04:40 AM,1.4202,www.integralife.com,10,https://deadline.com/wp-content/uploads/2019/05/amc.jpg,591
ben,2/11/2021 9:05:04 AM,1.0148,www.integralife.com,10,https://deadline.com/wp-content/uploads/2019/05/amc.jpg,896
```

