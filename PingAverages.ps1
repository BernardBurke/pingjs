$Iplist = Get-Content ips.txt
echo "Running"
$group = @()
foreach ($ip in $Iplist) {
  echo $ip
  $status = @{ "ServerIP Name" = $ip; "TimeStamp" = (Get-Date -f s) }
  $pings = Test-Connection $ip -Count 1 -ea 0
  if ($pings) {
    $status["AverageResponseTime"] =
        ($pings | Measure-Object -Property ResponseTime -Average).Average
    $status["Results"] = "Up"
  }
  else {
    $status["Results"] = "Down"
  }

  New-Object -TypeName PSObject -Property $status -OutVariable serverStatus
  $group += $serverStatus
}

$group | Export-Csv results.csv -NoTypeInformation