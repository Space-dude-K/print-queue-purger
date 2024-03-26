param
(
	[Parameter(Mandatory = $false)]
    $hostName = "127.0.0.1",
	[Parameter(Mandatory = $false)]
    $userNameForSID = "userName",
	[Parameter(Mandatory = $false)]
    $forceClear = 0
)

$version = "1.2"
[System.Console]::Title = "PrintQueuePurger [v. $version]"

Add-Type -AssemblyName "System.Printing"

# Получаем SID User-а
$userSid = (Get-AdUser -Identity $userNameForSID).SID
# Получаем имя принтера по умолчанию из реестра
$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users', $hostName)
$RegKey= $Reg.OpenSubKey("$userSid\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows")
$printName = ($RegKey.GetValue("Device")).Split(",")[0].Trim()

Write-Host $printName

Function ClearQueue()
{
	Write-Host "Clear -> $hostName for printer $printName"
	
    try
	{
		$ps = [System.Printing.PrintServer]::new("\\$hostName", [System.Printing.PrintSystemDesiredAccess]::AdministrateServer)
		$pq = [System.Printing.PrintQueue]::new($ps, "$printName", [System.Printing.PrintSystemDesiredAccess]::AdministratePrinter)

		# TODO: Проверки доступности хоста, статуса принтера
		Write-Host ("Current status for {0} -> {1}" -f $pq.FullName, $pq.QueueStatus)
		Write-Host ("Current jobs for {0} -> {1}" -f $pq.FullName, $pq.NumberOfJobs)

		if($forceClear -or $pq.NumberOfJobs -gt 0)
		{
			Write-Host "Purge!"
			#$pq.Purge()
		}	
	}
	catch
	{
		Write-Host "Access print error $_"
	}
}

ClearQueue

Write-Host "End"