$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
$DLDS = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path;

$dlfold = Get-Item -Path $DLDS;
$wasreadonly = $false;
if ($dlfold.Attributes -imatch ".*ReadOnly.*" ) 
{
    $wasreadonly = $true;
    $dlfold.Attributes = $dlfold.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly;
    
} 
                               
$DLDS += "/ohm.zip"
Invoke-WebRequest -Uri 'https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip' -OutFile $DLDS;
Expand-Archive ($DLDS)  -DestinationPath 'C:\Program Files\';

if ($wasreadonly)
{
    $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
}

$scheduleObject = New-Object -ComObject schedule.service;
$scheduleObject.connect();
$rootFolder = $scheduleObject.GetFolder("\");
$rootFolder.CreateFolder("Open Hardware Monitor");
Register-ScheduledTask -TaskName "Startup" -Action $(New-ScheduledTaskAction -Execute 'C:\Program Files\OpenHardwareMonitor\OpenHardwareMonitor.exe') -TaskPath "\Open Hardware Monitor\" -RunLevel Highest -User "System" -Trigger $(New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -minutes 3));
Start-ScheduledTask -TaskPath "\Open Hardware Monitor" -TaskName "Startup";
