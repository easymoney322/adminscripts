$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
$DLDS = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path;
Invoke-WebRequest -Uri 'https://openhardwaremonitor.org/files/openhardwaremonitor-v0.9.6.zip' -OutFile $DLDS;
Expand-Archive ($DLDS+ "ohm.zip")  -DestinationPath 'C:\Program Files\'
$scheduleObject = New-Object -ComObject schedule.service;
$scheduleObject.connect();
$rootFolder = $scheduleObject.GetFolder("\");
$rootFolder.CreateFolder("Open Hardware Monitor");
Register-ScheduledTask -TaskName "Startup" -Action $(New-ScheduledTaskAction -Execute 'C:\Program Files\OpenHardwareMonitor\OpenHardwareMonitor.exe') -TaskPath "\Open Hardware Monitor\" -RunLevel Highest -User "System" -Trigger $(New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -minutes 3));
Start-ScheduledTask -TaskPath "\Open Hardware Monitor" -TaskName "Startup";