try {
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope Process
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope LocalMachine
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope CurrentUser
}
catch
{
    echo "Error occured during the change of execution policy"
}

New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress @("192.168.100.73", "192.168.100.24", "192.168.100.1") -Direction Inbound -Enabled "True";    
New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress @("192.168.100.73", "192.168.100.24", "192.168.100.1") -Direction Outbound -Enabled "True";

winrm quickconfig -quiet
[System.Threading.Thread]::Sleep(2300)
Enable-PSRemoting –Force
[System.ServiceProcess.ServiceController] $SC = Get-Service -Name WinRM;
[string]$ST = $SC.StartType

if (!("Automatic".Equals($ST)))
{
    Set-Service -Name WinRM -StartupType Automatic
}
else
{
    echo "WinRM is already set to Automatic"
}
Set-PSSessionConfiguration -ShowSecurityDescriptorUI -Name Microsoft.PowerShell;