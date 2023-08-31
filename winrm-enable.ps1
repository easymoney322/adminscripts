$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
Clear-Host;

try 
{
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope Process;
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope LocalMachine;
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -Scope CurrentUser;
}catch
{
    echo "Error occured during the change of execution policy";
}


$whitelist = @("192.168.100.73", "192.168.100.24", "192.168.100.1", "192.168.0.24", "192.168.100.43");
$firewallRules = @();
$firewallRulesInbound = @();
$firewallRulesOutbound = @();
$rulesExist = $true;
$ruleInboundExists = $true;
$ruleOutboundExists = $true;

try
{
    $firewallRules = Get-NetFirewallrule -DisplayName 'NewFromADM' -ErrorAction Stop;
}catch  [Microsoft.PowerShell.Cmdletization.Cim.CimJobException] 
{
    if ($_.Exception.Message -imatch ".*DisplayName.*")
    {
        $rulesExist = $false;
    }
}


if ($rulesExist -eq $true)
{
  for ($i = 0; $i -lt $firewallRules.Count; $i++)
  {
    if ($firewallRules[$i].Direction.toString() -eq "Inbound")
    {
        $firewallRulesInbound += $firewallRules[$i];
        $ruleInboundExists = $true;
    }else
    {
        $firewallRulesOutbound += $firewallRules[$i];
        $ruleOutboundExists = $true;
    }
  }
}else 
{
    New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress $whitelist -Direction Inbound -Enabled "True";  
    New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress $whitelist -Direction Outbound -Enabled "True";
    Exit;
}

if ($ruleInboundExists -eq $false)
{
    New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress $whitelist -Direction Inbound -Enabled "True";   
}else
{
    $firewallRulesInbound | Set-NetFirewallRule -Action Allow -RemoteAddress $whitelist -Enabled "True";
}

if ($ruleOutboundExists -eq $false)
{
    New-NetFirewallRule -DisplayName NewFromADM -Action Allow -RemoteAddress $whitelist -Direction Outbound -Enabled "True";
}else 
{
    $firewallRulesOutbound | Set-NetFirewallRule -Action Allow -RemoteAddress $whitelist -Enabled "True";
}



winrm quickconfig -quiet;
[System.Threading.Thread]::Sleep(2300);
Enable-PSRemoting –Force;
[System.ServiceProcess.ServiceController] $SC = Get-Service -Name WinRM;
[string]$ST = $SC.StartType;

if (!("Automatic".Equals($ST)))
{
    Set-Service -Name WinRM -StartupType Automatic;
}else
{
    echo "WinRM is already set to Automatic";
}
Set-PSSessionConfiguration -ShowSecurityDescriptorUI -Name Microsoft.PowerShell;
