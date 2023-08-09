Clear-Host
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
$CI = Get-ComputerInfo
[Microsoft.Management.Infrastructure.CimInstance]$LOC = Get-NetIPAddress  -AddressFamily IPv4 -IPAddress 192.168.100*;
$II = ($LOC.InterfaceIndex);
$NA = (Get-NetAdapter -Physical -InterfaceIndex $II).MacAddress;
$DNSSUF = (Get-DnsClient -InterfaceIndex $II).ConnectionSpecificSuffix;
$SMBCONN = Get-SmbConnection;
$DNSIP = (Get-DnsClientServerAddress -InterfaceIndex $II);
[string] $HN = $CI.CsUserName;


$dirpath = ($env:HOMEPATH + "\Desktop\")
$PCInfo = $dirpath +  "PCInfo.txt"


[System.IO.Stream]$FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($PCInfo), [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [IO.FileShare]::Read))
while ($FileStreamPCinfo.CanWrite -eq $false) 
{
    try 
    {
       $FileStreamPCinfo.Dispose();
       $FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($PCInfo), [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
    }
    catch 
    {
        echo "File is busy";
        [System.Threading.Thread]::Sleep(2000);
        $FileStreamPCinfo.Dispose();
    }
}

$sw = New-Object System.IO.StreamWriter([System.IO.Stream]$FileStreamPCinfo, [Text.Encoding]::UTF8);
$sw.AutoFlush = $true;
$sw.WriteLine("======");
$sw.WriteLine("Host: "+ $HN);
$sw.WriteLine("Date: " + $(get-date))
$sw.WriteLine("MAC Address: " + $NA); 
$sw.WriteLine("IPv4: " + $LOC.IPAddress);
$sw.WriteLine("DNS INFO:");
$sw.WriteLine((Out-String -InputObject $DNSIP));
$sw.WriteLine("DNS Suffix: " + $DNSSUF);
$sw.Flush();
$sw.WriteLine("SMB Shares: " + (Out-String -InputObject $SMBCONN));
$srvcs = Get-Service | where{$_.StartType -match "Automati.*"} | Select DisplayName,Name,StartType,Status
$vault1 = (vaultcmd /listcreds:"Windows Credentials")
$sw.WriteLine((Out-String -InputObject $vault1));
$vault2=  (vaultcmd /listcreds:"Учетные данные Windows") 
$sw.WriteLine((Out-String -InputObject $vault2));
$sw.Flush();
$sw.WriteLine($(Out-String -InputObject $(Get-NetFirewallRule -All | where { $_.Enabled -eq "True"})));
$sw.Flush();
$sw.WriteLine("Services: " + (Out-String -InputObject $srvcs));
$sw.Flush();
$sw.WriteLine("Windows info:");
$sw.WriteLine($(Out-String -InputObject $CI));
$sw.Flush();

try
{
  $sw.Dispose();
}
catch 
{
    echo "Unable to dispose StreamWriter object";
}

try 
{
    $FileStreamPCinfo.Dispose();
}
catch 
{
    echo "Unable to dispose filestream object";
}

