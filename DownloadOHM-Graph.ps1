try 
{
    $OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
}
catch
{
    Write-Debug ("Errors occured while setting an encoding");
}
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
$PSDefaultParameterValues['Invoke-WebRequest:MaximumRetryCount'] = 5;
Clear-Host;
$INSTPATH = 'C:\Program Files\';
$graphiteVersion = '0.30.0';
$ExporterVersion = "0.23.1";
$SmartMonVersion = "7.4-1";
$SmartMonReleaseDir = "7_4";
$SmartCTLExporterVersion = "0.11.0";
$downloadDir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path;
$dlfold = Get-Item -Path $downloadDir;

function ServiceCheck([string] $ServiceName)
{
    $serviceQuery = "Select * from Win32_Service WHERE name like  '%" + $ServiceName + "%'";
    $SrvcList = (Get-CimInstance -Query $serviceQuery);
    if ($SrvcList -ne $null)
    {
        Write-Host "Found service " + $ServiceName;
        [System.Array]$ServicePath = @();
        $ServicePathString= $SrvcList.PathName -split " --";
        if ($ServicePathString -ne $null)
        {
            for ($i=0; $i -lt $ServicePathString.Length; $i++)
            {
                if (($ServicePathString[$i][2] -eq ':') -or ($ServicePathString[$i][1] -eq ':'))
                {
                    [System.Array]$indxList=@();
                    for ($indx =0;$indx -lt $ServicePathString[$i].Length; $indx++)
                    {
                        if ($ServicePathString[$i][$indx] -eq '"')
                        {
                            $indxList += $indx;
                        }
                    }
                    if ($indxList.Length -gt 1)
                    {
                        write-host "LVAL=" $lval;
                        [int]$lval = $indxList[0];
                        [int]$rval = $indxList[1];
                        $bubble = [string]($ServicePathString[$i]).Substring($lval+1, $rval-$lval-1);
                        [string]$ServicePathString[$i] = $bubble
                    }
                    $ServicePath += $ServicePathString[$i];
                }
            }
            if ($ServicePath -ne $null)
            {
                Write-Host "Succesfully extracted executable path from " + $ServiceName + " service";
                for ($k=0; $k -lt $ServicePath.Length; $k++)
                {
                    write-host $ServicePath[$k]
                    if (Test-Path -Path $($ServicePath[$k]) -PathType Leaf)
                    {
                        Write-Host ("Found " + $ServiceName + " executable at " + $ServicePath[$k]);
                        return $true;
                    }
                }
                Write-Host ("Exporter executables were not found"); 
                return $false;
            }
        }else
        {
            Write-Host ("Service path is empty");
            return $false;
        }
    }
    $serviceQuery = "Select * from Win32_Service WHERE DisplayName like  '%" + $ServiceName + "%'"; ##If service wasnt found, do the same but for DisplayName
    $SrvcList = (Get-CimInstance -Query $serviceQuery);
    if ($SrvcList -ne $null)
    {
        Write-Host ("Found service " + $ServiceName);
        [System.Array]$ServicePath = @();
        $ServicePathString= $SrvcList.PathName -split " --";
        if ($ServicePathString -ne $null)
        {
            for ($i=0; $i -lt $ServicePathString.Length; $i++)
            {
                if (($ServicePathString[$i][2] -eq ':') -or ($ServicePathString[$i][1] -eq ':'))
                {
                    [System.Array]$indxList=@();
                    for ($indx =0;$indx -lt $ServicePathString[$i].Length; $indx++)
                    {
                        if ($ServicePathString[$i][$indx] -eq '"')
                        {
                            $indxList += $indx;
                        }
                    }
                    if ($indxList.Length -gt 1)
                    {
                        write-host "LVAL=" $lval;
                        [int]$lval = $indxList[0];
                        [int]$rval = $indxList[1];
                        $bubble = [string]($ServicePathString[$i]).Substring($lval+1, $rval-$lval-1);
                        [string]$ServicePathString[$i] = $bubble
                    }
                    $ServicePath += $ServicePathString[$i];
                }
            }
            if ($ServicePath -ne $null)
            {
                Write-Host "Succesfully extracted executable path from " + $ServiceName + " service";
                for ($k=0; $k -lt $ServicePath.Length; $k++)
                {
                    write-host $ServicePath[$k]
                    if (Test-Path -Path $($ServicePath[$k]) -PathType Leaf)
                    {
                        Write-Host ("Found " + $ServiceName + " executable at " + $ServicePath[$k]);
                        return $true;
                    }
                }
                Write-Host "Exporter executables were not found"; 
                return $false;
            }
        }else
        {
            Write-Host "Service path is empty";
            return $false;
        }
    }
    Write-Host ($ServiceName + " service wasn't found");
    return $false;
}


$wasreadonly = $false;
if ($dlfold.Attributes -imatch '.*ReadOnly.*') 
{
    $wasreadonly = $true;
    $dlfold.Attributes = $dlfold.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly;
} 


if ($(ServiceCheck "OhmGraphite") -eq $false) 
{
    if ($false -eq (Test-Path -Path $($INSTPATH+'OHMGraphite\OhmGraphite.exe') -PathType Leaf))
    {   
        Write-Host "Graphite installation was not found. Starting installation process.";
        $fileGraphite = $downloadDir+"/GrOHM.zip";
        if ($false -eq (Test-Path -Path $fileGraphite -PathType Leaf))
        {
            $dlUrl = 'https://github.com/nickbabcock/OhmGraphite/releases/download/v' + $graphiteVersion + '/OhmGraphite-' + $graphiteVersion + '.zip';
            try
            {
                Write-Debug ($dlUrl);
                Invoke-WebRequest -Uri $dlUrl -OutFile $fileGraphite -UseBasicParsing; 
            }catch {
                try 
                {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    WebClient.DownloadFile($dlUrl, $fileGraphite);
                }catch 
                {
                    Write-Error ($_.Exception.Message);
                    if ($wasreadonly)
                    {
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    [System.Threading.Thread]::Sleep(5000);
                    Exit; 
                }
            }
        }else
        {
            Write-Host ("Found Graphite archive. Proceeding to unzip.");
        }

        if($false -eq (Test-Path ($INSTPATH+'OHMGraphite\')))
        {
            New-Item -Path ($INSTPATH+'OHMGraphite\') -ItemType Directory;
        }
        Expand-Archive ($fileGraphite)  -DestinationPath ($INSTPATH+'OHMGraphite\') -Force;
        $configFilePath = $INSTPATH+'OHMGraphite\OhmGraphite.exe.config';
        [System.IO.Stream]$FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($configFilePath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
        while ($FileStreamPCinfo.CanWrite -eq $false) 
        {
            try 
            {
               $FileStreamPCinfo.Dispose();
               $FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($configFilePath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
            }catch 
            {
                Write-Debug ("File is busy");
                [System.Threading.Thread]::Sleep(5000);
                $FileStreamPCinfo.Dispose();
            }
        }
        $sw = New-Object System.IO.StreamWriter([System.IO.Stream]$FileStreamPCinfo, [Text.Encoding]::UTF8);
        $sw.AutoFlush = $true;
        $sw.WriteLine('<?xml version="1.0" encoding="utf-8" ?>');
        $sw.WriteLine('<configuration>');
        $sw.WriteLine('  <appSettings>');
        $sw.WriteLine('    <add key="host" value="localhost" />');
        $sw.WriteLine('    <add key="port" value="2003" />');
        $sw.WriteLine('    <add key="type" value="prometheus" /> ');
        $sw.WriteLine('    <add key="prometheus_port" value="4445" />');
        $sw.WriteLine('    <!-- This is the host that OhmGraphite listens on.');
        $sw.WriteLine('    `*` means that it will listen on all interfaces.');
        $sw.WriteLine('    Consider restricting to a given IP address -->');
        $sw.WriteLine('    <add key="prometheus_host" value="*" /> ');
        $sw.WriteLine('    <add key="prometheus_path" value="metrics/" /> ');
        $sw.WriteLine('    <add key="interval" value="5" />');
        $sw.WriteLine('  </appSettings>');
        $sw.WriteLine('</configuration>');
        try
        {
            $sw.Dispose();
        }catch 
        {
            Write-Error "Unable to dispose StreamWriter object";
            Write-Error ($_.Exception.Message);
        }

        try 
        {
            $FileStreamPCinfo.Dispose();
        }catch 
        {
            Write-Error "Unable to dispose filestream object:";
            Write-Error ($_.Exception.Message);
        }

        Start-Process -FilePath $($INSTPATH+'OHMGraphite\OhmGraphite.exe') -ArgumentList "install" -WorkingDirectory $($INSTPATH+'OHMGraphite\');
        Write-Host ("Successfully installed Graphite");
        [System.Threading.Thread]::Sleep(2000);
        try {
            Start-Service OhmGraphite;
        }catch
        {
            Write-Error "Unable to start Graphite service:";
            Write-Error ($_.Exception.Message);
        }

    }else
    {
        Write-Host ("Found Graphite Installation at " + $($INSTPATH+'OHMGraphite\'));
    }
}

if ($(ServiceCheck "windows_exporter") -eq $false) 
{
    if ($false -eq (Test-Path -Path $($INSTPATH+'windows_exporter\windows_exporter.exe') -PathType Leaf))
    {
        Write-Host "Wasn't able to find Windows Exporter installation path. Starting the installation";
        $exporterPrefixLink = "https://github.com/prometheus-community/windows_exporter/releases/download/v" + $ExporterVersion+"/windows_exporter-"+$ExporterVersion;
        
        if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem)
        {
            $exporterPrefixLink+="-amd64.msi";
        }else 
        {
            $exporterPrefixLink+="-386.msi";
        }

        $filePrometheusExporterInstaller = $downloadDir+'\ExporterInstaller.msi';
        if ($false -eq (Test-Path -Path $filePrometheusExporterInstaller -PathType Leaf))
        {
            try
            {
                Write-Debug ($exporterPrefixLink);
                Invoke-WebRequest -Uri $ExporterPrefixLink -OutFile $filePrometheusExporterInstaller -UseBasicParsing;
            }catch 
            {
                try 
                {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    WebClient.DownloadFile($ExporterPrefixLink, $filePrometheusExporterInstaller);
                }catch 
                {
                    Write-Error ($_.Exception.Message);
                    if ($wasreadonly)
                    {
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    [System.Threading.Thread]::Sleep(5000);
                }   Exit; 
            }
        }
        $argument = "/i $('"' + $filePrometheusExporterInstaller + '"') /qn /norestart /log $('"' + $downloadDir+'WindowsExporterInstaller.log' + '"') INSTALLDIR=`"C:\Program Files\windows_exporter\`"";
        (Start-Process msiexec.exe -argumentlist $argument).ExitCode;
        [System.Threading.Thread]::Sleep(5000);
        try
        {
            Start-Service "windows_exporter"
        }catch
        {
            Write-Error ("Unable to start windows_exporter service:");
            Write-Error ($_.Exception.Message);
        }
    }
}


if ($false -eq (Test-Path -Path $($INSTPATH+"SmartMonTools/bin/smartctl.exe") -PathType Leaf))
{
    $fileSmartMonInstaller = $downloadDir+ '\smartmontools-' + $SmartMonVersion +".win32-setup.exe";
    if ($false -eq (Test-Path -Path $fileSmartMonInstaller -PathType Leaf))
    {
        $smartMonPrefixLink = "https://github.com/smartmontools/smartmontools/releases/download/RELEASE_" + $SmartMonReleaseDir+ "/smartmontools-" +$SmartMonVersion + ".win32-setup.exe";
        try {
                Invoke-WebRequest -Uri $smartMonPrefixLink -OutFile $fileSmartMonInstaller -UseBasicParsing;
            }catch
            {
                try 
                {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    WebClient.DownloadFile($smartMonPrefixLink, $fileSmartMonInstaller);
                }catch 
                {
                    Write-Error ($_.Exception.Message);
                    if ($wasreadonly)
                    {
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    [System.Threading.Thread]::Sleep(5000);
                    Exit;
                }
            }
    }else
    {
        Write-Host("Installer for SmartMonTools found, proceeding to install...");
    }
    $smartMonDir = $INSTPATH+"SmartMonTools1";
    try {
            Start-Process  -NoNewWindow -FilePath $fileSmartMonInstaller -ArgumentList "/S /D=`"$smartMonDir`" ";
    }catch
    {
        Write-Error ($_.Exception.Message);
        if ($wasreadonly)
        {
            $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
        }
        [System.Threading.Thread]::Sleep(5000);
        Exit;
    }
}


if ($(ServiceCheck "SmartCtlExporter") -eq $false) 
{
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem)   #Test Installation
    {
        $SmartCTLexporterSufffixLink=".windows-amd64";
    }else 
    {
        $SmartCTLexporterSufffixLink=".windows-386";
    }
    $SmartCTLExporterExec = $INSTPATH +"smartctl_exporter-"+$SmartCTLExporterVersion + $SmartCTLexporterSufffixLink + "/smartctl_exporter.exe";
    if ($false -eq (Test-Path -Path $SmartCTLExporterExec -PathType Leaf)) 
    {
        Write-Host ("SmartCTLExporter not found"); ##Installation not found
        $SmartCTLexporterSufffixLink +=".zip"; ##Search for an archive
        $SmartCTLExporterFile = $downloadDir + '/smartctl_exporter-' + $SmartCTLExporterVersion + $SmartCTLexporterSufffixLink ;

        if ($false -eq (Test-Path -Path $($INSTPATH+"smartctl_exporter-"+$SmartCTLExporterVersion + $SmartCTLexporterSufffixLink) -PathType Leaf)) ##Archive not found
        { 
            Write-Host ("Downloading SmartCtlExporter");
            $SmartCTLExporterURL = 'https://github.com/prometheus-community/smartctl_exporter/releases/download/v' + $SmartCTLExporterVersion +"/smartctl_exporter-" + $SmartCTLExporterVersion + $SmartCTLexporterSufffixLink;
            Write-Debug ($SmartCTLExporterURL);
            try{
                  Invoke-WebRequest -Uri $SmartCTLExporterURL -OutFile $SmartCTLExporterFile -UseBasicParsing; 
            }catch
            {
                try 
                {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    WebClient.DownloadFile($SmartCTLExporterURL, $SmartCTLExporterFile);
                }catch 
                {
                    Write-Error ($_.Exception.Message);
                    if ($wasreadonly)
                    {
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    [System.Threading.Thread]::Sleep(5000);
                    Exit;
                }
            }
        }
        Expand-Archive ($SmartCTLExporterFile) -DestinationPath ($INSTPATH); ##Install (or rather unzip)
    }
    $scheduleObject = New-Object -ComObject schedule.service; #Add service
    $scheduleObject.connect();
    $rootFolder = $scheduleObject.GetFolder('\');
    try 
    {
        $rootFolder.CreateFolder("Prometheus");
    }catch
    {
        write-host ($_.Exception.Message);
    }
    $action = New-ScheduledTaskAction -Execute $('"' + $SmartCTLExporterExec + '"') -Argument $("--smartctl.path='C:\Program Files\smartmontools\bin\smartctl.exe'");
    $trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -minutes 3);
    $settings = $(New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0);

    try
    {
        Register-ScheduledTask -TaskName "SmartCtlExporter" -TaskPath '\Prometheus\' -RunLevel Highest -User "System" -Action $action -Settings $settings -Trigger $trigger;
    }catch
    {
        Set-ScheduledTask -TaskPath '\Prometheus' -TaskName "SmartCtlExporter" -Settings $settings -Trigger $trigger -Action $action;
    }
    try
    {
        Start-ScheduledTask -TaskPath '\Prometheus' -TaskName "SmartCtlExporter";
    }catch
    {
        Write-Error ($_.Exception.Message);
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($scheduleObject)
}


if ($wasreadonly)
{
    $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
}
  
