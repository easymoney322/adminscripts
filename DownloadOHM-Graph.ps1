try {
    $OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8");
}catch{
    Write-Debug ("Errors occured while setting an encoding");
}
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
$PSDefaultParameterValues['Invoke-WebRequest:MaximumRetryCount'] = 5;
$INSTPATH = 'C:\Program Files\';
$graphiteVersion = '0.30.0';
$ExporterVersion = "0.24.0";
$SmartMonVersion = "7.4-1";
$SmartMonReleaseDir = "7_4";
$SmartCTLExporterVersion = "0.11.0";
$PromtailVersion = "2.9.0";
$PromtailListenPort = "9080";
$loki = "http://192.168.100.1:3100";
$lokiapi = $loki + '/loki/api/v1/push';
$IANATimezoneDatabasestring = 'Asia/Yekaterinburg';
$downloadDir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path;
$dlfold = Get-Item -Path $downloadDir;
$schedulerFolderExists = $false;
$arch=$null;
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem){
        $arch="amd64";
    }else {
        $arch="386";
    }
function ServiceCheck([string] $ServiceName){
    $serviceQuery = "Select * from Win32_Service WHERE name like  '%" + $ServiceName + "%'";
    $SrvcList = (Get-CimInstance -Query $serviceQuery);
    if ($SrvcList -ne $null){
        Write-Host ("Found service " + $ServiceName);
        [System.Array]$ServicePath = @();
        $ServicePathString= $SrvcList.PathName -split " --";
        if ($ServicePathString -ne $null){
            for ($i=0; $i -lt $ServicePathString.Length; $i++){
                if (($ServicePathString[$i][2] -eq ':') -or ($ServicePathString[$i][1] -eq ':')){
                    [System.Array]$indxList=@();
                    for ($indx =0;$indx -lt $ServicePathString[$i].Length; $indx++) {
                        if ($ServicePathString[$i][$indx] -eq '"'){
                            $indxList += $indx;
                        }
                    }
                    if ($indxList.Length -gt 1){
                        Write-Debug ("LVAL=" + $lval);
                        [int]$lval = $indxList[0];
                        [int]$rval = $indxList[1];
                        $bubble = [string]($ServicePathString[$i]).Substring($lval+1, $rval-$lval-1);
                        [string]$ServicePathString[$i] = $bubble
                    }
                    $ServicePath += $ServicePathString[$i];
                }
            }
            if ($ServicePath -ne $null) {
                Write-Host ("Succesfully extracted executable path from " + $ServiceName + " service");
                for ($k=0; $k -lt $ServicePath.Length; $k++){
                    if ($null -ne $ServicePath[$k]){
                        Write-Host ($ServicePath[$k]);
                        if (Test-Path -Path $($ServicePath[$k]) -PathType Leaf){
                            Write-Host ("Found " + $ServiceName + " executable at " + $ServicePath[$k]);
                            return $true;
                        }
                    }
                }
                Write-Host ("Exporter executables were not found"); 
                return $false;
            }
        }else {
            Write-Host ("Service path is empty");
            return $false;
        }
    }
    $serviceQuery = "Select * from Win32_Service WHERE DisplayName like  '%" + $ServiceName + "%'"; ##If service wasnt found, do the same but for DisplayName
    $SrvcList = (Get-CimInstance -Query $serviceQuery);
    if ($SrvcList -ne $null) {
        Write-Host ("Found service " + $ServiceName);
        [System.Array]$ServicePath = @();
        $ServicePathString= $SrvcList.PathName -split " --";
        if ($ServicePathString -ne $null)  {
            for ($i=0; $i -lt $ServicePathString.Length; $i++) {
                if (($ServicePathString[$i][2] -eq ':') -or ($ServicePathString[$i][1] -eq ':')) {
                    [System.Array]$indxList=@();
                    for ($indx =0;$indx -lt $ServicePathString[$i].Length; $indx++) {
                        if ($ServicePathString[$i][$indx] -eq '"') {
                            $indxList += $indx;
                        }
                    }
                    if ($indxList.Length -gt 1)  {
                        write-host "LVAL=" $lval;
                        [int]$lval = $indxList[0];
                        [int]$rval = $indxList[1];
                        $bubble = [string]($ServicePathString[$i]).Substring($lval+1, $rval-$lval-1);
                        [string]$ServicePathString[$i] = $bubble
                    }
                    $ServicePath += $ServicePathString[$i];
                }
            }
            if ($ServicePath -ne $null) {
                Write-Host "Succesfully extracted executable path from " + $ServiceName + " service";
                for ($k=0; $k -lt $ServicePath.Length; $k++) {
                    write-host $ServicePath[$k]
                    if (Test-Path -Path $($ServicePath[$k]) -PathType Leaf) {
                        Write-Host ("Found " + $ServiceName + " executable at " + $ServicePath[$k]);
                        return $true;
                    }
                }
                Write-Host "Exporter executables were not found"; 
                return $false;
            }
        }else{
            Write-Host "Service path is empty";
            return $false;
        }
    }
    Write-Host ($ServiceName + " service wasn't found");
    return $false;
}
function UnquoteAString([string] $inputString){
    $indxList=@();
    for([int]$i=0; $i -lt $inputString.Length; $i++){
       if (($inputString[0] -or $inputString[1]) -eq ("'" -or '"' )){
          $indxList+=New-Object PSObject -Property @{Character=$inputString[$i]; Position=$i;}
       }
    }
    if (0 -eq $indxList.Length){
        return $inputString;
    }else{
          $qtsmbl = $indxList[0].Character;
          $rpos = 0;
          for ([int]$i=1; $i -lt $indxList.Length; $i++){
              if ($qtsmbl -eq $indxList[$i].Character){
                  $rpos = $indxList[$i].Position;
                  Write-Debug ("String is being terminated at index " +$rpos);
                  break;
              }
          }
          if ($rpos -eq 0){
              Write-Error ("Unable to unquote the string for " +$inputString);
              [System.Threading.Thread]::Sleep(5000);
              Exit; 
            }else{
                [string]$retval = $inputString.Substring(1,($rpos-1));
                return $retval;
               }
         }
}
function SchedulerCheck([string] $TaskDir, $TaskName ){
    Write-Host ('Looking for ' + $TaskName + ' scheduler task');
    $schedulerInterface = New-Object -ComObject schedule.service; 
    $searchFolder = '\'+ $TaskDir;
    Write-Host ('Expected folder is ' + $searchFolder);
    $schedulerInterface.connect();
    try{
        $schedulerRoot = $schedulerInterface.GetFolder('\');
    } catch{
        Write-Warning ("Scheduler interface isn't working:" +  ($_.Exception.Message));
        Exit;
    }
    try{
    $schedulerRoot = $schedulerInterface.GetFolder($searchFolder);
    $schedulerFolderExists=$true;
    } catch{
        Write-Host ('Scheduler folder wasnt found:' + ($_.Exception.Message) );
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedulerInterface) | Out-Null;
        }catch{
            Write-Warning ("Unable to release Scheduler COM object:" + ($_.Exception.Message));
        }
            return $false;  
        }
    Write-Debug ("Found the folder");
    try {
        $result = Get-ScheduledTask -TaskPath ('\' + $TaskDir +'\') -TaskName $TaskName -ErrorAction "Stop";   
    } catch{
        Write-Host ('Unable to find schedulers task: ' +  ($_.Exception.Message));
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedulerInterface) | Out-Null;
        }catch{
            Write-Warning ("Unable to release Scheduler COM object:" + ($_.Exception.Message));
        }
        return $false;
    }
    Write-Debug ('Found scheduler job ' + '\' + $TaskDir + '\' + $TaskName);
    try {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedulerInterface) | Out-Null;
    } catch{
        Write-Warning ("Unable to release Scheduler COM object:" + ($_.Exception.Message));
        }
        $execPath = $result.Actions.Execute;
        if ($null -ne $execPath) {
        $execPathTrimmed = (UnquoteAString $execPath);
        if (Test-Path -Path $execPathTrimmed -PathType Leaf){
            return $true;
            }
        }else{
            Write-Host ('Scheduler task was found, but an executable - wasnt');
        }
return $false;
}
$wasreadonly = $false;
if ($dlfold.Attributes -imatch '.*ReadOnly.*') {
    $wasreadonly = $true;
    $dlfold.Attributes = $dlfold.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly;
} 
if ($(ServiceCheck "OhmGraphite") -eq $false) {
    if ($false -eq (Test-Path -Path $($INSTPATH+'OHMGraphite\OhmGraphite.exe') -PathType Leaf)){   
        Write-Host "Graphite installation was not found. Starting installation process.";
        $fileGraphite = $downloadDir+"/GrOHM.zip";
        if ($false -eq (Test-Path -Path $fileGraphite -PathType Leaf)){
            $dlUrl = 'https://github.com/nickbabcock/OhmGraphite/releases/download/v' + $graphiteVersion + '/OhmGraphite-' + $graphiteVersion + '.zip';
            try{
                Write-Debug ($dlUrl);
                Invoke-WebRequest -Uri $dlUrl -OutFile $fileGraphite -UseBasicParsing; 
            }catch {
                try{
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    $WebClient.DownloadFile($dlUrl, $fileGraphite);
                }catch {
                    if ($wasreadonly){
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    Write-Error ($_.Exception.Message);
                    [System.Threading.Thread]::Sleep(5000);
                    Exit; 
                }
            }
        }else{
            Write-Host ("Found Graphite archive. Proceeding to unzip.");
        }
        if($false -eq (Test-Path ($INSTPATH+'OHMGraphite\'))) {
            New-Item -Path ($INSTPATH+'OHMGraphite\') -ItemType Directory;
        }
        try{
        Expand-Archive ($fileGraphite)  -DestinationPath ($INSTPATH+'OHMGraphite\')  -Force -ErrorAction Stop;
        }catch{
        Write-Error ("Unable to unzip " +$fileGraphite + ": " +$_.Exception.Message);
        }
        $configFilePath = $INSTPATH+'OHMGraphite\OhmGraphite.exe.config';
        [System.IO.Stream]$FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($configFilePath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
        while ($FileStreamPCinfo.CanWrite -eq $false){
            try{
               $FileStreamPCinfo.Dispose();
               $FileStreamPCinfo  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($configFilePath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
            }catch {
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
        try{
            $sw.Dispose();
        }catch {
            Write-Error ("Unable to dispose StreamWriter object" + ($_.Exception.Message));
        }
        try {
            $FileStreamPCinfo.Dispose();
        }catch {
            Write-Error ("Unable to dispose filestream object:" + ($_.Exception.Message));
        }
        Start-Process -FilePath $($INSTPATH+'OHMGraphite\OhmGraphite.exe') -ArgumentList "install" -WorkingDirectory $($INSTPATH+'OHMGraphite\');
        Write-Host ("Successfully installed Graphite");
        [System.Threading.Thread]::Sleep(2000);
        try {
            Start-Service OhmGraphite;
        }catch {
            Write-Error ("Unable to start Graphite service:" + ($_.Exception.Message));
        }
    }else{
        Write-Host ("Found Graphite Installation at " + $($INSTPATH+'OHMGraphite\'));
    }
}
if ($(ServiceCheck "windows_exporter") -eq $false) {
    if ($false -eq (Test-Path -Path $($INSTPATH+'windows_exporter\windows_exporter.exe') -PathType Leaf)){
        Write-Host "Wasn't able to find Windows Exporter installation path. Starting the installation";
        $exporterPrefixLink = "https://github.com/prometheus-community/windows_exporter/releases/download/v" + $ExporterVersion+"/windows_exporter-"+$ExporterVersion + "-" +$arch+".msi";
        $filePrometheusExporterInstaller = $downloadDir+'\ExporterInstaller.msi';
        if ($false -eq (Test-Path -Path $filePrometheusExporterInstaller -PathType Leaf)){
            try{
                Write-Debug ($exporterPrefixLink);
                Invoke-WebRequest -Uri $ExporterPrefixLink -OutFile $filePrometheusExporterInstaller -UseBasicParsing;
            }catch{
                try{
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    $WebClient.DownloadFile($ExporterPrefixLink, $filePrometheusExporterInstaller);
                }catch{
                    if ($wasreadonly){
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    Write-Error ($_.Exception.Message);
                    [System.Threading.Thread]::Sleep(5000);
                }   Exit; 
            }
        }
        $argument = "/i $('"' + $filePrometheusExporterInstaller + '"') /qn /norestart /log $('"' + $downloadDir+'WindowsExporterInstaller.log' + '"') INSTALLDIR=`"C:\Program Files\windows_exporter\`"";
        (Start-Process msiexec.exe -argumentlist $argument).ExitCode;
        [System.Threading.Thread]::Sleep(5000);
        try{
            Start-Service "windows_exporter"
        }catch{
            Write-Error ("Unable to start windows_exporter service:" + ($_.Exception.Message));
        }
    }
}
if ($false -eq (Test-Path -Path $($INSTPATH+"SmartMonTools/bin/smartctl.exe") -PathType Leaf)){
    $fileSmartMonInstaller = $downloadDir+ '\smartmontools-' + $SmartMonVersion +".win32-setup.exe";
    if ($false -eq (Test-Path -Path $fileSmartMonInstaller -PathType Leaf)) {
        $smartMonPrefixLink = "https://github.com/smartmontools/smartmontools/releases/download/RELEASE_" + $SmartMonReleaseDir+ "/smartmontools-" +$SmartMonVersion + ".win32-setup.exe";
        try {
                Invoke-WebRequest -Uri $smartMonPrefixLink -OutFile $fileSmartMonInstaller -UseBasicParsing;
            }catch{
                try {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    $WebClient.DownloadFile($smartMonPrefixLink, $fileSmartMonInstaller);
                }catch{
                    if ($wasreadonly){
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    Write-Error ($_.Exception.Message);
                    [System.Threading.Thread]::Sleep(5000);
                    Exit;
                }
            }
    }else{
        Write-Host("Installer for SmartMonTools found, proceeding to install...");
    }
    $smartMonDir = $INSTPATH+"SmartMonTools1";
    try {
            Start-Process  -NoNewWindow -FilePath $fileSmartMonInstaller -ArgumentList "/S /D=`"$smartMonDir`" ";
    }catch{
        if ($wasreadonly) {
            $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
        }
        Write-Error ($_.Exception.Message);
        [System.Threading.Thread]::Sleep(5000);
        Exit;
    }
}
if ($false -eq (SchedulerCheck "Prometheus" "SmartCtlExporter" )){
    $SmartCTLexporterSufffixLink=(".windows-" + $arch);
    $SmartCTLExporterExec = $INSTPATH +"smartctl_exporter-"+$SmartCTLExporterVersion + $SmartCTLexporterSufffixLink + "/smartctl_exporter.exe";
    if ($false -eq (Test-Path -Path $SmartCTLExporterExec -PathType Leaf)){
        Write-Host ("SmartCTLExporter not found"); ##Installation not found
        $SmartCTLexporterSufffixLink +=".zip"; ##Search for an archive
        $SmartCTLExporterFile = $downloadDir + '/smartctl_exporter-' + $SmartCTLExporterVersion + $SmartCTLexporterSufffixLink ;
        if ($false -eq (Test-Path -Path $($INSTPATH+"smartctl_exporter-"+$SmartCTLExporterVersion + $SmartCTLexporterSufffixLink) -PathType Leaf)){ 
            Write-Host ("Downloading SmartCtlExporter");
            $SmartCTLExporterURL = 'https://github.com/prometheus-community/smartctl_exporter/releases/download/v' + $SmartCTLExporterVersion +"/smartctl_exporter-" + $SmartCTLExporterVersion + $SmartCTLexporterSufffixLink;
            Write-Debug ($SmartCTLExporterURL);
            try{
                  Invoke-WebRequest -Uri $SmartCTLExporterURL -OutFile $SmartCTLExporterFile -UseBasicParsing; 
            }catch{
                try {
                    Write-Host ($_.Exception.Message);
                    $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                    $WebClient.DownloadFile($SmartCTLExporterURL, $SmartCTLExporterFile);
                }catch {
                    if ($wasreadonly){
                        $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                    }
                    Write-Error ($_.Exception.Message);
                    [System.Threading.Thread]::Sleep(5000);
                    Exit;
                }
            }
        }
        try{
        Expand-Archive ($SmartCTLExporterFile) -DestinationPath ($INSTPATH)  -Force -ErrorAction Stop; ##Install (or rather unzip)
                }catch{
        Write-Error ("Unable to unzip " + $SmartCTLExporterFile + ": " +$_.Exception.Message);
        }
    }
    $scheduleObject = New-Object -ComObject schedule.service; #Add service
    $scheduleObject.connect();
    $rootFolder = $scheduleObject.GetFolder('\');
    if ($false -eq $schedulerFolderExists){ 
    try {
        $rootFolder.CreateFolder("Prometheus");
    }catch{
        Write-Warning ("An error occured when tried to create task scheduler folder: " + $_.Exception.Message);}
    }
    $action = New-ScheduledTaskAction -Execute $('"' + $SmartCTLExporterExec + '"') -Argument $('--smartctl.path="C:\Program Files\smartmontools\bin\smartctl.exe"');
    $trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -minutes 3);
    $settings = $(New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0);
    try{
        Register-ScheduledTask -TaskName "SmartCtlExporter" -TaskPath '\Prometheus\' -RunLevel Highest -User "System" -Action $action -Settings $settings -Trigger $trigger  -ErrorAction Stop;
    }catch{
        Set-ScheduledTask -TaskPath '\Prometheus\' -TaskName "SmartCtlExporter" -Settings $settings -Trigger $trigger -Action $action; ##Assume task is already registered
    }
    try{
        Start-ScheduledTask -TaskPath '\Prometheus\' -TaskName "SmartCtlExporter";
    }catch
    {
        Write-Error ($_.Exception.Message);
    }finally{
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($scheduleObject) | Out-Null;
    }
}else{
    Write-Host ("Found SmartCtlExporter scheduler task");
}
if ($false -eq (SchedulerCheck "Prometheus" "Promtail" )){
            $PromtailDir = $INSTPATH + "Promtail";
            $PromtailExecPath = $PromtailDir + '/promtail-windows-amd64.exe';
            if ($true -eq (Test-Path -Path $($INSTPATH + "Promtail") -PathType Container)) {
                Write-Debug ("Promtail directory exists");
            }else {
                New-Item -Path $INSTPATH -Name "Promtail" -ItemType "directory";
                  }
            if ($true -eq (Test-Path -Path $($PromtailDir + '\var') -PathType Container)) {
                Write-Debug ("Promtail var directory already exists");
            }else {
                New-Item -Path $PromtailDir -Name "var" -ItemType "directory";
                  }
            if ($true -eq (Test-Path -Path $($PromtailDir + '\log') -PathType Container)) {
                Write-Debug ("Promtail log directory already exists");
            }else {
                New-Item -Path $PromtailDir -Name "log" -ItemType "directory";
                  }
            if ($false -eq (Test-Path -Path $PromtailExecPath -PathType Leaf)){
                    $PromtailArchivePath = $downloadDir + '/promtail-v' + $PromtailVersion + ".zip";
                    if ($true -eq (Test-Path -Path $PromtailExecPath -PathType Leaf)){
                        Write-Debug ("Found promtail archive");
                    }else{
                        try{
                              $PromtailURL = "https://github.com/grafana/loki/releases/download/v" + $PromtailVersion + "/promtail-windows-" + $arch +".exe.zip";
                              Write-Debug ($PromtailURL);
                              Invoke-WebRequest -Uri $PromtailURL -OutFile $PromtailArchivePath -UseBasicParsing; 
                            }catch{
                                try {
                                        Write-Host ($_.Exception.Message);
                                        $WebClient = New-Object System.Net.WebClient; ##Assume that Invoke-WebRequest is not supported 
                                        $WebClient.DownloadFile($PromtailURL, $PromtailArchivePath);
                                    }catch {
                                        if ($wasreadonly){
                                                             $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
                                                         }
                                        Write-Error ($_.Exception.Message);
                                        [System.Threading.Thread]::Sleep(5000);
                                        Exit;
                                           }
                                  }
                         }
                   Write-Host ("Started promtail installation");
                   try{
                   Expand-Archive ($PromtailArchivePath) -DestinationPath ($PromtailDir) -Force -ErrorAction Stop; ##Install (or rather unzip)
                   }catch {
                      Write-Error ("Unable to expand promtail archive to the " + $PromtailDir + " directory:" + $_.Exception.Message);
                   }
                    $PromtailConfigPath=$PromtailDir+'\promtail-local-config.yaml';
                    [System.IO.Stream]$PromtailFS  = New-Object System.IO.FileStream (([System.IO.Path]::Combine( $PromtailConfigPath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
                    while ($PromtailFS.CanWrite -eq $false){
                        try{
                           $PromtailFS.Dispose();
                           $PromtailFS  = New-Object System.IO.FileStream (([System.IO.Path]::Combine( $PromtailConfigPath ), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
                        }catch {
                            Write-Debug ("File is busy");
                            [System.Threading.Thread]::Sleep(5000);
                            $PromtailFS.Dispose();
                        }
                    }
                    [string]$PromtailLogs = ($PromtailDir+'\log\*');
                    $PromtailBookmarkApplication = $PromtailDir +'\var\bookmark-application.xml'
                    $PromtailBookmarkSystem= $PromtailDir +'\var\bookmark-system.xml'
                    $Promtailsw = New-Object System.IO.StreamWriter([System.IO.Stream]$PromtailFS, [Text.Encoding]::UTF8);
                    $Promtailsw.AutoFlush = $true;
                    $Promtailsw.WriteLine('server:');
                    $Promtailsw.WriteLine('  http_listen_port: '+ $PromtailListenPort);
                    $Promtailsw.WriteLine('  grpc_listen_port: 0');
                    $Promtailsw.WriteLine('positions:');
                    $PromtailPositionsFile=$PromtailDir+'\var\positions.yaml';
                    $Promtailsw.WriteLine('  filename: ' + "'" + $PromtailPositionsFile + "'");
                    $Promtailsw.WriteLine('clients:');
                    $Promtailsw.WriteLine('  - url: ' + $lokiapi);
                    $Promtailsw.WriteLine('scrape_configs:');
                    $Promtailsw.WriteLine('- job_name: scrape');
                    $Promtailsw.WriteLine('  pipeline_stages:');
                    $Promtailsw.WriteLine('  - output:');
                    $Promtailsw.WriteLine('      source: message');
                    $Promtailsw.WriteLine('      action_on_failure: skip');
                    $Promtailsw.WriteLine('  static_configs:');
                    $Promtailsw.WriteLine('  - targets:');
                    $Promtailsw.WriteLine('      - localhost');
                    $Promtailsw.WriteLine('    labels:');
                    $Promtailsw.WriteLine('      job: filelogs');
                    $Promtailsw.WriteLine('      agent: promtail');
                    $Promtailsw.WriteLine('      __path__: ' + "'" + $PromtailLogs +"'");
                    $Promtailsw.WriteLine('- job_name: windows-application');
                    $Promtailsw.WriteLine('  pipeline_stages:');
                    $Promtailsw.WriteLine('  - timestamp:');
                    $Promtailsw.WriteLine('      location: ' + $IANATimezoneDatabasestring);
                    $Promtailsw.WriteLine('      source: timestamp');
                    $Promtailsw.WriteLine('      format: RFC1123');
                    $Promtailsw.WriteLine('      fallback_formats: [RFC3339Nano]');
                    $Promtailsw.WriteLine('  windows_events:');
                    $Promtailsw.WriteLine('    use_incoming_timestamp: true');
                    $Promtailsw.WriteLine('    exclude_event_data: true');
                    $Promtailsw.WriteLine('    exclude_user_data: true');
                    $Promtailsw.WriteLine('    bookmark_path: ''' + $PromtailBookmarkApplication + "'");
                    $Promtailsw.WriteLine('    eventlog_name: "Application"');
                    $Promtailsw.WriteLine('    xpath_query: ''*''');
                    $Promtailsw.WriteLine('    poll_interval: 1m');
                    $Promtailsw.WriteLine('    locale: 1033');
                    $Promtailsw.WriteLine('    labels:');
                    $Promtailsw.WriteLine('      job: windows');
                    $Promtailsw.WriteLine('      logsource: windows-eventlog');
                    $Promtailsw.WriteLine('- job_name: windows-system');
                    $Promtailsw.WriteLine('  pipeline_stages:');
                    $Promtailsw.WriteLine('  - timestamp:');
                    $Promtailsw.WriteLine('      location: ' + $IANATimezoneDatabasestring);
                    $Promtailsw.WriteLine('      source: timestamp');
                    $Promtailsw.WriteLine('      format: RFC1123');
                    $Promtailsw.WriteLine('      fallback_formats: [RFC3339Nano]');
                    $Promtailsw.WriteLine('  windows_events:');
                    $Promtailsw.WriteLine('    use_incoming_timestamp: true');
                    $Promtailsw.WriteLine('    exclude_event_data: true');
                    $Promtailsw.WriteLine('    exclude_user_data: true');
                    $Promtailsw.WriteLine('    bookmark_path: ' + "'" + $PromtailBookmarkSystem + "'");
                    $Promtailsw.WriteLine('    eventlog_name: "System"');
                    $Promtailsw.WriteLine('    xpath_query: ''*''');
                    $Promtailsw.WriteLine('    poll_interval: 1m');
                    $Promtailsw.WriteLine('    locale: 1033');
                    $Promtailsw.WriteLine('    labels:');
                    $Promtailsw.WriteLine('      job: windows');
                    $Promtailsw.WriteLine('      logsource: windows-eventlog');
                    try{
                        $Promtailsw.Dispose();
                    }catch {
                        Write-Error ("Unable to dispose StreamWriter object" + ($_.Exception.Message));
                    }
                    try {
                        $PromtailFS.Dispose();
                    }catch {
                        Write-Error ("Unable to dispose filestream object:" + ($_.Exception.Message));
                    }
                    $scheduleObject = New-Object -ComObject schedule.service; #Add service
                    $scheduleObject.connect();
                    $rootFolder = $scheduleObject.GetFolder('\');
                    if ($false -eq $schedulerFolderExists){ 
                       try {
																$rootFolder.CreateFolder("Prometheus");
													 }catch{
																		Write-Warning ("An error occured when tried to create task scheduler folder: " + $_.Exception.Message); 
																 }
                            $action = New-ScheduledTaskAction -Execute $('"' + $PromtailExecPath + '"') -Argument $('--config.file="'+ $PromtailDir +'\promtail-local-config.yaml" --config.expand-env=true');
                            $trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -minutes 3);
                            $settings = $(New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0);
                            try{
                                Register-ScheduledTask -TaskName "Promtail" -TaskPath '\Prometheus\' -RunLevel Highest -User "System" -Action $action -Settings $settings -Trigger $trigger -ErrorAction Stop;
                            }catch{
                                Set-ScheduledTask -TaskPath '\Prometheus\' -TaskName "Promtail" -Settings $settings -Trigger $trigger -Action $action; ##Assume task is already registered
                             }
                            try{
                                Start-ScheduledTask -TaskPath '\Prometheus\' -TaskName "Promtail";
                            }catch{
                                Write-Error ($_.Exception.Message);
                             }finally{
                                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($scheduleObject) | Out-Null;
                                     }
                                                        }
                                                                             }
}
if ($wasreadonly){
    $dlfold.Attributes = $dlfold.Attributes -bor[System.IO.FileAttributes]::ReadOnly;
}
  
