##Since PRTG Network Monitor only able to show ONE field from ONE line at the same time, 
##and since it also doesn't have anything to automatically create WQL scripts, I present you with this rofloscript 
Clear-Host
$islocal = $true
##No named parameters are available in 5.1 and I CBA parsing them
if ($args.Length -eq 0)
{
    $TGT_COMPUTERNAME = "localhost"
}
else 
{
    $TGT_COMPUTERNAME = $args[0];
    echo ("Target computer = " + $TGT_COMPUTERNAME + "(taken from a CLI launch Argument)");
}



#CHANGEME
$dir = "C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\WMI WQL scripts\" + $TGT_COMPUTERNAME; ##PRTG Doesn't search in subfolders for WQLs
#CHANGEME


function FileChecker($HWTYPEFold, $Collumn)
{
    $tmppath = $dir + "\" + $HWTYPEFold;
    if(!(test-path -PathType container $tmppath))
    {
      if ($tmppath.Length -lt 259)
      {
        New-Item -ItemType Directory -Path $tmppath | Out-Null;
      }
      else 
      {
        echo "Directory path is too long"
      } 
    }

    for ($i=0; $i -lt $Collumn.length; $i++)
    {
       $tmppath = $dir + "\" + $HWTYPEFold + "\" + $Collumn[$i];
       if(!(test-path -PathType container $tmppath))
       {
             if ($tmppath.Length -lt 259)
             {
                  New-Item -ItemType Directory -Path $tmppath | Out-Null;
             }
             else 
             {
                  echo "Directory path is too long"
             } 
        }
    }

}

function SVAL ( $inp, $ident )
{
    
    [string]$retval = "Select " + $inp + " from Sensor WHERE Identifier='" + $ident + "'";
    return $retval;
} 

function FileWrite ($dirls, $type, $idls, $mode)
{
    if ($idls.length -gt 0)
    {
        $qry = (SVAL $mode $idls)
        $tmpidls = ($idls.subString(1));
        $path = $dirls + "\" + $type + "\" + "$mode" + "-" + $tmpidls.replace('/','_') + ".wql";
        echo ("Processing: " + $path);
        if (!($($dirls+$type).length -lt 257))
        {
            if ($path.length -gt 259)
            {
                $tempstr = $path.subString(0,259);
                $path = $tempstr;
                echo "[Warning]: Filepath trimmed to "+ $path + " due to its length."
            }
        }
        try 
        {
            [System.IO.Stream]$FileStream  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($path), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read))
        }
        catch 
        {
            echo ("Failed to create filestream! The path was:" + $path);
        }
        while ($FileStream.CanWrite -eq $false) 
        {
            try 
            {
               $FileStream.Dispose();
               $FileStream  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($path), [System.IO.FileMode]::Append, [System.IO.FileAccess]::Create, [IO.FileShare]::Read));
            }
            catch 
            {
                echo "File is busy";
                [System.Threading.Thread]::Sleep(2000);
                $FileStream.Dispose();
            }
        }
        $sw = New-Object System.IO.StreamWriter([System.IO.Stream]$FileStream, [Text.Encoding]::UTF8);

        $sw.WriteLine($qry);
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
            $FileStream.Dispose();
        }
        catch 
        {
            echo "Unable to dispose filestream object";
        }
    }
}

function GetSensorInfo ($HWArray) 
{
    $retval = @();
    for ($i=0; $i -lt $HWArray.Identifier.Length; $i++)
    {
        $HWinstid = $HWArray[$i].Identifier;
        if ($islocal)
        {
            $retval += $(Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Sensor | Where-Object {($_.Parent -eq $HWinstid)}  | Select-Object -Property Name,Identifier,Parent,Value,Min,Max);
        }
        else 
        {
            $retval += $(Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Sensor  -ComputerName $TGT_COMPUTERNAME -Credential $CREDS  | Where-Object {($_.Parent -eq $HWinstid)} | Select-Object -Property Name,Identifier,Parent,Value,Min,Max);
        }   
    }
    return $retval;
}

function WriterHW ($HWType, $HWSensorContainer, $Collumn) 
{
    for ($i=0; $i -lt $HWSensorContainer.Length; $i++)
    {
       for ($k=0; $k -lt $Collumn.length; $k++)
       {
         FileWrite $dir $HWType $HWSensorContainer[$i].Identifier $Collumn[$k];
       }
    }
}

$MB = @();
$CPU = @();
$HDD = @();
$RAM = @();
$SIO = @();
if (($TGT_COMPUTERNAME -ilike "localhost") -or ($TGT_COMPUTERNAME -match '^127\.'))
{
    [System.Array]$CPU += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "CPU";
    [System.Array]$MB  += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "Mainboard";
    [System.Array]$HDD += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "HDD";
    [System.Array]$RAM += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "RAM";
    [System.Array]$SIO += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "SuperIO";
}
else 
{
    $islocal=$false;
    $Failed = $false;
    $CREDS = Get-Credential;
    do
    {
        $Failed = $false
        try 
        {
            [System.Array]$CPU += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "CPU";
        }
        catch [UnauthorizedAccessException] 
        {
            echo "Access denied.";
            $Failed = $true;
            $CREDS = Get-Credential;
        }
    } while ($Failed)
    [System.Array]$MB  += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "Mainboard";
    [System.Array]$HDD += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "HDD";
    [System.Array]$RAM += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "RAM";
    [System.Array]$SIO += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "SuperIO";
}


If(!(test-path -PathType container $dir))
{
      if ($dir.Length -lt 259)
      {
        New-Item -ItemType Directory -Path $dir | Out-Null
      }
      else 
      {
        echo "Directory path is too long";
      } 
}


###CPU
if ($CPU.Length -gt 0)
{
    echo ("CPUs: " + $CPU.Length);
    $CPUSens = GetSensorInfo $CPU;
    FileChecker "CPU" @("Value","Min","Max");
    WriterHW "CPU" $CPUSens @("Value","Min","Max");
}


###HDD
if ($HDD.Length -gt 0)
{
    echo ("HDDs: " + $HDD.Length)
    $HDDSens =  GetSensorInfo $HDD;
    FileChecker "HDD" @("Value","Min","Max");
    WriterHW "HDD" $HDDSens @("Value","Min","Max");
}


###RAM
if ($RAM.Length -gt 0)
{
    echo ("RAMs: " + $RAM.Length);
    $RAMSens =  GetSensorInfo $RAM;
    FileChecker "RAM" @("Value","Min","Max");
    WriterHW "RAM" $RAMSens @("Value","Min","Max");
}


###SuperIO
if ($SIO.Length -gt 0)
{
    echo ("SIOs: " + $SIO.Length);
    $SIOSens = GetSensorInfo $SIO;
    FileChecker "SuperIO" @("Value","Min","Max");
    WriterHW "SuperIO" $SIOSens @("Value","Min","Max");
}
