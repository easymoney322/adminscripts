Clear-Host
$islocal = $true
##No named parameters are available in 5.1 and I CBA parsing them
if ($args.Length -eq 0)
{
    $TGT_COMPUTERNAME = "localhost";
}
else 
{
    $TGT_COMPUTERNAME = $args[0];
    echo ("Target computer = " + $TGT_COMPUTERNAME + "(taken from a CLI launch Argument)");
}



#CHANGEME
$dir1 = $env:USERPROFILE + "\Desktop\Hosts\";
$dir2 = $dir1 + $TGT_COMPUTERNAME;
#CHANGEME

function CheckDir ($inputdir)
{
    if(!(test-path -PathType container $inputdir))
    {
        if ($inputdir.Length -lt 259)
        {
          New-Item -ItemType Directory -Path $inputdir | Out-Null;
        }
        else 
        {
            echo "Directory path " $inputdir " is too long"
            exit 2147549493;
        }
    }
}

function GetCredentialsMandatory ([ref]$_lsCREDS)
{
    try 
    {
       $_lsCREDS.Value = Get-Credential;
    }
    catch  
    {
        Write-Output "Ran into an issue inside of a function: $PSItem";
        $_lsCREDS.Value = $null;
        echo "Credentials are not full or missing!";
     
        Exit 401;
    }
}

CheckDir $dir1 ;
CheckDir $dir2 ;

$List = @();
$MB = @();
$CPU = @();
$HDD = @();
$RAM = @();
$SIO = @();

if (($TGT_COMPUTERNAME -ilike "localhost") -or ($TGT_COMPUTERNAME -match '^127\.'))
{
    [System.Array]$CPU += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "CPU"      |  Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$MB  += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "Mainboard"|  Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$HDD += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "HDD"      |  Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$RAM += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "RAM"      |  Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$SIO += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware | Where-Object -Property "HardwareType" -eq "SuperIO"  |  Select-Object -Property HardwareType,Name,Identifier;
}
else 
{
    $islocal=$false;
    $Failed = $false;
    [System.Management.Automation.PSCredential]$CREDS=$NULL;
    GetCredentialsMandatory ([ref] $CREDS)
    do
    {
        $Failed = $false
        try 
        {
            [System.Array]$CPU += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "CPU" | Select-Object -Property HardwareType,Name,Identifier;
        }
        catch [UnauthorizedAccessException] 
        {
            echo "Access denied.";
            $Failed = $true;
            GetCredentialsMandatory ([ref] $CREDS)
        }
    } while (($Failed) -and ($CREDS -ne $NULL))

    if ($CREDS -eq $null)
    {
        Write-Output "Ran into an issue: $PSItem"
        echo "Credentials are not full or missing!"
        Exit 401;
    }

    [System.Array]$MB  += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "Mainboard" | Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$HDD += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "HDD"       | Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$RAM += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "RAM"       | Select-Object -Property HardwareType,Name,Identifier;
    [System.Array]$SIO += Get-WmiObject -Namespace "root\OpenHardwareMonitor" -Class Hardware -ComputerName $TGT_COMPUTERNAME -Credential $CREDS | Where-Object -Property "HardwareType" -eq "SuperIO"   | Select-Object -Property HardwareType,Name,Identifier;
}

$List = $CPU + $MB + $SIO + $HDD + $RAM;
$path = ($dir2 + "\OHM-HWInfo.txt");
try
{
    [System.IO.Stream]$FileStream  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($path), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
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
            $FileStream  = New-Object System.IO.FileStream (([System.IO.Path]::Combine($path), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [IO.FileShare]::Read));
        }
    catch 
        {
            echo "File is busy";
            [System.Threading.Thread]::Sleep(2000);
            $FileStream.Dispose();
        }
}
$sw = New-Object System.IO.StreamWriter([System.IO.Stream]$FileStream, [Text.Encoding]::UTF8);
$sw.AutoFlush = $true;

$sw.WriteLine((Out-String -InputObject $List));
$sw.Flush();
echo "File '$path' written!";
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

explorer $dir1;
