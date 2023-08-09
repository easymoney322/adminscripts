clear
$dirpath = ($env:HOMEPATH + "\Desktop\");
$Filename_RU = "SrvcsRU.csv";
$Filename_EN = "SrvcsEN.csv";

$CSV_RU = Import-CSV $($dir + $Filename_RU) -Encoding UTF8;
$CSV_EN = Import-CSV $($dir + $Filename_EN) -Encoding UTF8;
$arrayofobj = @();
for ($i =0; $i -lt $CSV_RU.Count; $i++)
{
    $k =0; $found=$false;
    While (($k -lt $CSV_EN.Count) -and (!$found))
    {
      While (($CSV_EN[$k].Name -le $CSV_RU[$i].Name) -and ( $k -lt $CSV_EN.Count))
      {
         if ($CSV_RU[$i].Name -eq $CSV_EN[$k].Name)
         {
            $arrayofobj += [PSCustomObject]@{
                'Name' = $CSV_EN[$k].Name
                'RU' = $CSV_RU[$i].DisplayNameRU
                'EN' = $CSV_EN[$k].DisplayNameEN;
             } 
           $found=$true;
           break;
         }
         $k++
      }
      $k++
    }
    if (!$found)
    {
        $arrayofobj += [PSCustomObject]@{
        'Name' = $CSV_RU[$i].Name
        'RU' = $CSV_RU[$i].DisplayNameRU
        'EN' = ''
        }
    }
}
$arrayofobj| Export-Csv -Path $($dir+"Out.csv") -Encoding UTF8 -Delimiter ';' -NoTypeInformation