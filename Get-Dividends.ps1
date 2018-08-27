Add-Type -Path "C:\Users\User\Downloads\itextsharp-develop\src\core\bin\Debug_woDrawing\itextsharp.dll"
$targetObjs = New-Object System.Collections.ArrayList
Get-ChildItem C:\Users\User\Downloads\*Wertpapierertrag.pdf -File |  %{
$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $_.FullName
Write-Host "operating on " $_.Name

for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
 $lines = [char[]]$reader.GetPageContent($page) -join "" -split "`n"
 $targetLines = New-Object System.Collections.ArrayList
 for($lineIdx = 0; $lineIdx -le $lines.Count;$lineIdx++) {
 $line = $lines[$lineIdx]
  if ($line -match "^\[") {   
   $line = $line -replace "\\([\S])", $matches[1]
   $line = $line -replace "^\[\(|\)\] TJ$", "" -split "\)\-?\d+\.?\d*\(" -join ""
   $idx = $targetLines.Add($line)
  }
 }
  $obj = [PSCustomObject]@{
Datei = $_.FullName
Aktie = ''
ISIN = ''
Anteile = ''
DivPStueck = ''
DivPStueckWhrg = ''
Zeitraum = ''
Datum = ''
Start = ''
Ende = ''
Waehrg = ''
Dividende = ''
}
 for($lineIdx = 0; $lineIdx -le $targetLines.Count;$lineIdx++){
    if($targetLines[$lineIdx] -eq "Gattungsbezeichnung"){
    $targetLines[$lineIdx + 1] = $targetLines[$lineIdx + 1] -replace " Registered Shares DL -,06" -replace " Namens-Aktien o.N." -replace " Actions au Porteur o.N."
    $obj.Aktie = $targetLines[$lineIdx + 1]
    $obj.ISIN =  $targetLines[$lineIdx + 3]
    $obj.Anteile =  $targetLines[$lineIdx + 5] -replace "STK "
    $obj.DivPStueck =  $targetLines[$lineIdx + 11].ToString().Split(" ")[1]
    $obj.DivPStueckWhrg =  $targetLines[$lineIdx + 11].ToString().Split(" ")[0]
    if($targetLines[$lineIdx + 13].ToString().Length -lt 24) {
        $obj.Zeitraum =  $targetLines[$lineIdx + 13].ToString()
        $obj.Start =  $targetLines[$lineIdx + 13].ToString().Split(" - ")[0]
        $obj.Ende =  $targetLines[$lineIdx + 13].ToString().Split(" - ")[3]
    }
    $obj.Waehrg =  $targetLines[$lineIdx + 24]
    $obj.Dividende =  $targetLines[$lineIdx + 25]
   } elseif($targetLines[$lineIdx] -eq "Betrag zu Ihren Gunsten"){
    $obj.Datum =  $targetLines[$lineIdx - 3].ToString()
    if(-not $obj.Datum.Contains(".")){
        $obj.Datum =  $targetLines[$lineIdx - 5].ToString()}
    $obj.Waehrg =  $targetLines[$lineIdx + 1]
    $obj.Dividende =  $targetLines[$lineIdx + 2]
   }
   if($obj.Start -eq "") {$obj.Start = $obj.Datum}

 }

if($obj.Aktie -ne "" -and $obj.Aktie -ne "26koWorld-26koVision Classic Namens-Anteile C o.N."){
    $targetObjs.Add($obj)
}
}
}
$targetObjs | Export-Csv -Path "C:\Users\User\Downloads\dividenden.csv" -NoTypeInformation -Delimiter ";"
        $i = 0
$targetObjs | select Aktie -Unique | %{$val = $_ 
    $targetObjs | where Aktie -eq $val | %{
        $i = $i + $_.Dividende
        Write-Host $_.Aktie, $i
        }
        Write-Host $_.Aktie, $i
        $i = 0
       }
