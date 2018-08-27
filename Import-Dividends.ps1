Add-Type -Path "C:\Users\User\Downloads\itextsharp-develop\src\core\bin\Debug_woDrawing\itextsharp.dll"
$minNoOfCsvColumns = 5
$Directory = "C:\Users\User\Downloads\"

$DividendenObj = New-Object System.Collections.ArrayList

#reads pdf content from $pdfPath, removes unwanted PDF-characters and returns an ArrayList of lines of pdf content
function Get-PDFTargetLines($pdfPath){
$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $pdfPath

    $targetLines = New-Object System.Collections.ArrayList
    for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
     $lines = [char[]]$reader.GetPageContent($page) -join "" -split "`n"
     for($lineIdx = 0; $lineIdx -le $lines.Count;$lineIdx++) {
     $line = $lines[$lineIdx]
      if ($line -match "^\[") {   
       $line = $line -replace "\\([\S])", $matches[1]
       $line = $line -replace "^\[\(|\)\] TJ$", "" -split "\)\-?\d+\.?\d*\(" -join ""
       $idx = $targetLines.Add($line)
      }
     }
    }
    return $targetLines
}

#removes header lines in csv from $csvPath file as long as the no. of clumns is less than $minNoOfCsvColumns 
function cleanCsvHeader($csvPath){
    $skip = 0
    (Get-Content $csvPath).Replace("Gesamtperformance Currency;%", "Gesamtperformance Currency;Gesamtperformance_%").Replace("Tagesperformance Currency;%","Tagesperformance Currency;Tagesperformance_%") | Set-Content $csvPath
    while((Get-Content $csvPath | Select-Object -Skip $skip)[0].Split(";").Count -lt 5){
        $skip++
    }
    if($skip -gt 0){
        (Get-Content $csvPath | Select-Object -Skip $skip) | Set-Content $csvPath
    }
}

Get-ChildItem ($Directory + "*Wertpapierertrag.pdf") -File |  %{

	$targetLines = Get-PDFTargetLines($_.FullName)

    $obj = [PSCustomObject]@{
    Datei = $_.FullName
    Aktie = ''
    ISIN = 0
    Anteile = 0
    DivPStueck = 0
    DivPStueckWhrg = ''
    Zeitraum = ''
    Datum = ''
    Start = ''
    Ende = ''
    Waehrg = ''
    Dividende = 0
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
        $DividendenObj.Add($obj)
    }
}

$DividendenObj | Export-Csv -Path ($Directory + "dividenden.csv") -NoTypeInformation -Delimiter ";"

$KaufObjs = New-Object System.Collections.ArrayList

Get-ChildItem ($Directory + "*Wertpapierabrechnung.pdf") -File |  %{
	$targetLines = Get-PDFTargetLines($_.FullName)
    $obj = [PSCustomObject]@{
    Datei = $_.FullName
    Aktie = ''
    ISIN = 0
    Anteile = 0
    Kurswert = 0
    Datum = ''
    Waehrg = ''
    }

    for($lineIdx = 0; $lineIdx -le $targetLines.Count;$lineIdx++){
        if($targetLines[$lineIdx] -eq "Gattungsbezeichnung"){
            $targetLines[$lineIdx + 1] = $targetLines[$lineIdx + 1] -replace " Registered Shares DL -,06" -replace " Namens-Aktien o.N." -replace " Actions au Porteur o.N."
            $obj.Aktie = $targetLines[$lineIdx + 1]
            $obj.ISIN =  $targetLines[$lineIdx + 3]
            $obj.Anteile =  $targetLines[$lineIdx + 5] -replace "STK "
            $obj.Kurswert =  $targetLines[$lineIdx + 13]
            $obj.Waehrg =  $targetLines[$lineIdx + 12]
        } elseif ($targetLines[$lineIdx] -eq "Handelstag"){
            $obj.Datum =  $targetLines[$lineIdx + 1]
            if($obj.ISIN -eq "US7565771026"){$obj.Datum = "01.10.1995"}
            if($obj.Datum -ne ""){
                $KaufObjs.Add($obj)
                }
            $obj = [PSCustomObject]@{
    Datei = $_.FullName
    Aktie = ''
    ISIN = 0
    Anteile = 0
    Kurswert = 0
    Datum = ''
    Waehrg = ''
    }
        }

    }

    $obj = [PSCustomObject]@{
    Datei = "code-generated"
    Aktie = 'Red Hat'
    ISIN = "US7565771026"
    Anteile = 60
    Kurswert = 25
    Datum = '01.10.1995'
    Waehrg = 'EUR'
    }
    $KaufObjs.Add($obj)
    $obj = [PSCustomObject]@{
    Datei = "code-generated"
    Aktie = 'SAG Solrastrom'
    ISIN = "702100"
    Anteile = 1
    Kurswert = 25
    Datum = '01.10.2013'
    Waehrg = 'EUR'
    }
    $KaufObjs.Add($obj)


}

$KaufObjs | Export-Csv -Path ($Directory + "kaufdaten.csv") -NoTypeInformation -Delimiter ";"

$depotPath = (Get-ChildItem ($Directory + "Depotübersicht_Wertpapiere*.csv") -File | sort LastWriteTime | select -Last 1)[0].FullName
cleanCsvHeader($depotPath)

$depot = Import-Csv -Path $depotPath -Delimiter ";" 

#adds Kaufdatum & Dividende as new columns of $depotPath-csv.Based on Depot-, Dividende- & Kaufdatum-ISIN, corresponding columns values are inserted.
#in case of Dividende, the sum of all Dividende-values with the corresponding ISIN is calculated with Measure-Object after each values is converted to double 
$depot = Import-Csv -Path $depotPath -Delimiter ";" | Select-Object *,@{Name='Kaufdatum';Expression={$isin = $_.ISIN
($KaufObjs | where ISIN -EQ $isin | Select-Object Datum)[0].Datum}},@{Name='Dividende';Expression={$isin = $_.ISIN
($DividendenObj | where ISIN -EQ $isin | Select-Object  @{Name="Surname";Expression={[convert]::ToDouble($_.Dividende)}} | Measure-Object -Property "surname" -Sum).Sum}} 

$depot | Export-Csv -Path ($Directory + "Depotübersicht_Dividenden.csv") -NoTypeInformation -Delimiter ";"