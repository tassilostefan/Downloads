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
#removes double currency entries ( columns 
function cleanCsvHeader($csvPath){
    $skip = 0
    (Get-Content $csvPath).Replace("Gesamtperformance Currency;%", "Gesamtperformance Currency;Gesamtperformance_%").Replace("Tagesperformance Currency;%","Tagesperformance Currency;Tagesperformance_%") | Set-Content $csvPath
    while((Get-Content $csvPath | Select-Object -Skip $skip)[0].Split(";").Count -lt 5){
        $skip++
    }
    if($skip -gt 0){
        (Get-Content $csvPath | Select-Object -Skip $skip) | Set-Content $csvPath
    }
    $filteredCsv = Import-Csv $csvPath -Delimiter ";" | Select-Object Bestand, Name, ISIN, WKN, Typ, Datum, Zeit, 'Akt. Geldkurs', 'Akt. Geldkurs Currency', Boerse, Kaufkurs, Kaufwert, Gesamtperformance, Gesamtperformance_%, Tagesperformance, Tagesperformance_%, 'Aktueller Wert', Wert 
    $filteredCsv | Export-Csv $csvPath -Delimiter ";" -NoTypeInformation
}

function import2Xls($csvFile){
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    $Excel = New-Object -ComObject Excel.Application
    $xlLastCell = [Microsoft.Office.Interop.Excel.Constants]::xlLastCell
    $xlFile = $csvFile.Replace(".csv", ".xlsx")

    $Excel.visible = $false 
    $Excel.displayalerts=$False 

    $ExcelWordBook = $Excel.Workbooks.Open($xlFile)
    $ExcelWorkSheet = $Excel.WorkSheets.item('Depotübersicht_Dividenden')
    $ExcelWorkSheet.activate()

    $objRange = $ExcelWorkSheet.UsedRange
    $lastRow = 0
#    $lastRow = $objRange.SpecialCells($xlLastCell).Row
#    $ExcelWorkSheet.cells($lastRow,1).Select()

    $singers = Import-Csv -Path $csvFile
    $singers.foreach{
      $lastRow += 1
      $ExcelWorkSheet.cells($lastRow,1).value = $_.Name
      $ExcelWorkSheet.cells($lastRow,2).value = $_.ISIN
      $ExcelWorkSheet.cells($lastRow,3).value = $_.WKN
      $ExcelWorkSheet.cells($lastRow,4).value = $_.Typ
      $ExcelWorkSheet.cells($lastRow,5).value = $_.Datum
      $ExcelWorkSheet.cells($lastRow,6).value = $_.Zeit
      $ExcelWorkSheet.cells($lastRow,7).value = $_.Bestand
      $ExcelWorkSheet.cells($lastRow,8).value = $_.'Akt. Geldkurs'
      $ExcelWorkSheet.cells($lastRow,9).value = $_.'Akt. Geldkurs Currency'
      $ExcelWorkSheet.cells($lastRow,10).value = $_.Tagesperformance
      $ExcelWorkSheet.cells($lastRow,11).value = $_.'Tagesperformance_%'
      $ExcelWorkSheet.cells($lastRow,12).value = $_.'Aktueller Wert'
      $ExcelWorkSheet.cells($lastRow,13).value = $_.Wert
      $ExcelWorkSheet.cells($lastRow,14).value = $_.Kaufdatum
      $ExcelWorkSheet.cells($lastRow,15).value = $_.Kaufkurs
      $ExcelWorkSheet.cells($lastRow,16).value = $_.Kaufwert
      $ExcelWorkSheet.cells($lastRow,17).value = $_.Dividende
    }

    $ExcelWordBook.Save()
    $ExcelWordBook.Close()

    $Excel.Quit()
    stop-process -processname EXCEL   
}

Get-ChildItem ($Directory + "*Wertpapierertrag.pdf") -File |  %{

	$targetLines = Get-PDFTargetLines($_.FullName)

    $obj = [PSCustomObject]@{
    Datei = $_.FullName
    Aktie = ''
    ISIN = 0
    Anteile = 0
    DivPStueck = 0
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
    }

    for($lineIdx = 0; $lineIdx -le $targetLines.Count;$lineIdx++){
        if($targetLines[$lineIdx] -eq "Gattungsbezeichnung"){
            $targetLines[$lineIdx + 1] = $targetLines[$lineIdx + 1] -replace " Registered Shares DL -,06" -replace " Namens-Aktien o.N." -replace " Actions au Porteur o.N."
            $obj.Aktie = $targetLines[$lineIdx + 1]
            $obj.ISIN =  $targetLines[$lineIdx + 3]
            $obj.Anteile =  $targetLines[$lineIdx + 5] -replace "STK "
            $obj.Kurswert =  $targetLines[$lineIdx + 13]
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
    }
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
    }
$KaufObjs.Add($obj)
$obj = [PSCustomObject]@{
    Datei = "code-generated"
    Aktie = 'SAG Solrastrom'
    ISIN = "702100"
    Anteile = 1
    Kurswert = 25
    Datum = '01.10.2013'
    }
$KaufObjs.Add($obj)

$KaufObjs | Export-Csv -Path ($Directory + "kaufdaten.csv") -NoTypeInformation -Delimiter ";"

$depotPath = (Get-ChildItem ($Directory + "Depotübersicht_Wertpapiere*.csv") -File | sort LastWriteTime | select -Last 1)[0].FullName
cleanCsvHeader($depotPath)

$depot = Import-Csv -Path $depotPath -Delimiter ";" 

#adds Kaufdatum & Dividende as new columns of $depotPath-csv.Based on Depot-, Dividende- & Kaufdatum-ISIN, corresponding columns values are inserted.
#in case of Dividende, the sum of all Dividende-values with the corresponding ISIN is calculated with Measure-Object after each values is converted to double 
$depot = Import-Csv -Path $depotPath -Delimiter ";" | Select-Object *,@{Name='Kaufdatum';Expression={$isin = $_.ISIN
($KaufObjs | where ISIN -EQ $isin | Select-Object Datum)[0].Datum}},@{Name='Dividende';Expression={$isin = $_.ISIN
($DividendenObj | where ISIN -EQ $isin | Select-Object  @{Name="Surname";Expression={[convert]::ToDouble($_.Dividende)}} | Measure-Object -Property "surname" -Sum).Sum}} 
$csvPath = ($Directory + "Depotübersicht_Dividenden.csv")
$depot | Select-Object -First ($depot.Length-2) | Select-Object Name, ISIN, WKN, Typ, Datum, Zeit, Bestand, 'Akt. Geldkurs', 'Akt. Geldkurs Currency', Boerse, Tagesperformance, Tagesperformance_%, 'Aktueller Wert', Wert, Kaufdatum, Kaufkurs, Kaufwert, Dividende | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ";"
#import2Xls($csvPath)