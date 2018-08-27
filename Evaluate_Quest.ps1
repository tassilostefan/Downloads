$Directory = "C:\Users\User\Downloads\Blasendiagramme\"
$wsList = New-Object System.Collections.ArrayList
$resultCsv = New-Object System.Collections.ArrayList

#reads pdf content from $pdfPath, removes unwanted PDF-characters and returns an ArrayList of lines of pdf content
function Export-CsvData($xlsxPath){

    $xlCSV = 6
    $Excel = New-Object -Com Excel.Application 
    $Excel.visible = $False 
    $Excel.displayalerts=$False 
    $Wrkbk = $Excel.Workbooks.Open($xlsxPath)
    foreach ($Wrkst in $Wrkbk.Worksheets){
        $WsName = $xlsxPath + "_" + $Wrkst.Name + ".csv"
        $Wrkst.SaveAs($WsName, 6)
#        $Wrkst.SaveAs($WsName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText)
        $k = $wsList.Add($WsName)
    }
    $Excel.Quit()
    stop-process -processname EXCEL   

}

Get-ChildItem -Path $Directory -Filter "*.xlsx" | %{
    if($_.Name -ne "data.xlsx"){
#    if($_.Name -eq "FBHZ_Bla$sendiagramme.xlsx"){
        $wsList = New-Object System.Collections.ArrayList
        $xlsxFile = $Directory + $_
        Write-Host "Getting ", $_
        Export-CsvData($xlsxFile)
    
        foreach ($ws in $wsList){
            (Get-Content $ws).Replace(";Bereich","A;Bereich").Replace("N;;;","N1;;;").Replace(";N ;",";N2;") | Set-Content $ws
            $institution = $ws.Split("\")[5].Split("_")[0]
            $csv = Import-Csv $ws -Delimiter ";"
            $csv | Where Bereich -NE "" | Where H1 -NE "" | Select-Object Bereich, H1, N, Minimum, Maximum, Mittelwert, Standardabweichung, @{Name='Institution';Expression={$institution}} | Export-Csv $ws -Delimiter ";" -NoTypeInformation
        }
        foreach ($ws in $wsList){
            $isPersonal = $ws.IndexOf("nlich") -gt 0
            if($isPersonal){
                (Get-Content $ws).Replace("H1", "P_H1").Replace("Minimum","P_Minimum").Replace("Maximum","P_Maximum").Replace("Mittelwert","P_Mittelwert").Replace("Standardabweichung","P_Standardabweichung") | Set-Content $ws
                $csv = Import-Csv $ws -Delimiter ";"
                $csv | Select-Object @{Name='P_H1';Expression={$_.P_H1.Remove($_.P_H1.IndexOf("F?r mich pers?nlich")-1)}}, P_Minimum, P_Maximum, P_Mittelwert, P_Standardabweichung | Export-Csv $ws -Delimiter ";" -NoTypeInformation
                $csv | Export-Csv $ws.Replace(".csv", "t.csv") -Delimiter ";" -NoTypeInformation
            } else {
                $csv = Import-Csv $ws -Delimiter ";"
                $csv | Select-Object Institution, Bereich, @{Name='H1';Expression={$_.H1.Remove($_.H1.IndexOf("F?r die Einrichtung")-1)}}, Minimum, Maximum, Mittelwert, Standardabweichung | Export-Csv $ws -Delimiter ";" -NoTypeInformation
            }
        }
        $cont1 = Get-Content $wsList[0]
        $csv1 = Import-Csv $wsList[0] -Delimiter ";"
        $sum = $cont1
        $cont2 = Get-Content $wsList[1]
        $csv2 = Import-Csv $wsList[1] -Delimiter ";"

        $limit = $cont2.Count-1
        $smallerCont = $cont2
        $largerCont = $cont1
        if($limit -gt $cont1.Count-1){
            $limit = $cont1.Count-1
            $smallerCont = $cont1
            $largerCont = $cont2
        }

        $result = New-Object System.Collections.ArrayList
        for ($page = 0; $page -le $limit; $page++) {
            if($page -ne 0){
                $ele = $csv1 | where P_H1 -eq $csv2[$page].H1
                if($ele -eq $null -and $csv2[$page].H1 -eq ""){
                    Write-Host "Header empty in line: " $line " Bereich: " $csv2[$page].Bereich} 
                elseif($ele -eq $null){
                    Write-Host "Header not found: |" $csv2[$page].H1 "| Bereich: " $csv2[$page].Bereich} 
                else {
                    $resultObj = [PSCustomObject]@{
                    Institution = $csv2[$page].Institution
                    Bereich = $csv2[$page].Bereich
                    H1 = $csv2[$page].H1
                    N = $csv2[$page].N
                    Minimum = $csv2[$page].Minimum
                    Maximum = $csv2[$page].Maximum
                    Mittelwert = $csv2[$page].Mittelwert
                    Standardabweichung = $csv2[$page].Standardabweichung

                    P_Minimum = $ele.P_Minimum
                    P_Maximum = $ele.P_Maximum
                    P_Mittelwert = $ele.P_Mittelwert
                    P_Standardabweichung = $ele.P_Standardabweichung
                    }
                    $k = $result.Add($resultObj)
                }
            }
#            $sum[$page] = $cont2[$page] + ";" + $cont1[$page] 
        }
        $csvFile = $xlsxFile.Replace("xlsx", "csv")
#        $sum | Set-Content $csvFile
        $result | Export-Csv $csvFile -Delimiter ";" -NoTypeInformation
        $k = $resultCsv.Add($csvFile)
    }
    } 

$result = New-Object System.Collections.ArrayList

$resultCsv | %{
    $csv = Import-Csv $_ -Delimiter ";"
    $bereiche = ($csv | Select Bereich -Unique )


    $bereiche | %{
#        Write-Host "Bereich: " $_.Bereich
        $curBereich = $_.Bereich
        $obj = [PSCustomObject]@{
        Inst = ($csv | Select-Object Institution)[0].Institution
        Bereich = $curBereich
        N = 0
        Mittelwert = 0
        Standardabweichung = 0
        P_Mittelwert = 0
        P_Standardabweichung = 0
        }
        $obj.Mittelwert = ($csv | Where-Object Bereich -EQ $curBereich | Select-Object @{Name="Mittelwert";Expression={[double]::Parse($_.Mittelwert)}} | Measure-Object -Property Mittelwert -Average).Average
        $obj.Standardabweichung = ($csv | Where-Object Bereich -EQ $curBereich | Select-Object @{Name="Standardabweichung";Expression={[double]::Parse($_.Standardabweichung)}} | Measure-Object -Property Standardabweichung -Average).Average
        $obj.P_Mittelwert = ($csv | Where-Object Bereich -EQ $curBereich | Select-Object @{Name="P_Mittelwert";Expression={[double]::Parse($_.P_Mittelwert)}} | Measure-Object -Property P_Mittelwert -Average).Average
        $obj.P_Standardabweichung = ($csv | Where-Object Bereich -EQ $curBereich | Select-Object @{Name="P_Standardabweichung";Expression={[double]::Parse($_.P_Standardabweichung)}} | Measure-Object -Property P_Standardabweichung -Average).Average
        $obj.N = ($csv | Where-Object Bereich -EQ $curBereich).Count
    
        $k = $result.Add($obj)
    }
}
$result | Export-Csv ($Directory + "data.csv") -NoTypeInformation -Delimiter ";"

