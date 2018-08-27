$Directory = "C:\Users\User\Downloads\"

function cleanCsvHeader($csvPath){
    $skip = 0
    while((Get-Content $csvPath | Select-Object -Skip $skip)[0].Contains(";;") -or -not (Get-Content $csvPath | Select-Object -Skip $skip)[0].Contains(";")){
        $skip++
    }
    if($skip -gt 0){
        (Get-Content $csvPath | Select-Object -Skip $skip) | Set-Content $csvPath
    }
}


$depotPath = (Get-ChildItem ($Directory + "Depotbestand*.csv") -File | sort LastWriteTime | select -Last 1)[0].FullName
cleanCsvHeader($depotPath)

Get-Content $depotPath | Set-Content ($Directory + "fonds-daten.csv")
