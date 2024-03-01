#!ps
#timeout=600000

$companyFlag = 'MMS'
$url = 'https://raw.githubusercontent.com/digicentric/Thread-Resources/main/scripts/thread-deployment.ps1'
$outFile = 'C:\Windows\Temp\thread-deployment.ps1'

Invoke-WebRequest -Uri $url -OutFile $outFile
Start-Process -FilePath $outFile -ArgumentList "-$companyFlag", "-Verbose"